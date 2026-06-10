"""Flask web app for RVtools cluster filtering and export."""

from __future__ import annotations

import json
import secrets
import shutil
from datetime import datetime, timezone
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from customer_report import extract_customer_summary, generate_customer_doc
from vm_analytics import extract_vm_analytics
from rvtools_processor import (
    analyze_workbook,
    apply_edits,
    export_workbook,
    filter_by_clusters,
    load_workbook,
    sheet_to_records,
)

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024
app.config["SECRET_KEY"] = secrets.token_hex(16)

SESSIONS: dict[str, dict] = {}


def _session_path(session_id: str) -> Path:
    return UPLOAD_DIR / session_id


def _meta_path(session_id: str) -> Path:
    return _session_path(session_id) / "meta.json"


def _workbook_path(session_id: str) -> Path:
    return _session_path(session_id) / "workbook.xlsx"


def _filtered_path(session_id: str) -> Path:
    return _session_path(session_id) / "filtered.xlsx"


def _load_meta(session_id: str) -> dict | None:
    path = _meta_path(session_id)
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def _save_meta(session_id: str, meta: dict) -> None:
    _meta_path(session_id).write_text(json.dumps(meta, indent=2), encoding="utf-8")


def _get_session(session_id: str) -> dict | None:
    if session_id not in SESSIONS:
        meta = _load_meta(session_id)
        if meta:
            SESSIONS[session_id] = meta
    return SESSIONS.get(session_id)


def _active_workbook_path(session: dict, session_id: str) -> Path:
    if session.get("filtered"):
        return _filtered_path(session_id)
    return _workbook_path(session_id)


def _load_active_workbook(session: dict, session_id: str):
    path = _active_workbook_path(session, session_id)
    workbook = load_workbook(path)
    if session.get("edits"):
        workbook = apply_edits(workbook, session["edits"])
    return workbook


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "No filename provided."}), 400

    if not file.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "Only .xlsx files are supported."}), 400

    session_id = secrets.token_urlsafe(12)
    session_dir = _session_path(session_id)
    session_dir.mkdir(parents=True)

    filename = secure_filename(file.filename) or "rvtools.xlsx"
    dest = session_dir / "original.xlsx"
    file.save(dest)

    try:
        workbook = load_workbook(dest)
        analysis = analyze_workbook(workbook)
        export_workbook(workbook, _workbook_path(session_id))
    except Exception as exc:
        shutil.rmtree(session_dir, ignore_errors=True)
        return jsonify({"error": f"Could not read file: {exc}"}), 400

    meta = {
        "session_id": session_id,
        "filename": filename,
        "uploaded_at": datetime.now(timezone.utc).isoformat(),
        "selected_clusters": [],
        "edits": {},
        "filtered": False,
        "analysis": analysis,
    }
    _save_meta(session_id, meta)
    SESSIONS[session_id] = meta

    return jsonify({"session_id": session_id, "filename": filename, **analysis})


@app.route("/api/session/<session_id>/clusters", methods=["POST"])
def select_clusters(session_id: str):
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found."}), 404

    data = request.get_json(silent=True) or {}
    clusters = data.get("clusters", [])
    if not isinstance(clusters, list):
        return jsonify({"error": "Invalid cluster list."}), 400

    workbook = load_workbook(_workbook_path(session_id))
    filtered = filter_by_clusters(workbook, clusters)
    export_workbook(filtered, _filtered_path(session_id))

    session["selected_clusters"] = clusters
    session["filtered"] = True
    session["edits"] = {}
    _save_meta(session_id, session)

    filtered_stats = {
        sheet: len(df) for sheet, df in filtered.items()
    }
    return jsonify({
        "selected_clusters": clusters,
        "filtered_stats": filtered_stats,
        "total_rows": sum(filtered_stats.values()),
    })


@app.route("/api/session/<session_id>/sheets")
def list_sheets(session_id: str):
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found."}), 404

    path = _filtered_path(session_id) if session.get("filtered") else _workbook_path(session_id)
    if not path.exists():
        return jsonify({"error": "No workbook available."}), 404

    workbook = load_workbook(path)
    sheets = [
        {"name": name, "rows": len(df), "columns": len(df.columns)}
        for name, df in workbook.items()
    ]
    return jsonify({"sheets": sheets, "filtered": session.get("filtered", False)})


@app.route("/api/session/<session_id>/sheet/<sheet_name>")
def get_sheet(session_id: str, sheet_name: str):
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found."}), 404

    path = _filtered_path(session_id) if session.get("filtered") else _workbook_path(session_id)
    workbook = load_workbook(path)
    if sheet_name not in workbook:
        return jsonify({"error": f"Sheet '{sheet_name}' not found."}), 404

    offset = request.args.get("offset", 0, type=int)
    limit = request.args.get("limit", 100, type=int)
    limit = min(max(limit, 1), 500)

    df = workbook[sheet_name]
    if session.get("edits", {}).get(sheet_name):
        workbook = apply_edits(workbook, session["edits"])
        df = workbook[sheet_name]

    return jsonify(sheet_to_records(df, offset=offset, limit=limit))


@app.route("/api/session/<session_id>/edit", methods=["POST"])
def edit_cell(session_id: str):
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found."}), 404

    data = request.get_json(silent=True) or {}
    sheet = data.get("sheet")
    row = data.get("row")
    column = data.get("column")
    value = data.get("value")

    if not sheet or row is None or not column:
        return jsonify({"error": "sheet, row and column are required."}), 400

    edits = session.setdefault("edits", {})
    sheet_edits = edits.setdefault(sheet, [])

    for edit in sheet_edits:
        if edit.get("row") == row and edit.get("column") == column:
            edit["value"] = value
            break
    else:
        sheet_edits.append({"row": row, "column": column, "value": value})

    _save_meta(session_id, session)
    return jsonify({"ok": True})


@app.route("/api/session/<session_id>/export")
def export(session_id: str):
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found."}), 404

    if not session.get("filtered"):
        return jsonify({"error": "Please select and apply clusters first."}), 400

    path = _filtered_path(session_id)
    workbook = load_workbook(path)
    if session.get("edits"):
        workbook = apply_edits(workbook, session["edits"])

    export_path = _session_path(session_id) / "export.xlsx"
    export_workbook(workbook, export_path)

    clusters = session.get("selected_clusters", [])
    suffix = f"_{len(clusters)}clusters" if clusters else "_filtered"
    download_name = f"RVtools{suffix}.xlsx"

    return send_file(
        export_path,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/session/<session_id>/vm-analytics")
def vm_analytics(session_id: str):
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found."}), 404

    path = _active_workbook_path(session, session_id)
    if not path.exists():
        return jsonify({"error": "No workbook available."}), 404

    workbook = _load_active_workbook(session, session_id)
    analytics = extract_vm_analytics(workbook)
    return jsonify({
        **analytics,
        "filtered": session.get("filtered", False),
        "selected_clusters": session.get("selected_clusters", []),
    })


@app.route("/api/session/<session_id>/customer-summary")
def customer_summary(session_id: str):
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found."}), 404

    path = _active_workbook_path(session, session_id)
    if not path.exists():
        return jsonify({"error": "No workbook available."}), 404

    workbook = _load_active_workbook(session, session_id)
    summary = extract_customer_summary(workbook)
    return jsonify({
        **summary,
        "filtered": session.get("filtered", False),
        "selected_clusters": session.get("selected_clusters", []),
    })


@app.route("/api/session/<session_id>/customer-doc", methods=["POST"])
def customer_doc(session_id: str):
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found."}), 404

    data = request.get_json(silent=True) or {}
    customer_name = (data.get("customer_name") or "").strip()
    address = (data.get("address") or "").strip()
    contact_person = (data.get("contact_person") or "").strip()
    additional_info = (data.get("additional_info") or "").strip()
    software_info = (data.get("software_info") or "").strip()

    if not customer_name:
        return jsonify({"error": "Please provide a customer name."}), 400
    if not address:
        return jsonify({"error": "Please provide an address."}), 400

    path = _active_workbook_path(session, session_id)
    if not path.exists():
        return jsonify({"error": "No workbook available."}), 404

    workbook = _load_active_workbook(session, session_id)
    summary = extract_customer_summary(workbook)

    safe_name = secure_filename(customer_name) or "customer"
    doc_path = _session_path(session_id) / "customer_description.docx"
    generate_customer_doc(
        customer_name,
        address,
        summary,
        doc_path,
        contact_person=contact_person,
        additional_info=additional_info,
        software_info=software_info,
    )

    return send_file(
        doc_path,
        as_attachment=True,
        download_name=f"Customer_Description_{safe_name}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=8080)
