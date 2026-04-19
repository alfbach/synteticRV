from __future__ import annotations

import io
import re
import tempfile
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

from generator import generate_rvtools_xlsx

APP_DIR = Path(__file__).resolve().parent
TEMPLATE = APP_DIR / "customer.xlsx"

app = Flask(__name__)

_SAFE_NAME = re.compile(r"[^A-Za-z0-9._\-]+")


def _sanitize_filename(name: str) -> str:
    base = (name or "rvtools_synthetic").strip()
    base = _SAFE_NAME.sub("_", base).strip("._") or "rvtools_synthetic"
    if not base.lower().endswith(".xlsx"):
        base += ".xlsx"
    return base


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/generate")
def api_generate():
    if not TEMPLATE.exists():
        return jsonify({"error": "customer.xlsx is missing from the application directory."}), 500

    payload = request.get_json(silent=True) or {}
    size = str(payload.get("size", "m"))
    filename = _sanitize_filename(str(payload.get("filename", "rvtools_synthetic.xlsx")))
    export_dir = str(payload.get("export_dir") or "").strip()

    try:
        if export_dir:
            out_dir = Path(export_dir).expanduser()
            out_dir.mkdir(parents=True, exist_ok=True)
            out_path = out_dir / filename
            meta = generate_rvtools_xlsx(TEMPLATE, out_path, size)
            return jsonify({"ok": True, "written_path": str(out_path), **meta})
        buf = io.BytesIO()
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
            tmp_path = Path(tf.name)
        try:
            generate_rvtools_xlsx(TEMPLATE, tmp_path, size)
            buf.write(tmp_path.read_bytes())
        finally:
            tmp_path.unlink(missing_ok=True)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except OSError as e:
        return jsonify({"error": f"Filesystem error: {e}"}), 500


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8765, debug=True)
