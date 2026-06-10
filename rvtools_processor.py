"""RVtools Excel analysis, cluster filtering, and export."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd

VM_SHEETS = {
    "vInfo",
    "vCPU",
    "vMemory",
    "vDisk",
    "vPartition",
    "vNetwork",
    "vCD",
    "vUSB",
    "vSnapshot",
    "vTools",
}

HOST_SHEETS = {
    "vHost",
    "vHBA",
    "vNIC",
    "vSwitch",
    "vPort",
    "vSC_VMK",
    "vMultiPath",
}

Workbook = dict[str, pd.DataFrame]


def load_workbook(path: str | Path) -> Workbook:
    xl = pd.ExcelFile(path)
    return {sheet: pd.read_excel(xl, sheet_name=sheet) for sheet in xl.sheet_names}


def _host_set(workbook: Workbook, clusters: set[str]) -> set[str]:
    if "vHost" not in workbook:
        return set()
    df = workbook["vHost"]
    if "Cluster" not in df.columns or "Host" not in df.columns:
        return set()
    mask = df["Cluster"].astype(str).isin(clusters)
    return set(df.loc[mask, "Host"].dropna().astype(str))


def _vm_set(workbook: Workbook, clusters: set[str]) -> set[str]:
    for sheet in ("vInfo", "vCPU"):
        if sheet not in workbook:
            continue
        df = workbook[sheet]
        if "Cluster" not in df.columns or "VM" not in df.columns:
            continue
        mask = df["Cluster"].astype(str).isin(clusters)
        return set(df.loc[mask, "VM"].dropna().astype(str))
    return set()


def _split_hosts(value: Any) -> set[str]:
    if pd.isna(value) or value is None:
        return set()
    return {part.strip() for part in str(value).split(",") if part.strip()}


def discover_clusters(workbook: Workbook) -> list[dict[str, Any]]:
    clusters: dict[str, dict[str, Any]] = {}

    if "vCluster" in workbook and "Name" in workbook["vCluster"].columns:
        for name in workbook["vCluster"]["Name"].dropna().astype(str):
            clusters[name] = {
                "name": name,
                "vm_count": 0,
                "host_count": 0,
                "datacenter": "",
            }

    if "vInfo" in workbook:
        df = workbook["vInfo"]
        if "Cluster" in df.columns:
            for cluster, group in df.groupby(df["Cluster"].astype(str)):
                if cluster in ("nan", ""):
                    continue
                entry = clusters.setdefault(
                    cluster,
                    {"name": cluster, "vm_count": 0, "host_count": 0, "datacenter": ""},
                )
                entry["vm_count"] = int(len(group))
                if "Datacenter" in group.columns:
                    dcs = group["Datacenter"].dropna().astype(str).unique()
                    if len(dcs):
                        entry["datacenter"] = str(dcs[0])

    if "vHost" in workbook:
        df = workbook["vHost"]
        if "Cluster" in df.columns:
            for cluster, group in df.groupby(df["Cluster"].astype(str)):
                if cluster in ("nan", ""):
                    continue
                entry = clusters.setdefault(
                    cluster,
                    {"name": cluster, "vm_count": 0, "host_count": 0, "datacenter": ""},
                )
                entry["host_count"] = int(len(group))
                if not entry["datacenter"] and "Datacenter" in group.columns:
                    dcs = group["Datacenter"].dropna().astype(str).unique()
                    if len(dcs):
                        entry["datacenter"] = str(dcs[0])

    return sorted(clusters.values(), key=lambda item: item["name"].lower())


def analyze_workbook(workbook: Workbook) -> dict[str, Any]:
    clusters = discover_clusters(workbook)
    sheet_stats = {
        sheet: {"rows": int(len(df)), "columns": int(len(df.columns))}
        for sheet, df in workbook.items()
    }
    return {
        "clusters": clusters,
        "cluster_count": len(clusters),
        "sheets": list(workbook.keys()),
        "sheet_stats": sheet_stats,
        "vm_count": sheet_stats.get("vInfo", {}).get("rows", 0),
        "host_count": sheet_stats.get("vHost", {}).get("rows", 0),
    }


def filter_by_clusters(workbook: Workbook, selected: list[str]) -> Workbook:
    cluster_set = set(selected)
    if not cluster_set:
        return {sheet: df.iloc[0:0].copy() for sheet, df in workbook.items()}

    hosts = _host_set(workbook, cluster_set)
    vms = _vm_set(workbook, cluster_set)
    filtered: Workbook = {}

    for sheet, df in workbook.items():
        if sheet == "vSource":
            filtered[sheet] = df.copy()
            continue

        if sheet == "vCluster" and "Name" in df.columns:
            filtered[sheet] = df[df["Name"].astype(str).isin(cluster_set)].copy()
            continue

        if sheet in VM_SHEETS and "Cluster" in df.columns:
            filtered[sheet] = df[df["Cluster"].astype(str).isin(cluster_set)].copy()
            continue

        if sheet in HOST_SHEETS and "Cluster" in df.columns:
            filtered[sheet] = df[df["Cluster"].astype(str).isin(cluster_set)].copy()
            continue

        if sheet == "vRP" and "Resource Pool path" in df.columns:
            pattern = "|".join(re.escape(name) for name in cluster_set)
            mask = df["Resource Pool path"].astype(str).str.contains(pattern, regex=True, na=False)
            filtered[sheet] = df[mask].copy()
            continue

        if sheet == "vDatastore" and "Hosts" in df.columns and hosts:
            mask = df["Hosts"].apply(lambda value: bool(_split_hosts(value) & hosts))
            filtered[sheet] = df[mask].copy()
            continue

        if sheet == "dvSwitch" and "Host members" in df.columns and hosts:
            mask = df["Host members"].apply(lambda value: bool(_split_hosts(value) & hosts))
            filtered[sheet] = df[mask].copy()
            continue

        if sheet == "dvPort" and "Switch" in df.columns:
            if "dvSwitch" in filtered and "Switch" in filtered["dvSwitch"].columns:
                switches = set(filtered["dvSwitch"]["Switch"].dropna().astype(str))
                filtered[sheet] = df[df["Switch"].astype(str).isin(switches)].copy()
            else:
                filtered[sheet] = df.iloc[0:0].copy()
            continue

        filtered[sheet] = df.iloc[0:0].copy()

    return filtered


def apply_edits(workbook: Workbook, edits: dict[str, list[dict[str, Any]]]) -> Workbook:
    updated = {sheet: df.copy() for sheet, df in workbook.items()}
    for sheet, sheet_edits in edits.items():
        if sheet not in updated:
            continue
        df = updated[sheet]
        for edit in sheet_edits:
            row_idx = edit.get("row")
            col = edit.get("column")
            value = edit.get("value")
            if row_idx is None or col is None:
                continue
            if col not in df.columns:
                continue
            if 0 <= row_idx < len(df):
                df.iat[row_idx, df.columns.get_loc(col)] = value
    return updated


def sheet_to_records(df: pd.DataFrame, offset: int = 0, limit: int = 100) -> dict[str, Any]:
    total = len(df)
    end = min(offset + limit, total)
    slice_df = df.iloc[offset:end]
    records = []
    for row_idx, row in slice_df.iterrows():
        record = {"_row": int(row_idx)}
        for col in df.columns:
            val = row[col]
            if pd.isna(val):
                record[str(col)] = ""
            else:
                record[str(col)] = val
        records.append(record)
    return {
        "columns": [str(c) for c in df.columns],
        "rows": records,
        "total": total,
        "offset": offset,
        "limit": limit,
    }


def export_workbook(workbook: Workbook, path: str | Path) -> None:
    path = Path(path)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet, df in workbook.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)
