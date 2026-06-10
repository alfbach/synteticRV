"""VM OS and application analytics from RVtools vInfo data."""

from __future__ import annotations

import re
from typing import Any

import pandas as pd

from rvtools_processor import Workbook

MAX_LIST_ITEMS = 100


def _safe_str(value: Any) -> str:
    if pd.isna(value) or value is None:
        return ""
    return str(value).strip()


def _is_template(value: Any) -> bool:
    return _safe_str(value).lower() in ("true", "1", "yes")


def _is_powered_on(value: Any) -> bool:
    return _safe_str(value).lower() == "poweredon"


def _resolve_os(row: pd.Series) -> str:
    tools_os = _safe_str(row.get("OS according to the VMware Tools"))
    config_os = _safe_str(row.get("OS according to the configuration file"))
    if tools_os:
        return tools_os
    if config_os:
        return config_os
    return "Unknown"


def _os_family(os_name: str) -> str:
    lower = os_name.lower()
    if "windows" in lower:
        return "Windows"
    if any(token in lower for token in ("linux", "centos", "red hat", "ubuntu", "debian", "suse", "rhel")):
        return "Linux"
    if os_name == "Unknown":
        return "Unknown"
    return "Other"


def _simplify_os_label(os_name: str) -> str:
    if os_name == "Unknown":
        return os_name

    lower = os_name.lower()
    if "windows server 2022" in lower:
        return "Windows Server 2022"
    if "windows server 2019" in lower or "2016 or later" in lower:
        return "Windows Server 2016/2019"
    if "windows server 2012" in lower:
        return "Windows Server 2012"
    if "windows" in lower:
        return "Windows (other)"

    rhel = re.search(r"red hat enterprise linux (\d+)", lower)
    if rhel:
        return f"RHEL {rhel.group(1)}"
    centos = re.search(r"centos (\d+)", lower)
    if centos:
        return f"CentOS {centos.group(1)}"
    if "ubuntu" in lower:
        return "Ubuntu"
    if "debian" in lower:
        return "Debian"
    if "suse" in lower:
        return "SUSE Linux"
    if "linux" in lower:
        return "Linux (other)"
    return os_name


def _active_vms(df: pd.DataFrame) -> pd.DataFrame:
    if "Powerstate" not in df.columns:
        return df.iloc[0:0].copy()

    mask = df["Powerstate"].apply(_is_powered_on)
    if "Template" in df.columns:
        mask &= ~df["Template"].apply(_is_template)
    return df.loc[mask].copy()


def extract_vm_analytics(workbook: Workbook) -> dict[str, Any]:
    if "vInfo" not in workbook:
        return {
            "total_vms": 0,
            "powered_on_count": 0,
            "powered_off_count": 0,
            "os_summary": [],
            "os_families": [],
            "applications": [],
            "application_names": [],
            "application_count": 0,
            "application_names_truncated": False,
        }

    df = workbook["vInfo"]
    total_vms = len(df)
    active = _active_vms(df)
    powered_on_count = len(active)

    os_counts: dict[str, int] = {}
    family_counts: dict[str, int] = {}
    applications: list[dict[str, str]] = []
    app_names: list[str] = []

    for _, row in active.iterrows():
        os_name = _resolve_os(row)
        label = _simplify_os_label(os_name)
        os_counts[label] = os_counts.get(label, 0) + 1
        family = _os_family(os_name)
        family_counts[family] = family_counts.get(family, 0) + 1

        vm_name = _safe_str(row.get("VM"))
        dns_name = _safe_str(row.get("DNS Name"))
        display_name = dns_name or vm_name
        if not display_name:
            continue

        applications.append({
            "name": display_name,
            "vm": vm_name,
            "dns_name": dns_name,
            "os": label,
            "cluster": _safe_str(row.get("Cluster")),
            "host": _safe_str(row.get("Host")),
        })
        if display_name not in app_names:
            app_names.append(display_name)

    os_summary = [
        {"os": os_name, "count": count}
        for os_name, count in sorted(os_counts.items(), key=lambda item: (-item[1], item[0].lower()))
    ]
    os_families = [
        {"family": family, "count": count}
        for family, count in sorted(family_counts.items(), key=lambda item: (-item[1], item[0]))
    ]
    applications.sort(key=lambda item: item["name"].lower())
    app_names.sort(key=str.lower)

    listed_names = app_names[:MAX_LIST_ITEMS]
    return {
        "total_vms": total_vms,
        "powered_on_count": powered_on_count,
        "powered_off_count": total_vms - powered_on_count,
        "os_summary": os_summary,
        "os_families": os_families,
        "applications": applications,
        "application_names": listed_names,
        "application_count": len(app_names),
        "application_names_truncated": len(app_names) > MAX_LIST_ITEMS,
    }
