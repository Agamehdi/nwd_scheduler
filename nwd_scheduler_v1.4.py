# -*- coding: utf-8 -*-
"""
NWD Scheduler - Streamlit application

Run:
    streamlit run nwd_scheduler_app.py

Place NWD_Scheduler_Dummy_Input.xlsx in the same folder for automatic loading,
or upload an Excel/CSV file through the application. Outputs are CSV-based.
"""

from __future__ import annotations

import json
import math
import re
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# -----------------------------------------------------------------------------
# Page configuration and constants
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="NWD Scheduler",
    page_icon="🛢️",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_DIR = Path(__file__).resolve().parent
DEFAULT_EXCEL = APP_DIR / "NWD_Scheduler_Dummy_Input.xlsx"
LOCAL_SAVE_FILE = APP_DIR / "NWD_Scheduler_Latest.csv"

REQUIRED_COLUMNS = [
    "Target_ID",
    "Well_Name",
    "Reservoir",
    "Area",
    "Pad",
    "Rig",
    "Target_Type",
    "Start_Date",
    "End_Date",
    "Status",
    "Priority",
    "Progress_Pct",
    "Owner",
    "Campaign",
    "Color",
    "Notes",
]

TEXT_COLUMNS = [
    "Target_ID",
    "Well_Name",
    "Reservoir",
    "Area",
    "Pad",
    "Rig",
    "Target_Type",
    "Status",
    "Priority",
    "Owner",
    "Campaign",
    "Color",
    "Notes",
]

DATE_COLUMNS = ["Start_Date", "End_Date"]

STATUS_OPTIONS = ["Not Started", "Ready", "In Progress", "On Hold", "Completed", "Cancelled"]
PRIORITY_OPTIONS = ["Critical", "High", "Medium", "Low"]
TARGET_TYPE_OPTIONS = ["Producer", "Water Injector", "Observation", "Disposal", "Appraisal"]
RESERVOIR_OPTIONS = ["Main Pay", "Upper Shale", "Mishrif", "Nahr Umr"]
AREA_OPTIONS = ["North", "South"]

STATUS_COLORS = {
    "Not Started": "#94A3B8",
    "Ready": "#38BDF8",
    "In Progress": "#F59E0B",
    "On Hold": "#EF4444",
    "Completed": "#22C55E",
    "Cancelled": "#64748B",
}
PRIORITY_COLORS = {
    "Critical": "#DC2626",
    "High": "#F97316",
    "Medium": "#EAB308",
    "Low": "#3B82F6",
}
DEFAULT_COLOR = "#2563EB"
PALETTES = px.colors.qualitative.Safe + px.colors.qualitative.Set2 + px.colors.qualitative.Bold

COLUMN_ALIASES = {
    "target id": "Target_ID",
    "target_id": "Target_ID",
    "id": "Target_ID",
    "well": "Well_Name",
    "well name": "Well_Name",
    "well_name": "Well_Name",
    "reservoir": "Reservoir",
    "area": "Area",
    "pad": "Pad",
    "cluster": "Pad",
    "rig": "Rig",
    "target type": "Target_Type",
    "well type": "Target_Type",
    "type": "Target_Type",
    "start": "Start_Date",
    "start date": "Start_Date",
    "start_date": "Start_Date",
    "spud date": "Start_Date",
    "end": "End_Date",
    "end date": "End_Date",
    "end_date": "End_Date",
    "completion date": "End_Date",
    "status": "Status",
    "priority": "Priority",
    "progress": "Progress_Pct",
    "progress %": "Progress_Pct",
    "progress_pct": "Progress_Pct",
    "owner": "Owner",
    "engineer": "Owner",
    "campaign": "Campaign",
    "color": "Color",
    "colour": "Color",
    "notes": "Notes",
    "comments": "Notes",
}


# -----------------------------------------------------------------------------
# Styling
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
      .block-container {padding-top: 1.0rem; padding-bottom: 2rem;}
      [data-testid="stMetric"] {
          border: 1px solid rgba(128,128,128,.18);
          border-radius: 12px;
          padding: 10px 14px;
          background: rgba(127,127,127,.035);
      }
      .small-note {font-size: 0.86rem; opacity: .78;}
      .section-card {
          border: 1px solid rgba(128,128,128,.18);
          border-radius: 12px;
          padding: 0.8rem 1rem;
          margin-bottom: 0.75rem;
      }
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# Data helpers
# -----------------------------------------------------------------------------
def create_dummy_dataframe() -> pd.DataFrame:
    """Create deterministic, realistic NWD schedule data."""
    records = [
        ("NWD-001", "R-1401", "Main Pay", "North", "N-Pad-01", "Rig-101", "Producer", "2026-01-10", "2026-02-18", "Completed", "High", 100, "North MP Team", "2026 Base", "#2563EB", "Base development producer"),
        ("NWD-002", "R-1402", "Main Pay", "North", "N-Pad-01", "Rig-101", "Water Injector", "2026-02-22", "2026-04-02", "Completed", "Critical", 100, "WI Team", "2026 Base", "#0EA5E9", "Pattern pressure support"),
        ("NWD-003", "R-1403", "Upper Shale", "North", "N-Pad-02", "Rig-101", "Producer", "2026-04-08", "2026-05-22", "Completed", "Medium", 100, "US Team", "2026 Base", "#8B5CF6", "Upper Shale infill"),
        ("NWD-004", "R-1404", "Mishrif", "North", "N-Pad-03", "Rig-101", "Producer", "2026-06-01", "2026-07-16", "In Progress", "High", 62, "Mishrif Team", "2026 Base", "#F59E0B", "Current drilling target"),
        ("NWD-005", "R-1405", "Main Pay", "North", "N-Pad-03", "Rig-101", "Producer", "2026-07-20", "2026-09-03", "Ready", "High", 5, "North MP Team", "2026 Base", "#22C55E", "Materials confirmed"),
        ("NWD-006", "R-1406", "Main Pay", "North", "N-Pad-04", "Rig-101", "Water Injector", "2026-09-08", "2026-10-18", "Not Started", "Critical", 0, "WI Team", "2026 Base", "#06B6D4", "PWRI linkage required"),
        ("NWD-007", "R-1407", "Nahr Umr", "North", "N-Pad-05", "Rig-101", "Appraisal", "2026-10-25", "2026-12-15", "Not Started", "Medium", 0, "Nahr Umr Team", "2026 Base", "#EC4899", "Appraisal and coring"),
        ("NWD-008", "R-1408", "Main Pay", "North", "N-Pad-05", "Rig-101", "Producer", "2027-01-05", "2027-02-18", "Not Started", "High", 0, "North MP Team", "2027 Base", "#2563EB", "Carry-over candidate"),
        ("NWD-009", "R-2401", "Main Pay", "South", "S-Pad-01", "Rig-202", "Producer", "2026-01-15", "2026-02-28", "Completed", "High", 100, "South MP Team", "2026 Base", "#2563EB", "South development producer"),
        ("NWD-010", "R-2402", "Mishrif", "South", "S-Pad-02", "Rig-202", "Producer", "2026-03-05", "2026-04-20", "Completed", "Medium", 100, "Mishrif Team", "2026 Base", "#8B5CF6", "Mishrif infill"),
        ("NWD-011", "R-2403", "Main Pay", "South", "S-Pad-02", "Rig-202", "Water Injector", "2026-04-25", "2026-06-08", "Completed", "Critical", 100, "WI Team", "2026 Base", "#0EA5E9", "New injection pattern"),
        ("NWD-012", "R-2404", "Upper Shale", "South", "S-Pad-03", "Rig-202", "Producer", "2026-06-12", "2026-07-28", "In Progress", "High", 45, "US Team", "2026 Base", "#F59E0B", "Directional well"),
        ("NWD-013", "R-2405", "Main Pay", "South", "S-Pad-04", "Rig-202", "Producer", "2026-08-02", "2026-09-14", "Ready", "Medium", 0, "South MP Team", "2026 Base", "#22C55E", "Site ready"),
        ("NWD-014", "R-2406", "Mishrif", "South", "S-Pad-04", "Rig-202", "Observation", "2026-09-19", "2026-10-20", "Not Started", "Low", 0, "Surveillance", "2026 Base", "#64748B", "Permanent gauge planned"),
        ("NWD-015", "R-2407", "Main Pay", "South", "S-Pad-05", "Rig-202", "Producer", "2026-10-24", "2026-12-08", "Not Started", "High", 0, "South MP Team", "2026 Base", "#2563EB", "Facility dependency"),
        ("NWD-016", "R-2408", "Nahr Umr", "South", "S-Pad-06", "Rig-202", "Appraisal", "2026-12-12", "2027-02-02", "On Hold", "Medium", 0, "Nahr Umr Team", "2027 Base", "#EF4444", "Pending subsurface maturation"),
        ("NWD-017", "R-3401", "Main Pay", "North", "N-Pad-06", "Rig-303", "Producer", "2026-02-01", "2026-03-19", "Completed", "Medium", 100, "North MP Team", "2026 Base", "#2563EB", "Fast-track target"),
        ("NWD-018", "R-3402", "Upper Shale", "North", "N-Pad-07", "Rig-303", "Producer", "2026-03-24", "2026-05-10", "Completed", "High", 100, "US Team", "2026 Base", "#8B5CF6", "Upper Shale campaign"),
        ("NWD-019", "R-3403", "Main Pay", "North", "N-Pad-07", "Rig-303", "Water Injector", "2026-05-16", "2026-06-28", "Completed", "Critical", 100, "WI Team", "2026 Base", "#0EA5E9", "Injector conversion alternative"),
        ("NWD-020", "R-3404", "Mishrif", "North", "N-Pad-08", "Rig-303", "Producer", "2026-07-05", "2026-08-18", "In Progress", "High", 20, "Mishrif Team", "2026 Base", "#F59E0B", "Possible overlap for conflict test"),
        ("NWD-021", "R-3405", "Main Pay", "North", "N-Pad-08", "Rig-303", "Producer", "2026-08-12", "2026-09-25", "Ready", "High", 0, "North MP Team", "2026 Base", "#22C55E", "Intentional rig overlap example"),
        ("NWD-022", "R-3406", "Main Pay", "South", "S-Pad-07", "Rig-303", "Producer", "2026-10-01", "2026-11-15", "Not Started", "Medium", 0, "South MP Team", "2026 Base", "#2563EB", "Rig move north to south"),
        ("NWD-023", "R-3407", "Mishrif", "South", "S-Pad-08", "Rig-303", "Water Injector", "2026-11-20", "2027-01-05", "Not Started", "Critical", 0, "WI Team", "2027 Base", "#0EA5E9", "High value injector"),
        ("NWD-024", "R-4401", "Main Pay", "North", "N-Pad-09", "Rig-404", "Producer", "2027-01-08", "2027-02-22", "Not Started", "High", 0, "North MP Team", "2027 Base", "#2563EB", "2027 opening target"),
        ("NWD-025", "R-4402", "Upper Shale", "North", "N-Pad-10", "Rig-404", "Producer", "2027-03-01", "2027-04-14", "Not Started", "Medium", 0, "US Team", "2027 Base", "#8B5CF6", "Upper Shale target"),
        ("NWD-026", "R-4403", "Main Pay", "South", "S-Pad-09", "Rig-404", "Water Injector", "2027-04-20", "2027-06-03", "Not Started", "Critical", 0, "WI Team", "2027 Base", "#0EA5E9", "Pattern support"),
        ("NWD-027", "R-4404", "Mishrif", "South", "S-Pad-10", "Rig-404", "Producer", "2027-06-10", "2027-07-25", "Not Started", "High", 0, "Mishrif Team", "2027 Base", "#EC4899", "Facility tie-in required"),
        ("NWD-028", "R-4405", "Nahr Umr", "South", "S-Pad-11", "Rig-404", "Appraisal", "2027-08-01", "2027-09-20", "Not Started", "Medium", 0, "Nahr Umr Team", "2027 Base", "#64748B", "Data acquisition well"),
        ("NWD-029", "R-4406", "Main Pay", "North", "N-Pad-11", "Rig-404", "Producer", "2028-01-10", "2028-02-24", "Not Started", "High", 0, "North MP Team", "2028 Base", "#2563EB", "Long-range target"),
        ("NWD-030", "R-4407", "Main Pay", "South", "S-Pad-12", "Rig-404", "Producer", "2029-03-05", "2029-04-18", "Not Started", "Medium", 0, "South MP Team", "2029 Base", "#2563EB", "Long-range placeholder"),
    ]
    df = pd.DataFrame(records, columns=REQUIRED_COLUMNS)
    return normalize_dataframe(df)


def clean_column_name(value: object) -> str:
    text = re.sub(r"\s+", " ", str(value).strip())
    alias = COLUMN_ALIASES.get(text.lower())
    return alias if alias else text.replace(" ", "_")


def valid_hex_color(value: object) -> str:
    value = str(value or "").strip()
    if re.fullmatch(r"#[0-9A-Fa-f]{6}", value):
        return value.upper()
    return DEFAULT_COLOR


def next_target_id(existing: Iterable[str]) -> str:
    numbers = []
    for item in existing:
        match = re.search(r"(\d+)$", str(item))
        if match:
            numbers.append(int(match.group(1)))
    next_number = max(numbers, default=0) + 1
    return f"NWD-{next_number:03d}"


def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column names, data types, missing values and computed columns."""
    if df is None or df.empty:
        df = pd.DataFrame(columns=REQUIRED_COLUMNS)
    else:
        df = df.copy()

    df.columns = [clean_column_name(c) for c in df.columns]

    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            if col == "Progress_Pct":
                df[col] = 0
            elif col == "Color":
                df[col] = DEFAULT_COLOR
            else:
                df[col] = ""

    # Keep supported columns first, then any extra user columns.
    extra_cols = [c for c in df.columns if c not in REQUIRED_COLUMNS and c not in {"Duration_Days", "Start_Year", "End_Year"}]
    df = df[REQUIRED_COLUMNS + extra_cols]

    for col in DATE_COLUMNS:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.normalize()

    for col in TEXT_COLUMNS:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    df["Progress_Pct"] = pd.to_numeric(df["Progress_Pct"], errors="coerce").fillna(0).clip(0, 100).round(0).astype(int)
    df["Color"] = df["Color"].apply(valid_hex_color)

    # Assign IDs to blank/new rows and make duplicated IDs unique.
    existing: List[str] = []
    new_ids: List[str] = []
    for raw in df["Target_ID"].tolist():
        candidate = str(raw).strip()
        if not candidate or candidate.lower() == "nan" or candidate in existing:
            candidate = next_target_id(existing + df["Target_ID"].astype(str).tolist() + new_ids)
        existing.append(candidate)
        new_ids.append(candidate)
    df["Target_ID"] = new_ids

    # Fill sensible defaults.
    defaults = {
        "Status": "Not Started",
        "Priority": "Medium",
        "Target_Type": "Producer",
        "Reservoir": "Main Pay",
        "Area": "North",
        "Rig": "Unassigned",
        "Campaign": "Unassigned",
    }
    for col, value in defaults.items():
        df.loc[df[col].eq(""), col] = value

    df["Duration_Days"] = (df["End_Date"] - df["Start_Date"]).dt.days + 1
    df.loc[df["Duration_Days"] < 1, "Duration_Days"] = np.nan
    df["Start_Year"] = df["Start_Date"].dt.year.astype("Int64")
    df["End_Year"] = df["End_Date"].dt.year.astype("Int64")

    return df.reset_index(drop=True)


def load_schedule_file(source: object, filename: str = "") -> pd.DataFrame:
    """Load the NWD schedule from Excel or CSV."""
    suffix = Path(filename).suffix.lower() if filename else ""

    if suffix == ".csv":
        try:
            df = pd.read_csv(source, encoding="utf-8-sig")
        except UnicodeDecodeError:
            if hasattr(source, "seek"):
                source.seek(0)
            df = pd.read_csv(source, encoding="latin-1")
        return normalize_dataframe(df)

    excel = pd.ExcelFile(source)
    preferred = next(
        (s for s in excel.sheet_names if s.lower() in {"nwd_schedule", "schedule", "targets"}),
        excel.sheet_names[0],
    )
    df = pd.read_excel(excel, sheet_name=preferred)
    return normalize_dataframe(df)


def df_fingerprint(df: pd.DataFrame) -> str:
    canonical = normalize_dataframe(df).copy()
    for col in DATE_COLUMNS:
        canonical[col] = canonical[col].dt.strftime("%Y-%m-%d").fillna("")
    return canonical.to_json(orient="records", date_format="iso")


def push_undo(df: pd.DataFrame) -> None:
    stack = st.session_state.setdefault("undo_stack", [])
    stack.append(df.copy(deep=True))
    if len(stack) > 15:
        del stack[0]


def set_master(df: pd.DataFrame, add_undo: bool = True) -> None:
    new_df = normalize_dataframe(df)
    current = st.session_state.get("schedule_df")
    if add_undo and current is not None and df_fingerprint(current) != df_fingerprint(new_df):
        push_undo(current)
    st.session_state.schedule_df = new_df



def merge_filtered_edits(
    master_df: pd.DataFrame,
    original_filtered_df: pd.DataFrame,
    edited_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Merge edits made in the filtered data editor back into the full master schedule.

    Behaviour:
    - Rows outside the active filters remain unchanged.
    - Edited visible rows replace their original versions.
    - Rows deleted from the editor are removed from the master schedule.
    - Newly added editor rows are added to the master schedule.
    - Blank or duplicated Target_ID values are corrected by normalize_dataframe().
    """
    master = normalize_dataframe(master_df)
    original_visible = normalize_dataframe(original_filtered_df)

    if edited_df is None:
        edited = pd.DataFrame(columns=original_visible.columns)
    else:
        edited = edited_df.copy()

    original_visible_ids = set(
        original_visible["Target_ID"]
        .fillna("")
        .astype(str)
        .str.strip()
        .tolist()
    )

    untouched_rows = master[
        ~master["Target_ID"]
        .fillna("")
        .astype(str)
        .str.strip()
        .isin(original_visible_ids)
    ].copy()

    edited_rows = normalize_dataframe(edited)

    merged = pd.concat(
        [untouched_rows, edited_rows],
        ignore_index=True,
        sort=False,
    )

    return normalize_dataframe(merged)

def initialize_state() -> None:
    if "schedule_df" not in st.session_state:
        if DEFAULT_EXCEL.exists():
            try:
                st.session_state.schedule_df = load_schedule_file(DEFAULT_EXCEL, DEFAULT_EXCEL.name)
                st.session_state.source_name = DEFAULT_EXCEL.name
            except Exception:
                st.session_state.schedule_df = create_dummy_dataframe()
                st.session_state.source_name = "Built-in dummy schedule"
        else:
            st.session_state.schedule_df = create_dummy_dataframe()
            st.session_state.source_name = "Built-in dummy schedule"
    st.session_state.setdefault("undo_stack", [])
    st.session_state.setdefault("filter_reset_token", 0)


def select_options(df: pd.DataFrame, column: str) -> List[str]:
    return sorted([x for x in df[column].dropna().astype(str).unique().tolist() if x])


def apply_filters(
    df: pd.DataFrame,
    search: str,
    reservoirs: Sequence[str],
    areas: Sequence[str],
    rigs: Sequence[str],
    statuses: Sequence[str],
    priorities: Sequence[str],
    target_types: Sequence[str],
    campaigns: Sequence[str],
    owners: Sequence[str],
    date_window: Tuple[date, date],
    include_cancelled: bool,
) -> pd.DataFrame:
    filtered = df.copy()
    if search:
        search_lower = search.lower().strip()
        searchable = filtered.fillna("").astype(str).agg(" | ".join, axis=1).str.lower()
        filtered = filtered[searchable.str.contains(re.escape(search_lower), regex=True)]
    mappings = [
        ("Reservoir", reservoirs),
        ("Area", areas),
        ("Rig", rigs),
        ("Status", statuses),
        ("Priority", priorities),
        ("Target_Type", target_types),
        ("Campaign", campaigns),
        ("Owner", owners),
    ]
    for column, values in mappings:
        if values:
            filtered = filtered[filtered[column].isin(values)]
    if not include_cancelled:
        filtered = filtered[filtered["Status"] != "Cancelled"]

    start_filter = pd.Timestamp(date_window[0])
    end_filter = pd.Timestamp(date_window[1])
    # Include schedules that overlap the selected date window.
    filtered = filtered[(filtered["End_Date"] >= start_filter) & (filtered["Start_Date"] <= end_filter)]
    return filtered.reset_index(drop=True)


def validate_schedule(df: pd.DataFrame) -> pd.DataFrame:
    issues: List[Dict[str, object]] = []
    if df.empty:
        return pd.DataFrame(columns=["Severity", "Target_ID", "Field", "Issue"])

    duplicated = df["Target_ID"].duplicated(keep=False)
    for _, row in df[duplicated].iterrows():
        issues.append({"Severity": "Error", "Target_ID": row["Target_ID"], "Field": "Target_ID", "Issue": "Target ID is duplicated."})

    for _, row in df.iterrows():
        target = row["Target_ID"]
        for col in ["Well_Name", "Rig", "Start_Date", "End_Date"]:
            if pd.isna(row[col]) or str(row[col]).strip() == "":
                issues.append({"Severity": "Error", "Target_ID": target, "Field": col, "Issue": f"{col} is missing."})
        if pd.notna(row["Start_Date"]) and pd.notna(row["End_Date"]) and row["End_Date"] < row["Start_Date"]:
            issues.append({"Severity": "Error", "Target_ID": target, "Field": "End_Date", "Issue": "End date is before start date."})
        if not 0 <= int(row["Progress_Pct"]) <= 100:
            issues.append({"Severity": "Error", "Target_ID": target, "Field": "Progress_Pct", "Issue": "Progress must be between 0 and 100."})
        if row["Status"] == "Completed" and int(row["Progress_Pct"]) < 100:
            issues.append({"Severity": "Warning", "Target_ID": target, "Field": "Progress_Pct", "Issue": "Completed target has progress below 100%."})
        if row["Status"] != "Completed" and int(row["Progress_Pct"]) == 100:
            issues.append({"Severity": "Warning", "Target_ID": target, "Field": "Status", "Issue": "Progress is 100% but status is not Completed."})

    return pd.DataFrame(issues)


def find_rig_conflicts(df: pd.DataFrame) -> pd.DataFrame:
    conflicts: List[Dict[str, object]] = []
    valid = df.dropna(subset=["Start_Date", "End_Date"]).copy()
    valid = valid[valid["Status"] != "Cancelled"]
    for rig, group in valid.groupby("Rig"):
        group = group.sort_values(["Start_Date", "End_Date"]).reset_index(drop=True)
        for i in range(len(group)):
            a = group.iloc[i]
            for j in range(i + 1, len(group)):
                b = group.iloc[j]
                if b["Start_Date"] > a["End_Date"]:
                    break
                overlap_start = max(a["Start_Date"], b["Start_Date"])
                overlap_end = min(a["End_Date"], b["End_Date"])
                if overlap_start <= overlap_end:
                    conflicts.append(
                        {
                            "Rig": rig,
                            "Target_1": a["Target_ID"],
                            "Well_1": a["Well_Name"],
                            "Target_2": b["Target_ID"],
                            "Well_2": b["Well_Name"],
                            "Overlap_Start": overlap_start,
                            "Overlap_End": overlap_end,
                            "Overlap_Days": (overlap_end - overlap_start).days + 1,
                        }
                    )
    return pd.DataFrame(conflicts)


def union_days(group: pd.DataFrame, window_start: pd.Timestamp, window_end: pd.Timestamp) -> int:
    intervals: List[Tuple[pd.Timestamp, pd.Timestamp]] = []
    for _, row in group.dropna(subset=["Start_Date", "End_Date"]).iterrows():
        start = max(row["Start_Date"], window_start)
        end = min(row["End_Date"], window_end)
        if start <= end:
            intervals.append((start, end))
    if not intervals:
        return 0
    intervals.sort(key=lambda x: x[0])
    merged = [intervals[0]]
    for start, end in intervals[1:]:
        last_start, last_end = merged[-1]
        if start <= last_end + pd.Timedelta(days=1):
            merged[-1] = (last_start, max(last_end, end))
        else:
            merged.append((start, end))
    return sum((end - start).days + 1 for start, end in merged)


def calculate_rig_utilization(df: pd.DataFrame, window: Tuple[date, date]) -> pd.DataFrame:
    start = pd.Timestamp(window[0])
    end = pd.Timestamp(window[1])
    total_days = max((end - start).days + 1, 1)
    records = []
    for rig, group in df[df["Status"] != "Cancelled"].groupby("Rig"):
        scheduled = union_days(group, start, end)
        records.append({"Rig": rig, "Scheduled_Days": scheduled, "Window_Days": total_days, "Utilization_Pct": round(100 * scheduled / total_days, 1)})
    return pd.DataFrame(records).sort_values("Utilization_Pct", ascending=False) if records else pd.DataFrame(columns=["Rig", "Scheduled_Days", "Window_Days", "Utilization_Pct"])


def category_color_map(values: Iterable[str]) -> Dict[str, str]:
    unique = sorted(set(str(v) for v in values))
    return {value: PALETTES[i % len(PALETTES)] for i, value in enumerate(unique)}


def build_gantt(
    df: pd.DataFrame,
    color_mode: str,
    group_by: str,
    label_mode: str,
    gantt_type: str,
    y_axis_column: str,
    bar_text_mode: str,
    show_today: bool,
    show_progress: bool,
    show_horizontal_grid: bool,
    height_per_row: int,
) -> go.Figure:
    """
    Build an interactive Gantt chart.

    Supported chart types:
    - Detailed target rows:
        One Y-axis row per target, matching the original application behaviour.
    - Compact merged lanes:
        One Y-axis row per selected column value. For example, all targets for the
        same Rig, Pad or Reservoir are drawn on the same lane.
    - Grouped target rows:
        One row per target, with the selected Y-axis column shown as a prefix.

    The dataframe supplied to this function is already filtered, so merged lanes
    automatically contain only the currently visible schedule.
    """
    if df.empty:
        fig = go.Figure()
        fig.add_annotation(
            text="No targets match the selected filters",
            x=0.5,
            y=0.5,
            showarrow=False,
            font={"size": 18},
        )
        fig.update_layout(
            height=420,
            xaxis={"visible": False},
            yaxis={"visible": False},
        )
        return fig

    plot_df = df.dropna(subset=["Start_Date", "End_Date"]).copy()
    if plot_df.empty:
        fig = go.Figure()
        fig.add_annotation(
            text="Visible targets do not have valid start and end dates",
            x=0.5,
            y=0.5,
            showarrow=False,
            font={"size": 18},
        )
        fig.update_layout(
            height=420,
            xaxis={"visible": False},
            yaxis={"visible": False},
        )
        return fig

    if y_axis_column not in plot_df.columns:
        y_axis_column = "Rig"
    if group_by not in plot_df.columns:
        group_by = "Rig"

    # Keep blank grouping values visible instead of dropping them.
    plot_df[y_axis_column] = (
        plot_df[y_axis_column]
        .fillna("")
        .astype(str)
        .str.strip()
        .replace("", "Unassigned")
    )
    plot_df[group_by] = (
        plot_df[group_by]
        .fillna("")
        .astype(str)
        .str.strip()
        .replace("", "Unassigned")
    )

    # Target-label source used both on the detailed Y-axis and inside bars.
    if label_mode == "Well name":
        plot_df["Task_Label"] = plot_df["Well_Name"].fillna("").astype(str)
    elif label_mode == "Target ID + well":
        plot_df["Task_Label"] = (
            plot_df["Target_ID"].fillna("").astype(str)
            + " | "
            + plot_df["Well_Name"].fillna("").astype(str)
        )
    elif label_mode == "Rig + well":
        plot_df["Task_Label"] = (
            plot_df["Rig"].fillna("").astype(str)
            + " | "
            + plot_df["Well_Name"].fillna("").astype(str)
        )
    else:
        plot_df["Task_Label"] = plot_df["Target_ID"].fillna("").astype(str)

    sort_field = group_by if gantt_type == "Detailed target rows" else y_axis_column
    plot_df = plot_df.sort_values(
        [sort_field, "Start_Date", "End_Date", "Rig", "Well_Name"],
        ascending=[True, True, True, True, True],
    ).reset_index(drop=True)

    def make_unique_labels(values: Sequence[str]) -> List[str]:
        """Make duplicate detailed labels unique without changing the base text."""
        seen: Dict[str, int] = {}
        result: List[str] = []
        for raw_value in values:
            value = str(raw_value)
            seen[value] = seen.get(value, 0) + 1
            result.append(value if seen[value] == 1 else f"{value} ({seen[value]})")
        return result

    if gantt_type == "Compact merged lanes":
        # Repeated values intentionally remain repeated: Plotly draws every target
        # assigned to that value on one shared Y-axis lane.
        plot_df["Y_Label"] = plot_df[y_axis_column].astype(str)
        separator_field = y_axis_column
        type_description = f"Compact lanes by {y_axis_column.replace('_', ' ')}"
    elif gantt_type == "Grouped target rows":
        combined_labels = (
            plot_df[y_axis_column].astype(str)
            + "  |  "
            + plot_df["Task_Label"].astype(str)
        )
        plot_df["Y_Label"] = make_unique_labels(combined_labels.tolist())
        separator_field = y_axis_column
        type_description = f"Grouped rows by {y_axis_column.replace('_', ' ')}"
    else:
        plot_df["Y_Label"] = make_unique_labels(plot_df["Task_Label"].tolist())
        separator_field = group_by
        type_description = "Detailed target rows"

    # Preserve the first filtered/sorted appearance of each Y-axis category.
    y_categories = list(dict.fromkeys(plot_df["Y_Label"].astype(str).tolist()))
    row_count = len(y_categories)

    if color_mode == "Custom row color":
        colors = plot_df["Color"].apply(valid_hex_color).tolist()
        legend_title = None
    else:
        field = color_mode
        if field == "Status":
            cmap = {**category_color_map(plot_df[field]), **STATUS_COLORS}
        elif field == "Priority":
            cmap = {**category_color_map(plot_df[field]), **PRIORITY_COLORS}
        else:
            cmap = category_color_map(plot_df[field])
        colors = plot_df[field].map(cmap).fillna(DEFAULT_COLOR).tolist()
        legend_title = field.replace("_", " ")

    duration_ms = (
        (plot_df["End_Date"] - plot_df["Start_Date"]).dt.total_seconds() * 1000
    ).clip(lower=86_400_000)

    progress_text = plot_df["Progress_Pct"].astype(int).astype(str) + "%"
    target_name_text = (
        plot_df["Well_Name"]
        .fillna("")
        .astype(str)
        .str.strip()
    )
    target_name_text = target_name_text.where(
        target_name_text.ne(""),
        plot_df["Target_ID"].fillna("").astype(str),
    )

    if bar_text_mode == "Target name + progress %":
        bar_text = target_name_text + "<br>" + progress_text
    elif bar_text_mode == "Target name only":
        bar_text = target_name_text
    elif bar_text_mode == "Progress % only":
        bar_text = progress_text
    else:
        bar_text = pd.Series([""] * len(plot_df), index=plot_df.index)

    # Text is drawn in a separate final trace after progress shading, preventing
    # the darker progress overlay from hiding the target name.
    text_midpoint_dates = (
        plot_df["Start_Date"]
        + (plot_df["End_Date"] - plot_df["Start_Date"]) / 2
    )

    customdata = np.stack(
        [
            plot_df["Target_ID"],
            plot_df["Well_Name"],
            plot_df["Rig"],
            plot_df["Pad"],
            plot_df["Reservoir"],
            plot_df["Area"],
            plot_df["Status"],
            plot_df["Priority"],
            plot_df["Progress_Pct"],
            plot_df["Start_Date"].dt.strftime("%d-%b-%Y"),
            plot_df["End_Date"].dt.strftime("%d-%b-%Y"),
            plot_df["Duration_Days"].fillna(0).astype(int),
            plot_df["Owner"],
            plot_df["Campaign"],
            plot_df["Notes"],
        ],
        axis=-1,
    )

    fig = go.Figure()

    # Main schedule bars.
    fig.add_trace(
        go.Bar(
            x=duration_ms,
            base=plot_df["Start_Date"],
            y=plot_df["Y_Label"],
            orientation="h",
            marker={
                "color": colors,
                "line": {
                    "width": 0.7,
                    "color": "rgba(30,41,59,.52)",
                },
            },
            # Bar labels are added later as a dedicated top-layer text trace.
            text=None,
            customdata=customdata,
            hovertemplate=(
                "<b>%{customdata[0]} | %{customdata[1]}</b><br>"
                "Rig: %{customdata[2]}<br>"
                "Pad: %{customdata[3]}<br>"
                "Reservoir: %{customdata[4]}<br>"
                "Area: %{customdata[5]}<br>"
                "Status: %{customdata[6]}<br>"
                "Priority: %{customdata[7]}<br>"
                "Progress: %{customdata[8]}%<br>"
                "Start: %{customdata[9]}<br>"
                "End: %{customdata[10]}<br>"
                "Duration: %{customdata[11]} days<br>"
                "Owner: %{customdata[12]}<br>"
                "Campaign: %{customdata[13]}<br>"
                "Notes: %{customdata[14]}"
                "<extra></extra>"
            ),
            showlegend=False,
            name="Targets",
        )
    )

    # Optional progress shading from target start to current planned progress.
    if show_progress:
        progress_ms = (
            duration_ms
            * plot_df["Progress_Pct"].astype(float).clip(lower=0, upper=100)
            / 100.0
        )
        fig.add_trace(
            go.Bar(
                x=progress_ms,
                base=plot_df["Start_Date"],
                y=plot_df["Y_Label"],
                orientation="h",
                marker={
                    "color": "rgba(15,23,42,.32)",
                    "line": {"width": 0},
                },
                hoverinfo="skip",
                showlegend=False,
                name="Progress",
            )
        )

    # Dedicated top-layer labels: always above both the main bar and the
    # optional progress shading. This keeps target names visible.
    if bar_text_mode != "No text":
        fig.add_trace(
            go.Scatter(
                x=text_midpoint_dates,
                y=plot_df["Y_Label"],
                mode="text",
                text=bar_text,
                textposition="middle center",
                textfont={
                    "color": "white",
                    "size": 11,
                    "family": "Arial, sans-serif",
                },
                hoverinfo="skip",
                showlegend=False,
                cliponaxis=True,
                name="Bar labels",
            )
        )

    if show_today:
        today = pd.Timestamp.today().normalize()
        fig.add_vline(
            x=today,
            line_width=2,
            line_dash="dash",
            line_color="#DC2626",
        )
        fig.add_annotation(
            x=today,
            y=1.02,
            yref="paper",
            text="Today",
            showarrow=False,
            font={"color": "#DC2626", "size": 11},
        )

    # Full horizontal row/lane grid, including top and bottom boundaries.
    if show_horizontal_grid:
        for boundary_index in range(row_count + 1):
            fig.add_hline(
                y=boundary_index - 0.5,
                line_width=0.8,
                line_dash="solid",
                line_color="rgba(148,163,184,.28)",
                layer="below",
            )

    # Stronger separators when the selected grouping value changes.
    sorted_groups = plot_df[separator_field].astype(str).tolist()
    if gantt_type != "Compact merged lanes":
        for row_index in range(1, len(sorted_groups)):
            if sorted_groups[row_index] != sorted_groups[row_index - 1]:
                fig.add_hline(
                    y=row_index - 0.5,
                    line_width=1.6,
                    line_dash="solid",
                    line_color="rgba(71,85,105,.52)",
                    layer="below",
                )

    chart_height = max(
        420 if gantt_type == "Compact merged lanes" else 480,
        min(1800, 150 + max(row_count, 1) * height_per_row),
    )
    title = (
        f"NWD Drilling Schedule — {len(plot_df)} "
        f"target{'s' if len(plot_df) != 1 else ''}"
    )

    fig.update_layout(
        title={
            "text": title,
            "subtitle": {
                "text": type_description,
                "font": {"size": 12, "color": "#64748B"},
            },
            "x": 0.01,
            "xanchor": "left",
            "font": {"size": 20},
        },
        height=chart_height,
        barmode="overlay",
        bargap=0.12 if gantt_type == "Compact merged lanes" else 0.24,
        hoverlabel={"align": "left"},
        margin={"l": 20, "r": 25, "t": 88, "b": 25},
        uniformtext={"mode": "hide", "minsize": 8},
        xaxis={
            "title": "Calendar date",
            "type": "date",
            "showgrid": True,
            "gridcolor": "rgba(148,163,184,.20)",
            "gridwidth": 1,
            "rangeslider": {
                "visible": True,
                "thickness": 0.06,
            },
            "rangeselector": {
                "buttons": [
                    {
                        "count": 3,
                        "label": "3m",
                        "step": "month",
                        "stepmode": "backward",
                    },
                    {
                        "count": 6,
                        "label": "6m",
                        "step": "month",
                        "stepmode": "backward",
                    },
                    {
                        "count": 1,
                        "label": "1y",
                        "step": "year",
                        "stepmode": "backward",
                    },
                    {
                        "step": "all",
                        "label": "All",
                    },
                ]
            },
        },
        yaxis={
            "title": y_axis_column.replace("_", " ")
            if gantt_type != "Detailed target rows"
            else "",
            "autorange": "reversed",
            "showgrid": show_horizontal_grid,
            "gridcolor": "rgba(148,163,184,.18)",
            "gridwidth": 1,
            "tickfont": {"size": 11},
            "categoryorder": "array",
            "categoryarray": y_categories,
            "automargin": True,
        },
        font={"family": "Arial, sans-serif"},
    )

    # Lightweight categorical legend, preserving the original implementation.
    if color_mode != "Custom row color":
        field = color_mode
        categories = plot_df[field].dropna().astype(str).unique().tolist()
        if field == "Status":
            cmap = {**category_color_map(categories), **STATUS_COLORS}
        elif field == "Priority":
            cmap = {**category_color_map(categories), **PRIORITY_COLORS}
        else:
            cmap = category_color_map(categories)

        for category in categories:
            fig.add_trace(
                go.Bar(
                    x=[None],
                    y=[None],
                    marker={"color": cmap.get(category, DEFAULT_COLOR)},
                    name=category,
                    orientation="h",
                    showlegend=True,
                    hoverinfo="skip",
                )
            )

        fig.update_layout(
            legend={
                "title": {"text": legend_title},
                "orientation": "h",
                "y": -0.14,
                "x": 0,
            }
        )

    return fig


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    """Return a normalized dataframe as UTF-8-SIG CSV bytes for Excel compatibility."""
    export_df = normalize_dataframe(df).copy()
    for col in DATE_COLUMNS:
        export_df[col] = pd.to_datetime(export_df[col], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    return export_df.to_csv(index=False).encode("utf-8-sig")


def make_csv_package(
    master_df: pd.DataFrame,
    filtered_df: pd.DataFrame,
    conflicts_df: pd.DataFrame,
    issues_df: pd.DataFrame,
) -> bytes:
    """Create a ZIP package containing all scheduler outputs as CSV files."""
    output = __import__("io").BytesIO()

    files = {
        "NWD_Schedule.csv": dataframe_to_csv_bytes(master_df),
        "Filtered_View.csv": dataframe_to_csv_bytes(filtered_df),
        "Rig_Conflicts.csv": conflicts_df.to_csv(index=False).encode("utf-8-sig"),
        "Data_Quality.csv": issues_df.to_csv(index=False).encode("utf-8-sig"),
    }

    summary_df = pd.DataFrame(
        {
            "Metric": [
                "Total Targets",
                "Filtered Targets",
                "Unique Rigs",
                "Total Planned Days",
                "Rig Conflict Pairs",
                "Data Quality Issues",
                "Generated At",
            ],
            "Value": [
                len(master_df),
                len(filtered_df),
                master_df["Rig"].replace("", np.nan).nunique(),
                pd.to_numeric(master_df["Duration_Days"], errors="coerce").fillna(0).sum(),
                len(conflicts_df),
                len(issues_df),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ],
        }
    )
    files["Summary.csv"] = summary_df.to_csv(index=False).encode("utf-8-sig")

    with zipfile.ZipFile(output, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        for filename, content in files.items():
            archive.writestr(filename, content)

    output.seek(0)
    return output.getvalue()


# -----------------------------------------------------------------------------
# App state and header
# -----------------------------------------------------------------------------
initialize_state()
master_df = normalize_dataframe(st.session_state.schedule_df)
st.session_state.schedule_df = master_df

header_col, status_col = st.columns([4, 1.2])
with header_col:
    st.title("🛢️ NWD Scheduler")
    st.caption("Interactive drilling schedule editor, Gantt planner, rig-conflict checker and CSV exporter")
with status_col:
    st.markdown(f"**Source:**  \n{st.session_state.get('source_name', 'Session data')}")
    st.caption(f"Last refresh: {datetime.now():%d-%b-%Y %H:%M}")


# -----------------------------------------------------------------------------
# Sidebar: data source, filters and chart settings
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("Data & Controls")

    uploaded = st.file_uploader("Upload schedule Excel or CSV", type=["xlsx", "xls", "xlsm", "csv"], help="For Excel, the app looks for NWD_Schedule, Schedule or Targets; otherwise it reads the first sheet. CSV is read directly.")
    col_load, col_reset = st.columns(2)
    if col_load.button("Load file", use_container_width=True, disabled=uploaded is None):
        try:
            loaded = load_schedule_file(uploaded, uploaded.name)
            set_master(loaded, add_undo=True)
            st.session_state.source_name = uploaded.name
            st.success(f"Loaded {len(loaded)} targets")
            st.rerun()
        except Exception as exc:
            st.error(f"Could not load the schedule file: {exc}")

    if col_reset.button("Load dummy", use_container_width=True):
        set_master(create_dummy_dataframe(), add_undo=True)
        st.session_state.source_name = "Built-in dummy schedule"
        st.rerun()

    if st.button("↩ Undo last change", use_container_width=True, disabled=not st.session_state.undo_stack):
        previous = st.session_state.undo_stack.pop()
        st.session_state.schedule_df = normalize_dataframe(previous)
        st.toast("Last change undone")
        st.rerun()

    st.divider()
    st.subheader("Filters")
    reset_token = st.session_state.filter_reset_token

    search = st.text_input("Search all fields", key=f"search_{reset_token}", placeholder="Well, target, owner, note...")
    reservoirs = st.multiselect("Reservoir", select_options(master_df, "Reservoir"), key=f"reservoir_{reset_token}")
    areas = st.multiselect("Area", select_options(master_df, "Area"), key=f"area_{reset_token}")
    rigs = st.multiselect("Rig", select_options(master_df, "Rig"), key=f"rig_{reset_token}")
    statuses = st.multiselect("Status", select_options(master_df, "Status"), key=f"status_{reset_token}")
    priorities = st.multiselect("Priority", select_options(master_df, "Priority"), key=f"priority_{reset_token}")
    target_types = st.multiselect("Target type", select_options(master_df, "Target_Type"), key=f"type_{reset_token}")
    campaigns = st.multiselect("Campaign", select_options(master_df, "Campaign"), key=f"campaign_{reset_token}")
    owners = st.multiselect("Owner", select_options(master_df, "Owner"), key=f"owner_{reset_token}")
    include_cancelled = st.checkbox("Include cancelled targets", value=False, key=f"cancelled_{reset_token}")

    valid_dates = pd.concat([master_df["Start_Date"], master_df["End_Date"]]).dropna()
    if valid_dates.empty:
        min_date = date.today() - timedelta(days=30)
        max_date = date.today() + timedelta(days=365)
    else:
        min_date = valid_dates.min().date()
        max_date = valid_dates.max().date()
    date_window = st.date_input(
        "Schedule window",
        value=(min_date, max_date),
        min_value=min_date - timedelta(days=365),
        max_value=max_date + timedelta(days=365),
        key=f"date_{reset_token}",
    )
    if not isinstance(date_window, tuple) or len(date_window) != 2:
        date_window = (min_date, max_date)

    if st.button("Clear all filters", use_container_width=True):
        st.session_state.filter_reset_token += 1
        st.rerun()

    st.divider()
    st.subheader("Gantt settings")

    gantt_type = st.selectbox(
        "Gantt chart type",
        [
            "Detailed target rows",
            "Compact merged lanes",
            "Grouped target rows",
        ],
        index=0,
        help=(
            "Detailed gives one row per target. Compact merges all visible targets "
            "sharing the selected Y-axis value onto one lane. Grouped keeps one row "
            "per target and prefixes it with the selected Y-axis column."
        ),
    )

    y_axis_column = st.selectbox(
        "Left-side Y-axis column",
        [
            "Rig",
            "Pad",
            "Reservoir",
            "Area",
            "Campaign",
            "Owner",
            "Status",
            "Priority",
            "Target_Type",
        ],
        index=0,
        help=(
            "Used by Compact merged lanes and Grouped target rows. "
            "The lanes update automatically after every active filter."
        ),
    )

    color_mode = st.selectbox(
        "Color bars by",
        [
            "Custom row color",
            "Status",
            "Priority",
            "Rig",
            "Reservoir",
            "Area",
            "Target_Type",
            "Campaign",
        ],
        index=0,
    )

    group_by = st.selectbox(
        "Sort/group detailed rows by",
        ["Rig", "Pad", "Area", "Reservoir", "Campaign", "Owner", "Status"],
        index=0,
        help="Used as the main sorting and separator field in Detailed target rows.",
    )

    label_mode = st.selectbox(
        "Target label format",
        ["Target ID + well", "Well name", "Rig + well", "Target ID"],
        index=0,
        help=(
            "Controls detailed row labels and the target-name portion of text "
            "shown inside each Gantt bar."
        ),
    )

    bar_text_mode = st.selectbox(
        "Text inside bars",
        [
            "Target name + progress %",
            "Target name only",
            "Progress % only",
            "No text",
        ],
        index=0,
        help="Target name uses the Well_Name field and is drawn above progress shading so it remains visible.",
    )

    show_today = st.checkbox("Show today line", value=True)
    show_progress = st.checkbox("Show progress shading", value=True)
    show_horizontal_grid = st.checkbox("Show horizontal grid lines", value=True)
    height_per_row = st.slider(
        "Row/lane height",
        min_value=24,
        max_value=60,
        value=34,
        step=2,
    )

filtered_df = apply_filters(
    master_df,
    search,
    reservoirs,
    areas,
    rigs,
    statuses,
    priorities,
    target_types,
    campaigns,
    owners,
    date_window,
    include_cancelled,
)

conflicts_df = find_rig_conflicts(filtered_df)
issues_df = validate_schedule(master_df)
rig_util_df = calculate_rig_utilization(filtered_df, date_window)


# -----------------------------------------------------------------------------
# KPIs
# -----------------------------------------------------------------------------
k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Visible targets", f"{len(filtered_df):,}", delta=f"of {len(master_df):,} total")
k2.metric("Active rigs", f"{filtered_df['Rig'].nunique():,}")
k3.metric("Scheduled days", f"{int(filtered_df['Duration_Days'].fillna(0).sum()):,}")
k4.metric("Average progress", f"{filtered_df['Progress_Pct'].mean() if len(filtered_df) else 0:.0f}%")
k5.metric("Rig conflicts", f"{len(conflicts_df):,}", delta="Needs review" if len(conflicts_df) else "Clear", delta_color="inverse")
k6.metric("Data issues", f"{len(issues_df):,}", delta="Master data", delta_color="inverse")


# -----------------------------------------------------------------------------
# Main tabs
# -----------------------------------------------------------------------------
tab_gantt, tab_edit, tab_actions, tab_quality, tab_summary, tab_help = st.tabs(
    ["📅 Gantt", "✏️ Edit Schedule", "⚡ Bulk Actions", "⚠️ Quality & Conflicts", "📊 Summary", "ℹ️ Guide"]
)

with tab_gantt:
    gantt = build_gantt(
        filtered_df,
        color_mode,
        group_by,
        label_mode,
        gantt_type,
        y_axis_column,
        bar_text_mode,
        show_today,
        show_progress,
        show_horizontal_grid,
        height_per_row,
    )
    st.plotly_chart(
        gantt,
        use_container_width=True,
        config={
            "displaylogo": False,
            "scrollZoom": True,
            "toImageButtonOptions": {
                "format": "png",
                "filename": "NWD_Gantt",
                "scale": 2,
            },
        },
    )
    st.caption(
        "Tip: choose Compact merged lanes to place all visible targets for the "
        "same Rig, Pad, Reservoir or other selected column on one Y-axis lane. "
        "Use the range slider, hover details and camera icon as before."
    )

with tab_edit:
    st.subheader("Editable filtered schedule")
    st.caption("Every visible cell is editable. Type completely new values for rig, reservoir, area, target type, status or priority; edit dates and progress; add/remove rows; then click Apply table edits.")

    editor_columns = [
        "Target_ID", "Well_Name", "Reservoir", "Area", "Pad", "Rig", "Target_Type",
        "Start_Date", "End_Date", "Status", "Priority", "Progress_Pct", "Owner",
        "Campaign", "Color", "Notes",
    ]
    table_input = filtered_df[editor_columns].copy()

    edited = st.data_editor(
        table_input,
        key=f"schedule_editor_{st.session_state.filter_reset_token}",
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        height=min(720, 150 + max(len(table_input), 8) * 35),
        column_config={
            # All categorical columns use free-text editors so new values can be
            # entered directly without being restricted to an existing dropdown.
            "Target_ID": st.column_config.TextColumn(
                "Target ID",
                required=False,
                width="small",
                help="Fully editable. Blank or duplicate IDs are corrected automatically when Apply is pressed.",
            ),
            "Well_Name": st.column_config.TextColumn(
                "Well name",
                required=False,
                width="small",
            ),
            "Reservoir": st.column_config.TextColumn(
                "Reservoir",
                required=False,
                width="medium",
                help="Free text: existing or completely new reservoir names are accepted.",
            ),
            "Area": st.column_config.TextColumn(
                "Area",
                required=False,
                width="small",
            ),
            "Pad": st.column_config.TextColumn(
                "Pad / cluster",
                required=False,
                width="small",
            ),
            "Rig": st.column_config.TextColumn(
                "Rig",
                required=False,
                width="small",
                help="Free text: type a new rig name directly.",
            ),
            "Target_Type": st.column_config.TextColumn(
                "Target type",
                required=False,
                width="medium",
                help="Free text: Producer, Injector, TAR, Break, Maintenance, or any custom activity.",
            ),
            "Start_Date": st.column_config.DateColumn(
                "Start date",
                format="DD-MMM-YYYY",
                required=False,
            ),
            "End_Date": st.column_config.DateColumn(
                "End date",
                format="DD-MMM-YYYY",
                required=False,
            ),
            "Status": st.column_config.TextColumn(
                "Status",
                required=False,
                width="small",
            ),
            "Priority": st.column_config.TextColumn(
                "Priority",
                required=False,
                width="small",
            ),
            "Progress_Pct": st.column_config.NumberColumn(
                "Progress %",
                min_value=0,
                max_value=100,
                step=1,
                format="%d%%",
                required=False,
                width="small",
                help="Enter any whole-number progress value from 0 to 100.",
            ),
            "Owner": st.column_config.TextColumn(
                "Owner",
                required=False,
                width="medium",
            ),
            "Campaign": st.column_config.TextColumn(
                "Campaign",
                required=False,
                width="medium",
            ),
            "Color": st.column_config.TextColumn(
                "Hex color",
                required=False,
                help="Example: #2563EB",
                width="small",
            ),
            "Notes": st.column_config.TextColumn(
                "Notes",
                required=False,
                width="large",
            ),
        },
        disabled=False,
    )

    apply_col, download_col, local_col = st.columns([1, 1, 1])
    if apply_col.button("✅ Apply table edits", type="primary", use_container_width=True):
        try:
            updated = merge_filtered_edits(master_df, table_input, edited)
            set_master(updated, add_undo=True)
            st.session_state.source_name = "Edited in NWD Scheduler"
            st.success("Edits applied to the master schedule.")
            st.rerun()
        except Exception as exc:
            st.error(f"Could not apply edits: {exc}")

    current_export = make_csv_package(master_df, filtered_df, conflicts_df, issues_df)
    download_col.download_button(
        "⬇ Download CSV package",
        data=current_export,
        file_name=f"NWD_Scheduler_CSV_{datetime.now():%Y%m%d_%H%M}.zip",
        mime="application/zip",
        use_container_width=True,
        help="Downloads master schedule, filtered view, conflict report, quality report and summary as CSV files.",
    )

    if local_col.button("💾 Save master CSV", use_container_width=True, help=f"Saves to {LOCAL_SAVE_FILE}"):
        try:
            LOCAL_SAVE_FILE.write_bytes(dataframe_to_csv_bytes(master_df))
            st.success(f"Saved: {LOCAL_SAVE_FILE}")
        except Exception as exc:
            st.error(f"Local save failed: {exc}")

with tab_actions:
    left, right = st.columns(2)

    with left:
        st.subheader("Add a new target")
        with st.form("add_target_form", clear_on_submit=True):
            a1, a2 = st.columns(2)
            well_name = a1.text_input("Well name", placeholder="R-1501")
            target_id = a2.text_input("Target ID (optional)", placeholder="Auto-generated")
            reservoir = a1.selectbox("Reservoir", sorted(set(RESERVOIR_OPTIONS + select_options(master_df, "Reservoir"))))
            area = a2.selectbox("Area", sorted(set(AREA_OPTIONS + select_options(master_df, "Area"))))
            rig = a1.selectbox("Rig", sorted(set(select_options(master_df, "Rig") + ["Unassigned"])))
            pad = a2.text_input("Pad / cluster")
            target_type = a1.selectbox("Target type", TARGET_TYPE_OPTIONS)
            priority = a2.selectbox("Priority", PRIORITY_OPTIONS, index=2)
            start_date = a1.date_input("Start date", value=max(date.today(), min_date))
            planned_days = a2.number_input("Planned duration (days)", min_value=1, max_value=365, value=45, step=1)
            status = a1.selectbox("Status", STATUS_OPTIONS, index=0)
            progress = a2.number_input("Progress %", min_value=0, max_value=100, value=0, step=5)
            owner = a1.text_input("Owner")
            campaign = a2.text_input("Campaign", value=f"{start_date.year} Base")
            color = a1.color_picker("Bar color", value=DEFAULT_COLOR)
            notes = st.text_area("Notes")
            submitted = st.form_submit_button("Add target", type="primary", use_container_width=True)

        if submitted:
            new_id = target_id.strip() or next_target_id(master_df["Target_ID"].tolist())
            new_row = {
                "Target_ID": new_id,
                "Well_Name": well_name.strip() or f"New-Well-{len(master_df)+1}",
                "Reservoir": reservoir,
                "Area": area,
                "Pad": pad,
                "Rig": rig,
                "Target_Type": target_type,
                "Start_Date": pd.Timestamp(start_date),
                "End_Date": pd.Timestamp(start_date) + pd.Timedelta(days=int(planned_days) - 1),
                "Status": status,
                "Priority": priority,
                "Progress_Pct": int(progress),
                "Owner": owner,
                "Campaign": campaign,
                "Color": color,
                "Notes": notes,
            }
            set_master(pd.concat([master_df, pd.DataFrame([new_row])], ignore_index=True), add_undo=True)
            st.session_state.source_name = "Edited in NWD Scheduler"
            st.success(f"Added {new_id}")
            st.rerun()

    with right:
        st.subheader("Selected-target actions")
        selected_ids = st.multiselect("Choose target(s)", master_df["Target_ID"].tolist(), help="Actions below apply to these targets.")

        c1, c2 = st.columns(2)
        if c1.button("Duplicate first selected", use_container_width=True, disabled=not selected_ids):
            source = master_df[master_df["Target_ID"] == selected_ids[0]].iloc[0].copy()
            source["Target_ID"] = next_target_id(master_df["Target_ID"].tolist())
            source["Well_Name"] = f"{source['Well_Name']}-COPY"
            source["Start_Date"] = source["Start_Date"] + pd.Timedelta(days=7)
            source["End_Date"] = source["End_Date"] + pd.Timedelta(days=7)
            set_master(pd.concat([master_df, pd.DataFrame([source])], ignore_index=True), add_undo=True)
            st.rerun()

        if c2.button("🗑 Delete selected", use_container_width=True, disabled=not selected_ids):
            set_master(master_df[~master_df["Target_ID"].isin(selected_ids)], add_undo=True)
            st.success(f"Deleted {len(selected_ids)} target(s)")
            st.rerun()

        st.markdown("#### Shift dates")
        shift_days = st.number_input("Shift start and end dates by days", min_value=-730, max_value=730, value=0, step=1)
        if st.button("Apply date shift", use_container_width=True, disabled=not selected_ids or shift_days == 0):
            updated = master_df.copy()
            mask = updated["Target_ID"].isin(selected_ids)
            updated.loc[mask, "Start_Date"] = updated.loc[mask, "Start_Date"] + pd.Timedelta(days=int(shift_days))
            updated.loc[mask, "End_Date"] = updated.loc[mask, "End_Date"] + pd.Timedelta(days=int(shift_days))
            set_master(updated, add_undo=True)
            st.rerun()

        st.markdown("#### Set exact dates")
        exact_start = st.date_input("New start date", value=date.today(), key="bulk_exact_start")
        preserve_duration = st.checkbox("Preserve each target's duration", value=True)
        exact_end = st.date_input("New end date", value=date.today() + timedelta(days=44), disabled=preserve_duration, key="bulk_exact_end")
        if st.button("Apply exact dates", use_container_width=True, disabled=not selected_ids):
            updated = master_df.copy()
            for idx in updated.index[updated["Target_ID"].isin(selected_ids)]:
                duration = max(int(updated.at[idx, "Duration_Days"] if pd.notna(updated.at[idx, "Duration_Days"]) else 1), 1)
                updated.at[idx, "Start_Date"] = pd.Timestamp(exact_start)
                updated.at[idx, "End_Date"] = pd.Timestamp(exact_start) + pd.Timedelta(days=duration - 1) if preserve_duration else pd.Timestamp(exact_end)
            set_master(updated, add_undo=True)
            st.rerun()

        st.markdown("#### Update attributes")
        bulk_status = st.selectbox("New status", ["— Keep current —"] + STATUS_OPTIONS)
        bulk_priority = st.selectbox("New priority", ["— Keep current —"] + PRIORITY_OPTIONS)
        bulk_rig = st.selectbox("New rig", ["— Keep current —"] + sorted(set(select_options(master_df, "Rig") + ["Unassigned"])))
        bulk_color = st.color_picker("New custom color", value=DEFAULT_COLOR)
        change_color = st.checkbox("Apply the selected color", value=False)
        if st.button("Apply attribute changes", use_container_width=True, disabled=not selected_ids):
            updated = master_df.copy()
            mask = updated["Target_ID"].isin(selected_ids)
            if bulk_status != "— Keep current —":
                updated.loc[mask, "Status"] = bulk_status
                if bulk_status == "Completed":
                    updated.loc[mask, "Progress_Pct"] = 100
            if bulk_priority != "— Keep current —":
                updated.loc[mask, "Priority"] = bulk_priority
            if bulk_rig != "— Keep current —":
                updated.loc[mask, "Rig"] = bulk_rig
            if change_color:
                updated.loc[mask, "Color"] = bulk_color
            set_master(updated, add_undo=True)
            st.rerun()

with tab_quality:
    q1, q2 = st.columns(2)
    with q1:
        st.subheader("Rig overlap conflicts")
        if conflicts_df.empty:
            st.success("No rig overlaps in the filtered schedule.")
        else:
            st.error(f"{len(conflicts_df)} overlapping rig assignment(s) found.")
            st.dataframe(conflicts_df, use_container_width=True, hide_index=True)
    with q2:
        st.subheader("Master-data validation")
        if issues_df.empty:
            st.success("No data-quality issues found.")
        else:
            error_count = int((issues_df["Severity"] == "Error").sum())
            warning_count = int((issues_df["Severity"] == "Warning").sum())
            st.warning(f"{error_count} error(s), {warning_count} warning(s)")
            st.dataframe(issues_df, use_container_width=True, hide_index=True)

with tab_summary:
    s1, s2 = st.columns(2)
    with s1:
        status_counts = filtered_df.groupby("Status").size().reset_index(name="Targets")
        if not status_counts.empty:
            status_fig = px.bar(status_counts, x="Status", y="Targets", title="Targets by status", text_auto=True, color="Status", color_discrete_map=STATUS_COLORS)
            status_fig.update_layout(showlegend=False, margin={"l": 10, "r": 10, "t": 50, "b": 10})
            st.plotly_chart(status_fig, use_container_width=True, config={"displaylogo": False})
    with s2:
        reservoir_counts = filtered_df.groupby("Reservoir").size().reset_index(name="Targets").sort_values("Targets", ascending=True)
        if not reservoir_counts.empty:
            res_fig = px.bar(reservoir_counts, x="Targets", y="Reservoir", orientation="h", title="Targets by reservoir", text_auto=True)
            res_fig.update_layout(showlegend=False, margin={"l": 10, "r": 10, "t": 50, "b": 10})
            st.plotly_chart(res_fig, use_container_width=True, config={"displaylogo": False})

    st.subheader("Rig utilization in selected window")
    if rig_util_df.empty:
        st.info("No rig utilization data for the selected window.")
    else:
        utilization_fig = px.bar(rig_util_df, x="Rig", y="Utilization_Pct", text="Utilization_Pct", title="Calendar utilization by rig")
        utilization_fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        utilization_fig.update_yaxes(title="Utilization %", range=[0, max(105, rig_util_df["Utilization_Pct"].max() * 1.15)])
        utilization_fig.update_layout(margin={"l": 10, "r": 10, "t": 50, "b": 10})
        st.plotly_chart(utilization_fig, use_container_width=True, config={"displaylogo": False})
        st.dataframe(rig_util_df, use_container_width=True, hide_index=True)

    st.subheader("Filtered schedule data")
    st.dataframe(filtered_df, use_container_width=True, hide_index=True)
    st.download_button(
        "Download filtered CSV",
        data=filtered_df.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"NWD_Filtered_{datetime.now():%Y%m%d_%H%M}.csv",
        mime="text/csv",
    )

with tab_help:
    st.markdown(
        """
        ### Recommended workflow
        1. Upload your NWD Excel/CSV file, or start with the supplied dummy workbook.
        2. Use sidebar filters to isolate a year, reservoir, rig, status, owner or campaign.
        3. Review the **Gantt** and hover over any bar for complete target details.
        4. Open **Edit Schedule** to change dates, rigs, status, progress, colors and notes. Add/remove rows directly in the table and click **Apply table edits**.
        5. Use **Bulk Actions** for mass date shifts, rig reassignment, status updates, duplication or deletion.
        6. Check **Quality & Conflicts** before publishing the schedule.
        7. Download the CSV package. It contains the master schedule, filtered view, conflicts, data-quality issues and summary as separate CSV files.

        ### Minimum useful input columns
        `Target_ID`, `Well_Name`, `Rig`, `Start_Date`, `End_Date`.
        Missing optional columns are automatically created. Common headings such as **Well**, **Start Date**, **Spud Date**, **End Date**, **Progress %**, **Comments** and **Colour** are recognized.

        ### Gantt chart views
        - **Detailed target rows:** preserves the original one-target-per-row chart.
        - **Compact merged lanes:** merges visible targets onto one lane for each selected Rig, Pad, Reservoir, Area, Campaign or other Y-axis column.
        - **Grouped target rows:** keeps individual target rows while showing the selected Y-axis grouping value on the left.
        - **Text inside bars:** can show the Well Name plus progress, Well Name only, percentage only, or no text. Labels are drawn above progress shading so they remain visible.
        - **Editable schedule:** all visible cells accept direct edits, including completely new Rig, Reservoir, Area, Target Type, Status and Priority values.
        - **Horizontal grid lines:** can be switched on or off from the sidebar.

        ### Color handling
        Keep **Color bars by = Custom row color** to use each row's hex color. Edit the `Color` column or use **Bulk Actions → Apply the selected color**. You can also color dynamically by status, priority, rig, reservoir, area, target type or campaign.

        ### Important behavior
        Filters use date overlap logic: a target is visible when any portion of its duration overlaps the selected schedule window. Edits made in a filtered table are merged back into the complete master schedule; rows outside the filter remain untouched.
        """
    )

st.divider()
st.caption("NWD Scheduler • Session-based editing with undo • Download the CSV package regularly to preserve approved schedule versions")
