# -*- coding: utf-8 -*-
"""
NWD Scheduler v1.5.1 - Streamlit application

Run:
    streamlit run nwd_scheduler_v1.5.1.py

The application opens with an empty schedule. Upload an Excel/CSV schedule
when required. Place SoR PDF documents in an "input" folder beside this Python
file, or upload PDFs temporarily from the sidebar.

Outputs are CSV-based.
"""

from __future__ import annotations

import base64
import html
import io
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
import streamlit.components.v1 as components


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
DEFAULT_SOR_DIR = APP_DIR / "input"
COMPUTED_COLUMNS = ["Duration_Days", "Start_Year", "End_Year"]
INTERNAL_ROW_KEY = "__Original_Target_ID"
HIGHLIGHT_COLOR = "#7C3AED"

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
    "Highlight",
    "Highlight_Label",
    "SoR_File",
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
    "Highlight_Label",
    "SoR_File",
    "Notes",
]

DATE_COLUMNS = ["Start_Date", "End_Date"]
BOOLEAN_COLUMNS = ["Highlight"]

STATUS_OPTIONS = ["Not Started", "Ready", "In Progress", "On Hold", "Completed", "Cancelled"]
PRIORITY_OPTIONS = ["Critical", "High", "Medium", "Low"]
TARGET_TYPE_OPTIONS = [
    "Producer", "Water Injector", "Observation", "Disposal", "Appraisal",
    "TAR", "Drilling Break", "Rig Maintenance", "Rig Move", "New Technology",
]
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
    "highlight": "Highlight",
    "highlighted": "Highlight",
    "new technology": "Highlight",
    "new_technology": "Highlight",
    "highlight label": "Highlight_Label",
    "highlight_label": "Highlight_Label",
    "sor": "SoR_File",
    "sor file": "SoR_File",
    "sor_file": "SoR_File",
    "sor document": "SoR_File",
    "sor_document": "SoR_File",
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
def create_empty_dataframe() -> pd.DataFrame:
    """Return an empty, correctly structured NWD schedule."""
    return normalize_dataframe(pd.DataFrame(columns=REQUIRED_COLUMNS))


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
            elif col == "Highlight":
                df[col] = False
            else:
                df[col] = ""

    # Keep supported columns first, then any extra user columns.
    extra_cols = [c for c in df.columns if c not in REQUIRED_COLUMNS and c not in set(COMPUTED_COLUMNS)]
    df = df[REQUIRED_COLUMNS + extra_cols]

    for col in DATE_COLUMNS:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.normalize()

    for col in TEXT_COLUMNS:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    df["Progress_Pct"] = (
        pd.to_numeric(df["Progress_Pct"], errors="coerce")
        .fillna(0)
        .clip(0, 100)
        .round(0)
        .astype(int)
    )
    df["Color"] = df["Color"].apply(valid_hex_color)

    truthy_values = {"true", "1", "yes", "y", "on", "highlight", "highlighted"}
    df["Highlight"] = (
        df["Highlight"]
        .fillna(False)
        .apply(
            lambda value: value
            if isinstance(value, (bool, np.bool_))
            else str(value).strip().lower() in truthy_values
        )
        .astype(bool)
    )

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
    visible_columns: Sequence[str],
) -> pd.DataFrame:
    """
    Merge visible editor changes into the full master schedule.

    Hidden columns are preserved from the original rows. Deleted visible rows are
    removed, newly added rows are included, and rows outside active filters remain
    untouched.
    """
    master = normalize_dataframe(master_df)
    original_visible = normalize_dataframe(original_filtered_df)

    if edited_df is None:
        edited = pd.DataFrame(columns=[INTERNAL_ROW_KEY] + list(visible_columns))
    else:
        edited = edited_df.copy()

    original_ids = set(
        original_visible["Target_ID"].fillna("").astype(str).str.strip().tolist()
    )

    untouched_rows = master[
        ~master["Target_ID"].fillna("").astype(str).str.strip().isin(original_ids)
    ].copy()

    original_lookup = {
        str(row["Target_ID"]).strip(): row.to_dict()
        for _, row in original_visible.iterrows()
    }

    rebuilt_rows: List[Dict[str, object]] = []
    for _, edited_row in edited.iterrows():
        row_key = str(edited_row.get(INTERNAL_ROW_KEY, "") or "").strip()
        if not row_key:
            index_key = str(edited_row.name or "").strip()
            if index_key in original_lookup:
                row_key = index_key
        base_row = dict(original_lookup.get(row_key, {}))

        for column in visible_columns:
            if column in edited_row.index:
                base_row[column] = edited_row[column]

        rebuilt_rows.append(base_row)

    edited_rows = pd.DataFrame(rebuilt_rows)
    merged = pd.concat([untouched_rows, edited_rows], ignore_index=True, sort=False)
    return normalize_dataframe(merged)


def initialize_state() -> None:
    if "schedule_df" not in st.session_state:
        st.session_state.schedule_df = create_empty_dataframe()
        st.session_state.source_name = "Empty schedule"
        st.session_state.source_file_name = ""
    st.session_state.setdefault("undo_stack", [])
    st.session_state.setdefault("filter_reset_token", 0)
    st.session_state.setdefault("uploaded_sor_files", {})
    st.session_state.setdefault("selected_sor_target_id", "")
    st.session_state.setdefault("gantt_selection_token", 0)


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


def editable_schedule_columns(df: pd.DataFrame) -> List[str]:
    """Return all editable input columns, including user-supplied extra columns."""
    return [column for column in df.columns if column not in set(COMPUTED_COLUMNS)]


def column_display_name(column: str) -> str:
    labels = {
        "Target_ID": "Target ID",
        "Well_Name": "Well name",
        "Target_Type": "Target type",
        "Start_Date": "Start date",
        "End_Date": "End date",
        "Progress_Pct": "Progress %",
        "Highlight_Label": "Highlight label",
        "SoR_File": "SoR file",
    }
    return labels.get(column, column.replace("_", " "))


def format_hover_value(value: object, column: str) -> str:
    """Format one tooltip value without exposing columns hidden by the user."""
    if pd.isna(value):
        return ""
    if column in DATE_COLUMNS:
        parsed = pd.to_datetime(value, errors="coerce")
        return "" if pd.isna(parsed) else parsed.strftime("%d-%b-%Y")
    if column == "Progress_Pct":
        try:
            return f"{int(float(value))}%"
        except (TypeError, ValueError):
            return str(value)
    if column == "Highlight":
        return "Yes" if bool(value) else "No"
    return html.escape(str(value))


def build_editor_column_config(
    visible_columns: Sequence[str],
) -> Dict[str, object]:
    """
    Build column configuration only for columns actually sent to st.data_editor.

    Avoid configuring hidden or absent columns because Streamlit validates each
    configured type against the dataframe schema.
    """
    known_config: Dict[str, object] = {
        "Target_ID": st.column_config.TextColumn(
            "Target ID",
            required=False,
            width="small",
            help="Blank or duplicate IDs are corrected automatically after Apply.",
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
        ),
        "Target_Type": st.column_config.TextColumn(
            "Target type",
            required=False,
            width="medium",
            help=(
                "Producer, Injector, TAR, Break, Maintenance, "
                "New Technology, or any custom item."
            ),
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
            width="small",
            help="Example: #2563EB",
        ),
        "Highlight": st.column_config.CheckboxColumn(
            "Highlight",
            default=False,
            help="Adds a purple dashed border around the Gantt bar.",
        ),
        "Highlight_Label": st.column_config.TextColumn(
            "Highlight label",
            required=False,
            width="medium",
            help="Example: New Technology",
        ),
        "SoR_File": st.column_config.TextColumn(
            "SoR file",
            required=False,
            width="large",
            help="PDF filename stored in the configured input folder.",
        ),
        "Notes": st.column_config.TextColumn(
            "Notes",
            required=False,
            width="large",
        ),
    }

    config = {
        column: known_config[column]
        for column in visible_columns
        if column in known_config
    }
    config[INTERNAL_ROW_KEY] = None
    return config


def prepare_editor_dataframe(
    filtered_df: pd.DataFrame,
    visible_columns: Sequence[str],
) -> pd.DataFrame:
    """
    Create a Streamlit-safe editor dataframe.

    Only visible columns are passed to the widget. Hidden columns are preserved
    separately in the master dataframe and merged back after Apply.
    """
    editor_df = filtered_df[list(visible_columns)].copy()

    # Explicit dtypes prevent Streamlit from inferring incompatible schemas,
    # especially when the uploaded file contains mixed or blank values.
    for column in editor_df.columns:
        if column in DATE_COLUMNS:
            editor_df[column] = pd.to_datetime(
                editor_df[column],
                errors="coerce",
            )
        elif column == "Progress_Pct":
            editor_df[column] = (
                pd.to_numeric(editor_df[column], errors="coerce")
                .fillna(0)
                .clip(0, 100)
                .round(0)
                .astype("int64")
            )
        elif column == "Highlight":
            editor_df[column] = (
                editor_df[column]
                .fillna(False)
                .astype(bool)
            )
        elif column in TEXT_COLUMNS:
            editor_df[column] = (
                editor_df[column]
                .fillna("")
                .astype(str)
            )

    editor_df[INTERNAL_ROW_KEY] = (
        filtered_df["Target_ID"]
        .fillna("")
        .astype(str)
        .to_numpy()
    )

    # A RangeIndex avoids the empty/custom-index issue in dynamic data editors.
    editor_df.index = pd.RangeIndex(start=0, stop=len(editor_df), step=1)
    return editor_df


def safe_filename_stem(name: str) -> str:
    raw_stem = Path(str(name or "")).stem.strip()
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", raw_stem).strip("._-")
    return stem or "NWD_Schedule"


def updated_input_filename(source_name: str) -> str:
    return f"{safe_filename_stem(source_name)}_updated.csv"


def resolve_sor_directory(folder_text: str) -> Path:
    folder_text = str(folder_text or "").strip()
    if not folder_text:
        return DEFAULT_SOR_DIR
    candidate = Path(folder_text).expanduser()
    return candidate if candidate.is_absolute() else APP_DIR / candidate


def normalized_document_key(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def local_pdf_files(folder: Path) -> List[Path]:
    if not folder.exists() or not folder.is_dir():
        return []
    try:
        return sorted(
            [
                path
                for path in folder.rglob("*")
                if path.is_file() and path.suffix.lower() == ".pdf"
            ],
            key=lambda item: item.name.lower(),
        )
    except OSError:
        return []


def resolve_sor_document(
    target_row: pd.Series,
    sor_directory: Path,
    uploaded_files: Dict[str, bytes],
) -> Optional[Dict[str, object]]:
    """
    Resolve an SoR PDF using:
    1. The row's SoR_File value.
    2. Exact Target_ID or Well_Name filename matching.
    3. Partial Target_ID or Well_Name filename matching.
    """
    uploaded_lookup = {
        name.lower(): {"name": name, "bytes": content, "source": "Uploaded PDF"}
        for name, content in uploaded_files.items()
        if str(name).lower().endswith(".pdf")
    }

    local_files = local_pdf_files(sor_directory)
    local_lookup = {path.name.lower(): path for path in local_files}

    requested = str(target_row.get("SoR_File", "") or "").strip()
    if requested:
        requested_name = Path(requested).name.lower()
        if requested_name in uploaded_lookup:
            return uploaded_lookup[requested_name]

        requested_path = Path(requested).expanduser()
        possible_paths = []
        if requested_path.is_absolute():
            possible_paths.append(requested_path)
        else:
            possible_paths.extend(
                [
                    sor_directory / requested_path,
                    APP_DIR / requested_path,
                ]
            )

        for path in possible_paths:
            if path.exists() and path.is_file() and path.suffix.lower() == ".pdf":
                try:
                    pdf_bytes = path.read_bytes()
                except OSError:
                    continue
                return {
                    "name": path.name,
                    "bytes": pdf_bytes,
                    "source": str(path),
                }

        if requested_name in local_lookup:
            path = local_lookup[requested_name]
            try:
                pdf_bytes = path.read_bytes()
            except OSError:
                pdf_bytes = None
            if pdf_bytes is not None:
                return {
                    "name": path.name,
                    "bytes": pdf_bytes,
                    "source": str(path),
                }

    target_keys = [
        normalized_document_key(target_row.get("Target_ID", "")),
        normalized_document_key(target_row.get("Well_Name", "")),
    ]
    target_keys = [key for key in target_keys if key]

    candidates: List[Dict[str, object]] = list(uploaded_lookup.values())
    for path in local_files:
        candidates.append(
            {
                "name": path.name,
                "path": path,
                "source": str(path),
            }
        )

    def candidate_key(candidate: Dict[str, object]) -> str:
        return normalized_document_key(Path(str(candidate["name"])).stem)

    exact_matches = [
        candidate
        for candidate in candidates
        if candidate_key(candidate) in target_keys
    ]
    partial_matches = [
        candidate
        for candidate in candidates
        if any(key in candidate_key(candidate) for key in target_keys)
    ]

    matches = exact_matches or partial_matches
    if not matches:
        return None

    selected = sorted(matches, key=lambda item: len(str(item["name"])))[0]
    if "bytes" not in selected:
        selected = dict(selected)
        try:
            selected["bytes"] = Path(selected["path"]).read_bytes()
        except OSError:
            return None
    return selected


def extract_selected_target_id(event: object) -> str:
    try:
        points = event.selection.points
    except (AttributeError, TypeError):
        try:
            points = event.get("selection", {}).get("points", [])
        except AttributeError:
            points = []

    if not points:
        return ""

    customdata = points[0].get("customdata", [])
    if isinstance(customdata, (list, tuple)) and customdata:
        return str(customdata[0] or "").strip()
    return ""


def show_pdf_viewer(document: Dict[str, object], target_id: str) -> None:
    pdf_bytes = bytes(document["bytes"])
    file_name = str(document["name"])

    info_col, download_col = st.columns([3, 1])
    info_col.success(f"SoR found for {target_id}: {file_name}")
    download_col.download_button(
        "Download SoR PDF",
        data=pdf_bytes,
        file_name=file_name,
        mime="application/pdf",
        use_container_width=True,
        key=f"download_sor_{target_id}_{normalized_document_key(file_name)}",
    )

    if hasattr(st, "pdf"):
        try:
            st.pdf(
                pdf_bytes,
                height=850,
                key=f"sor_pdf_{target_id}_{normalized_document_key(file_name)}",
            )
            return
        except Exception:
            pass

    encoded = base64.b64encode(pdf_bytes).decode("ascii")
    components.html(
        (
            '<iframe '
            f'src="data:application/pdf;base64,{encoded}" '
            'width="100%" height="850" '
            'style="border:1px solid rgba(128,128,128,.3);border-radius:8px;">'
            "</iframe>"
        ),
        height=870,
        scrolling=True,
    )


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
    tooltip_columns: Sequence[str],
) -> go.Figure:
    """Build the interactive and selectable NWD Gantt chart."""
    if df.empty:
        fig = go.Figure()
        fig.add_annotation(
            text="The schedule is empty. Upload a file or add a target.",
            x=0.5,
            y=0.5,
            showarrow=False,
            font={"size": 18},
        )
        fig.update_layout(height=420, xaxis={"visible": False}, yaxis={"visible": False})
        return fig

    plot_df = df.dropna(subset=["Start_Date", "End_Date"]).copy()
    if plot_df.empty:
        fig = go.Figure()
        fig.add_annotation(
            text="Visible items do not have valid start and end dates.",
            x=0.5,
            y=0.5,
            showarrow=False,
            font={"size": 18},
        )
        fig.update_layout(height=420, xaxis={"visible": False}, yaxis={"visible": False})
        return fig

    if y_axis_column not in plot_df.columns:
        y_axis_column = "Rig"
    if group_by not in plot_df.columns:
        group_by = "Rig"

    for column in [y_axis_column, group_by]:
        plot_df[column] = (
            plot_df[column]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace("", "Unassigned")
        )

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

    def unique_labels(values: Sequence[str]) -> List[str]:
        seen: Dict[str, int] = {}
        result: List[str] = []
        for raw_value in values:
            value = str(raw_value)
            seen[value] = seen.get(value, 0) + 1
            result.append(value if seen[value] == 1 else f"{value} ({seen[value]})")
        return result

    if gantt_type == "Compact merged lanes":
        plot_df["Y_Label"] = plot_df[y_axis_column].astype(str)
        separator_field = y_axis_column
        type_description = f"Compact lanes by {column_display_name(y_axis_column)}"
    elif gantt_type == "Grouped target rows":
        plot_df["Y_Label"] = unique_labels(
            (
                plot_df[y_axis_column].astype(str)
                + "  |  "
                + plot_df["Task_Label"].astype(str)
            ).tolist()
        )
        separator_field = y_axis_column
        type_description = f"Grouped rows by {column_display_name(y_axis_column)}"
    else:
        plot_df["Y_Label"] = unique_labels(plot_df["Task_Label"].tolist())
        separator_field = group_by
        type_description = "Detailed target rows"

    y_categories = list(dict.fromkeys(plot_df["Y_Label"].astype(str).tolist()))
    y_position = {category: index for index, category in enumerate(y_categories)}
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
        legend_title = column_display_name(field)

    duration_ms = (
        (plot_df["End_Date"] - plot_df["Start_Date"]).dt.total_seconds() * 1000
    ).clip(lower=86_400_000)

    progress_text = plot_df["Progress_Pct"].astype(int).astype(str) + "%"
    target_text = plot_df["Well_Name"].fillna("").astype(str).str.strip()
    target_text = target_text.where(
        target_text.ne(""),
        plot_df["Target_ID"].fillna("").astype(str),
    )

    if bar_text_mode == "Target name + progress %":
        bar_text = target_text + "<br>" + progress_text
    elif bar_text_mode == "Target name only":
        bar_text = target_text
    elif bar_text_mode == "Progress % only":
        bar_text = progress_text
    else:
        bar_text = pd.Series([""] * len(plot_df), index=plot_df.index)

    text_midpoints = (
        plot_df["Start_Date"]
        + (plot_df["End_Date"] - plot_df["Start_Date"]) / 2
    )

    tooltip_columns = [
        column
        for column in tooltip_columns
        if column in plot_df.columns
    ]

    customdata: List[List[object]] = []
    for _, row in plot_df.iterrows():
        values: List[object] = [
            str(row.get("Target_ID", "")),
            str(row.get("Well_Name", "")),
            str(row.get("SoR_File", "")),
        ]
        values.extend(
            format_hover_value(row.get(column, ""), column)
            for column in tooltip_columns
        )
        customdata.append(values)

    hover_lines = []
    for data_index, column in enumerate(tooltip_columns, start=3):
        hover_lines.append(
            f"<b>{html.escape(column_display_name(column))}:</b> "
            f"%{{customdata[{data_index}]}}<br>"
        )
    hover_lines.append("<i>Click to open the SoR PDF</i>")
    hovertemplate = "".join(hover_lines) + "<extra></extra>"

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=duration_ms,
            base=plot_df["Start_Date"],
            y=plot_df["Y_Label"],
            orientation="h",
            marker={
                "color": colors,
                "line": {"width": 0.8, "color": "rgba(30,41,59,.52)"},
            },
            text=None,
            customdata=customdata,
            hovertemplate=hovertemplate,
            showlegend=False,
            name="Schedule items",
        )
    )

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
                marker={"color": "rgba(15,23,42,.32)", "line": {"width": 0}},
                customdata=customdata,
                hovertemplate=hovertemplate,
                showlegend=False,
                name="Progress",
            )
        )

    if bar_text_mode != "No text":
        fig.add_trace(
            go.Scatter(
                x=text_midpoints,
                y=plot_df["Y_Label"],
                mode="text",
                text=bar_text,
                textposition="middle center",
                textfont={"color": "white", "size": 11, "family": "Arial, sans-serif"},
                customdata=customdata,
                hovertemplate=hovertemplate,
                showlegend=False,
                cliponaxis=True,
                name="Bar labels",
            )
        )

    # Purple dashed border for highlighted items, such as New Technology.
    highlighted = plot_df[plot_df["Highlight"].fillna(False).astype(bool)]
    half_height = 0.42 if gantt_type == "Compact merged lanes" else 0.37
    for _, row in highlighted.iterrows():
        category = str(row["Y_Label"])
        category_position = y_position.get(category)
        if category_position is None:
            continue
        fig.add_shape(
            type="rect",
            xref="x",
            yref="y",
            x0=row["Start_Date"],
            x1=row["End_Date"],
            y0=category_position - half_height,
            y1=category_position + half_height,
            line={
                "color": HIGHLIGHT_COLOR,
                "width": 3,
                "dash": "dash",
            },
            fillcolor="rgba(0,0,0,0)",
            layer="above",
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

    if show_horizontal_grid:
        for boundary_index in range(row_count + 1):
            fig.add_hline(
                y=boundary_index - 0.5,
                line_width=0.8,
                line_dash="solid",
                line_color="rgba(148,163,184,.28)",
                layer="below",
            )

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

    fig.update_layout(
        title={
            "text": (
                f"NWD Drilling Schedule — {len(plot_df)} "
                f"item{'s' if len(plot_df) != 1 else ''}"
            ),
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
        clickmode="event+select",
        selectionrevision="nwd-v1.5.1",
        margin={"l": 20, "r": 25, "t": 88, "b": 25},
        uniformtext={"mode": "hide", "minsize": 8},
        xaxis={
            "title": "Calendar date",
            "type": "date",
            "showgrid": True,
            "gridcolor": "rgba(148,163,184,.20)",
            "gridwidth": 1,
            "rangeslider": {"visible": True, "thickness": 0.06},
            "rangeselector": {
                "buttons": [
                    {"count": 3, "label": "3m", "step": "month", "stepmode": "backward"},
                    {"count": 6, "label": "6m", "step": "month", "stepmode": "backward"},
                    {"count": 1, "label": "1y", "step": "year", "stepmode": "backward"},
                    {"step": "all", "label": "All"},
                ]
            },
        },
        yaxis={
            "title": column_display_name(y_axis_column)
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

    if not highlighted.empty:
        fig.add_trace(
            go.Scatter(
                x=[None],
                y=[None],
                mode="lines",
                line={"color": HIGHLIGHT_COLOR, "width": 3, "dash": "dash"},
                name="Highlighted item",
                hoverinfo="skip",
                showlegend=True,
            )
        )

    return fig


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    """Return a normalized dataframe as UTF-8-SIG CSV bytes."""
    export_df = normalize_dataframe(df).copy()
    for column in DATE_COLUMNS:
        export_df[column] = (
            pd.to_datetime(export_df[column], errors="coerce")
            .dt.strftime("%Y-%m-%d")
            .fillna("")
        )
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
all_editable_columns = editable_schedule_columns(master_df)

if "visible_columns_widget" not in st.session_state:
    st.session_state.visible_columns_widget = list(all_editable_columns)
else:
    st.session_state.visible_columns_widget = [
        column
        for column in st.session_state.visible_columns_widget
        if column in all_editable_columns
    ]

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

    uploaded = st.file_uploader(
        "Upload schedule Excel or CSV",
        type=["xlsx", "xls", "xlsm", "csv"],
        help=(
            "For Excel, the app looks for NWD_Schedule, Schedule or Targets; "
            "otherwise it reads the first sheet. CSV is read directly."
        ),
    )

    col_load, col_empty = st.columns(2)
    if col_load.button(
        "Load file",
        use_container_width=True,
        disabled=uploaded is None,
    ):
        try:
            loaded = load_schedule_file(uploaded, uploaded.name)
            set_master(loaded, add_undo=True)
            st.session_state.source_name = uploaded.name
            st.session_state.source_file_name = uploaded.name
            st.session_state.filter_reset_token += 1
            st.session_state.pop("visible_columns_widget", None)
            st.success(f"Loaded {len(loaded)} schedule items")
            st.rerun()
        except Exception as exc:
            st.error(f"Could not load the schedule file: {exc}")

    if col_empty.button("New empty", use_container_width=True):
        set_master(create_empty_dataframe(), add_undo=True)
        st.session_state.source_name = "Empty schedule"
        st.session_state.source_file_name = ""
        st.session_state.selected_sor_target_id = ""
        st.session_state.filter_reset_token += 1
        st.session_state.pop("visible_columns_widget", None)
        st.rerun()

    if st.button(
        "↩ Undo last change",
        use_container_width=True,
        disabled=not st.session_state.undo_stack,
    ):
        previous = st.session_state.undo_stack.pop()
        st.session_state.schedule_df = normalize_dataframe(previous)
        st.toast("Last change undone")
        st.rerun()

    st.download_button(
        "⬇ Download updated input CSV",
        data=dataframe_to_csv_bytes(master_df),
        file_name=updated_input_filename(
            st.session_state.get("source_file_name", "")
        ),
        mime="text/csv",
        use_container_width=True,
        help="Downloads the complete current master schedule after all applied edits.",
    )

    with st.expander("SoR PDF settings", expanded=False):
        sor_folder_text = st.text_input(
            "SoR PDF folder",
            value="input",
            help=(
                "Relative paths are resolved beside the Python file. "
                "Example: input. An absolute local path also works when running locally."
            ),
        )

        uploaded_sor_documents = st.file_uploader(
            "Upload SoR PDFs for this session",
            type=["pdf"],
            accept_multiple_files=True,
            help=(
                "Useful on Streamlit Cloud. Uploaded PDFs remain available only "
                "for the current app session."
            ),
        )
        for pdf_file in uploaded_sor_documents:
            st.session_state.uploaded_sor_files[pdf_file.name] = pdf_file.getvalue()

        st.caption(
            f"Session PDFs: {len(st.session_state.uploaded_sor_files)} • "
            f"Folder: {resolve_sor_directory(sor_folder_text)}"
        )

        if st.button(
            "Clear uploaded SoR PDFs",
            use_container_width=True,
            disabled=not st.session_state.uploaded_sor_files,
        ):
            st.session_state.uploaded_sor_files = {}
            st.rerun()

    st.divider()
    st.subheader("Table & tooltip columns")
    visible_editor_columns = st.multiselect(
        "Visible columns",
        options=all_editable_columns,
        key="visible_columns_widget",
        help=(
            "A hidden table column is also removed from the Gantt mouse-over tooltip. "
            "Hidden values remain preserved in the master schedule."
        ),
    )
    if not visible_editor_columns:
        st.warning("Select at least one column to edit the schedule table.")

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
        visible_editor_columns,
    )

    gantt_event = st.plotly_chart(
        gantt,
        use_container_width=True,
        key=f"nwd_gantt_{st.session_state.gantt_selection_token}",
        on_select="rerun",
        selection_mode="points",
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

    clicked_target_id = extract_selected_target_id(gantt_event)
    if clicked_target_id:
        st.session_state.selected_sor_target_id = clicked_target_id

    st.caption(
        "Click a target bar, its progress section, or its text label to open the "
        "matching SoR PDF. Hidden table columns are also hidden from mouse-over text."
    )

    manual_col, clear_col = st.columns([3, 1])
    manual_target_id = manual_col.selectbox(
        "SoR fallback selection",
        options=[""] + master_df["Target_ID"].astype(str).tolist(),
        format_func=lambda value: "Select a target manually..." if value == "" else value,
        key="manual_sor_target",
    )
    if manual_target_id:
        st.session_state.selected_sor_target_id = manual_target_id

    if clear_col.button(
        "Close SoR",
        use_container_width=True,
        disabled=not st.session_state.selected_sor_target_id,
    ):
        st.session_state.selected_sor_target_id = ""
        st.session_state.gantt_selection_token += 1
        st.rerun()

    selected_target_id = st.session_state.get("selected_sor_target_id", "")
    if selected_target_id:
        selected_rows = master_df[
            master_df["Target_ID"].astype(str) == str(selected_target_id)
        ]

        if selected_rows.empty:
            st.warning(f"Selected target {selected_target_id} is not in the current schedule.")
        else:
            selected_row = selected_rows.iloc[0]
            document = resolve_sor_document(
                selected_row,
                resolve_sor_directory(sor_folder_text),
                st.session_state.uploaded_sor_files,
            )

            st.markdown("---")
            st.subheader(
                f"SoR — {selected_row['Target_ID']} | {selected_row['Well_Name']}"
            )
            if document is None:
                expected_name = str(selected_row.get("SoR_File", "") or "").strip()
                detail = (
                    f" Expected file: `{expected_name}`."
                    if expected_name
                    else ""
                )
                st.warning(
                    f"SoR not found for {selected_target_id}.{detail} "
                    "Add the PDF to the configured input folder, upload it from the "
                    "sidebar, or populate the SoR_File column."
                )
            else:
                show_pdf_viewer(document, selected_target_id)

with tab_edit:
    st.subheader("Editable filtered schedule")
    st.caption(
        "All visible cells are editable. Hidden columns remain preserved and "
        "are also excluded from Gantt tooltips."
    )

    if master_df.empty:
        st.info(
            "The schedule is currently empty. Upload an Excel/CSV file from the "
            "sidebar, add a target from Bulk Actions, or create the first blank row."
        )

        first_row_col, template_col = st.columns([1, 1])

        if first_row_col.button(
            "➕ Create first blank row",
            type="primary",
            use_container_width=True,
        ):
            first_row = {
                "Target_ID": "",
                "Well_Name": "",
                "Reservoir": "",
                "Area": "",
                "Pad": "",
                "Rig": "",
                "Target_Type": "",
                "Start_Date": pd.NaT,
                "End_Date": pd.NaT,
                "Status": "",
                "Priority": "",
                "Progress_Pct": 0,
                "Owner": "",
                "Campaign": "",
                "Color": DEFAULT_COLOR,
                "Highlight": False,
                "Highlight_Label": "",
                "SoR_File": "",
                "Notes": "",
            }
            set_master(pd.DataFrame([first_row]), add_undo=True)
            st.session_state.source_name = "New schedule"
            st.rerun()

        template_col.download_button(
            "⬇ Download empty template",
            data=dataframe_to_csv_bytes(create_empty_dataframe()),
            file_name="NWD_Schedule_Empty_Template.csv",
            mime="text/csv",
            use_container_width=True,
        )

        edited = pd.DataFrame()
    elif not visible_editor_columns:
        st.info(
            "No columns are currently visible. Select at least one column under "
            "Table & tooltip columns in the sidebar."
        )
        edited = pd.DataFrame()
    else:
        table_input = prepare_editor_dataframe(
            filtered_df,
            visible_editor_columns,
        )

        if table_input.empty:
            st.info(
                "The master schedule contains data, but no rows match the current "
                "filters. Clear or change the filters to edit rows."
            )
            edited = table_input
        else:
            edited = st.data_editor(
                table_input,
                key=f"schedule_editor_{st.session_state.filter_reset_token}",
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                height=min(
                    720,
                    150 + max(len(table_input), 8) * 35,
                ),
                column_config=build_editor_column_config(
                    visible_editor_columns,
                ),
                disabled=False,
            )

    apply_col, input_col, package_col, local_col = st.columns(4)

    can_apply = (
        not master_df.empty
        and bool(visible_editor_columns)
        and not filtered_df.empty
    )

    if apply_col.button(
        "✅ Apply edits",
        type="primary",
        use_container_width=True,
        disabled=not can_apply,
    ):
        try:
            updated = merge_filtered_edits(
                master_df,
                filtered_df,
                edited,
                visible_editor_columns,
            )
            set_master(updated, add_undo=True)
            st.session_state.source_name = "Edited in NWD Scheduler v1.5.1"
            st.success("Edits applied to the complete master schedule.")
            st.rerun()
        except Exception as exc:
            st.error(
                f"Could not apply edits: {type(exc).__name__}: {exc}"
            )

    updated_file_name = updated_input_filename(
        st.session_state.get("source_file_name", "")
    )
    input_col.download_button(
        "⬇ Updated input",
        data=dataframe_to_csv_bytes(master_df),
        file_name=updated_file_name,
        mime="text/csv",
        use_container_width=True,
        help="Complete master schedule with all applied edits.",
    )

    current_export = make_csv_package(
        master_df,
        filtered_df,
        conflicts_df,
        issues_df,
    )
    package_col.download_button(
        "⬇ CSV package",
        data=current_export,
        file_name=(
            f"NWD_Scheduler_v1.5.1_"
            f"{datetime.now():%Y%m%d_%H%M}.zip"
        ),
        mime="application/zip",
        use_container_width=True,
    )

    local_save_path = APP_DIR / updated_file_name
    if local_col.button(
        "💾 Save beside app",
        use_container_width=True,
        help=f"Saves to {local_save_path}",
    ):
        try:
            local_save_path.write_bytes(
                dataframe_to_csv_bytes(master_df)
            )
            st.success(f"Saved: {local_save_path}")
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
            highlight = a2.checkbox(
                "Highlight item",
                value=False,
                help="Adds a purple dashed border around the Gantt bar.",
            )
            highlight_label = a1.text_input(
                "Highlight label",
                value="",
                placeholder="New Technology",
            )
            sor_file = a2.text_input(
                "SoR PDF filename",
                value="",
                placeholder="NWD-001_SoR.pdf",
            )
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
                "Highlight": bool(highlight),
                "Highlight_Label": highlight_label,
                "SoR_File": sor_file,
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
        bulk_highlight = st.selectbox(
            "Highlight border",
            ["— Keep current —", "Set highlight", "Remove highlight"],
        )
        bulk_highlight_label = st.text_input(
            "Highlight label",
            value="",
            placeholder="New Technology",
        )
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
            if bulk_highlight == "Set highlight":
                updated.loc[mask, "Highlight"] = True
                if bulk_highlight_label.strip():
                    updated.loc[mask, "Highlight_Label"] = bulk_highlight_label.strip()
            elif bulk_highlight == "Remove highlight":
                updated.loc[mask, "Highlight"] = False
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
        ### NWD Scheduler v1.5.1 workflow
        1. The application opens with an **empty schedule**.
        2. Upload an Excel/CSV schedule or add items manually.
        3. Choose visible columns under **Table & tooltip columns**. Hidden table
           columns are automatically excluded from Gantt mouse-over text.
        4. Mark `Highlight = True` for New Technology or another special item.
           It receives a purple dashed Gantt border.
        5. Apply edits and download **Updated input** to preserve the complete
           amended master schedule.
        6. Put SoR PDFs in the `input` folder beside the Python file, upload PDFs
           temporarily from the sidebar, and optionally populate `SoR_File`.
        7. Click a Gantt target to open its matching SoR PDF. If no match exists,
           the application displays **SoR not found**.

        ### SoR filename matching
        The most reliable method is to populate `SoR_File`, for example
        `R-1501_SoR.pdf`. If it is blank, the app searches PDF filenames using
        `Target_ID` and `Well_Name`.

        ### Streamlit Cloud
        A cloud app cannot read a folder on your personal computer. For cloud use,
        either:
        - add permitted PDFs to the repository's `input` folder, or
        - upload PDFs from the application's sidebar for the current session.

        Do not put confidential operational documents in a public repository.

        ### Important editing behavior
        Filtered edits are merged into the complete master schedule. Rows outside
        active filters and all hidden column values remain unchanged.
        """
    )

st.divider()
st.caption("NWD Scheduler v1.5.1 • Empty start • Highlighted technology items • Dynamic tooltips • Click-to-open SoR")
