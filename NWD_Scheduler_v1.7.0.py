# -*- coding: utf-8 -*-
"""
NWD Scheduler v1.7.0 - Streamlit application

Run:
    streamlit run NWD_Scheduler_v1.7.0.py

The application opens with an empty schedule. Upload an Excel/CSV schedule
when required. Place linked PDF documents in an "input" folder beside this
Python file, or upload PDFs temporarily from the sidebar. Each schedule item
can use its own document name, such as SoR, WCS, ToR, or any custom label.

Outputs are CSV-based.
"""

from __future__ import annotations

import base64
import gzip
import html
import io
import json
import math
import re
import sqlite3
import zipfile
import uuid
from datetime import date, datetime, timedelta, timezone
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
DEFAULT_DOCUMENT_DIR = APP_DIR / "input"
COMPUTED_COLUMNS = ["Duration_Days", "Start_Year", "End_Year"]
INTERNAL_ROW_KEY = "__Original_Target_ID"
HIGHLIGHT_COLOR = "#7C3AED"
SCHEDULE_STATE_DIR = Path.home() / ".nwd_scheduler_state"
SCHEDULE_SNAPSHOT_PATTERN = re.compile(r"^[a-f0-9]{32}$")
USAGE_DB_PATH = SCHEDULE_STATE_DIR / "usage.sqlite3"
ACTIVE_USER_MINUTES = 15

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
    "Document_Name",
    "Document_File",
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
    "Document_Name",
    "Document_File",
    "Notes",
]

DATE_COLUMNS = ["Start_Date", "End_Date"]
BOOLEAN_COLUMNS = ["Highlight"]

STATUS_OPTIONS = ["Not Started", "Ready", "In Progress", "On Hold", "Completed", "Cancelled"]
PRIORITY_OPTIONS = ["Critical", "High", "Medium", "Low"]
TARGET_TYPE_OPTIONS = [
    "Producer", "Water Injector", "Gas Injector", "Observation", "Disposal",
    "Appraisal", "TAR", "Drilling Break", "Rig Maintenance", "Rig Move",
    "New Technology", "Other Event",
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
    "document name": "Document_Name",
    "document_name": "Document_Name",
    "document title": "Document_Name",
    "document_title": "Document_Name",
    "document type": "Document_Name",
    "document_type": "Document_Name",
    "document file": "Document_File",
    "document_file": "Document_File",
    "linked document": "Document_File",
    "linked_document": "Document_File",
    # Backward compatibility with existing schedule files.
    "sor": "Document_File",
    "sor file": "Document_File",
    "sor_file": "Document_File",
    "sor document": "Document_File",
    "sor_document": "Document_File",
    "event id": "Target_ID",
    "event_id": "Target_ID",
    "item id": "Target_ID",
    "item_id": "Target_ID",
    "event name": "Well_Name",
    "event_name": "Well_Name",
    "item name": "Well_Name",
    "item_name": "Well_Name",
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

      /* Make Gantt bars visibly clickable for linked-document opening. */
      [data-testid="stPlotlyChart"] .plotly .barlayer path,
      [data-testid="stPlotlyChart"] .plotly .scatterlayer text,
      [data-testid="stPlotlyChart"] .plotly .scatterlayer .point,
      [data-testid="stPlotlyChart"] .plotly .nsewdrag,
      [data-testid="stPlotlyChart"] .plotly .cursor-crosshair {
          cursor: pointer !important;
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

    original_column_names = [str(column).strip().lower() for column in df.columns]
    legacy_sor_column = any(
        name in {"sor", "sor file", "sor_file", "sor document", "sor_document"}
        for name in original_column_names
    )
    df.columns = [clean_column_name(c) for c in df.columns]

    # If old and new document columns are both present, combine them safely.
    if df.columns.duplicated().any():
        combined_columns: Dict[str, pd.Series] = {}
        for column in dict.fromkeys(df.columns):
            matches = df.loc[:, df.columns == column]
            combined = matches.iloc[:, 0]
            for duplicate_index in range(1, matches.shape[1]):
                fallback = matches.iloc[:, duplicate_index]
                populated = combined.notna() & combined.astype(str).str.strip().ne("")
                combined = combined.where(populated, fallback)
            combined_columns[column] = combined
        df = pd.DataFrame(combined_columns)

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

    if legacy_sor_column:
        legacy_document_rows = df["Document_File"].ne("") & df["Document_Name"].eq("")
        df.loc[legacy_document_rows, "Document_Name"] = "SoR"

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


def authenticated_user_identity() -> Tuple[str, str, str]:
    """Return a stable Streamlit-auth identity without exposing auth tokens."""
    try:
        user = getattr(st, "user", None)
        if user is None:
            return "", "", ""
        if hasattr(user, "to_dict"):
            values = user.to_dict()
        elif isinstance(user, dict):
            values = dict(user)
        else:
            values = {}
        if not bool(values.get("is_logged_in", False)):
            return "", "", ""

        email = str(
            values.get("email")
            or values.get("preferred_username")
            or ""
        ).strip()
        display_name = str(
            values.get("name")
            or values.get("given_name")
            or email
            or "Authenticated user"
        ).strip()
        stable_claim = str(
            values.get("oid")
            or values.get("sub")
            or email
            or display_name
        ).strip()
        return f"auth:{stable_claim}", display_name[:120], email[:160]
    except Exception:
        return "", "", ""


def record_user_activity(
    session_id: str,
    user_key: str,
    display_name: str,
    email: str = "",
) -> None:
    """Record one session heartbeat in the app server's lightweight database."""
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    try:
        SCHEDULE_STATE_DIR.mkdir(parents=True, exist_ok=True)
        with sqlite3.connect(USAGE_DB_PATH, timeout=5) as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS usage_sessions (
                    session_id TEXT PRIMARY KEY,
                    user_key TEXT NOT NULL,
                    display_name TEXT NOT NULL,
                    email TEXT NOT NULL DEFAULT '',
                    first_seen TEXT NOT NULL,
                    last_seen TEXT NOT NULL
                )
                """
            )
            connection.execute(
                """
                INSERT INTO usage_sessions (
                    session_id, user_key, display_name, email,
                    first_seen, last_seen
                ) VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(session_id) DO UPDATE SET
                    user_key = excluded.user_key,
                    display_name = excluded.display_name,
                    email = excluded.email,
                    last_seen = excluded.last_seen
                """,
                (
                    str(session_id),
                    str(user_key),
                    str(display_name)[:120],
                    str(email)[:160],
                    now,
                    now,
                ),
            )
            connection.commit()
    except (OSError, sqlite3.Error):
        # Usage visibility is optional and must never interrupt the scheduler.
        pass


def usage_summary() -> Tuple[List[Dict[str, object]], int, int]:
    """Return active identities, total identities and total opened sessions."""
    if not USAGE_DB_PATH.exists():
        return [], 0, 0

    active_after = (
        datetime.now(timezone.utc) - timedelta(minutes=ACTIVE_USER_MINUTES)
    ).isoformat(timespec="seconds")
    try:
        with sqlite3.connect(USAGE_DB_PATH, timeout=5) as connection:
            connection.row_factory = sqlite3.Row
            rows = connection.execute(
                """
                SELECT
                    user_key,
                    MAX(display_name) AS display_name,
                    MAX(email) AS email,
                    MIN(first_seen) AS first_seen,
                    MAX(last_seen) AS last_seen,
                    COUNT(*) AS session_count
                FROM usage_sessions
                GROUP BY user_key
                ORDER BY last_seen DESC
                """
            ).fetchall()
            total_sessions = int(
                connection.execute(
                    "SELECT COUNT(*) FROM usage_sessions"
                ).fetchone()[0]
            )
    except sqlite3.Error:
        return [], 0, 0

    users = [dict(row) for row in rows]
    active_users = [
        user for user in users if str(user.get("last_seen", "")) >= active_after
    ]
    return active_users, len(users), total_sessions



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
        persisted_settings = read_persisted_view_settings()
        restored_df, restored_source_name = load_schedule_snapshot(
            persisted_settings.get("schedule_snapshot_id", "")
        )

        if restored_df is not None and not restored_df.empty:
            st.session_state.schedule_df = restored_df
            st.session_state.source_file_name = str(
                persisted_settings.get(
                    "source_file_name",
                    restored_source_name,
                )
                or restored_source_name
                or ""
            )
            st.session_state.source_name = (
                st.session_state.source_file_name
                or "Restored schedule"
            )
            st.session_state.schedule_snapshot_id = valid_schedule_snapshot_id(
                persisted_settings.get("schedule_snapshot_id", "")
            )
        else:
            st.session_state.schedule_df = create_empty_dataframe()
            st.session_state.source_name = "Empty schedule"
            st.session_state.source_file_name = ""
            st.session_state.schedule_snapshot_id = ""

    st.session_state.setdefault("undo_stack", [])
    st.session_state.setdefault("filter_reset_token", 0)
    st.session_state.setdefault(
        "uploaded_document_files",
        st.session_state.pop("uploaded_sor_files", {}),
    )
    st.session_state.setdefault(
        "selected_document_item_id",
        st.session_state.pop("selected_sor_target_id", ""),
    )
    st.session_state.setdefault("gantt_selection_token", 0)
    st.session_state.setdefault("schedule_snapshot_id", "")
    st.session_state.setdefault("usage_session_id", uuid.uuid4().hex)



VIEW_QUERY_PARAM = "nwd_view"


def valid_schedule_snapshot_id(value: object) -> str:
    snapshot_id = str(value or "").strip().lower()
    return snapshot_id if SCHEDULE_SNAPSHOT_PATTERN.fullmatch(snapshot_id) else ""


def schedule_snapshot_path(snapshot_id: str) -> Path:
    return SCHEDULE_STATE_DIR / f"{snapshot_id}.csv"


def save_schedule_snapshot(
    df: pd.DataFrame,
    source_file_name: str,
    existing_snapshot_id: str = "",
) -> str:
    """
    Save the current master schedule on the app server.

    The URL stores only a random snapshot identifier, not the schedule contents.
    This allows browser refresh to restore the current schedule during the
    lifetime of the Streamlit deployment/container.
    """
    normalized = normalize_dataframe(df)
    if normalized.empty:
        return ""

    snapshot_id = valid_schedule_snapshot_id(existing_snapshot_id)
    if not snapshot_id:
        snapshot_id = uuid.uuid4().hex

    try:
        SCHEDULE_STATE_DIR.mkdir(parents=True, exist_ok=True)
        schedule_snapshot_path(snapshot_id).write_bytes(
            dataframe_to_csv_bytes(normalized)
        )

        metadata = {
            "source_file_name": str(source_file_name or ""),
            "saved_at": datetime.now().isoformat(timespec="seconds"),
        }
        (SCHEDULE_STATE_DIR / f"{snapshot_id}.json").write_text(
            json.dumps(metadata, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return snapshot_id
    except OSError:
        return ""


def load_schedule_snapshot(
    snapshot_id: object,
) -> Tuple[Optional[pd.DataFrame], str]:
    """Load a previously persisted schedule snapshot."""
    valid_id = valid_schedule_snapshot_id(snapshot_id)
    if not valid_id:
        return None, ""

    csv_path = schedule_snapshot_path(valid_id)
    if not csv_path.exists():
        return None, ""

    try:
        restored = load_schedule_file(csv_path, csv_path.name)
        source_file_name = ""
        metadata_path = SCHEDULE_STATE_DIR / f"{valid_id}.json"
        if metadata_path.exists():
            metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
            source_file_name = str(metadata.get("source_file_name", "") or "")
        return restored, source_file_name
    except Exception:
        return None, ""


def embedded_schedule_payload(
    df: pd.DataFrame,
) -> Dict[str, object]:
    """
    Compress the complete schedule into a portable JSON-template payload.

    A browser cannot silently reopen a local CSV path after refresh, so the
    template stores the actual CSV data instead of an inaccessible local link.
    """
    normalized = normalize_dataframe(df)
    csv_bytes = dataframe_to_csv_bytes(normalized)
    compressed = gzip.compress(csv_bytes, compresslevel=9)
    return {
        "encoding": "gzip+base64",
        "row_count": int(len(normalized)),
        "data": base64.b64encode(compressed).decode("ascii"),
    }


def dataframe_from_embedded_payload(
    payload: object,
) -> Optional[pd.DataFrame]:
    """Restore a schedule dataframe embedded in a JSON view template."""
    if not isinstance(payload, dict):
        return None

    if payload.get("encoding") != "gzip+base64":
        return None

    encoded = str(payload.get("data", "") or "")
    if not encoded:
        return None

    try:
        compressed = base64.b64decode(encoded)
        csv_bytes = gzip.decompress(compressed)
        return load_schedule_file(
            io.BytesIO(csv_bytes),
            "template_schedule.csv",
        )
    except Exception as exc:
        raise ValueError(
            f"Could not restore the schedule embedded in the template: {exc}"
        ) from exc


VIEW_WIDGET_KEYS = [
    "visible_columns_widget",
    "document_folder_widget",
    "gantt_type_widget",
    "y_axis_column_widget",
    "color_mode_widget",
    "group_by_widget",
    "label_mode_widget",
    "bar_text_mode_widget",
    "show_today_widget",
    "show_progress_widget",
    "show_horizontal_grid_widget",
    "height_per_row_widget",
]


def encode_view_settings(settings: Dict[str, object]) -> str:
    """Encode view settings into a compact URL-safe string."""
    payload = json.dumps(
        settings,
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")
    return base64.urlsafe_b64encode(payload).decode("ascii").rstrip("=")


def decode_view_settings(encoded: object) -> Dict[str, object]:
    """Decode URL or template settings safely."""
    value = str(encoded or "").strip()
    if not value:
        return {}

    try:
        padding = "=" * (-len(value) % 4)
        decoded = base64.urlsafe_b64decode(value + padding).decode("utf-8")
        result = json.loads(decoded)
        return result if isinstance(result, dict) else {}
    except Exception:
        return {}


def read_persisted_view_settings() -> Dict[str, object]:
    """Read the saved view from the page URL after a browser refresh."""
    try:
        encoded = st.query_params.get(VIEW_QUERY_PARAM, "")
    except Exception:
        return {}
    return decode_view_settings(encoded)


def write_persisted_view_settings(settings: Dict[str, object]) -> None:
    """Persist current settings in the URL so they survive page refresh."""
    encoded = encode_view_settings(settings)
    try:
        current = str(st.query_params.get(VIEW_QUERY_PARAM, "") or "")
        if current != encoded:
            st.query_params[VIEW_QUERY_PARAM] = encoded
    except Exception:
        # The app still works if query-parameter persistence is unavailable.
        pass


def clear_persisted_view_settings() -> None:
    """Remove the saved view from the URL."""
    try:
        if VIEW_QUERY_PARAM in st.query_params:
            del st.query_params[VIEW_QUERY_PARAM]
    except Exception:
        try:
            st.query_params.clear()
        except Exception:
            pass


def clear_view_widget_state() -> None:
    """Clear stable widget state before loading a template or resetting defaults."""
    for key in VIEW_WIDGET_KEYS:
        st.session_state.pop(key, None)


def safe_setting_list(
    settings: Dict[str, object],
    key: str,
    allowed_values: Sequence[str],
) -> List[str]:
    """Return saved multiselect values that still exist in the current dataset."""
    raw = settings.get(key, [])
    if not isinstance(raw, list):
        return []
    allowed = set(str(value) for value in allowed_values)
    return [str(value) for value in raw if str(value) in allowed]


def safe_select_index(
    options: Sequence[str],
    saved_value: object,
    default_index: int = 0,
) -> int:
    """Return a safe selectbox index."""
    try:
        return list(options).index(str(saved_value))
    except ValueError:
        return default_index


def safe_bool(settings: Dict[str, object], key: str, default: bool) -> bool:
    value = settings.get(key, default)
    return value if isinstance(value, bool) else default


def parse_saved_date(value: object) -> Optional[date]:
    try:
        parsed = pd.to_datetime(value, errors="coerce")
        return None if pd.isna(parsed) else parsed.date()
    except Exception:
        return None


def create_view_settings(
    *,
    visible_columns: Sequence[str],
    search: str,
    reservoirs: Sequence[str],
    areas: Sequence[str],
    rigs: Sequence[str],
    statuses: Sequence[str],
    priorities: Sequence[str],
    target_types: Sequence[str],
    campaigns: Sequence[str],
    owners: Sequence[str],
    include_cancelled: bool,
    date_window: Tuple[date, date],
    document_folder_text: str,
    gantt_type: str,
    y_axis_column: str,
    color_mode: str,
    group_by: str,
    label_mode: str,
    bar_text_mode: str,
    show_today: bool,
    show_progress: bool,
    show_horizontal_grid: bool,
    height_per_row: int,
    source_file_name: str,
    schedule_snapshot_id: str,
) -> Dict[str, object]:
    """Build the complete reusable sidebar/filter template."""
    return {
        "template_version": 1,
        "visible_columns": list(visible_columns),
        "search": str(search),
        "reservoirs": list(reservoirs),
        "areas": list(areas),
        "rigs": list(rigs),
        "statuses": list(statuses),
        "priorities": list(priorities),
        "target_types": list(target_types),
        "campaigns": list(campaigns),
        "owners": list(owners),
        "include_cancelled": bool(include_cancelled),
        "date_window": [
            date_window[0].isoformat(),
            date_window[1].isoformat(),
        ],
        "document_folder": str(document_folder_text),
        "gantt_type": str(gantt_type),
        "y_axis_column": str(y_axis_column),
        "color_mode": str(color_mode),
        "group_by": str(group_by),
        "label_mode": str(label_mode),
        "bar_text_mode": str(bar_text_mode),
        "show_today": bool(show_today),
        "show_progress": bool(show_progress),
        "show_horizontal_grid": bool(show_horizontal_grid),
        "height_per_row": int(height_per_row),
        "source_file_name": str(source_file_name or ""),
        "schedule_snapshot_id": valid_schedule_snapshot_id(
            schedule_snapshot_id
        ),
    }


def view_template_bytes(
    settings: Dict[str, object],
    template_name: str,
    schedule_df: pd.DataFrame,
    source_file_name: str,
) -> bytes:
    """
    Create a portable JSON template containing:
    - current filters and chart settings;
    - original input filename;
    - the complete current schedule as compressed CSV data.
    """
    payload = {
        "template_name": str(template_name or "NWD View"),
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "source_file_name": str(source_file_name or ""),
        "settings": settings,
        "schedule": embedded_schedule_payload(schedule_df),
    }
    return json.dumps(
        payload,
        ensure_ascii=False,
        indent=2,
    ).encode("utf-8")


def read_uploaded_view_template(
    uploaded_file: object,
) -> Tuple[Dict[str, object], Optional[pd.DataFrame], str]:
    """Read settings and the optional embedded schedule from a JSON template."""
    if uploaded_file is None:
        return {}, None, ""

    raw = uploaded_file.getvalue().decode("utf-8-sig")
    payload = json.loads(raw)
    if not isinstance(payload, dict):
        raise ValueError("The template JSON must contain an object.")

    settings = payload.get("settings", payload)
    if not isinstance(settings, dict):
        raise ValueError("The template does not contain valid settings.")

    restored_schedule = dataframe_from_embedded_payload(
        payload.get("schedule")
    )
    source_file_name = str(
        payload.get(
            "source_file_name",
            settings.get("source_file_name", ""),
        )
        or ""
    )

    return settings, restored_schedule, source_file_name


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
        issues.append({"Severity": "Error", "Target_ID": row["Target_ID"], "Field": "Target_ID", "Issue": "Item / event ID is duplicated."})

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
        "Target_ID": "Item / event ID",
        "Well_Name": "Well / event name",
        "Target_Type": "Item type",
        "Start_Date": "Start date",
        "End_Date": "End date",
        "Progress_Pct": "Progress %",
        "Highlight_Label": "Highlight label",
        "Document_Name": "Document name",
        "Document_File": "Document file",
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
            "Item / event ID",
            required=False,
            width="small",
            help="Blank or duplicate IDs are corrected automatically after Apply.",
        ),
        "Well_Name": st.column_config.TextColumn(
            "Well / event name",
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
            "Item type",
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
            "Stored bar color",
            required=False,
            width="small",
            help=(
                "Normally keep this technical value hidden and use the visual "
                "color pickers below the table or in Bulk Actions."
            ),
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
        "Document_Name": st.column_config.TextColumn(
            "Document name",
            required=False,
            width="medium",
            help="User-defined label, for example SoR, WCS, ToR, or another name.",
        ),
        "Document_File": st.column_config.TextColumn(
            "Document file",
            required=False,
            width="large",
            help="PDF filename stored in the configured document folder.",
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


def resolve_document_directory(folder_text: str) -> Path:
    folder_text = str(folder_text or "").strip()
    if not folder_text:
        return DEFAULT_DOCUMENT_DIR
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


def resolve_document(
    item_row: pd.Series,
    document_directory: Path,
    uploaded_files: Dict[str, bytes],
) -> Optional[Dict[str, object]]:
    """
    Resolve a linked PDF using:
    1. The row's Document_File value.
    2. Exact item/event ID or well/event name filename matching.
    3. Partial item/event ID or well/event name filename matching.
    """
    uploaded_lookup = {
        name.lower(): {"name": name, "bytes": content, "source": "Uploaded PDF"}
        for name, content in uploaded_files.items()
        if str(name).lower().endswith(".pdf")
    }

    local_files = local_pdf_files(document_directory)
    local_lookup = {path.name.lower(): path for path in local_files}

    requested = str(item_row.get("Document_File", "") or "").strip()
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
                    document_directory / requested_path,
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

    item_keys = [
        normalized_document_key(item_row.get("Target_ID", "")),
        normalized_document_key(item_row.get("Well_Name", "")),
    ]
    item_keys = [key for key in item_keys if key]

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
        if candidate_key(candidate) in item_keys
    ]
    partial_matches = [
        candidate
        for candidate in candidates
        if any(key in candidate_key(candidate) for key in item_keys)
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


def document_display_name(
    item_row: pd.Series,
    document: Optional[Dict[str, object]] = None,
) -> str:
    """Return the user's label, with a sensible fallback for older inputs."""
    explicit_name = str(item_row.get("Document_Name", "") or "").strip()
    if explicit_name:
        return explicit_name

    requested_file = str(item_row.get("Document_File", "") or "").strip()
    resolved_file = str((document or {}).get("name", "") or "").strip()
    fallback_file = requested_file or resolved_file
    return Path(fallback_file).stem if fallback_file else "Linked document"


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


def show_pdf_viewer(
    document: Dict[str, object],
    item_id: str,
    document_name: str,
) -> None:
    pdf_bytes = bytes(document["bytes"])
    file_name = str(document["name"])

    info_col, download_col = st.columns([3, 1])
    info_col.success(f"{document_name} found for {item_id}: {file_name}")
    download_col.download_button(
        f"Download {document_name}",
        data=pdf_bytes,
        file_name=file_name,
        mime="application/pdf",
        use_container_width=True,
        key=f"download_document_{item_id}_{normalized_document_key(file_name)}",
    )

    if hasattr(st, "pdf"):
        try:
            st.pdf(
                pdf_bytes,
                height=850,
                key=f"document_pdf_{item_id}_{normalized_document_key(file_name)}",
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


def enable_zoom_aware_tooltips(base_font_size: int = 13) -> None:
    """
    Enlarge Plotly hover labels as the visible calendar range is zoomed in.

    Streamlit's native Plotly event bridge currently reports point selections,
    but not relayout/zoom events. This small browser-side listener keeps the
    existing click-to-open behavior and adjusts only the hover-label font.
    """
    components.html(
        f"""
        <script>
        (() => {{
          const BASE_SIZE = {int(base_font_size)};
          const MAX_SCALE = 2.4;

          function rangeSpan(axis) {{
            if (!axis || !Array.isArray(axis.range) || axis.range.length < 2) return 0;
            const start = new Date(axis.range[0]).getTime();
            const end = new Date(axis.range[1]).getTime();
            return Number.isFinite(start) && Number.isFinite(end) ? Math.abs(end - start) : 0;
          }}

          function attach(attempt = 0) {{
            try {{
              const charts = window.parent.document.querySelectorAll(
                '[data-testid="stPlotlyChart"] .js-plotly-plot'
              );
              const chart = charts[charts.length - 1];
              if (!chart || !chart._fullLayout || typeof chart.on !== 'function') {{
                if (attempt < 30) setTimeout(() => attach(attempt + 1), 100);
                return;
              }}

              const initialSpan = rangeSpan(chart._fullLayout.xaxis);
              if (!initialSpan) return;
              chart.dataset.nwdTooltipInitialSpan = String(initialSpan);

              const updateFont = () => {{
                const baseline = Number(chart.dataset.nwdTooltipInitialSpan) || initialSpan;
                const current = rangeSpan(chart._fullLayout.xaxis) || baseline;
                const scale = Math.min(MAX_SCALE, Math.max(1, Math.sqrt(baseline / current)));
                const size = Math.round(BASE_SIZE * scale);

                chart.layout.hoverlabel = chart.layout.hoverlabel || {{}};
                chart.layout.hoverlabel.font = chart.layout.hoverlabel.font || {{}};
                chart.layout.hoverlabel.font.size = size;
                if (chart._fullLayout.hoverlabel && chart._fullLayout.hoverlabel.font) {{
                  chart._fullLayout.hoverlabel.font.size = size;
                }}
              }};

              if (chart.dataset.nwdTooltipZoomAttached !== 'true') {{
                chart.on('plotly_relayout', updateFont);
                chart.on('plotly_hover', updateFont);
                chart.dataset.nwdTooltipZoomAttached = 'true';
              }}
              updateFont();
            }} catch (error) {{
              if (attempt < 30) setTimeout(() => attach(attempt + 1), 100);
            }}
          }}

          attach();
        }})();
        </script>
        """,
        height=0,
        width=0,
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
            text="The schedule is empty. Upload a file or add an item.",
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

    if label_mode in {"Well name", "Item / event name"}:
        plot_df["Task_Label"] = plot_df["Well_Name"].fillna("").astype(str)
    elif label_mode in {"Target ID + well", "Item ID + name"}:
        plot_df["Task_Label"] = (
            plot_df["Target_ID"].fillna("").astype(str)
            + " | "
            + plot_df["Well_Name"].fillna("").astype(str)
        )
    elif label_mode in {"Rig + well", "Rig + name"}:
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
    else:
        field = color_mode
        if field == "Status":
            cmap = {**category_color_map(plot_df[field]), **STATUS_COLORS}
        elif field == "Priority":
            cmap = {**category_color_map(plot_df[field]), **PRIORITY_COLORS}
        else:
            cmap = category_color_map(plot_df[field])
        colors = plot_df[field].map(cmap).fillna(DEFAULT_COLOR).tolist()
    plot_df["_Bar_Color"] = colors

    duration_ms = (
        (plot_df["End_Date"] - plot_df["Start_Date"]).dt.total_seconds() * 1000
    ).clip(lower=86_400_000)

    progress_text = plot_df["Progress_Pct"].astype(int).astype(str) + "%"
    target_text = plot_df["Well_Name"].fillna("").astype(str).str.strip()
    target_text = target_text.where(
        target_text.ne(""),
        plot_df["Target_ID"].fillna("").astype(str),
    )

    if bar_text_mode in {"Target name + progress %", "Item name + progress %"}:
        bar_text = target_text + "<br>" + progress_text
    elif bar_text_mode in {"Target name only", "Item name only"}:
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
            str(row.get("Document_File", "")),
            str(row.get("Document_Name", "")),
            (
                "Click to open the linked PDF"
                if str(row.get("Document_File", "") or "").strip()
                else "Click to check for a matching PDF"
            ),
        ]
        values.extend(
            format_hover_value(row.get(column, ""), column)
            for column in tooltip_columns
        )
        customdata.append(values)

    hover_lines = []
    for data_index, column in enumerate(tooltip_columns, start=5):
        hover_lines.append(
            f"<b>{html.escape(column_display_name(column))}:</b> "
            f"%{{customdata[{data_index}]}}<br>"
        )
    hover_lines.append("<i>%{customdata[4]}</i>")
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
        height=chart_height + 75,
        barmode="overlay",
        bargap=0.12 if gantt_type == "Compact merged lanes" else 0.24,
        hoverlabel={"align": "left", "font": {"size": 13}},
        clickmode="event+select",
        selectionrevision="nwd-v1.7.0",
        margin={"l": 20, "r": 25, "t": 88, "b": 95},
        uniformtext={"mode": "hide", "minsize": 8},
        xaxis={
            "title": "Calendar date",
            "type": "date",
            "showgrid": True,
            "gridcolor": "rgba(148,163,184,.20)",
            "gridwidth": 1,
            "rangeslider": {"visible": False},
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

    if color_mode == "Custom row color":
        # In custom mode, explain each visible color using the item types that
        # actually use it. This keeps the legend accurate without showing hex.
        for color in plot_df["_Bar_Color"].drop_duplicates().tolist():
            color_rows = plot_df[plot_df["_Bar_Color"] == color]
            item_types = [
                value
                for value in color_rows["Target_Type"]
                .fillna("")
                .astype(str)
                .str.strip()
                .drop_duplicates()
                .tolist()
                if value
            ]
            legend_label = " / ".join(item_types[:4]) or "Other schedule item"
            if len(item_types) > 4:
                legend_label += f" +{len(item_types) - 4} more"
            fig.add_trace(
                go.Bar(
                    x=[None],
                    y=[None],
                    marker={"color": color},
                    name=legend_label,
                    orientation="h",
                    showlegend=True,
                    hoverinfo="skip",
                )
            )
    else:
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

    fig.update_layout(
        legend={
            "title": {"text": "Color legend"},
            "orientation": "h",
            "y": -0.20,
            "x": 0,
            "xanchor": "left",
            "yanchor": "top",
            "traceorder": "normal",
        }
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

if master_df.empty:
    st.session_state.schedule_snapshot_id = ""
else:
    current_snapshot_id = save_schedule_snapshot(
        master_df,
        st.session_state.get("source_file_name", ""),
        st.session_state.get("schedule_snapshot_id", ""),
    )
    if current_snapshot_id:
        st.session_state.schedule_snapshot_id = current_snapshot_id

all_editable_columns = editable_schedule_columns(master_df)
persisted_view_settings = read_persisted_view_settings()
default_visible_columns = [
    column for column in all_editable_columns if column != "Color"
]

if "visible_columns_widget" not in st.session_state:
    saved_visible_columns = persisted_view_settings.get(
        "visible_columns",
        list(default_visible_columns),
    )
    if not isinstance(saved_visible_columns, list):
        saved_visible_columns = list(default_visible_columns)

    st.session_state.visible_columns_widget = [
        column
        for column in saved_visible_columns
        if column in all_editable_columns
    ] or list(default_visible_columns)
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

    with st.expander("👥 Usage", expanded=False):
        auth_key, auth_name, auth_email = authenticated_user_identity()
        if auth_key:
            current_user_key = auth_key
            current_user_name = auth_name
            current_user_email = auth_email
            st.caption(f"Signed in as {current_user_name}")
        else:
            manual_name = st.text_input(
                "Your name",
                key="usage_manual_name",
                placeholder="Enter name for usage tracking",
                help=(
                    "If app authentication is not configured, this name identifies "
                    "you in the local usage list."
                ),
            ).strip()
            current_user_name = manual_name or "Anonymous user"
            current_user_email = ""
            current_user_key = (
                f"manual:{manual_name.casefold()}"
                if manual_name
                else f"anonymous:{st.session_state.usage_session_id}"
            )

        record_user_activity(
            st.session_state.usage_session_id,
            current_user_key,
            current_user_name,
            current_user_email,
        )
        active_users, total_users, total_sessions = usage_summary()
        usage_col_1, usage_col_2, usage_col_3 = st.columns(3)
        usage_col_1.metric("Active", len(active_users))
        usage_col_2.metric("Users", total_users)
        usage_col_3.metric("Opens", total_sessions)

        if active_users:
            activity_rows = []
            for user in active_users:
                display = str(user.get("display_name", "User"))
                email = str(user.get("email", "") or "")
                if email and email.casefold() not in display.casefold():
                    display = f"{display} ({email})"
                last_active = pd.to_datetime(
                    user.get("last_seen", ""), errors="coerce"
                )
                activity_rows.append(
                    {
                        "Active user": display,
                        "Last active": (
                            last_active.strftime("%d-%b %H:%M")
                            if pd.notna(last_active)
                            else ""
                        ),
                    }
                )
            st.dataframe(
                pd.DataFrame(activity_rows),
                hide_index=True,
                use_container_width=True,
            )
        st.caption(
            f"Active means activity within {ACTIVE_USER_MINUTES} minutes. "
            "Records are stored on this app server."
        )

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
            st.session_state.schedule_snapshot_id = save_schedule_snapshot(
                loaded,
                uploaded.name,
            )
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
        st.session_state.schedule_snapshot_id = ""
        st.session_state.selected_document_item_id = ""
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

    with st.expander("Linked PDF settings", expanded=False):
        document_folder_text = st.text_input(
            "Document PDF folder",
            value=str(
                persisted_view_settings.get(
                    "document_folder",
                    persisted_view_settings.get("sor_folder", "input"),
                )
            ),
            key="document_folder_widget",
            help=(
                "Relative paths are resolved beside the Python file. "
                "Example: input. An absolute local path also works when running locally."
            ),
        )

        uploaded_documents = st.file_uploader(
            "Upload linked PDFs for this session",
            type=["pdf"],
            accept_multiple_files=True,
            help=(
                "Useful on Streamlit Cloud. Uploaded PDFs remain available only "
                "for the current app session."
            ),
        )
        for pdf_file in uploaded_documents:
            st.session_state.uploaded_document_files[pdf_file.name] = pdf_file.getvalue()

        st.caption(
            f"Session PDFs: {len(st.session_state.uploaded_document_files)} • "
            f"Folder: {resolve_document_directory(document_folder_text)}"
        )

        if st.button(
            "Clear uploaded PDFs",
            use_container_width=True,
            disabled=not st.session_state.uploaded_document_files,
        ):
            st.session_state.uploaded_document_files = {}
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

    reservoir_options = select_options(master_df, "Reservoir")
    area_options = select_options(master_df, "Area")
    rig_options = select_options(master_df, "Rig")
    status_options = select_options(master_df, "Status")
    priority_options = select_options(master_df, "Priority")
    target_type_options = select_options(master_df, "Target_Type")
    campaign_options = select_options(master_df, "Campaign")
    owner_options = select_options(master_df, "Owner")

    search = st.text_input(
        "Search all fields",
        value=str(persisted_view_settings.get("search", "")),
        key=f"search_{reset_token}",
        placeholder="Well, event, item ID, owner, note...",
    )
    reservoirs = st.multiselect(
        "Reservoir",
        reservoir_options,
        default=safe_setting_list(
            persisted_view_settings,
            "reservoirs",
            reservoir_options,
        ),
        key=f"reservoir_{reset_token}",
    )
    areas = st.multiselect(
        "Area",
        area_options,
        default=safe_setting_list(
            persisted_view_settings,
            "areas",
            area_options,
        ),
        key=f"area_{reset_token}",
    )
    rigs = st.multiselect(
        "Rig",
        rig_options,
        default=safe_setting_list(
            persisted_view_settings,
            "rigs",
            rig_options,
        ),
        key=f"rig_{reset_token}",
    )
    statuses = st.multiselect(
        "Status",
        status_options,
        default=safe_setting_list(
            persisted_view_settings,
            "statuses",
            status_options,
        ),
        key=f"status_{reset_token}",
    )
    priorities = st.multiselect(
        "Priority",
        priority_options,
        default=safe_setting_list(
            persisted_view_settings,
            "priorities",
            priority_options,
        ),
        key=f"priority_{reset_token}",
    )
    target_types = st.multiselect(
        "Target type",
        target_type_options,
        default=safe_setting_list(
            persisted_view_settings,
            "target_types",
            target_type_options,
        ),
        key=f"type_{reset_token}",
    )
    campaigns = st.multiselect(
        "Campaign",
        campaign_options,
        default=safe_setting_list(
            persisted_view_settings,
            "campaigns",
            campaign_options,
        ),
        key=f"campaign_{reset_token}",
    )
    owners = st.multiselect(
        "Owner",
        owner_options,
        default=safe_setting_list(
            persisted_view_settings,
            "owners",
            owner_options,
        ),
        key=f"owner_{reset_token}",
    )
    include_cancelled = st.checkbox(
        "Include cancelled items",
        value=safe_bool(
            persisted_view_settings,
            "include_cancelled",
            False,
        ),
        key=f"cancelled_{reset_token}",
    )

    valid_dates = pd.concat(
        [master_df["Start_Date"], master_df["End_Date"]]
    ).dropna()
    if valid_dates.empty:
        min_date = date.today() - timedelta(days=30)
        max_date = date.today() + timedelta(days=365)
    else:
        min_date = valid_dates.min().date()
        max_date = valid_dates.max().date()

    saved_date_window = persisted_view_settings.get("date_window", [])
    saved_start = (
        parse_saved_date(saved_date_window[0])
        if isinstance(saved_date_window, list)
        and len(saved_date_window) == 2
        else None
    )
    saved_end = (
        parse_saved_date(saved_date_window[1])
        if isinstance(saved_date_window, list)
        and len(saved_date_window) == 2
        else None
    )

    default_window_start = saved_start or min_date
    default_window_end = saved_end or max_date
    if default_window_end < default_window_start:
        default_window_start, default_window_end = (
            default_window_end,
            default_window_start,
        )

    date_minimum = min(
        min_date - timedelta(days=365),
        default_window_start,
    )
    date_maximum = max(
        max_date + timedelta(days=365),
        default_window_end,
    )

    date_window = st.date_input(
        "Schedule window",
        value=(default_window_start, default_window_end),
        min_value=date_minimum,
        max_value=date_maximum,
        key=f"date_{reset_token}",
    )
    if not isinstance(date_window, tuple) or len(date_window) != 2:
        date_window = (min_date, max_date)

    if st.button("Clear all filters", use_container_width=True):
        cleared_settings = dict(persisted_view_settings)
        for key in [
            "search",
            "reservoirs",
            "areas",
            "rigs",
            "statuses",
            "priorities",
            "target_types",
            "campaigns",
            "owners",
            "include_cancelled",
            "date_window",
        ]:
            cleared_settings.pop(key, None)

        write_persisted_view_settings(cleared_settings)
        st.session_state.filter_reset_token += 1
        st.rerun()

    st.divider()
    st.subheader("Gantt settings")

    gantt_type_options = [
        "Detailed target rows",
        "Compact merged lanes",
        "Grouped target rows",
    ]
    gantt_type = st.selectbox(
        "Gantt chart type",
        gantt_type_options,
        index=safe_select_index(
            gantt_type_options,
            persisted_view_settings.get("gantt_type"),
            0,
        ),
        key="gantt_type_widget",
        help=(
            "Detailed gives one row per item. Compact merges all visible "
            "items sharing the selected Y-axis value onto one lane. "
            "Grouped keeps one row per item."
        ),
    )

    y_axis_options = [
        "Rig",
        "Pad",
        "Reservoir",
        "Area",
        "Campaign",
        "Owner",
        "Status",
        "Priority",
        "Target_Type",
    ]
    y_axis_column = st.selectbox(
        "Left-side Y-axis column",
        y_axis_options,
        index=safe_select_index(
            y_axis_options,
            persisted_view_settings.get("y_axis_column"),
            0,
        ),
        key="y_axis_column_widget",
    )

    color_mode_options = [
        "Custom row color",
        "Status",
        "Priority",
        "Rig",
        "Reservoir",
        "Area",
        "Target_Type",
        "Campaign",
    ]
    color_mode = st.selectbox(
        "Color bars by",
        color_mode_options,
        index=safe_select_index(
            color_mode_options,
            persisted_view_settings.get("color_mode"),
            0,
        ),
        key="color_mode_widget",
    )

    group_by_options = [
        "Rig",
        "Pad",
        "Area",
        "Reservoir",
        "Campaign",
        "Owner",
        "Status",
    ]
    group_by = st.selectbox(
        "Sort/group detailed rows by",
        group_by_options,
        index=safe_select_index(
            group_by_options,
            persisted_view_settings.get("group_by"),
            0,
        ),
        key="group_by_widget",
    )

    label_mode_options = [
        "Item ID + name",
        "Item / event name",
        "Rig + name",
        "Item ID",
    ]
    label_mode = st.selectbox(
        "Target label format",
        label_mode_options,
        index=safe_select_index(
            label_mode_options,
            persisted_view_settings.get("label_mode"),
            0,
        ),
        key="label_mode_widget",
    )

    bar_text_options = [
        "Item name + progress %",
        "Item name only",
        "Progress % only",
        "No text",
    ]
    bar_text_mode = st.selectbox(
        "Text inside bars",
        bar_text_options,
        index=safe_select_index(
            bar_text_options,
            persisted_view_settings.get("bar_text_mode"),
            0,
        ),
        key="bar_text_mode_widget",
        help=(
            "The item name is drawn above progress shading and remains visible."
        ),
    )

    show_today = st.checkbox(
        "Show today line",
        value=safe_bool(
            persisted_view_settings,
            "show_today",
            True,
        ),
        key="show_today_widget",
    )
    show_progress = st.checkbox(
        "Show progress shading",
        value=safe_bool(
            persisted_view_settings,
            "show_progress",
            True,
        ),
        key="show_progress_widget",
    )
    show_horizontal_grid = st.checkbox(
        "Show horizontal grid lines",
        value=safe_bool(
            persisted_view_settings,
            "show_horizontal_grid",
            True,
        ),
        key="show_horizontal_grid_widget",
    )

    saved_height = persisted_view_settings.get("height_per_row", 34)
    try:
        saved_height = int(saved_height)
    except (TypeError, ValueError):
        saved_height = 34
    saved_height = min(max(saved_height, 24), 60)

    height_per_row = st.slider(
        "Row/lane height",
        min_value=24,
        max_value=60,
        value=saved_height,
        step=2,
        key="height_per_row_widget",
    )

    current_view_settings = create_view_settings(
        visible_columns=visible_editor_columns,
        search=search,
        reservoirs=reservoirs,
        areas=areas,
        rigs=rigs,
        statuses=statuses,
        priorities=priorities,
        target_types=target_types,
        campaigns=campaigns,
        owners=owners,
        include_cancelled=include_cancelled,
        date_window=date_window,
        document_folder_text=document_folder_text,
        gantt_type=gantt_type,
        y_axis_column=y_axis_column,
        color_mode=color_mode,
        group_by=group_by,
        label_mode=label_mode,
        bar_text_mode=bar_text_mode,
        show_today=show_today,
        show_progress=show_progress,
        show_horizontal_grid=show_horizontal_grid,
        height_per_row=height_per_row,
        source_file_name=st.session_state.get(
            "source_file_name",
            "",
        ),
        schedule_snapshot_id=st.session_state.get(
            "schedule_snapshot_id",
            "",
        ),
    )

    # Keep the current filter/Gantt view after browser refresh.
    write_persisted_view_settings(current_view_settings)

    st.divider()
    with st.expander("Save / load view template", expanded=False):
        st.caption(
            "Filters and chart settings are saved in the page URL. The JSON "
            "template also contains the complete current schedule, so loading it "
            "restores both the CSV data and the view."
        )

        template_name = st.text_input(
            "Template name",
            value="NWD_View",
            key="view_template_name",
        )
        safe_template_name = re.sub(
            r"[^A-Za-z0-9._-]+",
            "_",
            template_name.strip(),
        ).strip("._-") or "NWD_View"

        st.download_button(
            "⬇ Download settings template",
            data=view_template_bytes(
                current_view_settings,
                template_name,
                master_df,
                st.session_state.get("source_file_name", ""),
            ),
            file_name=f"{safe_template_name}.json",
            mime="application/json",
            use_container_width=True,
        )

        uploaded_view_template = st.file_uploader(
            "Load settings template",
            type=["json"],
            key="view_template_uploader",
        )

        apply_template_col, reset_view_col = st.columns(2)

        if apply_template_col.button(
            "Apply template",
            use_container_width=True,
            disabled=uploaded_view_template is None,
        ):
            try:
                (
                    loaded_settings,
                    restored_schedule,
                    restored_source_file_name,
                ) = read_uploaded_view_template(
                    uploaded_view_template
                )

                if restored_schedule is not None:
                    set_master(restored_schedule, add_undo=True)
                    st.session_state.source_file_name = (
                        restored_source_file_name
                    )
                    st.session_state.source_name = (
                        restored_source_file_name
                        or "Schedule restored from template"
                    )
                    snapshot_id = save_schedule_snapshot(
                        restored_schedule,
                        restored_source_file_name,
                    )
                    st.session_state.schedule_snapshot_id = snapshot_id
                    loaded_settings["schedule_snapshot_id"] = snapshot_id
                    loaded_settings["source_file_name"] = (
                        restored_source_file_name
                    )

                write_persisted_view_settings(loaded_settings)
                clear_view_widget_state()
                st.session_state.filter_reset_token += 1
                st.success("Schedule and view template applied.")
                st.rerun()
            except Exception as exc:
                st.error(
                    f"Could not load settings template: "
                    f"{type(exc).__name__}: {exc}"
                )

        if reset_view_col.button(
            "Reset defaults",
            use_container_width=True,
        ):
            schedule_reference = {
                "source_file_name": st.session_state.get(
                    "source_file_name",
                    "",
                ),
                "schedule_snapshot_id": st.session_state.get(
                    "schedule_snapshot_id",
                    "",
                ),
            }
            write_persisted_view_settings(schedule_reference)
            clear_view_widget_state()
            st.session_state.filter_reset_token += 1
            st.session_state.pop("view_template_uploader", None)
            st.rerun()


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
k1.metric("Visible items", f"{len(filtered_df):,}", delta=f"of {len(master_df):,} total")
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
    enable_zoom_aware_tooltips(base_font_size=13)

    clicked_item_id = extract_selected_target_id(gantt_event)
    if clicked_item_id:
        st.session_state.selected_document_item_id = clicked_item_id

    st.caption(
        "Move the mouse over an item: the pointer changes to a hand. "
        "Click the bar, progress section or item text to open its linked PDF."
    )

    selected_item_id = st.session_state.get("selected_document_item_id", "")
    if selected_item_id:
        selected_rows = master_df[
            master_df["Target_ID"].astype(str) == str(selected_item_id)
        ]

        if selected_rows.empty:
            st.warning(f"Selected item {selected_item_id} is not in the current schedule.")
        else:
            selected_row = selected_rows.iloc[0]
            document = resolve_document(
                selected_row,
                resolve_document_directory(document_folder_text),
                st.session_state.uploaded_document_files,
            )
            document_name = document_display_name(selected_row, document)

            st.markdown("---")
            document_title_col, document_close_col = st.columns([5, 1])
            document_title_col.subheader(
                f"{document_name} — {selected_row['Target_ID']} | "
                f"{selected_row['Well_Name']}"
            )
            if document_close_col.button(
                "Close document",
                use_container_width=True,
                key=f"close_document_{selected_item_id}",
            ):
                st.session_state.selected_document_item_id = ""
                st.session_state.gantt_selection_token += 1
                st.rerun()

            if document is None:
                expected_name = str(
                    selected_row.get("Document_File", "") or ""
                ).strip()
                detail = (
                    f" Expected file: `{expected_name}`."
                    if expected_name
                    else ""
                )
                st.warning(
                    f"No linked PDF was found for {selected_item_id}.{detail}"
                )
                st.caption(
                    "Add the PDF to the configured input folder, upload it "
                    "from the sidebar, or populate the Document_File column."
                )
            else:
                show_pdf_viewer(document, selected_item_id, document_name)

with tab_edit:
    st.subheader("Editable filtered schedule")
    st.caption(
        "All visible cells are editable. Hidden columns remain preserved and "
        "are also excluded from Gantt tooltips."
    )

    if master_df.empty:
        st.info(
            "The schedule is currently empty. Upload an Excel/CSV file from the "
            "sidebar, add an item from Bulk Actions, or create the first blank row."
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
                "Document_Name": "",
                "Document_File": "",
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

    if not master_df.empty:
        with st.expander("🎨 Visual bar color editor", expanded=False):
            color_label_lookup = {
                str(row["Target_ID"]): (
                    f"{row['Target_ID']} | {row['Well_Name']} | {row['Target_Type']}"
                )
                for _, row in master_df.iterrows()
            }
            color_item_id = st.selectbox(
                "Schedule item",
                options=master_df["Target_ID"].astype(str).tolist(),
                format_func=lambda item: color_label_lookup.get(str(item), str(item)),
                key="visual_color_item_id",
            )
            current_color = valid_hex_color(
                master_df.loc[
                    master_df["Target_ID"].astype(str) == str(color_item_id),
                    "Color",
                ].iloc[0]
            )
            picked_color = st.color_picker(
                "Choose bar color",
                value=current_color,
                key=f"visual_color_picker_{normalized_document_key(color_item_id)}",
            )
            if st.button(
                "Apply selected color",
                use_container_width=True,
                key="apply_visual_item_color",
            ):
                updated = master_df.copy()
                updated.loc[
                    updated["Target_ID"].astype(str) == str(color_item_id),
                    "Color",
                ] = picked_color
                set_master(updated, add_undo=True)
                st.success(f"Color updated for {color_item_id}.")
                st.rerun()
            st.caption(
                "Use this picker instead of typing a color code. For several "
                "items at once, use Bulk Actions."
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
            st.session_state.source_name = "Edited in NWD Scheduler v1.7.0"
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
            f"NWD_Scheduler_v1.7.0_"
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
        st.subheader("Add a new schedule item")
        with st.form("add_target_form", clear_on_submit=True):
            a1, a2 = st.columns(2)
            well_name = a1.text_input("Well / event name", placeholder="R-1501 or Rig Maintenance")
            target_id = a2.text_input("Item / event ID (optional)", placeholder="Auto-generated")
            reservoir = a1.selectbox("Reservoir", sorted(set(RESERVOIR_OPTIONS + select_options(master_df, "Reservoir"))))
            area = a2.selectbox("Area", sorted(set(AREA_OPTIONS + select_options(master_df, "Area"))))
            rig = a1.selectbox("Rig", sorted(set(select_options(master_df, "Rig") + ["Unassigned"])))
            pad = a2.text_input("Pad / cluster")
            target_type = a1.selectbox("Item type", TARGET_TYPE_OPTIONS)
            priority = a2.selectbox("Priority", PRIORITY_OPTIONS, index=2)
            custom_item_type = st.text_input(
                "Custom item type (optional)",
                value="",
                placeholder="Enter any event type not listed above",
            )
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
            document_name = a2.text_input(
                "Document name",
                value="",
                placeholder="SoR, WCS, ToR, or custom name",
            )
            document_file = st.text_input(
                "Document PDF filename",
                value="",
                placeholder="NWD-001_SoR.pdf",
            )
            notes = st.text_area("Notes")
            submitted = st.form_submit_button("Add item", type="primary", use_container_width=True)

        if submitted:
            new_id = target_id.strip() or next_target_id(master_df["Target_ID"].tolist())
            new_row = {
                "Target_ID": new_id,
                "Well_Name": well_name.strip() or f"New-Item-{len(master_df)+1}",
                "Reservoir": reservoir,
                "Area": area,
                "Pad": pad,
                "Rig": rig,
                "Target_Type": custom_item_type.strip() or target_type,
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
                "Document_Name": document_name,
                "Document_File": document_file,
                "Notes": notes,
            }
            set_master(pd.concat([master_df, pd.DataFrame([new_row])], ignore_index=True), add_undo=True)
            st.session_state.source_name = "Edited in NWD Scheduler"
            st.success(f"Added {new_id}")
            st.rerun()

    with right:
        st.subheader("Selected-item actions")
        selected_ids = st.multiselect("Choose item(s)", master_df["Target_ID"].tolist(), help="Actions below apply to these schedule items.")

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
            st.success(f"Deleted {len(selected_ids)} item(s)")
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
        preserve_duration = st.checkbox("Preserve each item's duration", value=True)
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
        bulk_color = st.color_picker("Pick bar color", value=DEFAULT_COLOR)
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
        ### NWD Scheduler v1.7.0 workflow
        1. The application opens with an **empty schedule**.
        2. Upload an Excel/CSV schedule or add items manually.
        3. Choose visible columns under **Table & tooltip columns**. Hidden table
           columns are automatically excluded from Gantt mouse-over text.
        4. Mark `Highlight = True` for New Technology or another special item.
           It receives a purple dashed Gantt border.
        5. Apply edits and download **Updated input** to preserve the complete
           amended master schedule.
        6. Put linked PDFs in the `input` folder beside the Python file, upload PDFs
           temporarily from the sidebar, and populate `Document_File` when possible.
        7. Set `Document_Name` to any user-defined label, such as **SoR**, **WCS**,
           **ToR**, or another document name. Click a Gantt item to open its PDF.
           The item may be a well target or any other schedule event. The mouse pointer
           changes to a hand over clickable chart items.
        8. Use **Save / load view template** to download a portable JSON file.
           It contains the complete current schedule plus all filters and chart settings.
        9. The active schedule is linked to the page URL through a server snapshot,
           so browser refresh restores the Gantt instead of opening empty.
        10. Use **Reset defaults** to clear filters and chart settings while keeping
            the current schedule loaded.

        ### Linked-document filename matching
        The most reliable method is to populate `Document_File`, for example
        `R-1501_SoR.pdf`, and enter the preferred display label in `Document_Name`.
        If `Document_File` is blank, the app searches PDF filenames using the item ID
        and well/event name. Existing inputs containing `SoR_File` remain compatible
        and are automatically treated as `Document_File` with the name **SoR**.

        ### Colors and legend
        The Gantt always shows a **Color legend** underneath the chart. In custom-color
        mode, the legend groups the visible item types by the colors actually used.
        Use the visual color picker below the editable table for one item, or use
        **Bulk Actions** for several items; typing hex codes is not required.

        ### Usage panel
        The **Usage** panel shows active users, total known users and opened sessions.
        If Streamlit authentication is configured, it uses the signed-in identity;
        otherwise each user can enter a display name. The lightweight database is local
        to the running app server, so redeployment or an ephemeral cloud restart can
        clear the history.

        ### Streamlit Cloud
        A browser cannot automatically reopen the original CSV path on your
        computer. Therefore the JSON template embeds the current CSV data itself.

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
st.caption("NWD Scheduler v1.7.0 • Generic linked PDFs • Zoom-aware tooltips • Color legend • Usage visibility")
