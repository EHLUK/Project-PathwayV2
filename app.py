"""
P6 XER Project Manager Planning Tool
=====================================
A Streamlit app for interrogating Primavera P6 XER schedules
without needing to open P6. Designed for Project Managers.
"""

import io
import re
import math
import warnings
from collections import defaultdict, deque
from datetime import datetime, timedelta

import networkx as nx
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# -----------------------------------------------------------------------------
# PAGE CONFIG
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="P6 Planner Tool",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------------------------------------------------------
# CUSTOM CSS
# -----------------------------------------------------------------------------
st.markdown("""
<style>
    [data-testid="stSidebar"] { background-color: #1a2332; }
    [data-testid="stSidebar"] * { color: #e0e6f0 !important; }
    [data-testid="stSidebar"] .stSelectbox label { color: #a0b0c8 !important; }
    .metric-card {
        background: #f0f4f8; border-radius: 8px; padding: 16px;
        border-left: 4px solid #2563eb; margin: 4px 0;
    }
    .critical-badge {
        background: #dc2626; color: white; padding: 2px 8px;
        border-radius: 12px; font-size: 12px; font-weight: bold;
    }
    .near-critical-badge {
        background: #f59e0b; color: white; padding: 2px 8px;
        border-radius: 12px; font-size: 12px; font-weight: bold;
    }
    .ok-badge {
        background: #16a34a; color: white; padding: 2px 8px;
        border-radius: 12px; font-size: 12px; font-weight: bold;
    }
    .warn-box {
        background: #fffbeb; border: 1px solid #f59e0b; border-radius: 8px;
        padding: 12px; margin: 8px 0;
    }
    .info-box {
        background: #eff6ff; border: 1px solid #3b82f6; border-radius: 8px;
        padding: 12px; margin: 8px 0;
    }
    h1, h2, h3 { color: #1e3a5f; }
    .stDataFrame { border-radius: 8px; }
    div[data-testid="metric-container"] {
        background: #f8fafc; border-radius: 8px; padding: 12px;
        border: 1px solid #e2e8f0;
    }
</style>
""", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# XER PARSING  (xerparser + manual fallback)
# -----------------------------------------------------------------------------

def parse_xer_fallback(raw_text: str) -> dict:
    """
    Manual fallback parser that reads XER table format:
    %T TABLE_NAME  /  %F col1 col2 ...  /  %R val1 val2 ...
    Returns dict of {table_name: list_of_dicts}
    """
    tables = {}
    current_table = None
    current_fields = []

    for line in raw_text.splitlines():
        line = line.rstrip("\r")
        if line.startswith("%T\t"):
            current_table = line[3:].strip()
            current_fields = []
            tables[current_table] = []
        elif line.startswith("%F\t") and current_table:
            current_fields = line[3:].split("\t")
        elif line.startswith("%R\t") and current_table and current_fields:
            values = line[3:].split("\t")
            # Pad values if shorter than fields
            while len(values) < len(current_fields):
                values.append("")
            row = {current_fields[i]: values[i] for i in range(len(current_fields))}
            tables[current_table].append(row)

    return tables


def hours_to_days(hours, hours_per_day=8.0):
    """Convert hours to working days."""
    if hours is None:
        return None
    try:
        return round(float(hours) / hours_per_day, 1)
    except (TypeError, ValueError):
        return None


def safe_float(val, default=None):
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def safe_date(val):
    if val is None or str(val).strip() in ("", "None"):
        return None
    if isinstance(val, datetime):
        return val
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(val).strip(), fmt)
        except ValueError:
            pass
    return None


def parse_xer(file_bytes: bytes):
    """
    Parse an XER file. Uses xerparser library first; falls back to manual parsing.
    Returns a dict with keys: tasks_df, relationships_df, wbs_df, resources_df,
    task_resources_df, project_info, calendars_df, parse_method
    """
    # Try to decode the file
    for codec in ("cp1252", "utf-8", "latin-1"):
        try:
            raw_text = file_bytes.decode(codec)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("Cannot decode XER file. Please check the file encoding.")

    result = {
        "tasks_df": pd.DataFrame(),
        "relationships_df": pd.DataFrame(),
        "wbs_df": pd.DataFrame(),
        "resources_df": pd.DataFrame(),
        "task_resources_df": pd.DataFrame(),
        "project_info": {},
        "calendars_df": pd.DataFrame(),
        "parse_method": "unknown",
    }

    # -- Try xerparser library -------------------------------------------------
    try:
        from xerparser.src.xer import Xer
        xer = Xer(raw_text)

        # Project info
        proj = None
        if xer.projects:
            proj_id = next(iter(xer.projects))
            proj = xer.projects[proj_id]
            result["project_info"] = {
                "name": getattr(proj, "name", ""),
                "data_date": getattr(proj, "last_recalc_date", None),
                "project_id": proj_id,
                "plan_start": getattr(proj, "plan_start_date", None),
                "scd_end": getattr(proj, "scd_end_date", None),
            }

        # Tasks DataFrame
        rows = []
        for uid, task in xer.tasks.items():
            tf = task.total_float_hr_cnt
            ff = task.free_float_hr_cnt
            # Effective start/finish (actual if done, early if not)
            eff_start = task.act_start_date or task.early_start_date or task.target_start_date
            eff_finish = task.act_end_date or task.early_end_date or task.target_end_date

            # WBS path
            wbs_node = xer.wbs_nodes.get(task.wbs_id)
            wbs_path = ""
            if wbs_node:
                parts = []
                n = wbs_node
                while n:
                    parts.append(getattr(n, "name", ""))
                    n = getattr(n, "parent", None)
                wbs_path = " > ".join(reversed(parts))

            # Calendar name
            cal = xer.calendars.get(task.clndr_id)
            cal_name = getattr(cal, "name", "") if cal else ""

            rows.append({
                "task_id": uid,
                "task_code": task.task_code,
                "task_name": task.name,
                "wbs_id": task.wbs_id,
                "wbs_path": wbs_path,
                "status": task.status.value if task.status else "",
                "task_type": task.type.value if task.type else "",
                "calendar": cal_name,
                "early_start": task.early_start_date,
                "early_finish": task.early_end_date,
                "late_start": task.late_start_date,
                "late_finish": task.late_end_date,
                "act_start": task.act_start_date,
                "act_finish": task.act_end_date,
                "target_start": task.target_start_date,
                "target_finish": task.target_end_date,
                "eff_start": eff_start,
                "eff_finish": eff_finish,
                "orig_dur_days": hours_to_days(task.target_drtn_hr_cnt),
                "rem_dur_days": hours_to_days(task.remain_drtn_hr_cnt),
                "total_float_days": hours_to_days(tf),
                "free_float_days": hours_to_days(ff),
                "total_float_hrs": tf,
                "is_longest_path": task.is_longest_path,
                "cstr_type": task.cstr_type,
                "cstr_date": task.cstr_date,
                "cstr_type2": task.cstr_type2,
                "cstr_date2": task.cstr_date2,
                "phys_pct": round(task.phys_complete_pct * 100, 1),
                "float_path": task.float_path,
            })

        result["tasks_df"] = pd.DataFrame(rows)

        # Relationships DataFrame
        rel_rows = []
        for uid, rel in xer.relationships.items():
            rel_rows.append({
                "pred_id": uid,
                "pred_task_id": rel.predecessor.uid if rel.predecessor else "",
                "pred_task_code": rel.predecessor.task_code if rel.predecessor else "",
                "pred_task_name": rel.predecessor.name if rel.predecessor else "",
                "succ_task_id": rel.successor.uid if rel.successor else "",
                "succ_task_code": rel.successor.task_code if rel.successor else "",
                "succ_task_name": rel.successor.name if rel.successor else "",
                "rel_type": rel.link,
                "lag_days": rel.lag,
                "lag_hrs": rel.lag_hr_cnt,
            })
        result["relationships_df"] = pd.DataFrame(rel_rows)

        # WBS DataFrame
        wbs_rows = []
        for uid, wbs in xer.wbs_nodes.items():
            wbs_rows.append({
                "wbs_id": uid,
                "wbs_code": getattr(wbs, "short_name", ""),
                "wbs_name": getattr(wbs, "name", ""),
                "parent_wbs_id": getattr(wbs, "parent_wbs_id", ""),
                "proj_id": getattr(wbs, "proj_id", ""),
            })
        result["wbs_df"] = pd.DataFrame(wbs_rows)

        # Resources & task resources
        if xer.resources:
            res_rows = []
            for uid, r in xer.resources.items():
                res_rows.append({
                    "rsrc_id": uid,
                    "rsrc_name": getattr(r, "name", ""),
                    "rsrc_short": getattr(r, "rsrc_short_name", ""),
                    "rsrc_type": getattr(r, "rsrc_type", ""),
                })
            result["resources_df"] = pd.DataFrame(res_rows)

        # Task resources (loading)
        taskrsrc_rows = []
        for uid, task in xer.tasks.items():
            for tr in getattr(task, "resources", []):
                taskrsrc_rows.append({
                    "task_id": uid,
                    "task_code": task.task_code,
                    "rsrc_id": getattr(tr, "rsrc_id", ""),
                    "target_qty": safe_float(getattr(tr, "target_qty", 0), 0),
                    "remain_qty": safe_float(getattr(tr, "remain_qty", 0), 0),
                    "act_reg_qty": safe_float(getattr(tr, "act_reg_qty", 0), 0),
                    "target_start": safe_date(getattr(tr, "target_start_date", None)),
                    "target_finish": safe_date(getattr(tr, "target_end_date", None)),
                })
        result["task_resources_df"] = pd.DataFrame(taskrsrc_rows)

        result["parse_method"] = "xerparser"
        return result

    except Exception as e:
        st.warning(f"xerparser failed ({e}), using fallback parser...")

    # -- Manual fallback -------------------------------------------------------
    try:
        tables = parse_xer_fallback(raw_text)
        return _build_from_raw_tables(tables)
    except Exception as e2:
        raise ValueError(f"Both parsers failed. Last error: {e2}")


def _build_from_raw_tables(tables: dict) -> dict:
    """Build result dict from raw parsed tables (fallback)."""
    result = {
        "tasks_df": pd.DataFrame(),
        "relationships_df": pd.DataFrame(),
        "wbs_df": pd.DataFrame(),
        "resources_df": pd.DataFrame(),
        "task_resources_df": pd.DataFrame(),
        "project_info": {},
        "calendars_df": pd.DataFrame(),
        "parse_method": "manual_fallback",
    }

    # Project info
    if "PROJECT" in tables and tables["PROJECT"]:
        proj = tables["PROJECT"][0]
        result["project_info"] = {
            "name": proj.get("proj_short_name", proj.get("proj_id", "")),
            "data_date": safe_date(proj.get("last_recalc_date")),
            "plan_start": safe_date(proj.get("plan_start_date")),
            "scd_end": safe_date(proj.get("scd_end_date")),
        }

    # Tasks
    if "TASK" in tables:
        df = pd.DataFrame(tables["TASK"])
        # Normalise date columns
        for col in ["early_start_date", "early_end_date", "late_start_date",
                    "late_end_date", "act_start_date", "act_end_date",
                    "target_start_date", "target_end_date", "cstr_date", "cstr_date2"]:
            if col in df.columns:
                df[col] = df[col].apply(safe_date)
        # Float
        for col in ["total_float_hr_cnt", "free_float_hr_cnt",
                    "target_drtn_hr_cnt", "remain_drtn_hr_cnt"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")

        # Build normalised columns
        df["eff_start"] = df.get("act_start_date", df.get("early_start_date"))
        df["eff_finish"] = df.get("act_end_date", df.get("early_end_date"))
        df["total_float_days"] = df.get("total_float_hr_cnt", pd.Series(dtype=float)).apply(hours_to_days)
        df["free_float_days"] = df.get("free_float_hr_cnt", pd.Series(dtype=float)).apply(hours_to_days)
        df["orig_dur_days"] = df.get("target_drtn_hr_cnt", pd.Series(dtype=float)).apply(hours_to_days)
        df["rem_dur_days"] = df.get("remain_drtn_hr_cnt", pd.Series(dtype=float)).apply(hours_to_days)

        # Rename for consistency
        rename = {
            "task_id": "task_id", "task_code": "task_code",
            "task_name": "task_name", "wbs_id": "wbs_id",
            "status_code": "status", "task_type": "task_type",
            "early_start_date": "early_start", "early_end_date": "early_finish",
            "late_start_date": "late_start", "late_end_date": "late_finish",
            "act_start_date": "act_start", "act_end_date": "act_finish",
            "target_start_date": "target_start", "target_end_date": "target_finish",
            "cstr_type": "cstr_type", "cstr_date": "cstr_date",
            "cstr_type2": "cstr_type2", "cstr_date2": "cstr_date2",
            "driving_path_flag": "is_longest_path_flag",
            "phys_complete_pct": "phys_pct",
        }
        df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
        if "is_longest_path_flag" in df.columns:
            df["is_longest_path"] = df["is_longest_path_flag"] == "Y"
        df["wbs_path"] = df.get("wbs_id", "")
        result["tasks_df"] = df

    # Relationships
    if "TASKPRED" in tables:
        df = pd.DataFrame(tables["TASKPRED"])
        if "lag_hr_cnt" in df.columns:
            df["lag_days"] = pd.to_numeric(df["lag_hr_cnt"], errors="coerce").apply(hours_to_days)
        rename_r = {"pred_type": "rel_type", "task_id": "succ_task_id",
                    "pred_task_id": "pred_task_id"}
        df = df.rename(columns={k: v for k, v in rename_r.items() if k in df.columns})
        result["relationships_df"] = df

    # WBS
    if "PROJWBS" in tables:
        df = pd.DataFrame(tables["PROJWBS"])
        rename_w = {"wbs_id": "wbs_id", "wbs_short_name": "wbs_code",
                    "wbs_name": "wbs_name", "parent_wbs_id": "parent_wbs_id"}
        df = df.rename(columns={k: v for k, v in rename_w.items() if k in df.columns})
        result["wbs_df"] = df

    # Resources
    if "RSRC" in tables:
        df = pd.DataFrame(tables["RSRC"])
        rename_rs = {"rsrc_id": "rsrc_id", "rsrc_name": "rsrc_name",
                     "rsrc_short_name": "rsrc_short"}
        df = df.rename(columns={k: v for k, v in rename_rs.items() if k in df.columns})
        result["resources_df"] = df

    # Task resources
    if "TASKRSRC" in tables:
        df = pd.DataFrame(tables["TASKRSRC"])
        for col in ["target_qty", "remain_qty", "act_reg_qty"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        for col in ["target_start_date", "target_end_date"]:
            if col in df.columns:
                df[col] = df[col].apply(safe_date)
                df = df.rename(columns={col: col.replace("_date", "")})
        result["task_resources_df"] = df

    return result


# -----------------------------------------------------------------------------
# GRAPH BUILDING
# -----------------------------------------------------------------------------

def build_graph(tasks_df: pd.DataFrame, rels_df: pd.DataFrame) -> nx.DiGraph:
    """Build a networkx directed graph from tasks and relationships."""
    G = nx.DiGraph()
    for _, row in tasks_df.iterrows():
        G.add_node(row["task_id"], **row.to_dict())
    for _, row in rels_df.iterrows():
        if row.get("pred_task_id") and row.get("succ_task_id"):
            G.add_edge(
                row["pred_task_id"],
                row["succ_task_id"],
                rel_type=row.get("rel_type", "FS"),
                lag_days=row.get("lag_days", 0),
            )
    return G


# -----------------------------------------------------------------------------
# CRITICAL PATH HELPERS
# -----------------------------------------------------------------------------

def get_critical_threshold(tasks_df: pd.DataFrame, near_crit_days: float = 10.0):
    """Classify activities as critical / near-critical / float."""
    df = tasks_df.copy()
    df["is_critical"] = df["total_float_days"].apply(
        lambda f: f is not None and f <= 0
    )
    df["is_near_critical"] = df["total_float_days"].apply(
        lambda f: f is not None and 0 < f <= near_crit_days
    )
    return df


def float_status_badge(f):
    if f is None:
        return "-"
    elif f <= 0:
        return "🔴 Critical"
    elif f <= 10:
        return "🟡 Near-Critical"
    else:
        return "🟢 Float"


# -----------------------------------------------------------------------------
# LOGIC TRACE HELPERS
# -----------------------------------------------------------------------------

def trace_predecessors(G: nx.DiGraph, task_id: str, max_depth=100) -> list:
    """BFS backwards through predecessors. Returns list of (task_id, depth)."""
    visited = {}
    queue = deque([(task_id, 0)])
    result = []
    while queue:
        node, depth = queue.popleft()
        if node in visited or depth > max_depth:
            continue
        visited[node] = depth
        if node != task_id:
            result.append((node, depth))
        for pred in G.predecessors(node):
            if pred not in visited:
                queue.append((pred, depth + 1))
    return result


def trace_successors(G: nx.DiGraph, task_id: str, max_depth=100) -> list:
    """BFS forwards through successors."""
    visited = {}
    queue = deque([(task_id, 0)])
    result = []
    while queue:
        node, depth = queue.popleft()
        if node in visited or depth > max_depth:
            continue
        visited[node] = depth
        if node != task_id:
            result.append((node, depth))
        for succ in G.successors(node):
            if succ not in visited:
                queue.append((succ, depth + 1))
    return result


def driving_path_to_activity(G: nx.DiGraph, tasks_df: pd.DataFrame, target_id: str):
    """
    Identify the most likely driving predecessor chain into a target activity.
    Selects path through predecessors with lowest total float at each step.
    Returns ordered list of task_ids from chain start to target.
    """
    task_lookup = tasks_df.set_index("task_id").to_dict("index") if not tasks_df.empty else {}

    def get_float(tid):
        t = task_lookup.get(tid, {})
        f = t.get("total_float_days")
        return f if f is not None else 9999

    path = [target_id]
    visited = {target_id}
    current = target_id

    for _ in range(200):
        preds = list(G.predecessors(current))
        if not preds:
            break
        # Pick predecessor with lowest float (most critical)
        preds_sorted = sorted(preds, key=get_float)
        best = preds_sorted[0]
        if best in visited:
            break
        path.insert(0, best)
        visited.add(best)
        current = best

    return path


# -----------------------------------------------------------------------------
# EXPORT HELPERS
# -----------------------------------------------------------------------------

def style_header_row(ws, row_idx, fill_color="1e3a5f", font_color="FFFFFF"):
    fill = PatternFill("solid", start_color=fill_color, fgColor=fill_color)
    font = Font(bold=True, color=font_color)
    for cell in ws[row_idx]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")


def df_to_sheet(ws, df, sheet_title=None):
    """Write a DataFrame to an openpyxl worksheet with formatting."""
    if sheet_title:
        ws.title = sheet_title[:31]
    ws.append(list(df.columns))
    style_header_row(ws, 1)
    for r in df.itertuples(index=False):
        ws.append(list(r))
    # Auto-width
    for col_cells in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col_cells), default=10)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max_len + 4, 50)


def export_df_to_excel(sheets: dict) -> bytes:
    """sheets = {sheet_name: dataframe}. Returns Excel bytes."""
    wb = Workbook()
    first = True
    for name, df in sheets.items():
        if first:
            ws = wb.active
            first = False
        else:
            ws = wb.create_sheet()
        df_to_sheet(ws, df, name)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def format_date(d):
    if d is None:
        return "-"
    try:
        return d.strftime("%d %b %Y")
    except Exception:
        return str(d)


# -----------------------------------------------------------------------------
# PAGE: PROJECT SUMMARY
# -----------------------------------------------------------------------------

def page_project_summary(data: dict, near_crit_days: float):
    st.title("📊 Project Summary")

    proj = data["project_info"]
    tasks = data["tasks_df"]
    rels = data["relationships_df"]

    if tasks.empty:
        st.warning("No activities found in this file.")
        return

    tasks = get_critical_threshold(tasks, near_crit_days)

    # Header metrics
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Project Name", proj.get("name", "Unknown"))
    c2.metric("Data Date", format_date(proj.get("data_date")))
    c3.metric("Parse Method", data.get("parse_method", "-"))
    c4.metric("Activities", len(tasks))

    st.divider()
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("🔴 Critical", int(tasks["is_critical"].sum()))
    c2.metric("🟡 Near-Critical", int(tasks["is_near_critical"].sum()))
    neg_float = tasks["total_float_days"].apply(lambda f: f is not None and f < 0).sum()
    c3.metric("⚠️ Negative Float", int(neg_float))
    c4.metric("🔗 Relationships", len(rels))

    # Open-ended
    if not rels.empty and "pred_task_id" in rels.columns:
        tasks_with_pred = set(rels["succ_task_id"].dropna())
        tasks_with_succ = set(rels["pred_task_id"].dropna())
        task_ids = set(tasks["task_id"])
        no_pred = len(task_ids - tasks_with_pred)
        no_succ = len(task_ids - tasks_with_succ)
        c5.metric("Open-Ended Activities", no_pred + no_succ)
    else:
        c5.metric("Open-Ended Activities", "-")

    # Date range
    valid_starts = tasks["eff_start"].dropna()
    valid_finishes = tasks["eff_finish"].dropna()
    if not valid_starts.empty and not valid_finishes.empty:
        earliest = min(valid_starts)
        latest = max(valid_finishes)
        st.info(f"**Schedule Span:** {format_date(earliest)} -> {format_date(latest)}")

    # Constraint count
    constrained = tasks["cstr_type"].apply(lambda x: bool(x) and str(x).strip() not in ("", "None")).sum() if "cstr_type" in tasks.columns else 0
    st.info(f"**Constrained Activities:** {int(constrained)}")

    st.divider()

    # Charts
    tab1, tab2, tab3 = st.tabs(["Float Distribution", "Activities by WBS", "Status Breakdown"])

    with tab1:
        float_vals = tasks["total_float_days"].dropna()
        if not float_vals.empty:
            fig = px.histogram(
                float_vals, nbins=40, title="Total Float Distribution (Days)",
                labels={"value": "Float (days)", "count": "Activities"},
                color_discrete_sequence=["#2563eb"],
            )
            fig.add_vline(x=0, line_dash="dash", line_color="red", annotation_text="Critical")
            fig.add_vline(x=near_crit_days, line_dash="dot", line_color="orange",
                          annotation_text=f"Near-Critical ({near_crit_days}d)")
            st.plotly_chart(fig, use_container_width=True)

    with tab2:
        if "wbs_path" in tasks.columns:
            # Show top-level WBS only
            tasks["wbs_top"] = tasks["wbs_path"].apply(
                lambda x: str(x).split(" > ")[0] if pd.notna(x) and x else "Unknown"
            )
            wbs_counts = tasks.groupby("wbs_top").size().reset_index(name="count")
            wbs_counts = wbs_counts.sort_values("count", ascending=False).head(20)
            fig = px.bar(wbs_counts, x="count", y="wbs_top", orientation="h",
                         title="Activities by Top-Level WBS",
                         color_discrete_sequence=["#1e3a5f"])
            fig.update_layout(yaxis_title="", xaxis_title="Activity Count")
            st.plotly_chart(fig, use_container_width=True)

    with tab3:
        if "status" in tasks.columns:
            status_counts = tasks["status"].value_counts().reset_index()
            status_counts.columns = ["Status", "Count"]
            fig = px.pie(status_counts, values="Count", names="Status",
                         title="Activity Status Breakdown",
                         color_discrete_sequence=px.colors.qualitative.Set2)
            st.plotly_chart(fig, use_container_width=True)

    # Summary table
    st.subheader("Activity Summary Table")
    display_cols = ["task_code", "task_name", "wbs_path", "eff_start", "eff_finish",
                    "total_float_days", "status", "is_critical"]
    avail = [c for c in display_cols if c in tasks.columns]
    st.dataframe(tasks[avail].head(100), use_container_width=True)


# -----------------------------------------------------------------------------
# PAGE: ACTIVITY SEARCH
# -----------------------------------------------------------------------------

def page_activity_search(data: dict, near_crit_days: float):
    st.title("🔍 Activity Search")

    tasks = data["tasks_df"]
    rels = data["relationships_df"]

    if tasks.empty:
        st.warning("No activities loaded.")
        return

    tasks = get_critical_threshold(tasks, near_crit_days)

    # -- Filters ---------------------------------------------------------------
    with st.expander("🔎 Search & Filter", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            search_code = st.text_input("Activity ID (partial match)")
            search_name = st.text_input("Activity Name (partial match)")
        with col2:
            search_wbs = st.text_input("WBS (partial match)")
            crit_filter = st.selectbox(
                "Float Status",
                ["All", "Critical (<=0d)", "Near-Critical", "Float >0", "Negative Float"]
            )

        # Date range
        valid_dates = tasks["eff_start"].dropna()
        if not valid_dates.empty:
            min_d = min(valid_dates).date()
            max_d = max(tasks["eff_finish"].dropna()).date()
            d1, d2 = st.columns(2)
            date_from = d1.date_input("Start From", value=min_d, min_value=min_d, max_value=max_d)
            date_to = d2.date_input("Start To", value=max_d, min_value=min_d, max_value=max_d)
        else:
            date_from = date_to = None

        status_opts = ["All"] + sorted(tasks["status"].dropna().unique().tolist()) if "status" in tasks.columns else ["All"]
        status_filter = st.selectbox("Status", status_opts)

    # Apply filters
    filtered = tasks.copy()
    if search_code:
        filtered = filtered[filtered["task_code"].str.contains(search_code, case=False, na=False)]
    if search_name:
        filtered = filtered[filtered["task_name"].str.contains(search_name, case=False, na=False)]
    if search_wbs:
        filtered = filtered[filtered["wbs_path"].str.contains(search_wbs, case=False, na=False)]
    if crit_filter == "Critical (<=0d)":
        filtered = filtered[filtered["is_critical"]]
    elif crit_filter == "Near-Critical":
        filtered = filtered[filtered["is_near_critical"]]
    elif crit_filter == "Float >0":
        filtered = filtered[filtered["total_float_days"].apply(lambda f: f is not None and f > 0)]
    elif crit_filter == "Negative Float":
        filtered = filtered[filtered["total_float_days"].apply(lambda f: f is not None and f < 0)]
    if status_filter != "All" and "status" in filtered.columns:
        filtered = filtered[filtered["status"] == status_filter]
    if date_from and "eff_start" in filtered.columns:
        filtered = filtered[
            filtered["eff_start"].apply(
                lambda d: d is not None and d.date() >= date_from
            )
        ]
    if date_to and "eff_finish" in filtered.columns:
        filtered = filtered[
            filtered["eff_finish"].apply(
                lambda d: d is not None and d.date() <= date_to
            )
        ]

    st.caption(f"Showing {len(filtered)} of {len(tasks)} activities")

    # Show table
    display_cols = ["task_code", "task_name", "wbs_path", "eff_start", "eff_finish",
                    "orig_dur_days", "rem_dur_days", "total_float_days", "free_float_days",
                    "status", "task_type", "is_critical"]
    avail = [c for c in display_cols if c in filtered.columns]
    st.dataframe(filtered[avail], use_container_width=True, height=300)

    # Activity detail
    st.divider()
    st.subheader("Activity Detail")

    if not filtered.empty:
        act_options = filtered.apply(
            lambda r: f"{r.get('task_code','?')} - {r.get('task_name','?')}", axis=1
        ).tolist()
        selected_str = st.selectbox("Select Activity", act_options)
        sel_idx = act_options.index(selected_str)
        row = filtered.iloc[sel_idx]

        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"**Activity ID:** `{row.get('task_code','-')}`")
            st.markdown(f"**Name:** {row.get('task_name','-')}")
            st.markdown(f"**WBS:** {row.get('wbs_path','-')}")
            st.markdown(f"**Type:** {row.get('task_type','-')}")
            st.markdown(f"**Calendar:** {row.get('calendar','-')}")
            st.markdown(f"**Status:** {row.get('status','-')}")
            st.markdown(f"**% Complete:** {row.get('phys_pct','-')}")
        with c2:
            st.markdown(f"**Early Start:** {format_date(row.get('early_start'))}")
            st.markdown(f"**Early Finish:** {format_date(row.get('early_finish'))}")
            st.markdown(f"**Late Start:** {format_date(row.get('late_start'))}")
            st.markdown(f"**Late Finish:** {format_date(row.get('late_finish'))}")
            st.markdown(f"**Orig Duration:** {row.get('orig_dur_days','-')} days")
            st.markdown(f"**Rem Duration:** {row.get('rem_dur_days','-')} days")
            tf = row.get("total_float_days")
            ff = row.get("free_float_days")
            st.markdown(f"**Total Float:** {tf} days  {float_status_badge(tf)}")
            st.markdown(f"**Free Float:** {ff} days")
            cstr = row.get("cstr_type","")
            if cstr and str(cstr).strip() not in ("", "None"):
                st.markdown(f"**Constraint:** {cstr} on {format_date(row.get('cstr_date'))}")

        # Predecessors / Successors
        if not rels.empty:
            task_id = row["task_id"]
            preds = rels[rels["succ_task_id"] == task_id] if "succ_task_id" in rels.columns else pd.DataFrame()
            succs = rels[rels["pred_task_id"] == task_id] if "pred_task_id" in rels.columns else pd.DataFrame()

            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("**Predecessors**")
                if not preds.empty:
                    disp_cols = [c for c in ["pred_task_code", "pred_task_name", "rel_type", "lag_days"] if c in preds.columns]
                    st.dataframe(preds[disp_cols], use_container_width=True)
                else:
                    st.info("No predecessors.")
            with col_b:
                st.markdown("**Successors**")
                if not succs.empty:
                    disp_cols = [c for c in ["succ_task_code", "succ_task_name", "rel_type", "lag_days"] if c in succs.columns]
                    st.dataframe(succs[disp_cols], use_container_width=True)
                else:
                    st.info("No successors.")


# -----------------------------------------------------------------------------
# PAGE: LOGIC TRACE
# -----------------------------------------------------------------------------

def page_logic_trace(data: dict, near_crit_days: float):
    st.title("🔗 Logic Trace")
    st.markdown("> Trace predecessors and successors through the schedule network.")

    tasks = data["tasks_df"]
    rels = data["relationships_df"]

    if tasks.empty or rels.empty:
        st.warning("Tasks or relationships not available.")
        return

    tasks = get_critical_threshold(tasks, near_crit_days)
    G = build_graph(tasks, rels)
    task_lookup = tasks.set_index("task_id").to_dict("index")

    # Activity selector
    act_options = tasks.apply(
        lambda r: f"{r.get('task_code','?')} - {r.get('task_name','?')}", axis=1
    ).tolist()
    selected_str = st.selectbox("Select Activity to Trace", act_options)
    sel_idx = act_options.index(selected_str)
    selected_row = tasks.iloc[sel_idx]
    selected_id = selected_row["task_id"]

    col1, col2, col3 = st.columns(3)
    show_dir_preds = col1.button("◀ Direct Predecessors")
    show_dir_succs = col2.button("▶ Direct Successors")
    show_all_preds = col1.button("◀◀ All Predecessors")
    show_all_succs = col2.button("▶▶ All Successors")
    show_full = col3.button("⚡ Full Logic Chain")

    def build_trace_df(task_ids_depths: list, direction="pred") -> pd.DataFrame:
        rows = []
        for tid, depth in task_ids_depths:
            t = task_lookup.get(tid, {})
            # Get relationship info
            if direction == "pred":
                rel_row = rels[(rels["succ_task_id"] == tid) | (rels["pred_task_id"] == tid)]
            else:
                rel_row = rels[(rels["pred_task_id"] == tid)]
            rel_type = rel_row["rel_type"].iloc[0] if not rel_row.empty and "rel_type" in rel_row.columns else "-"
            lag = rel_row["lag_days"].iloc[0] if not rel_row.empty and "lag_days" in rel_row.columns else 0
            tf = t.get("total_float_days")
            rows.append({
                "Level": depth,
                "Activity ID": t.get("task_code", tid),
                "Activity Name": t.get("task_name", ""),
                "Rel Type": rel_type,
                "Lag (days)": lag,
                "Start": format_date(t.get("eff_start")),
                "Finish": format_date(t.get("eff_finish")),
                "Total Float": tf,
                "Critical": "🔴" if tf is not None and tf <= 0 else "🟢",
            })
        return pd.DataFrame(rows)

    result_df = pd.DataFrame()

    if show_dir_preds:
        preds = [(p, 1) for p in G.predecessors(selected_id)]
        result_df = build_trace_df(preds, "pred")
        st.session_state["trace_label"] = "Direct Predecessors"

    elif show_dir_succs:
        succs = [(s, 1) for s in G.successors(selected_id)]
        result_df = build_trace_df(succs, "succ")
        st.session_state["trace_label"] = "Direct Successors"

    elif show_all_preds:
        all_preds = trace_predecessors(G, selected_id)
        result_df = build_trace_df(all_preds, "pred")
        st.session_state["trace_label"] = "All Predecessors"

    elif show_all_succs:
        all_succs = trace_successors(G, selected_id)
        result_df = build_trace_df(all_succs, "succ")
        st.session_state["trace_label"] = "All Successors"

    elif show_full:
        all_preds = trace_predecessors(G, selected_id)
        all_succs = trace_successors(G, selected_id)
        combined = [(tid, -depth) for tid, depth in all_preds] + [(tid, depth) for tid, depth in all_succs]
        result_df = build_trace_df(combined, "pred")
        st.session_state["trace_label"] = "Full Logic Chain"

    # Retrieve from session if already set
    if not result_df.empty:
        st.session_state["trace_df"] = result_df

    if "trace_df" in st.session_state and not st.session_state["trace_df"].empty:
        label = st.session_state.get("trace_label", "Trace Results")
        st.subheader(f"{label} -- {len(st.session_state['trace_df'])} activities")
        st.dataframe(st.session_state["trace_df"], use_container_width=True)

        # Export
        xls = export_df_to_excel({"Logic Trace": st.session_state["trace_df"]})
        st.download_button(
            "📥 Export Logic Trace to Excel",
            data=xls,
            file_name=f"logic_trace_{selected_row.get('task_code','act')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# -----------------------------------------------------------------------------
# PAGE: CRITICAL PATH ANALYSIS
# -----------------------------------------------------------------------------

def page_critical_path(data: dict, near_crit_days: float):
    st.title("🚨 Critical Path Analysis")

    tasks = data["tasks_df"]
    rels = data["relationships_df"]

    if tasks.empty:
        st.warning("No activities loaded.")
        return

    tasks = get_critical_threshold(tasks, near_crit_days)

    tab1, tab2, tab3, tab4 = st.tabs(
        ["Critical Activities", "Near-Critical", "Negative Float", "By WBS / Package"]
    )

    with tab1:
        critical = tasks[tasks["is_critical"]].sort_values("total_float_days")
        st.metric("Critical Activities", len(critical))
        disp = ["task_code", "task_name", "wbs_path", "eff_start", "eff_finish",
                "total_float_days", "status"]
        avail = [c for c in disp if c in critical.columns]
        st.dataframe(critical[avail], use_container_width=True)

        if not critical.empty and "eff_start" in critical.columns:
            st.subheader("Critical Path Gantt")
            gantt_df = critical.dropna(subset=["eff_start", "eff_finish"]).copy()
            gantt_df["Start"] = gantt_df["eff_start"]
            gantt_df["Finish"] = gantt_df["eff_finish"]
            gantt_df["Task"] = gantt_df["task_code"] + " - " + gantt_df["task_name"]
            if len(gantt_df) > 0:
                fig = px.timeline(
                    gantt_df.head(50),
                    x_start="Start", x_end="Finish", y="Task",
                    title="Critical Path Activities (top 50)",
                    color_discrete_sequence=["#dc2626"],
                )
                fig.update_yaxes(autorange="reversed")
                st.plotly_chart(fig, use_container_width=True)

        xls = export_df_to_excel({"Critical Path": critical[avail]})
        st.download_button("📥 Export Critical Path", xls,
                           "critical_path.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab2:
        near_crit = tasks[tasks["is_near_critical"]].sort_values("total_float_days")
        st.metric(f"Near-Critical (0 < float <= {near_crit_days}d)", len(near_crit))
        avail = [c for c in ["task_code","task_name","wbs_path","eff_start","eff_finish",
                              "total_float_days","status"] if c in near_crit.columns]
        st.dataframe(near_crit[avail], use_container_width=True)

    with tab3:
        neg = tasks[tasks["total_float_days"].apply(lambda f: f is not None and f < 0)].sort_values("total_float_days")
        st.metric("Negative Float Activities", len(neg))
        if not neg.empty:
            st.warning("⚠️ Activities with negative float indicate the schedule cannot be met -- investigate immediately.")
            avail = [c for c in ["task_code","task_name","total_float_days","eff_start",
                                  "eff_finish","status"] if c in neg.columns]
            st.dataframe(neg[avail], use_container_width=True)

    with tab4:
        if "wbs_path" not in tasks.columns:
            st.info("WBS data not available.")
            return
        tasks["wbs_top"] = tasks["wbs_path"].apply(
            lambda x: str(x).split(" > ")[0] if pd.notna(x) else "Unknown"
        )
        wbs_crit = tasks.groupby("wbs_top").agg(
            total=("task_id", "count"),
            critical=("is_critical", "sum"),
            near_critical=("is_near_critical", "sum"),
        ).reset_index()
        wbs_crit["crit_%"] = (wbs_crit["critical"] / wbs_crit["total"] * 100).round(1)
        fig = px.bar(
            wbs_crit, x="wbs_top", y=["critical", "near_critical"],
            title="Critical & Near-Critical by WBS",
            labels={"value": "Activities", "wbs_top": "WBS"},
            color_discrete_map={"critical": "#dc2626", "near_critical": "#f59e0b"},
            barmode="group",
        )
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(wbs_crit, use_container_width=True)


# -----------------------------------------------------------------------------
# PAGE: CRITICAL PATH TO SELECTED ACTIVITY
# -----------------------------------------------------------------------------

def page_critical_path_to_activity(data: dict, near_crit_days: float):
    st.title("🎯 Critical Path to Selected Activity")
    st.markdown(
        "> **What is driving this activity?**  \n"
        "> Select a target activity or milestone and this page will identify "
        "the most critical predecessor chain driving it."
    )

    tasks = data["tasks_df"]
    rels = data["relationships_df"]

    if tasks.empty or rels.empty:
        st.warning("Tasks or relationships not available.")
        return

    tasks = get_critical_threshold(tasks, near_crit_days)
    G = build_graph(tasks, rels)
    task_lookup = tasks.set_index("task_id").to_dict("index")

    # Select target
    act_options = tasks.apply(
        lambda r: f"{r.get('task_code','?')} - {r.get('task_name','?')}", axis=1
    ).tolist()
    selected_str = st.selectbox("Select Target Activity / Milestone", act_options)
    sel_idx = act_options.index(selected_str)
    target_row = tasks.iloc[sel_idx]
    target_id = target_row["task_id"]

    if st.button("🔍 Find Driving Path"):
        path = driving_path_to_activity(G, tasks, target_id)

        # Also gather all predecessor branches
        all_preds = trace_predecessors(G, target_id)
        all_pred_ids = [p for p, _ in all_preds]
        all_pred_tasks = tasks[tasks["task_id"].isin(all_pred_ids)].copy()

        st.subheader(f"Driving Path to: {target_row.get('task_code','?')} - {target_row.get('task_name','?')}")

        # Build path table
        path_rows = []
        for i, tid in enumerate(path):
            t = task_lookup.get(tid, {})
            tf = t.get("total_float_days")
            is_target = tid == target_id

            # Get relationship with previous in chain
            rel_type = "-"
            lag = 0
            if i > 0:
                prev_id = path[i-1]
                rel_row = rels[(rels["pred_task_id"] == prev_id) & (rels["succ_task_id"] == tid)]
                if not rel_row.empty:
                    rel_type = rel_row["rel_type"].iloc[0] if "rel_type" in rel_row.columns else "FS"
                    lag = rel_row["lag_days"].iloc[0] if "lag_days" in rel_row.columns else 0

            path_rows.append({
                "Pos": i + 1,
                "Activity ID": t.get("task_code", tid),
                "Activity Name": t.get("task_name", ""),
                "Link": rel_type,
                "Lag (days)": lag,
                "Start": format_date(t.get("eff_start")),
                "Finish": format_date(t.get("eff_finish")),
                "Float (days)": tf,
                "Status": t.get("status", ""),
                "🎯 Target": "✅" if is_target else "",
            })

        path_df = pd.DataFrame(path_rows)
        st.dataframe(path_df, use_container_width=True)

        # Key stats
        c1, c2, c3 = st.columns(3)
        c1.metric("Activities in Driving Chain", len(path))
        chain_tasks = tasks[tasks["task_id"].isin(path)]
        min_float = chain_tasks["total_float_days"].min()
        c2.metric("Lowest Float in Chain", f"{min_float} days" if min_float is not None else "-")
        c3.metric("Total Predecessor Network", len(all_pred_ids))

        # Gantt for driving chain
        gantt_data = chain_tasks.dropna(subset=["eff_start", "eff_finish"]).copy()
        if not gantt_data.empty:
            gantt_data["Task"] = gantt_data["task_code"] + " - " + gantt_data["task_name"]
            gantt_data["Color"] = gantt_data["task_id"].apply(
                lambda tid: "Driving Path" if tid == target_id else "Predecessor"
            )
            fig = px.timeline(
                gantt_data, x_start="eff_start", x_end="eff_finish", y="Task",
                color="Color",
                color_discrete_map={"Driving Path": "#dc2626", "Predecessor": "#2563eb"},
                title="Driving Path Gantt",
            )
            fig.update_yaxes(autorange="reversed")
            st.plotly_chart(fig, use_container_width=True)

        # Constraints in chain
        if "cstr_type" in chain_tasks.columns:
            constrained = chain_tasks[chain_tasks["cstr_type"].apply(
                lambda x: bool(x) and str(x).strip() not in ("", "None")
            )]
            if not constrained.empty:
                st.warning(f"⚠️ {len(constrained)} constrained activities in driving chain:")
                st.dataframe(constrained[["task_code","task_name","cstr_type","cstr_date"]], use_container_width=True)

        # All predecessors
        st.divider()
        st.subheader(f"All Predecessor Activities ({len(all_pred_tasks)} total)")
        all_pred_tasks_sorted = all_pred_tasks.sort_values("total_float_days")
        avail = [c for c in ["task_code","task_name","total_float_days","eff_start","eff_finish","status"] if c in all_pred_tasks_sorted.columns]
        st.dataframe(all_pred_tasks_sorted[avail], use_container_width=True)

        # Export
        xls = export_df_to_excel({
            "Driving Path": path_df,
            "All Predecessors": all_pred_tasks_sorted[avail] if not all_pred_tasks_sorted.empty else pd.DataFrame(),
        })
        st.download_button(
            "📥 Export Driving Path Report", xls,
            f"driving_path_{target_row.get('task_code','act')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# -----------------------------------------------------------------------------
# PAGE: LABOUR HISTOGRAM
# -----------------------------------------------------------------------------

def page_labour_histogram(data: dict):
    st.title("👷 Labour Histogram")

    task_res = data["task_resources_df"]
    tasks = data["tasks_df"]
    resources = data["resources_df"]

    if task_res.empty:
        st.markdown("""
        <div class="warn-box">
        ⚠️ <strong>No resource loading found in this XER file.</strong><br>
        This usually means the programme was not resourced in P6, or resource data was not exported.
        <br><br>
        You can upload a separate resource CSV or Excel file below.
        </div>
        """, unsafe_allow_html=True)

        st.subheader("Upload Resource Loading File")
        res_file = st.file_uploader("Upload CSV or Excel (columns: task_code, rsrc_name, target_qty, target_start, target_finish)", type=["csv","xlsx"])
        if res_file:
            try:
                if res_file.name.endswith(".csv"):
                    task_res = pd.read_csv(res_file)
                else:
                    task_res = pd.read_excel(res_file)
                for col in ["target_start","target_finish"]:
                    if col in task_res.columns:
                        task_res[col] = pd.to_datetime(task_res[col], errors="coerce")
                st.success(f"Loaded {len(task_res)} resource rows.")
            except Exception as e:
                st.error(f"Could not read resource file: {e}")
                return
        else:
            return

    # Merge with task info and resource names
    if not tasks.empty and "task_id" in task_res.columns:
        task_res = task_res.merge(
            tasks[["task_id","task_code","task_name","wbs_path","is_critical" if "is_critical" in tasks.columns else "task_id"]].drop_duplicates(),
            on="task_id", how="left", suffixes=("","_task")
        )
    if not resources.empty and "rsrc_id" in task_res.columns:
        task_res = task_res.merge(resources[["rsrc_id","rsrc_name"]], on="rsrc_id", how="left", suffixes=("","_res"))
        if "rsrc_name_res" in task_res.columns:
            task_res["rsrc_name"] = task_res["rsrc_name_res"].fillna(task_res.get("rsrc_name",""))

    # Expand resource loading to weekly intervals
    def expand_to_weeks(df):
        rows = []
        for _, r in df.iterrows():
            s = pd.to_datetime(r.get("target_start") or r.get("target_start_date"))
            e = pd.to_datetime(r.get("target_finish") or r.get("target_end_date"))
            if pd.isna(s) or pd.isna(e) or s > e:
                continue
            qty = safe_float(r.get("target_qty", 0), 0)
            if qty == 0:
                continue
            weeks = max(1, math.ceil((e - s).days / 7))
            qty_per_week = qty / weeks
            current = s
            for _ in range(weeks):
                rows.append({
                    "week": current.to_period("W").start_time,
                    "month": current.to_period("M").start_time,
                    "qty": qty_per_week,
                    "rsrc_name": r.get("rsrc_name","Unknown"),
                    "task_code": r.get("task_code",""),
                    "task_name": r.get("task_name",""),
                    "wbs_path": r.get("wbs_path",""),
                })
                current += timedelta(weeks=1)
        return pd.DataFrame(rows)

    weekly = expand_to_weeks(task_res)

    if weekly.empty:
        st.warning("Could not generate histogram -- resource dates or quantities may be missing.")
        return

    # Filters
    st.sidebar.divider()
    st.sidebar.subheader("Labour Filters")
    all_resources = sorted(weekly["rsrc_name"].unique().tolist())
    sel_res = st.sidebar.multiselect("Resource / Trade", all_resources, default=all_resources)
    if sel_res:
        weekly = weekly[weekly["rsrc_name"].isin(sel_res)]

    # Metrics
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Planned Hours", f"{weekly['qty'].sum():,.0f}")
    weekly_totals = weekly.groupby("week")["qty"].sum()
    c2.metric("Peak Week (hrs)", f"{weekly_totals.max():,.0f}" if not weekly_totals.empty else "-")
    c3.metric("Average Week (hrs)", f"{weekly_totals.mean():,.0f}" if not weekly_totals.empty else "-")

    tab1, tab2, tab3, tab4 = st.tabs(["By Week", "By Month", "By Resource", "By WBS"])

    with tab1:
        weekly_sum = weekly.groupby("week")["qty"].sum().reset_index()
        fig = px.bar(weekly_sum, x="week", y="qty",
                     title="Labour Loading by Week (Hours)",
                     labels={"week":"Week","qty":"Hours"},
                     color_discrete_sequence=["#2563eb"])
        st.plotly_chart(fig, use_container_width=True)

    with tab2:
        monthly_sum = weekly.groupby("month")["qty"].sum().reset_index()
        fig = px.bar(monthly_sum, x="month", y="qty",
                     title="Labour Loading by Month (Hours)",
                     labels={"month":"Month","qty":"Hours"},
                     color_discrete_sequence=["#1e3a5f"])
        st.plotly_chart(fig, use_container_width=True)

    with tab3:
        res_sum = weekly.groupby("rsrc_name")["qty"].sum().reset_index().sort_values("qty", ascending=False)
        fig = px.bar(res_sum, x="rsrc_name", y="qty",
                     title="Total Hours by Resource / Trade",
                     labels={"rsrc_name":"Resource","qty":"Hours"},
                     color_discrete_sequence=["#7c3aed"])
        st.plotly_chart(fig, use_container_width=True)

        # By week and resource stacked
        if len(sel_res) <= 10:
            by_res_week = weekly.groupby(["week","rsrc_name"])["qty"].sum().reset_index()
            fig2 = px.bar(by_res_week, x="week", y="qty", color="rsrc_name",
                          title="Weekly Labour by Resource",
                          labels={"week":"Week","qty":"Hours","rsrc_name":"Resource"})
            st.plotly_chart(fig2, use_container_width=True)

    with tab4:
        if "wbs_path" in weekly.columns:
            weekly["wbs_top"] = weekly["wbs_path"].apply(
                lambda x: str(x).split(" > ")[0] if pd.notna(x) and x else "Unknown"
            )
            wbs_sum = weekly.groupby("wbs_top")["qty"].sum().reset_index().sort_values("qty", ascending=False)
            fig = px.bar(wbs_sum, x="qty", y="wbs_top", orientation="h",
                         title="Total Hours by WBS",
                         color_discrete_sequence=["#059669"])
            st.plotly_chart(fig, use_container_width=True)

    # Export
    xls = export_df_to_excel({
        "Weekly Labour": weekly.groupby(["week","rsrc_name"])["qty"].sum().reset_index(),
        "Monthly Labour": weekly.groupby(["month","rsrc_name"])["qty"].sum().reset_index(),
        "By Resource": res_sum,
    })
    st.download_button("📥 Export Labour Data", xls, "labour_histogram.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# -----------------------------------------------------------------------------
# PAGE: SCHEDULE HEALTH CHECK
# -----------------------------------------------------------------------------

def page_health_check(data: dict, near_crit_days: float):
    st.title("🩺 Schedule Health Check")
    st.markdown("> Automated quality checks to identify common schedule issues.")

    tasks = data["tasks_df"]
    rels = data["relationships_df"]

    if tasks.empty:
        st.warning("No activities loaded.")
        return

    tasks = get_critical_threshold(tasks, near_crit_days)

    # Build predecessor/successor sets
    tasks_with_pred = set()
    tasks_with_succ = set()
    if not rels.empty:
        tasks_with_pred = set(rels["succ_task_id"].dropna()) if "succ_task_id" in rels.columns else set()
        tasks_with_succ = set(rels["pred_task_id"].dropna()) if "pred_task_id" in rels.columns else set()

    # Define checks
    checks = []

    # 1. No predecessors (excl. milestones at start)
    no_pred = tasks[~tasks["task_id"].isin(tasks_with_pred)]
    checks.append({
        "Check": "No Predecessors",
        "Count": len(no_pred),
        "Severity": "⚠️ Warning",
        "Why It Matters": "Activities with no predecessors are open-ended. They cannot be driven by logic and may cause float calculation issues.",
        "df": no_pred,
    })

    # 2. No successors
    no_succ = tasks[~tasks["task_id"].isin(tasks_with_succ)]
    checks.append({
        "Check": "No Successors",
        "Count": len(no_succ),
        "Severity": "⚠️ Warning",
        "Why It Matters": "Activities with no successors are open-ended and may have artificially high float.",
        "df": no_succ,
    })

    # 3. Negative float
    neg_float = tasks[tasks["total_float_days"].apply(lambda f: f is not None and f < 0)]
    checks.append({
        "Check": "Negative Float",
        "Count": len(neg_float),
        "Severity": "🔴 Critical",
        "Why It Matters": "Negative float means the current schedule cannot meet its target dates. Immediate attention required.",
        "df": neg_float,
    })

    # 4. High float (> 60 days)
    high_float = tasks[tasks["total_float_days"].apply(lambda f: f is not None and f > 60)]
    checks.append({
        "Check": "Very High Float (>60 days)",
        "Count": len(high_float),
        "Severity": "ℹ️ Info",
        "Why It Matters": "Activities with very high float may have missing logic or may not be properly constrained.",
        "df": high_float,
    })

    # 5. Excessive duration (> 60 working days)
    excess_dur = tasks[tasks["orig_dur_days"].apply(lambda d: d is not None and d > 60)]
    checks.append({
        "Check": "Excessive Duration (>60 days)",
        "Count": len(excess_dur),
        "Severity": "⚠️ Warning",
        "Why It Matters": "Very long activities are difficult to control and should usually be broken down into smaller work packages.",
        "df": excess_dur,
    })

    # 6. Constraints
    constrained = tasks[tasks["cstr_type"].apply(
        lambda x: bool(x) and str(x).strip() not in ("", "None")
    )] if "cstr_type" in tasks.columns else pd.DataFrame()
    checks.append({
        "Check": "Constrained Activities",
        "Count": len(constrained),
        "Severity": "⚠️ Warning",
        "Why It Matters": "Constraints override schedule logic and can create artificial float or negative float. Each constraint should be justified.",
        "df": constrained,
    })

    # 7. Excessive lag (> 10 days)
    if not rels.empty and "lag_days" in rels.columns:
        high_lag = rels[rels["lag_days"].apply(lambda l: l is not None and abs(safe_float(l,0)) > 10)]
        checks.append({
            "Check": "Excessive Lag (|lag| > 10 days)",
            "Count": len(high_lag),
            "Severity": "⚠️ Warning",
            "Why It Matters": "Excessive lag can hide critical path issues. Lag should be replaced with properly sequenced activities.",
            "df": high_lag,
        })

    # 8. Missing dates
    missing_dates = tasks[tasks["eff_start"].isna() | tasks["eff_finish"].isna()]
    checks.append({
        "Check": "Missing Start or Finish Dates",
        "Count": len(missing_dates),
        "Severity": "🔴 Critical",
        "Why It Matters": "Activities with no dates cannot be scheduled or reported on.",
        "df": missing_dates,
    })

    # 9. Actual dates in future
    now = datetime.now()
    future_actuals = tasks[
        tasks["act_start"].apply(lambda d: d is not None and d > now) |
        tasks["act_finish"].apply(lambda d: d is not None and d > now)
    ] if "act_start" in tasks.columns else pd.DataFrame()
    checks.append({
        "Check": "Future Actual Dates",
        "Count": len(future_actuals),
        "Severity": "🔴 Critical",
        "Why It Matters": "Actual start/finish dates should not be in the future. This indicates data entry errors.",
        "df": future_actuals,
    })

    # 10. Critical not started
    crit_not_started = tasks[
        tasks["is_critical"] &
        tasks["status"].apply(lambda s: str(s) in ("TK_NotStart", "Not Started") if pd.notna(s) else False)
    ] if "status" in tasks.columns else pd.DataFrame()
    checks.append({
        "Check": "Critical Activities Not Started",
        "Count": len(crit_not_started),
        "Severity": "🔴 Critical",
        "Why It Matters": "Critical activities that haven't started need immediate attention to avoid slippage.",
        "df": crit_not_started,
    })

    # 11. Near-critical due in 8 weeks
    eight_weeks = now + timedelta(weeks=8)
    near_due = tasks[
        tasks["is_near_critical"] &
        tasks["eff_finish"].apply(lambda d: d is not None and d <= eight_weeks)
    ] if "eff_finish" in tasks.columns else pd.DataFrame()
    checks.append({
        "Check": "Near-Critical Due in 8 Weeks",
        "Count": len(near_due),
        "Severity": "⚠️ Warning",
        "Why It Matters": "Near-critical activities finishing soon may become critical if not progressed.",
        "df": near_due,
    })

    # Scorecard
    st.subheader("Health Check Scorecard")
    score_data = [
        {"Check": c["Check"], "Count": c["Count"], "Severity": c["Severity"]}
        for c in checks
    ]
    score_df = pd.DataFrame(score_data)
    st.dataframe(score_df, use_container_width=True)

    # Detail per check
    st.divider()
    for chk in checks:
        with st.expander(f"{chk['Severity']} -- {chk['Check']} ({chk['Count']})"):
            st.markdown(f"**Why it matters:** {chk['Why It Matters']}")
            df = chk["df"]
            if not df.empty:
                disp = [c for c in ["task_code","task_name","wbs_path","eff_start",
                                     "eff_finish","total_float_days","status",
                                     "cstr_type","lag_days"] if c in df.columns]
                st.dataframe(df[disp].head(100), use_container_width=True)
                # Export individual check
                xls = export_df_to_excel({chk["Check"][:31]: df[disp]})
                st.download_button(
                    f"📥 Export: {chk['Check']}", xls,
                    f"health_{chk['Check'][:20].replace(' ','_')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.success("✅ No issues found for this check.")

    # Full export
    all_export = {chk["Check"][:31]: chk["df"][[c for c in ["task_code","task_name","total_float_days","status"] if c in chk["df"].columns]] if not chk["df"].empty else pd.DataFrame(columns=["No issues"]) for chk in checks}
    xls_all = export_df_to_excel(all_export)
    st.download_button("📥 Export Full Health Check Report", xls_all, "schedule_health_check.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# -----------------------------------------------------------------------------
# PAGE: PLANNING NOTES
# -----------------------------------------------------------------------------

HIGHLIGHT_WORDS = [
    "risk", "delay", "delayed", "blocked", "constraint", "access",
    "design", "procurement", "client", "instruction", "CE", "EWN",
    "change", "issue", "hold", "pending", "late", "overrun",
]

def highlight_text(text: str) -> str:
    """Wrap highlight words in HTML span."""
    for word in HIGHLIGHT_WORDS:
        pattern = re.compile(r"\b(" + re.escape(word) + r")\b", re.IGNORECASE)
        text = pattern.sub(r'<span style="background:#fef08a;font-weight:bold;">\1</span>', text)
    return text


def page_planning_notes(data: dict):
    st.title("📝 Planning Notes")
    st.markdown("> Upload planning notes and link them to activities in the programme.")

    tasks = data["tasks_df"]
    notes_file = st.file_uploader("Upload Planning Notes (CSV, Excel, TXT, or DOCX)",
                                   type=["csv","xlsx","txt","docx"])

    if notes_file is None:
        st.info("Upload a notes file to get started. The file should contain free-text notes referencing activity IDs.")
        return

    # Read notes
    notes_text = ""
    notes_rows = []

    try:
        if notes_file.name.endswith(".csv"):
            df = pd.read_csv(notes_file)
            notes_text = " ".join(df.astype(str).values.flatten())
            notes_rows = df.to_dict("records")
        elif notes_file.name.endswith(".xlsx"):
            df = pd.read_excel(notes_file)
            notes_text = " ".join(df.astype(str).values.flatten())
            notes_rows = df.to_dict("records")
        elif notes_file.name.endswith(".txt"):
            notes_text = notes_file.read().decode("utf-8", errors="replace")
            notes_rows = [{"line": i+1, "text": line} for i, line in enumerate(notes_text.splitlines()) if line.strip()]
        elif notes_file.name.endswith(".docx"):
            from docx import Document
            doc = Document(io.BytesIO(notes_file.read()))
            lines = [p.text for p in doc.paragraphs if p.text.strip()]
            notes_text = "\n".join(lines)
            notes_rows = [{"paragraph": i+1, "text": line} for i, line in enumerate(lines)]
        else:
            st.error("Unsupported file format.")
            return
        st.success(f"Loaded notes file: {notes_file.name}")
    except Exception as e:
        st.error(f"Could not read notes file: {e}")
        return

    # Find activity IDs mentioned in notes
    if not tasks.empty and "task_code" in tasks.columns:
        task_codes = tasks["task_code"].dropna().tolist()
        found_codes = [code for code in task_codes if code in notes_text]

        st.subheader(f"Activity IDs Found in Notes: {len(found_codes)}")
        if found_codes:
            matched_tasks = tasks[tasks["task_code"].isin(found_codes)][
                ["task_code","task_name","eff_start","eff_finish","total_float_days","status"]
            ]
            st.dataframe(matched_tasks, use_container_width=True)
        else:
            st.info("No activity IDs from the programme were found in the notes.")

        # Not found
        not_found = [code for code in task_codes if code not in notes_text]
        st.caption(f"{len(not_found)} activities not mentioned in notes.")

    # Keyword search
    st.divider()
    st.subheader("Keyword Search")
    keyword = st.text_input("Search notes for keyword")

    display_rows = notes_rows
    if keyword:
        display_rows = [r for r in notes_rows if keyword.lower() in str(r).lower()]
        st.caption(f"{len(display_rows)} matching entries")

    # Display with highlights
    for row in display_rows[:100]:
        text = str(row.get("text","") or list(row.values())[-1])
        highlighted = highlight_text(text)
        st.markdown(f"<div style='background:#f8fafc;border-left:3px solid #2563eb;padding:8px;margin:4px 0;font-size:13px;'>{highlighted}</div>", unsafe_allow_html=True)

    # Full highlighted dump
    st.divider()
    st.subheader("Full Notes (with keyword highlighting)")
    highlighted_full = highlight_text(notes_text.replace("\n","<br>"))
    st.markdown(f"<div style='background:white;border:1px solid #e2e8f0;padding:16px;border-radius:8px;max-height:400px;overflow-y:auto;font-size:12px;'>{highlighted_full}</div>", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# PAGE: PROGRAMME COMPARISON
# -----------------------------------------------------------------------------

def page_programme_comparison():
    st.title("📅 Programme Comparison")
    st.markdown("> Compare two programme revisions to identify changes in dates, float, and status.")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Previous Programme")
        prev_file = st.file_uploader("Upload Previous XER", type=["xer"], key="prev_xer")
    with col2:
        st.subheader("Current Programme")
        curr_file = st.file_uploader("Upload Current XER", type=["xer"], key="curr_xer")

    if not prev_file or not curr_file:
        st.info("Upload both XER files above to compare programmes.")
        return

    with st.spinner("Parsing both programmes..."):
        prev_data = parse_xer(prev_file.read())
        curr_data = parse_xer(curr_file.read())

    prev_tasks = prev_data["tasks_df"]
    curr_tasks = curr_data["tasks_df"]

    if prev_tasks.empty or curr_tasks.empty:
        st.error("Could not parse one or both files.")
        return

    prev_tasks = get_critical_threshold(prev_tasks)
    curr_tasks = get_critical_threshold(curr_tasks)

    # Merge on task_code
    merged = prev_tasks.merge(
        curr_tasks, on="task_code", how="outer", suffixes=("_prev","_curr")
    )

    # Added / deleted
    added = curr_tasks[~curr_tasks["task_code"].isin(prev_tasks["task_code"])]
    deleted = prev_tasks[~prev_tasks["task_code"].isin(curr_tasks["task_code"])]

    # Changed activities
    common = merged.dropna(subset=["task_code"])

    def date_diff_days(d1, d2):
        if pd.isna(d1) or pd.isna(d2):
            return None
        try:
            return int((pd.Timestamp(d2) - pd.Timestamp(d1)).days)
        except Exception:
            return None

    common = common.copy()
    common["start_movement"] = common.apply(
        lambda r: date_diff_days(r.get("eff_start_prev"), r.get("eff_start_curr")), axis=1
    )
    common["finish_movement"] = common.apply(
        lambda r: date_diff_days(r.get("eff_finish_prev"), r.get("eff_finish_curr")), axis=1
    )
    common["float_movement"] = common.apply(
        lambda r: safe_float(r.get("total_float_days_curr"), 0) - safe_float(r.get("total_float_days_prev"), 0), axis=1
    )

    # Became critical / stopped being critical
    if "is_critical_prev" in common.columns and "is_critical_curr" in common.columns:
        became_crit = common[~common["is_critical_prev"].fillna(False) & common["is_critical_curr"].fillna(False)]
        stopped_crit = common[common["is_critical_prev"].fillna(False) & ~common["is_critical_curr"].fillna(False)]
    else:
        became_crit = pd.DataFrame()
        stopped_crit = pd.DataFrame()

    # Tabs
    tab1, tab2, tab3, tab4 = st.tabs(["Summary", "Added/Deleted", "Date Movement", "Critical Changes"])

    with tab1:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Added Activities", len(added))
        c2.metric("Deleted Activities", len(deleted))
        c3.metric("Became Critical", len(became_crit))
        c4.metric("Stopped Being Critical", len(stopped_crit))

        slipped = common[common["finish_movement"].apply(lambda x: x is not None and x > 0)]
        brought_fwd = common[common["finish_movement"].apply(lambda x: x is not None and x < 0)]
        c1.metric("Finish Slipped", len(slipped))
        c2.metric("Finish Brought Forward", len(brought_fwd))

    with tab2:
        st.subheader(f"Added Activities ({len(added)})")
        if not added.empty:
            avail = [c for c in ["task_code","task_name","eff_start","eff_finish","total_float_days"] if c in added.columns]
            st.dataframe(added[avail], use_container_width=True)
        st.subheader(f"Deleted Activities ({len(deleted)})")
        if not deleted.empty:
            avail = [c for c in ["task_code","task_name","eff_start","eff_finish","total_float_days"] if c in deleted.columns]
            st.dataframe(deleted[avail], use_container_width=True)

    with tab3:
        st.subheader("Date & Float Movement")
        move_cols = ["task_code","task_name_curr","start_movement","finish_movement","float_movement",
                     "eff_start_prev","eff_start_curr","eff_finish_prev","eff_finish_curr"]
        avail = [c for c in move_cols if c in common.columns]
        st.dataframe(common[avail].sort_values("finish_movement", ascending=False, na_position="last"), use_container_width=True)

        if "finish_movement" in common.columns:
            fig = px.histogram(common["finish_movement"].dropna(), nbins=30,
                               title="Finish Date Movement Distribution (days, positive = slipped)",
                               color_discrete_sequence=["#2563eb"])
            fig.add_vline(x=0, line_dash="dash", line_color="green")
            st.plotly_chart(fig, use_container_width=True)

    with tab4:
        st.subheader(f"Became Critical ({len(became_crit)})")
        if not became_crit.empty:
            st.dataframe(became_crit[[c for c in ["task_code","task_name_curr","eff_finish_prev","eff_finish_curr","float_movement"] if c in became_crit.columns]], use_container_width=True)
        st.subheader(f"Stopped Being Critical ({len(stopped_crit)})")
        if not stopped_crit.empty:
            st.dataframe(stopped_crit[[c for c in ["task_code","task_name_curr","eff_finish_prev","eff_finish_curr","float_movement"] if c in stopped_crit.columns]], use_container_width=True)

    # Export
    xls = export_df_to_excel({
        "Added": added[[c for c in ["task_code","task_name","eff_start","eff_finish"] if c in added.columns]] if not added.empty else pd.DataFrame(columns=["No data"]),
        "Deleted": deleted[[c for c in ["task_code","task_name","eff_start","eff_finish"] if c in deleted.columns]] if not deleted.empty else pd.DataFrame(columns=["No data"]),
        "Date Movement": common[[c for c in move_cols if c in common.columns]],
        "Became Critical": became_crit[[c for c in ["task_code","task_name_curr","finish_movement"] if c in became_crit.columns]] if not became_crit.empty else pd.DataFrame(columns=["No data"]),
    })
    st.download_button("📥 Export Comparison Report", xls, "programme_comparison.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# -----------------------------------------------------------------------------
# PAGE: EXPORT REPORTS
# -----------------------------------------------------------------------------

def page_export_reports(data: dict, near_crit_days: float):
    st.title("📥 Export Reports")
    st.markdown("> Download all schedule data as formatted Excel reports.")

    tasks = data["tasks_df"]
    rels = data["relationships_df"]
    wbs = data["wbs_df"]
    resources = data["resources_df"]

    if tasks.empty:
        st.warning("No data loaded to export.")
        return

    tasks = get_critical_threshold(tasks, near_crit_days)
    critical = tasks[tasks["is_critical"]]
    neg_float = tasks[tasks["total_float_days"].apply(lambda f: f is not None and f < 0)]

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Single-Sheet Exports")

        # All activities
        avail = [c for c in ["task_code","task_name","wbs_path","eff_start","eff_finish",
                              "orig_dur_days","rem_dur_days","total_float_days","free_float_days",
                              "status","task_type","is_critical","cstr_type"] if c in tasks.columns]
        xls = export_df_to_excel({"All Activities": tasks[avail]})
        st.download_button("📄 All Activities", xls, "all_activities.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Critical path
        avail_c = [c for c in avail if c in critical.columns]
        xls2 = export_df_to_excel({"Critical Path": critical[avail_c]})
        st.download_button("🔴 Critical Path Activities", xls2, "critical_path.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Relationships
        if not rels.empty:
            xls3 = export_df_to_excel({"Relationships": rels})
            st.download_button("🔗 All Relationships", xls3, "relationships.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with col2:
        st.subheader("Multi-Sheet Reports")

        # Full schedule pack
        sheets = {"All Activities": tasks[avail]}
        if not critical.empty:
            sheets["Critical Path"] = critical[avail_c]
        if not neg_float.empty:
            sheets["Negative Float"] = neg_float[[c for c in avail if c in neg_float.columns]]
        if not rels.empty:
            sheets["Relationships"] = rels
        if not wbs.empty:
            sheets["WBS"] = wbs
        if not resources.empty:
            sheets["Resources"] = resources

        xls_full = export_df_to_excel(sheets)
        st.download_button("📦 Full Schedule Data Pack", xls_full, "schedule_data_pack.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # WBS summary
        if "wbs_path" in tasks.columns:
            tasks["wbs_top"] = tasks["wbs_path"].apply(
                lambda x: str(x).split(" > ")[0] if pd.notna(x) and x else "Unknown"
            )
            wbs_summary = tasks.groupby("wbs_top").agg(
                total=("task_id","count"),
                critical=("is_critical","sum"),
                near_critical=("is_near_critical","sum"),
            ).reset_index()
            xls_wbs = export_df_to_excel({"WBS Summary": wbs_summary})
            st.download_button("🌲 WBS Summary", xls_wbs, "wbs_summary.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# -----------------------------------------------------------------------------
# SIDEBAR & MAIN APP
# -----------------------------------------------------------------------------

def sidebar_upload():
    """Handle file upload in sidebar and return parsed data."""
    with st.sidebar:
        st.markdown("## 🏗️ P6 Planner Tool")
        st.markdown("*Primavera P6 XER Analyser*")
        st.divider()

        xer_file = st.file_uploader("📂 Upload XER File", type=["xer"])

        st.divider()
        st.subheader("⚙️ Settings")
        near_crit_days = st.slider(
            "Near-Critical Float Threshold (days)",
            min_value=1, max_value=30, value=10, step=1,
        )

        st.divider()
        page = st.selectbox(
            "📋 Navigation",
            [
                "📊 Project Summary",
                "🔍 Activity Search",
                "🔗 Logic Trace",
                "🚨 Critical Path Analysis",
                "🎯 Critical Path to Activity",
                "👷 Labour Histogram",
                "🩺 Schedule Health Check",
                "📝 Planning Notes",
                "📅 Programme Comparison",
                "📥 Export Reports",
            ]
        )

        st.divider()
        st.caption("Upload a .xer file exported from Primavera P6. For best results use a fully scheduled programme with resources assigned.")

    return xer_file, near_crit_days, page


def main():
    xer_file, near_crit_days, page = sidebar_upload()

    # Programme comparison doesn't need the main file loaded
    if page == "📅 Programme Comparison":
        page_programme_comparison()
        return

    # Load XER
    if xer_file is None:
        st.title("🏗️ P6 XER Project Manager Planning Tool")
        st.markdown("""
        <div class="info-box">
        <h3>Welcome</h3>
        This tool allows Project Managers and operational teams to interrogate Primavera P6 schedules 
        without needing to open P6.

        <h4>Getting Started</h4>
        <ol>
        <li>Export your programme from P6 as an <strong>.xer file</strong> (File -> Export -> Primavera P6 XER)</li>
        <li>Upload it using the <strong>sidebar on the left</strong></li>
        <li>Navigate between pages using the sidebar menu</li>
        </ol>

        <h4>What You Can Do</h4>
        <ul>
        <li>📊 Review project summary statistics and charts</li>
        <li>🔍 Search and filter activities</li>
        <li>🔗 Trace predecessor and successor logic chains</li>
        <li>🚨 Analyse the critical path and near-critical activities</li>
        <li>🎯 Identify what is driving any activity or milestone</li>
        <li>👷 View labour histograms (if resourced)</li>
        <li>🩺 Run automated schedule health checks</li>
        <li>📝 Link planning notes to activities</li>
        <li>📅 Compare two programme revisions</li>
        <li>📥 Export reports to Excel</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)
        return

    # Cache parsed data in session state
    cache_key = f"xer_data_{xer_file.name}_{xer_file.size}"
    if cache_key not in st.session_state:
        with st.spinner(f"Parsing {xer_file.name}..."):
            try:
                data = parse_xer(xer_file.read())
                st.session_state[cache_key] = data
                st.session_state["current_xer_key"] = cache_key
            except Exception as e:
                st.error(f"Failed to parse XER file: {e}")
                return
    else:
        data = st.session_state[cache_key]
        st.session_state["current_xer_key"] = cache_key

    # Show parse method info
    method = data.get("parse_method","-")
    n_tasks = len(data["tasks_df"])
    n_rels = len(data["relationships_df"])
    st.sidebar.success(f"✅ Loaded: {n_tasks} activities, {n_rels} relationships")
    st.sidebar.caption(f"Parser: {method}")

    # Route to pages
    if page == "📊 Project Summary":
        page_project_summary(data, near_crit_days)
    elif page == "🔍 Activity Search":
        page_activity_search(data, near_crit_days)
    elif page == "🔗 Logic Trace":
        page_logic_trace(data, near_crit_days)
    elif page == "🚨 Critical Path Analysis":
        page_critical_path(data, near_crit_days)
    elif page == "🎯 Critical Path to Activity":
        page_critical_path_to_activity(data, near_crit_days)
    elif page == "👷 Labour Histogram":
        page_labour_histogram(data)
    elif page == "🩺 Schedule Health Check":
        page_health_check(data, near_crit_days)
    elif page == "📝 Planning Notes":
        page_planning_notes(data)
    elif page == "📥 Export Reports":
        page_export_reports(data, near_crit_days)


if __name__ == "__main__":
    main()
