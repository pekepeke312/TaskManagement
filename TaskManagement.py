from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List, Tuple, Dict

import pandas as pd
import numpy as np
import webbrowser
import threading
from io import StringIO

import dash
from dash import Dash, html, dcc, dash_table, Input, Output, State, no_update
from dash import callback_context
import plotly.express as px
import plotly.graph_objects as go


# =========================
# UI Text (English)
# =========================

UI = {
    "title_app": "Excel → Gantt (Edit → Reflect)",
    "title_table": "Task Table (Edit Here)",
    "title_gantt": "Gantt Chart",
    "title_gantt_full": "Gantt (Category + Progress + Status + BLOCKED + Dependencies)",
    "btn_reload": "Reload Excel",
    "btn_export": "Export Updated Excel (Server Side)",
    "msg_no_data": "No data.",
    "msg_saved_prefix": "Saved:",
    "xaxis": "Date",
    "yaxis": "Task",
    "legend_category": "Category",
    "blocked_hover": "BLOCKED (Parent task incomplete)",
}

# Allowed values for the new Status column
STATUS_TODO = "To Do"
STATUS_INPROGRESS = "In progress"
STATUS_REVIEW = "Review"
STATUS_DONE = "Done"
ALLOWED_STATUS = {STATUS_TODO, STATUS_INPROGRESS, STATUS_REVIEW, STATUS_DONE}


# =========================
# Domain Model (optional)
# =========================

@dataclass(frozen=True)
class Task:
    name: str
    task_id: str
    start: pd.Timestamp
    end: pd.Timestamp
    progress: float  # 0..100
    parent_id: str
    category: str
    status: str


# =========================
# Schema (English columns)
# =========================

class TaskSchema:
    """Excel column names (English)."""

    COL_NAME = "Task Name"
    COL_ID = "Task ID"
    COL_START = "Start Date"
    COL_END = "End Date"
    COL_PROGRESS = "Progress"
    COL_PARENT = "Parent Task"
    COL_CATEGORY = "Category"
    COL_STATUS = "Status"  # ★ new

    REQUIRED = [
        COL_NAME, COL_ID, COL_START, COL_END,
        COL_PROGRESS, COL_PARENT, COL_CATEGORY,
        COL_STATUS,  # ★ new
    ]


# Optional: auto-migrate Japanese header Excel to English
JP_TO_EN: Dict[str, str] = {
    "項目名": TaskSchema.COL_NAME,
    "タスク管理ID": TaskSchema.COL_ID,
    "開始日": TaskSchema.COL_START,
    "期限": TaskSchema.COL_END,
    "進捗": TaskSchema.COL_PROGRESS,
    "親タスク": TaskSchema.COL_PARENT,
    "カテゴリ": TaskSchema.COL_CATEGORY,
    "ステータス": TaskSchema.COL_STATUS,  # (if you used this in JP)
}


# =========================
# Repository (Excel I/O)
# =========================

class ExcelTaskRepository:
    """Read/Write Excel. Internally normalizes types."""
    def __init__(self, xlsx_path: str | Path, sheet_name: str | int = 0):
        self.xlsx_path = Path(xlsx_path)
        self.sheet_name = sheet_name

    def load(self) -> pd.DataFrame:
        df = pd.read_excel(self.xlsx_path, sheet_name=self.sheet_name)
        df = self._maybe_rename_jp_to_en(df)

        # If Status column doesn't exist, create default
        if TaskSchema.COL_STATUS not in df.columns:
            df[TaskSchema.COL_STATUS] = STATUS_TODO

        self._validate_columns(df)
        return self._normalize(df)

    def save(self, df: pd.DataFrame, out_path: str | Path) -> Path:
        """Write updated xlsx with English headers; dates saved as YYYY-MM-DD strings (no time)."""
        out_path = Path(out_path)
        df_out = df.copy()

        self._validate_columns(df_out)

        for col in (TaskSchema.COL_START, TaskSchema.COL_END):
            df_out[col] = pd.to_datetime(df_out[col], errors="coerce").dt.strftime("%Y-%m-%d")

        df_out = df_out[TaskSchema.REQUIRED]
        df_out.to_excel(out_path, index=False)
        return out_path

    def _maybe_rename_jp_to_en(self, df: pd.DataFrame) -> pd.DataFrame:
        cols = set(df.columns.astype(str))
        jp_cols = set(JP_TO_EN.keys())
        if cols & jp_cols:
            df = df.rename(columns={c: JP_TO_EN[c] for c in df.columns if c in JP_TO_EN})
        return df

    def _validate_columns(self, df: pd.DataFrame) -> None:
        missing = [c for c in TaskSchema.REQUIRED if c not in df.columns]
        if missing:
            raise ValueError(f"Excel is missing required columns: {missing}")

    def _normalize(self, df: pd.DataFrame) -> pd.DataFrame:
        """Normalize types for internal use (dates as Timestamp normalized; progress numeric)."""
        df = df.copy()

        df[TaskSchema.COL_ID] = df[TaskSchema.COL_ID].astype(str).str.strip()
        df[TaskSchema.COL_PARENT] = df[TaskSchema.COL_PARENT].fillna("").astype(str).str.strip()
        df[TaskSchema.COL_CATEGORY] = df[TaskSchema.COL_CATEGORY].fillna("Uncategorized").astype(str).str.strip()

        df[TaskSchema.COL_START] = pd.to_datetime(df[TaskSchema.COL_START], errors="coerce").dt.normalize()
        df[TaskSchema.COL_END] = pd.to_datetime(df[TaskSchema.COL_END], errors="coerce").dt.normalize()

        df[TaskSchema.COL_PROGRESS] = (
            pd.to_numeric(df[TaskSchema.COL_PROGRESS], errors="coerce")
            .fillna(0)
            .clip(0, 100)
        )

        # Status normalization
        df[TaskSchema.COL_STATUS] = df[TaskSchema.COL_STATUS].fillna(STATUS_TODO).astype(str).str.strip()
        df.loc[~df[TaskSchema.COL_STATUS].isin(ALLOWED_STATUS), TaskSchema.COL_STATUS] = STATUS_TODO

        # Optional sort for readability
        df = df.sort_values([TaskSchema.COL_START, TaskSchema.COL_CATEGORY, TaskSchema.COL_NAME]).reset_index(drop=True)
        return df


# =========================
# Services
# =========================

class DependencyService:
    COL_BLOCK = "Blocked"  # OK / BLOCKED

    def add_blocked(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        id_to_prog = df.set_index(TaskSchema.COL_ID)[TaskSchema.COL_PROGRESS].to_dict()
        id_exists = set(id_to_prog.keys())

        def blocked(row) -> bool:
            parent = row[TaskSchema.COL_PARENT]
            if parent == "":
                return False
            if parent not in id_exists:
                return True
            return float(id_to_prog[parent]) < 100

        df[self.COL_BLOCK] = np.where(df.apply(blocked, axis=1), "BLOCKED", "OK")
        return df

    def iter_dependencies(self, df: pd.DataFrame) -> List[Tuple[str, str]]:
        deps: List[Tuple[str, str]] = []
        id_set = set(df[TaskSchema.COL_ID].astype(str))
        for _, r in df.iterrows():
            child = str(r[TaskSchema.COL_ID])
            parent = str(r[TaskSchema.COL_PARENT]).strip()
            if parent and parent in id_set:
                deps.append((parent, child))
        return deps


# =========================
# Plot Builder
# =========================

class GanttFigureBuilder:
    """
    Category + Progress overlay + Status coloring:
      - Review: gray
      - Done: dark gray
      - To Do / In progress: keep category color
    plus BLOCKED icon + dependency arrows + weekend bands.
    """
    def __init__(self, dependency_service: DependencyService):
        self.dep = dependency_service

    def add_weekend_vrects(self, fig: go.Figure, start_date, end_date) -> None:
        current = pd.to_datetime(start_date).normalize()
        end = pd.to_datetime(end_date).normalize()

        while current <= end:
            if current.weekday() == 5:  # Saturday
                fig.add_vrect(
                    x0=current, x1=current + pd.Timedelta(days=1),
                    fillcolor="rgba(173, 216, 230, 0.25)",
                    line_width=0, layer="below",
                )
            elif current.weekday() == 6:  # Sunday
                fig.add_vrect(
                    x0=current, x1=current + pd.Timedelta(days=1),
                    fillcolor="rgba(255, 182, 193, 0.30)",
                    line_width=0, layer="below",
                )
            current += pd.Timedelta(days=1)

    def _is_override_gray(self, status: str) -> bool:
        return status in (STATUS_REVIEW, STATUS_DONE)

    def build(self, df_in: pd.DataFrame) -> go.Figure:
        # Normalize dates safely
        df = df_in.copy()
        df[TaskSchema.COL_START] = pd.to_datetime(df[TaskSchema.COL_START], errors="coerce").dt.normalize()
        df[TaskSchema.COL_END] = pd.to_datetime(df[TaskSchema.COL_END], errors="coerce").dt.normalize()

        df_chart = df.dropna(subset=[TaskSchema.COL_START, TaskSchema.COL_END]).copy()
        df_chart = self.dep.add_blocked(df_chart)

        # Split by Status:
        # - normal: To Do / In progress -> category color
        # - review/done: override gray colors
        df_normal = df_chart[df_chart[TaskSchema.COL_STATUS].isin([STATUS_TODO, STATUS_INPROGRESS])].copy()
        df_review = df_chart[df_chart[TaskSchema.COL_STATUS] == STATUS_REVIEW].copy()
        df_done = df_chart[df_chart[TaskSchema.COL_STATUS] == STATUS_DONE].copy()

        fig = go.Figure()

        # 1) Normal tasks (category colors)
        if not df_normal.empty:
            fig_base = px.timeline(
                df_normal,
                x_start=TaskSchema.COL_START,
                x_end=TaskSchema.COL_END,
                y=TaskSchema.COL_NAME,
                color=TaskSchema.COL_CATEGORY,
                hover_data=[
                    TaskSchema.COL_ID, TaskSchema.COL_PARENT, TaskSchema.COL_PROGRESS,
                    TaskSchema.COL_START, TaskSchema.COL_END,
                    TaskSchema.COL_STATUS, DependencyService.COL_BLOCK, TaskSchema.COL_CATEGORY,
                ],
            )
            fig.add_traces(fig_base.data)

        # 2) Review tasks (gray)
        if not df_review.empty:
            fig_review = px.timeline(
                df_review,
                x_start=TaskSchema.COL_START,
                x_end=TaskSchema.COL_END,
                y=TaskSchema.COL_NAME,
            )
            for tr in fig_review.data:
                tr.name = "Review"
                tr.showlegend = True
                tr.legendgroup = "Status"
                tr.marker.color = "rgba(160,160,160,0.85)"
            fig.add_traces(fig_review.data)

        # 3) Done tasks (dark gray)
        if not df_done.empty:
            fig_done = px.timeline(
                df_done,
                x_start=TaskSchema.COL_START,
                x_end=TaskSchema.COL_END,
                y=TaskSchema.COL_NAME,
            )
            for tr in fig_done.data:
                tr.name = "Done"
                tr.showlegend = True
                tr.legendgroup = "Status"
                tr.marker.color = "rgba(90,90,90,0.90)"
            fig.add_traces(fig_done.data)

        # Common layout
        fig.update_yaxes(autorange="reversed")
        fig.update_layout(
            title=UI["title_gantt_full"],
            height=max(520, 28 * max(len(df_chart), 1) + 240),
            xaxis_title=UI["xaxis"],
            yaxis_title=UI["yaxis"],
            legend_title_text=UI["legend_category"],
            barmode="overlay",
        )
        fig.update_xaxes(tickformat="%Y-%m-%d")

        # Progress overlay as timeline for all rows, but color depends on status:
        # - Review: gray overlay
        # - Done: darker gray overlay
        # - Normal: black-ish overlay (lets category show)
        if not df_chart.empty:
            df_prog = df_chart.copy()
            df_prog["_progress_end"] = (
                df_prog[TaskSchema.COL_START]
                + (df_prog[TaskSchema.COL_END] - df_prog[TaskSchema.COL_START]) * (df_prog[TaskSchema.COL_PROGRESS] / 100.0)
            )
            df_prog[TaskSchema.COL_END] = df_prog["_progress_end"]

            fig_prog = px.timeline(
                df_prog,
                x_start=TaskSchema.COL_START,
                x_end=TaskSchema.COL_END,
                y=TaskSchema.COL_NAME,
            )

            prog_custom = np.stack([df_prog[TaskSchema.COL_PROGRESS].to_numpy()], axis=-1)

            # timeline returns 1 trace; style per-point is tricky.
            # We keep a single overlay style that works for all:
            # (visual rule mainly needed on base bar; progress overlay can stay subtle)
            for tr in fig_prog.data:
                tr.showlegend = False
                tr.marker.opacity = 0.30
                tr.marker.color = "rgba(0,0,0,0.35)"
                tr.customdata = prog_custom
                tr.hovertemplate = "Progress: %{customdata[0]}%<extra></extra>"

            fig.add_traces(fig_prog.data)

        # BLOCKED lock icon
        blocked_df = df_chart[df_chart[DependencyService.COL_BLOCK] == "BLOCKED"]
        if not blocked_df.empty:
            fig.add_trace(
                go.Scatter(
                    x=blocked_df[TaskSchema.COL_START],
                    y=blocked_df[TaskSchema.COL_NAME],
                    mode="text",
                    text=["🔒"] * len(blocked_df),
                    textposition="middle left",
                    hovertext=[UI["blocked_hover"]] * len(blocked_df),
                    hoverinfo="text",
                    showlegend=False,
                )
            )

        # Dependency arrows: Parent end -> Child start
        id_to_row = df_chart.set_index(TaskSchema.COL_ID).to_dict(orient="index")
        for parent_id, child_id in self.dep.iter_dependencies(df_chart):
            parent = id_to_row[parent_id]
            child = id_to_row[child_id]
            fig.add_annotation(
                x=child[TaskSchema.COL_START],
                y=child[TaskSchema.COL_NAME],
                ax=parent[TaskSchema.COL_END],
                ay=parent[TaskSchema.COL_NAME],
                xref="x",
                yref="y",
                axref="x",
                ayref="y",
                showarrow=True,
                arrowhead=3,
                arrowsize=1,
                arrowwidth=1,
                opacity=0.85,
            )

        # Weekend bands
        if not df_chart.empty:
            self.add_weekend_vrects(fig, df_chart[TaskSchema.COL_START].min(), df_chart[TaskSchema.COL_END].max())

        # Force x-axis range (prevents "blank chart" when autorange gets weird)
        if not df_chart.empty:
            x0 = df_chart[TaskSchema.COL_START].min()
            x1 = df_chart[TaskSchema.COL_END].max() + pd.Timedelta(days=1)
            fig.update_xaxes(range=[x0, x1], type="date")

        # =========================
        # Legend entries for Status (Review / Done)
        # =========================
        # Dummy traces (do not appear on chart, legend only)
        fig.add_trace(
            go.Scatter(
                x=[None],
                y=[None],
                mode="markers",
                marker=dict(color="rgba(160,160,160,0.85)", size=10),
                name="Review",
                showlegend=True,
                legendgroup="Status",
            )
        )
        fig.add_trace(
            go.Scatter(
                x=[None],
                y=[None],
                mode="markers",
                marker=dict(color="rgba(90,90,90,0.90)", size=10),
                name="Done",
                showlegend=True,
                legendgroup="Status",
            )
        )

        print("rows:", len(df_chart), "min:", df_chart[TaskSchema.COL_START].min(), "max:",
              df_chart[TaskSchema.COL_END].max())
        
        return fig


# =========================
# Dash App
# =========================

class GanttDashApp:
    STORE_KEY = "tasks-store"

    def __init__(self, repo: ExcelTaskRepository):
        self.repo = repo
        self.dep = DependencyService()
        self.fig_builder = GanttFigureBuilder(self.dep)
        self.app: Dash = dash.Dash(__name__)
        self._build_layout()
        self._register_callbacks()

    def _to_table_rows(self, df: pd.DataFrame) -> list[dict]:
        """Table display: show dates as YYYY-MM-DD strings."""
        d = df.copy()
        d[TaskSchema.COL_START] = pd.to_datetime(d[TaskSchema.COL_START], errors="coerce").dt.strftime("%Y-%m-%d")
        d[TaskSchema.COL_END] = pd.to_datetime(d[TaskSchema.COL_END], errors="coerce").dt.strftime("%Y-%m-%d")
        return d.to_dict("records")

    def _build_layout(self) -> None:
        df = self.repo.load()

        self.app.layout = html.Div(
            [
                html.H2(UI["title_app"]),
                html.Div(
                    [
                        html.Button(UI["btn_reload"], id="btn-reload"),
                        html.Span("  "),
                        html.Button(UI["btn_export"], id="btn-export"),
                        html.Span(id="export-msg", style={"marginLeft": "12px"}),
                    ],
                    style={"marginBottom": "10px"},
                ),

                dcc.Store(id=self.STORE_KEY, data=df.to_json(date_format="iso", orient="records")),

                html.Div(
                    [
                        html.Div(
                            [
                                html.H4(UI["title_table"]),
                                dash_table.DataTable(
                                    id="tasks-table",
                                    columns=[
                                        {"name": TaskSchema.COL_NAME, "id": TaskSchema.COL_NAME, "editable": True},
                                        {"name": TaskSchema.COL_ID, "id": TaskSchema.COL_ID, "editable": True},
                                        {"name": TaskSchema.COL_START, "id": TaskSchema.COL_START, "editable": True},
                                        {"name": TaskSchema.COL_END, "id": TaskSchema.COL_END, "editable": True},
                                        {"name": TaskSchema.COL_PROGRESS, "id": TaskSchema.COL_PROGRESS, "editable": True},
                                        {"name": TaskSchema.COL_PARENT, "id": TaskSchema.COL_PARENT, "editable": True},
                                        {"name": TaskSchema.COL_CATEGORY, "id": TaskSchema.COL_CATEGORY, "editable": True},
                                        {
                                            "name": TaskSchema.COL_STATUS,
                                            "id": TaskSchema.COL_STATUS,
                                            "editable": True,
                                            "presentation": "dropdown",
                                        },
                                    ],
                                    data=self._to_table_rows(df),
                                    editable=True,
                                    sort_action="native",
                                    row_deletable=True,
                                    page_action="none",
                                    dropdown={
                                        TaskSchema.COL_STATUS: {
                                            "options": [{"label": s, "value": s} for s in [STATUS_TODO, STATUS_INPROGRESS, STATUS_REVIEW, STATUS_DONE]]
                                        }
                                    },
                                    style_table={"height": "380px", "overflowY": "auto"},
                                    style_cell={"textAlign": "left", "padding": "6px", "minWidth": "120px", "maxWidth": "260px"},
                                ),
                            ],
                            style={"width": "48%", "display": "inline-block", "verticalAlign": "top"},
                        ),

                        html.Div(
                            [
                                html.H4(UI["title_gantt"]),
                                dcc.Graph(id="gantt-graph", figure=self.fig_builder.build(df)),
                            ],
                            style={"width": "51%", "display": "inline-block", "marginLeft": "1%", "verticalAlign": "top"},
                        ),
                    ]
                ),
            ],
            style={"padding": "14px"},
        )

    def _register_callbacks(self) -> None:
        app = self.app

        @app.callback(
            Output(self.STORE_KEY, "data"),
            Output("tasks-table", "data"),
            Input("btn-reload", "n_clicks"),
            Input("tasks-table", "data_timestamp"),
            State("tasks-table", "data"),
            prevent_initial_call=True,
        )
        def sync_store_and_table(n_reload, ts, table_rows):
            trigger = callback_context.triggered[0]["prop_id"] if callback_context.triggered else ""

            if trigger == "btn-reload.n_clicks":
                df = self.repo.load()
                return df.to_json(date_format="iso", orient="records"), self._to_table_rows(df)

            if trigger == "tasks-table.data_timestamp":
                if table_rows is None:
                    return no_update, no_update
                df = pd.DataFrame(table_rows)
                df = self.repo._normalize(df)
                return df.to_json(date_format="iso", orient="records"), self._to_table_rows(df)

            return no_update, no_update

        @app.callback(
            Output("gantt-graph", "figure"),
            Input(self.STORE_KEY, "data"),
        )
        def store_to_gantt(store_json):
            if not store_json:
                return no_update
            df = pd.read_json(StringIO(store_json), orient="records")
            df = self.repo._normalize(df)
            return self.fig_builder.build(df)

        @app.callback(
            Output("export-msg", "children"),
            Input("btn-export", "n_clicks"),
            State(self.STORE_KEY, "data"),
            prevent_initial_call=True,
        )
        def export_excel(_n, store_json):
            if not store_json:
                return UI["msg_no_data"]

            df = pd.read_json(StringIO(store_json), orient="records")
            df = self.repo._normalize(df)

            out = self.repo.xlsx_path.with_name(self.repo.xlsx_path.stem + "_updated.xlsx")
            self.repo.save(df, out)
            return f'{UI["msg_saved_prefix"]} {out.name}'

    def run(self, host="127.0.0.1", port=8050, debug=True):
        url = f"http://{host}:{port}/"
        threading.Timer(1.0, lambda: webbrowser.open_new(url)).start()
        self.app.run(host=host, port=port, debug=debug, use_reloader=False)


if __name__ == "__main__":
    repo = ExcelTaskRepository("gantt_sample.xlsx", sheet_name=0)
    app = GanttDashApp(repo)
    app.run()
