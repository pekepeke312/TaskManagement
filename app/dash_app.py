import pandas as pd
from io import StringIO
import webbrowser
import threading

import dash
from dash import Dash, html, dcc, dash_table, Input, Output, State, no_update, callback_context

from .constants import (
    STATUS_TODO, STATUS_INPROGRESS, STATUS_REVIEW, STATUS_DONE
)
from .schema import TaskSchema
from .repository import ExcelTaskRepository
from .dependency_service import DependencyService
from .gantt_figure import GanttFigureBuilder
from .ui_text import UI

from dash import dcc, html

upload_box = dcc.Upload(
    id="upload-xlsx",
    children=html.Div(["Drag and Drop or ", html.A("Select Excel (.xlsx)")]),
    multiple=False,
    accept=".xlsx,.xls",
    style={
        "width": "100%",
        "height": "60px",
        "lineHeight": "60px",
        "borderWidth": "1px",
        "borderStyle": "dashed",
        "borderRadius": "8px",
        "textAlign": "center",
        "marginBottom": "10px",
    },
)

import base64
from io import BytesIO

def _df_from_upload(contents: str, filename: str) -> pd.DataFrame:
    """
    contents: "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,...."
    """
    if not contents:
        raise ValueError("No upload contents")

    header, b64 = contents.split(",", 1)
    raw = base64.b64decode(b64)

    # Excelをメモリから読む
    bio = BytesIO(raw)
    df = pd.read_excel(bio, sheet_name=0)  # 必要なら sheet選択UIも作れる
    return df


class GanttDashApp:
    STORE_KEY = "tasks-store"
    HIDDEN_KEY = "hidden-legendgroups"

    def __init__(self, repo: ExcelTaskRepository):
        self.repo = repo
        self.dep = DependencyService()
        self.fig_builder = GanttFigureBuilder(self.dep)

        self.app: Dash = dash.Dash(__name__)
        self._build_layout()
        self._register_callbacks()

    # =========================
    # Helpers
    # =========================
    def _to_table_rows(self, df: pd.DataFrame) -> list[dict]:
        d = df.copy()
        d[TaskSchema.COL_START] = pd.to_datetime(
            d[TaskSchema.COL_START], errors="coerce"
        ).dt.strftime("%Y-%m-%d %H:%M")
        d[TaskSchema.COL_END] = pd.to_datetime(
            d[TaskSchema.COL_END], errors="coerce"
        ).dt.strftime("%Y-%m-%d %H:%M")
        return d.to_dict("records")


    upload_box = dcc.Upload(
        id="upload-xlsx",
        children=html.Div(["Drag and Drop or ", html.A("Select Excel (.xlsx)")]),
        multiple=False,
        accept=".xlsx,.xls",
        style={
            "width": "100%",
            "height": "60px",
            "lineHeight": "60px",
            "borderWidth": "1px",
            "borderStyle": "dashed",
            "borderRadius": "8px",
            "textAlign": "center",
            "marginBottom": "10px",
        },
    )

    # =========================
    # Layout
    # =========================
    def _build_layout(self) -> None:
        df = self.repo.load()

        self.app.layout = html.Div(
            [
                # ===== Title =====
                html.H2(UI["title_app"]),

                upload_box,

                # ===== Buttons =====
                html.Div(
                    [
                        html.Button(UI["btn_reload"], id="btn-reload"),
                        html.Span("  "),
                        html.Button(UI["btn_export"], id="btn-export"),
                        html.Span(id="export-msg", style={"marginLeft": "12px"}),
                    ],
                    style={"marginBottom": "10px"},
                ),

                # ===== Stores =====
                dcc.Store(
                    id=self.STORE_KEY,
                    data=df.to_json(date_format="iso", orient="records"),
                ),
                dcc.Store(id=self.HIDDEN_KEY, data=[]),
                dcc.Store(id="uploaded-filename", data=None),

                # =========================
                # Gantt (TOP)
                # =========================
                html.Div(
                    [
                        html.H4(UI["title_gantt"]),
                        dcc.Graph(
                            id="gantt-graph",
                            figure=self.fig_builder.build(df),
                            style={"height": "520px"},
                        ),
                    ],
                    style={"marginBottom": "16px"},
                ),

                # =========================
                # Table (BOTTOM)
                # =========================
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
                                    "options": [
                                        {"label": s, "value": s}
                                        for s in [
                                            STATUS_TODO,
                                            STATUS_INPROGRESS,
                                            STATUS_REVIEW,
                                            STATUS_DONE,
                                        ]
                                    ]
                                }
                            },
                            style_table={
                                "height": "360px",
                                "overflowY": "auto",
                            },
                            style_cell={
                                "textAlign": "left",
                                "padding": "6px",
                                "minWidth": "120px",
                                "maxWidth": "260px",
                                "whiteSpace": "normal",
                            },
                        ),
                    ],
                ),
            ],
            style={"padding": "14px"},
        )

    # =========================
    # Callbacks
    # =========================
    def _register_callbacks(self) -> None:
        app = self.app

        # ---- Table ↔ Store ----
        @app.callback(
            Output(self.STORE_KEY, "data"),
            Output("tasks-table", "data"),
            Output("uploaded-filename", "data"),  # ★追加
            Input("btn-reload", "n_clicks"),
            Input("tasks-table", "data_timestamp"),
            Input("upload-xlsx", "contents"),
            State("upload-xlsx", "filename"),  # ★filename受け取り
            State("tasks-table", "data"),
            prevent_initial_call=True,
        )
        def sync_store_and_table(n_reload, ts, upload_contents, upload_filename, table_rows):
            trigger = callback_context.triggered[0]["prop_id"] if callback_context.triggered else ""

            # Upload
            if trigger == "upload-xlsx.contents":
                if not upload_contents:
                    return no_update, no_update, no_update

                df = _df_from_upload(upload_contents, upload_filename or "")
                df = self.repo._maybe_rename_jp_to_en(df)

                if TaskSchema.COL_STATUS not in df.columns:
                    df[TaskSchema.COL_STATUS] = STATUS_TODO

                self.repo._validate_columns(df)
                df = self.repo._normalize(df)

                # ★ここで filename を保存
                return (
                    df.to_json(date_format="iso", orient="records"),
                    self._to_table_rows(df),
                    upload_filename,  # ← これが "uploaded-filename" store に入る
                )

            # Reload（元の固定ファイルから）
            if trigger == "btn-reload.n_clicks":
                df = self.repo.load()
                return df.to_json(date_format="iso", orient="records"), self._to_table_rows(df), no_update

            # Table edit
            if trigger == "tasks-table.data_timestamp":
                if table_rows is None:
                    return no_update, no_update, no_update
                df = pd.DataFrame(table_rows)
                df = self.repo._normalize(df)
                return df.to_json(date_format="iso", orient="records"), self._to_table_rows(df), no_update

            return no_update, no_update, no_update

        # ---- Gantt redraw ----
        @app.callback(
            Output("gantt-graph", "figure"),
            Input(self.STORE_KEY, "data"),
            Input(self.HIDDEN_KEY, "data"),
        )
        def update_gantt(store_json, hidden_groups):
            if not store_json:
                return no_update

            df = pd.read_json(StringIO(store_json), orient="records")
            df = self.repo._normalize(df)
            fig = self.fig_builder.build(df)

            hidden = set(hidden_groups or [])
            for tr in fig.data:
                lg = getattr(tr, "legendgroup", None)

                if lg and lg in hidden:
                    tr.visible = "legendonly"
                    continue

                meta = getattr(tr, "meta", None) or {}
                groups = meta.get("hide_if_any_hidden")
                if groups and any(g in hidden for g in groups):
                    tr.visible = "legendonly"

            fig.update_layout(uirevision="gantt")
            return fig

        # ---- Track legend hide/show ----
        @app.callback(
            Output(self.HIDDEN_KEY, "data"),
            Input("gantt-graph", "restyleData"),
            State("gantt-graph", "figure"),
            State(self.HIDDEN_KEY, "data"),
            prevent_initial_call=True,
        )
        def update_hidden_groups(restyle, fig, hidden):
            if not restyle or not fig:
                return hidden

            hidden_set = set(hidden or [])
            changes, idxs = restyle

            if "visible" not in changes:
                return list(hidden_set)

            new_vis = changes["visible"]
            for i in idxs:
                lg = fig["data"][i].get("legendgroup")
                if not lg:
                    continue
                if new_vis == "legendonly":
                    hidden_set.add(lg)
                else:
                    hidden_set.discard(lg)

            return list(hidden_set)

        # ---- Export Excel ----
        @app.callback(
            Output("export-msg", "children"),
            Input("btn-export", "n_clicks"),
            State(self.STORE_KEY, "data"),
            State("uploaded-filename", "data"),  # ★追加
            prevent_initial_call=True,
        )
        def export_excel(_n, store_json, uploaded_filename):
            if not store_json:
                return UI["msg_no_data"]

            df = pd.read_json(StringIO(store_json), orient="records")
            df = self.repo._normalize(df)

            # ★保存名を決める（アップロードがあればそれ優先）
            if uploaded_filename:
                src = Path(uploaded_filename).name  # 念のため basename のみ
                stem = Path(src).stem
                out_name = f"{stem}_updated.xlsx"
            else:
                # 従来通り（固定入力ファイル名ベース）
                out_name = f"{self.repo.xlsx_path.stem}_updated.xlsx"

            # サーバ側の保存先（プロジェクトフォルダに出す例）
            out_path = self.repo.xlsx_path.with_name(out_name)

            self.repo.save(df, out_path)
            return f'{UI["msg_saved_prefix"]} {out_path.name}'

    # =========================
    # Run
    # =========================
    def run(self, host="127.0.0.1", port=8050, debug=True):
        url = f"http://{host}:{port}/"
        threading.Timer(1.0, lambda: webbrowser.open_new(url)).start()
        self.app.run(host=host, port=port, debug=debug, use_reloader=False)
