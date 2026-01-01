from __future__ import annotations

import threading
import webbrowser
from io import StringIO

import pandas as pd
import dash
from dash import Dash, html, dcc, dash_table, Input, Output, State, no_update
from dash import callback_context

from .repository import ExcelTaskRepository
from .services import DependencyService
from .figure_builder import GanttFigureBuilder
from .schema import TaskSchema, STATUS_TODO, STATUS_INPROGRESS, STATUS_REVIEW, STATUS_DONE
from .ui_text import UI


class GanttDashApp:
    STORE_KEY = "tasks-store"
    HIDDEN_KEY = "hidden-groups"

    def __init__(self, repo: ExcelTaskRepository):
        self.repo = repo
        self.dep = DependencyService()
        self.fig_builder = GanttFigureBuilder(self.dep)
        self.app: Dash = dash.Dash(__name__)
        self._build_layout()
        self._register_callbacks()

    def _to_table_rows(self, df: pd.DataFrame) -> list[dict]:
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
                dcc.Store(id=self.HIDDEN_KEY, data=[]),

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
                                        {"name": TaskSchema.COL_STATUS, "id": TaskSchema.COL_STATUS, "editable": True, "presentation": "dropdown"},
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
                lg = None
                try:
                    lg = fig["data"][i].get("legendgroup")
                except Exception:
                    lg = None
                if not lg:
                    continue

                if new_vis == "legendonly":
                    hidden_set.add(lg)
                else:
                    hidden_set.discard(lg)

            return list(hidden_set)

        @app.callback(
            Output("gantt-graph", "figure"),
            Input(self.STORE_KEY, "data"),
            Input(self.HIDDEN_KEY, "data"),
        )
        def render_figure(store_json, hidden_groups):
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
                else:
                    tr.visible = True

            fig.update_layout(uirevision="gantt")
            return fig

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
