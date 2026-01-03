# app/gantt_figure.py
from __future__ import annotations

from typing import List, Tuple, Dict, Any

import numpy as np
import pandas as pd
from datetime import datetime

import plotly.express as px
import plotly.graph_objects as go

from .ui_text import UI
from .constants import (
    STATUS_TODO, STATUS_INPROGRESS, STATUS_REVIEW, STATUS_DONE
)

from .schema import TaskSchema
from .dependency_service import DependencyService


# from .schema import TaskSchema
# from .services import DependencyService
# from .constants import (
#     UI,
#     STATUS_TODO,
#     STATUS_INPROGRESS,
#     STATUS_REVIEW,
#     STATUS_DONE,
# )


class GanttFigureBuilder:
    """
    Build a Gantt chart with:
      - Base bars (Category colors for To Do / In progress)
      - Review / Done bars (gray tones)
      - Progress overlay (dark overlay portion) that follows legendgroup toggles
      - BLOCKED lock icon (🔒) that follows legendgroup toggles
      - Dependency arrows (line + arrowhead) that are hidden when either side is hidden
      - Weekend bands (Sat/Sun)

    Important:
      - Dependencies / locks are added as traces (NOT annotations) so we can control visibility.
      - Each dependency trace has meta={"hide_if_any_hidden":[parent_group, child_group]}.
      - Each lock trace has meta={"hide_if_any_hidden":[its_group]}.
    """

    def __init__(self, dependency_service: DependencyService):
        self.dep = dependency_service

    # -------- helpers --------
    @staticmethod
    def task_legendgroup(row: Dict[str, Any]) -> str:
        st = str(row.get(TaskSchema.COL_STATUS, "")).strip()
        if st == STATUS_REVIEW:
            return "status:Review"
        if st == STATUS_DONE:
            return "status:Done"
        cat = str(row.get(TaskSchema.COL_CATEGORY, "")).strip()
        return f"cat:{cat}"

    def add_weekend_vrects(self, fig: go.Figure, start_date, end_date) -> None:
        current = pd.to_datetime(start_date).normalize()
        end = pd.to_datetime(end_date).normalize()

        while current <= end:
            if current.weekday() == 5:  # Saturday
                fig.add_vrect(
                    x0=current,
                    x1=current + pd.Timedelta(days=1),
                    fillcolor="rgba(173, 216, 230, 0.25)",
                    line_width=0,
                    layer="below",
                )
            elif current.weekday() == 6:  # Sunday
                fig.add_vrect(
                    x0=current,
                    x1=current + pd.Timedelta(days=1),
                    fillcolor="rgba(255, 182, 193, 0.30)",
                    line_width=0,
                    layer="below",
                )
            current += pd.Timedelta(days=1)

    # -------- main --------
    def build(self, df_in: pd.DataFrame) -> go.Figure:
        # Normalize dates
        df = df_in.copy()
        df[TaskSchema.COL_START] = pd.to_datetime(df[TaskSchema.COL_START], errors="coerce")
        df[TaskSchema.COL_END] = pd.to_datetime(df[TaskSchema.COL_END], errors="coerce")

        df_chart = df.dropna(subset=[TaskSchema.COL_START, TaskSchema.COL_END]).copy()
        df_chart = self.dep.add_blocked(df_chart)

        # Split by status
        df_normal = df_chart[df_chart[TaskSchema.COL_STATUS].isin([STATUS_TODO, STATUS_INPROGRESS])].copy()
        df_review = df_chart[df_chart[TaskSchema.COL_STATUS] == STATUS_REVIEW].copy()
        df_done = df_chart[df_chart[TaskSchema.COL_STATUS] == STATUS_DONE].copy()

        fig = go.Figure()

        # -------------------------
        # 1) Base bars
        # -------------------------
        if not df_normal.empty:
            fig_base = px.timeline(
                df_normal,
                x_start=TaskSchema.COL_START,
                x_end=TaskSchema.COL_END,
                y=TaskSchema.COL_NAME,
                color=TaskSchema.COL_CATEGORY,
                hover_data=[
                    TaskSchema.COL_ID,
                    TaskSchema.COL_PARENT,
                    TaskSchema.COL_PROGRESS,
                    TaskSchema.COL_START,
                    TaskSchema.COL_END,
                    TaskSchema.COL_STATUS,
                    DependencyService.COL_BLOCK,
                    TaskSchema.COL_CATEGORY,
                ],
            )
            for tr in fig_base.data:
                cat = tr.name
                tr.legendgroup = f"cat:{cat}"
            fig.add_traces(fig_base.data)

        if not df_review.empty:
            fig_review = px.timeline(
                df_review,
                x_start=TaskSchema.COL_START,
                x_end=TaskSchema.COL_END,
                y=TaskSchema.COL_NAME,
            )
            for tr in fig_review.data:
                tr.showlegend = False
                tr.marker.color = "rgba(160,160,160,0.85)"
                tr.legendgroup = "status:Review"
            fig.add_traces(fig_review.data)

        if not df_done.empty:
            fig_done = px.timeline(
                df_done,
                x_start=TaskSchema.COL_START,
                x_end=TaskSchema.COL_END,
                y=TaskSchema.COL_NAME,
            )
            for tr in fig_done.data:
                tr.showlegend = False
                tr.marker.color = "rgba(90,90,90,0.90)"
                tr.legendgroup = "status:Done"
            fig.add_traces(fig_done.data)

        # -------------------------
        # Layout
        # -------------------------
        fig.update_yaxes(type="category", autorange="reversed")
        fig.update_xaxes(type="date",
                         tickformat="%Y-%m-%d\n%H:%M",
                         dtick=24*60*60*1000,
                         )
        fig.update_layout(
            title=UI["title_gantt_full"],
            height=max(520, 28 * max(len(df_chart), 1) + 260),
            xaxis_title=UI["xaxis"],
            yaxis_title=UI["yaxis"],
            legend_title_text=UI["legend_category"],
            barmode="overlay",
            legend=dict(groupclick="togglegroup"),
            # keep legend toggles even if figure is regenerated
            uirevision="gantt",
        )

        # -------------------------
        # 2) Progress overlay (same legendgroup)
        # -------------------------
        def add_progress_overlay(df_subset: pd.DataFrame, legendgroup: str):
            if df_subset.empty:
                return

            d = df_subset.copy()
            # progress end time
            d["_progress_end"] = d[TaskSchema.COL_START] + (d[TaskSchema.COL_END] - d[TaskSchema.COL_START]) * (
                pd.to_numeric(d[TaskSchema.COL_PROGRESS], errors="coerce").fillna(0) / 100.0
            )
            d_end = d.copy()
            d_end[TaskSchema.COL_END] = d_end["_progress_end"]

            fig_prog = px.timeline(
                d_end,
                x_start=TaskSchema.COL_START,
                x_end=TaskSchema.COL_END,
                y=TaskSchema.COL_NAME,
            )

            prog_custom = np.stack([pd.to_numeric(d[TaskSchema.COL_PROGRESS], errors="coerce").fillna(0).to_numpy()], axis=-1)

            for tr in fig_prog.data:
                tr.showlegend = False
                tr.marker.opacity = 0.30
                tr.marker.color = "rgba(0,0,0,0.35)"
                tr.legendgroup = legendgroup
                tr.customdata = prog_custom
                tr.hovertemplate = "Progress: %{customdata[0]}%<extra></extra>"

            fig.add_traces(fig_prog.data)

        # Normal grouped by category
        if not df_normal.empty:
            for cat, df_cat in df_normal.groupby(TaskSchema.COL_CATEGORY):
                add_progress_overlay(df_cat, legendgroup=f"cat:{cat}")

        add_progress_overlay(df_review, legendgroup="status:Review")
        add_progress_overlay(df_done, legendgroup="status:Done")

        # -------------------------
        # 3) LOCK icons (legend-linked + meta)
        # -------------------------
        blocked_df = df_chart[df_chart[DependencyService.COL_BLOCK] == "BLOCKED"]
        if not blocked_df.empty:
            for _, r in blocked_df.iterrows():
                lg = self.task_legendgroup(r.to_dict())
                fig.add_trace(
                    go.Scatter(
                        x=[r[TaskSchema.COL_START]],
                        y=[r[TaskSchema.COL_NAME]],
                        mode="text",
                        text=["🔒"],
                        textposition="middle left",
                        hovertext=[UI["blocked_hover"]],
                        hoverinfo="text",
                        showlegend=False,
                        legendgroup=lg,
                        meta={"kind": "lock", "hide_if_any_hidden": [lg]},
                    )
                )

        # -------------------------
        # 4) Dependencies (as traces + meta parent/child)
        # -------------------------
        id_to_row = df_chart.set_index(TaskSchema.COL_ID).to_dict(orient="index")
        for parent_id, child_id in self.dep.iter_dependencies(df_chart):
            parent = id_to_row[parent_id]
            child = id_to_row[child_id]

            parent_lg = self.task_legendgroup(parent)
            child_lg = self.task_legendgroup(child)

            meta_dep = {"kind": "dep", "hide_if_any_hidden": [parent_lg, child_lg]}

            # line
            fig.add_trace(
                go.Scatter(
                    x=[parent[TaskSchema.COL_END], child[TaskSchema.COL_START]],
                    y=[parent[TaskSchema.COL_NAME], child[TaskSchema.COL_NAME]],
                    mode="lines",
                    line=dict(width=1),
                    opacity=0.85,
                    showlegend=False,
                    hoverinfo="skip",
                    meta=meta_dep,
                )
            )

            # arrowhead marker at child start
            fig.add_trace(
                go.Scatter(
                    x=[child[TaskSchema.COL_START]],
                    y=[child[TaskSchema.COL_NAME]],
                    mode="markers",
                    marker=dict(size=8, symbol="triangle-right"),
                    opacity=0.85,
                    showlegend=False,
                    hoverinfo="skip",
                    meta=meta_dep,
                )
            )

        # -------------------------
        # 5) Weekend bands
        # -------------------------
        if not df_chart.empty:
            self.add_weekend_vrects(fig, df_chart[TaskSchema.COL_START].min(), df_chart[TaskSchema.COL_END].max())

        # -------------------------
        # 6) Status legend entries (clickable toggles)
        # -------------------------
        fig.add_trace(
            go.Scatter(
                x=[None],
                y=[None],
                mode="markers",
                marker=dict(color="rgba(160,160,160,0.85)", size=10),
                name="Review",
                showlegend=True,
                legendgroup="status:Review",
                legendgrouptitle=dict(text="Status"),
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
                legendgroup="status:Done",
                legendgrouptitle=dict(text="Status"),
            )
        )

        # -------------------------
        # Current time (NOW) vertical line  ※add_shape版（安全）
        # -------------------------
        if not df_chart.empty:
            now_ts = pd.Timestamp.now()
            xmin_ts = pd.to_datetime(df_chart[TaskSchema.COL_START].min())
            xmax_ts = pd.to_datetime(df_chart[TaskSchema.COL_END].max())

            if pd.notna(xmin_ts) and pd.notna(xmax_ts) and (xmin_ts <= now_ts <= xmax_ts):
                x_now = now_ts.to_pydatetime()  # datetime.datetime

                # 縦線（y方向はプロット全体=paper 0..1）
                fig.add_shape(
                    type="line",
                    x0=x_now, x1=x_now,
                    y0=0, y1=1,
                    xref="x",
                    yref="paper",
                    line=dict(color="red", width=2),
                    opacity=0.8,
                    layer="above",
                )

                # ラベルが欲しければ annotation を別で追加（これなら安全）
                fig.add_annotation(
                    x=x_now,
                    y=1,
                    xref="x",
                    yref="paper",
                    text="NOW",
                    showarrow=False,
                    font=dict(color="red"),
                    yanchor="bottom",
                )

        return fig
