# app/constants.py
from __future__ import annotations

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

# =========================
# Status values
# =========================

STATUS_TODO = "To Do"
STATUS_INPROGRESS = "In progress"
STATUS_REVIEW = "Review"
STATUS_DONE = "Done"

ALLOWED_STATUS = {STATUS_TODO, STATUS_INPROGRESS, STATUS_REVIEW, STATUS_DONE}