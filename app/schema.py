from __future__ import annotations
from typing import Dict

# Status
STATUS_TODO = "To Do"
STATUS_INPROGRESS = "In progress"
STATUS_REVIEW = "Review"
STATUS_DONE = "Done"
ALLOWED_STATUS = {STATUS_TODO, STATUS_INPROGRESS, STATUS_REVIEW, STATUS_DONE}


class TaskSchema:
    COL_NAME = "Task Name"
    COL_ID = "Task ID"
    COL_START = "Start Date"
    COL_END = "End Date"
    COL_PROGRESS = "Progress"
    COL_PARENT = "Parent Task"
    COL_CATEGORY = "Category"
    COL_STATUS = "Status"

    REQUIRED = [
        COL_NAME, COL_ID, COL_START, COL_END,
        COL_PROGRESS, COL_PARENT, COL_CATEGORY,
        COL_STATUS,
    ]


JP_TO_EN: Dict[str, str] = {
    "項目名": TaskSchema.COL_NAME,
    "タスク管理ID": TaskSchema.COL_ID,
    "開始日": TaskSchema.COL_START,
    "期限": TaskSchema.COL_END,
    "進捗": TaskSchema.COL_PROGRESS,
    "親タスク": TaskSchema.COL_PARENT,
    "カテゴリ": TaskSchema.COL_CATEGORY,
    "ステータス": TaskSchema.COL_STATUS,
}
