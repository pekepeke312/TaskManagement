from app.repository import ExcelTaskRepository
from app.dash_app import GanttDashApp

if __name__ == "__main__":
    repo = ExcelTaskRepository("Japan_Vacation_2026.xlsx", sheet_name=0)
    app = GanttDashApp(repo)
    app.run()
