from __future__ import annotations

from pathlib import Path
import pandas as pd

from .schema import TaskSchema, JP_TO_EN, STATUS_TODO, ALLOWED_STATUS


class ExcelTaskRepository:
    """Read/Write Excel. Internally normalizes types."""
    def __init__(self, xlsx_path: str | Path, sheet_name: str | int = 0):
        self.xlsx_path = Path(xlsx_path)
        self.sheet_name = sheet_name

    def load(self) -> pd.DataFrame:
        df = pd.read_excel(self.xlsx_path, sheet_name=self.sheet_name)
        df = self._maybe_rename_jp_to_en(df)

        if TaskSchema.COL_STATUS not in df.columns:
            df[TaskSchema.COL_STATUS] = STATUS_TODO

        self._validate_columns(df)
        return self._normalize(df)

    def save(self, df: pd.DataFrame, out_path: str | Path) -> Path:
        """Write updated xlsx; dates saved as YYYY-MM-DD strings (no time)."""
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
        df = df.copy()

        df[TaskSchema.COL_ID] = df[TaskSchema.COL_ID].astype(str).str.strip()
        df[TaskSchema.COL_PARENT] = df[TaskSchema.COL_PARENT].fillna("").astype(str).str.strip()
        df[TaskSchema.COL_CATEGORY] = df[TaskSchema.COL_CATEGORY].fillna("Uncategorized").astype(str).str.strip()

        df[TaskSchema.COL_START] = pd.to_datetime(df[TaskSchema.COL_START], errors="coerce")
        df[TaskSchema.COL_END] = pd.to_datetime(df[TaskSchema.COL_END], errors="coerce")

        df[TaskSchema.COL_PROGRESS] = (
            pd.to_numeric(df[TaskSchema.COL_PROGRESS], errors="coerce")
            .fillna(0)
            .clip(0, 100)
        )

        df[TaskSchema.COL_STATUS] = df[TaskSchema.COL_STATUS].fillna(STATUS_TODO).astype(str).str.strip()
        df.loc[~df[TaskSchema.COL_STATUS].isin(ALLOWED_STATUS), TaskSchema.COL_STATUS] = STATUS_TODO

        df = df.sort_values([TaskSchema.COL_START, TaskSchema.COL_CATEGORY, TaskSchema.COL_NAME]).reset_index(drop=True)
        return df
