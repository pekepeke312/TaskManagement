from __future__ import annotations

from typing import List, Tuple
import numpy as np
import pandas as pd

from .schema import TaskSchema


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
