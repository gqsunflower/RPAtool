"""
RunLogger: マクロ実行の証跡を残すための単純なCSVロガー。

「いつ・どのマクロの・何番目の手順を・成功/失敗/スキップ/手動補完したか」を
1行ずつ記録する。業務の実施記録・監査対応・不具合調査に使うことを想定し、
export_to_excel() でExcelファイルへ変換もできる。
"""
from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path

FIELDNAMES = ["timestamp", "macro", "step_number", "handler", "action", "status", "detail", "screenshot"]


class RunLogger:
    def __init__(self, csv_path: Path):
        self.csv_path = Path(csv_path)
        self.csv_path.parent.mkdir(parents=True, exist_ok=True)
        if not self.csv_path.exists():
            with open(self.csv_path, "w", newline="", encoding="utf-8-sig") as f:
                csv.DictWriter(f, fieldnames=FIELDNAMES).writeheader()

    def log(
        self,
        macro: str,
        step_number: int,
        handler: str,
        action: str,
        status: str,
        detail: str = "",
        screenshot: str = "",
    ) -> None:
        """status: 'success' | 'failure' | 'manual' | 'skip' など"""
        with open(self.csv_path, "a", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
            writer.writerow({
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "macro": macro,
                "step_number": step_number,
                "handler": handler,
                "action": action,
                "status": status,
                "detail": (detail or "")[:500],
                "screenshot": screenshot,
            })

    def export_to_excel(self, xlsx_path: Path) -> Path:
        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        ws.title = "実行ログ"
        ws.append(FIELDNAMES)

        with open(self.csv_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                ws.append([row.get(k, "") for k in FIELDNAMES])

        xlsx_path = Path(xlsx_path)
        xlsx_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(xlsx_path)
        return xlsx_path
