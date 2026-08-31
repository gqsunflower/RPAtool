"""
PipelineRunner: config/pipelines.json に登録した「マクロを実行する順番」に
沿って、複数のマクロを連続で実行する。

例:「Excelから注文データを転記 → 出来上がったPDFを結合 → 登録済みサイトに
アップロード」のような一連の定型業務をまとめて自動化する用途を想定。

途中のマクロが失敗した場合は、MacroExecutor.run() に渡した on_failure の
判断(即時中断/手動補完して再開/修正画面を開く)がそのまま適用される。
「中断」が選ばれた場合、パイプライン全体もそこで止まる
(=あるマクロの失敗を無視して次のマクロには進まない)。
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Callable

from engine.backup import backup_file
from engine.executor import MacroExecutor


class PipelineRunner:
    def __init__(self, config_dir: Path, executor: MacroExecutor):
        self.config_dir = Path(config_dir)
        self.path = self.config_dir / "pipelines.json"
        self.executor = executor
        self._load()

    def _load(self) -> None:
        with open(self.path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.pipelines: dict[str, dict] = data.get("pipelines", {})

    def reload(self) -> None:
        self._load()

    def list_names(self) -> list[str]:
        return list(self.pipelines.keys())

    def get_pipeline(self, name: str) -> dict:
        if name not in self.pipelines:
            raise KeyError(f"未登録のパイプラインです: {name}")
        return self.pipelines[name]

    def run(
        self,
        name: str,
        slot_prompt_fn: Callable[[str, str], Any],
        dry_run: bool = False,
        before_macro: Callable[[str], None] | None = None,
        **executor_kwargs: Any,
    ) -> dict[str, list[Any]]:
        """slot_prompt_fn(macro_name, slot_name) -> 値  を呼んで、固定値が
        無いスロットをその場で確認しながら、登録順にマクロを実行する。
        before_macro(macro_name) を渡しておくと、各マクロの実行直前に呼ばれる
        (マクロごとに異なるブラウザ(chrome/edge)へ切り替える等に使う)。
        """
        pipeline = self.get_pipeline(name)
        results: dict[str, list[Any]] = {}

        for entry in pipeline.get("macros", []):
            macro_name = entry["macro"]
            if before_macro:
                before_macro(macro_name)
            slots: dict = dict(entry.get("slots", {}))
            for slot_name in self.executor.required_slots(macro_name):
                if slot_name not in slots:
                    slots[slot_name] = slot_prompt_fn(macro_name, slot_name)

            results[macro_name] = self.executor.run(
                macro_name, slots, dry_run=dry_run, **executor_kwargs
            )
        return results

    def save_pipeline(self, name: str, description: str, macro_names: list[str]) -> None:
        backup_file(self.path, self.config_dir / "backups")

        with open(self.path, "r", encoding="utf-8") as f:
            data = json.load(f)
        data.setdefault("pipelines", {})[name] = {
            "description": description,
            "macros": [{"macro": m, "slots": {}} for m in macro_names],
        }
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.reload()
