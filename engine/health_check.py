"""
HealthChecker: 登録済みマクロで使っている「表示テキストによる目印」が
今もサイト上に存在するかどうかを、実際にクリック・入力せずに確認する。

相手サイトのUI変更を、業務で実際に使う前に検知するための仕組み。

制約(重要):
  この確認は同一画面内の要素検出に限られる。1つ目の open_registered_site の
  あとに続く click_by_text 等は、実際にはクリックしていない(=ページ遷移が
  発生しない)前提で、同じ画面の中に該当要素があるかどうかだけを見ている。
  そのため、複数ページ・複数画面にまたがる手順の後半ステップは正しく
  検証できないことがある(false negativeが出ることがある)。厳密に確認したい
  場合はステップ実行(main.pyの'ステップ実行')を使うこと。
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Callable


class HealthChecker:
    def __init__(self, config_dir: Path, browser_factory: Callable[[str | None], object]):
        """
        browser_factory: マクロに保存されたbrowser名("chrome"/"edge"、未保存ならNone)
                          を渡すと新しい BrowserHandler を返す関数
                          (例: lambda b: BrowserHandler(whitelist_path, headless=True,
                          browser=b or "chrome"))。Noneの場合の既定browserの決定は
                          呼び出し側(factory)に委ねる。
        """
        self.config_dir = Path(config_dir)
        self.macros_path = self.config_dir / "macros.json"
        self.browser_factory = browser_factory

    def _load_macros(self) -> dict:
        with open(self.macros_path, "r", encoding="utf-8") as f:
            return json.load(f).get("macros", {})

    def list_macro_names(self) -> list[str]:
        return list(self._load_macros().keys())

    def check_macro(self, macro_name: str) -> list[dict]:
        """1つのマクロについて、各ブラウザ手順の目印が見つかるかを確認する。
        戻り値: [{"action":..., "params":..., "status": "OK"|"NG"|"SKIP", "detail": ...}, ...]
        """
        macros = self._load_macros()
        if macro_name not in macros:
            raise KeyError(f"未登録のマクロです: {macro_name}")

        steps = macros[macro_name].get("steps", [])
        report: list[dict] = []
        browser = self.browser_factory(macros[macro_name].get("browser"))
        opened = False

        try:
            for step in steps:
                if step.get("handler") != "browser":
                    continue
                action = step.get("action")
                params = step.get("params", {})
                entry = {"action": action, "params": params, "status": "SKIP", "detail": ""}

                try:
                    if action == "open_registered_site":
                        browser.open_registered_site(params["site_key"])
                        opened = True
                        entry["status"] = "OK"
                        entry["detail"] = "サイトを開けました"
                    elif action == "click_by_text" and opened:
                        ok = browser.check_clickable_exists(params["text_hint"])
                        entry["status"] = "OK" if ok else "NG"
                        entry["detail"] = "" if ok else "表示テキストに一致する要素が見つかりません"
                    elif action == "type_by_text" and opened:
                        ok = browser.check_input_exists(params["label_hint"])
                        entry["status"] = "OK" if ok else "NG"
                        entry["detail"] = "" if ok else "入力欄が見つかりません"
                    elif action == "select_by_text" and opened:
                        ok = browser.check_select_exists(params["label_hint"])
                        entry["status"] = "OK" if ok else "NG"
                        entry["detail"] = "" if ok else "ドロップダウンが見つかりません"
                    else:
                        entry["detail"] = "この確認方式では検査できない手順です(実行して確認してください)"
                except Exception as e:  # noqa: BLE001
                    entry["status"] = "NG"
                    entry["detail"] = str(e)

                report.append(entry)
        finally:
            try:
                browser.close()
            except Exception:  # noqa: BLE001
                pass

        return report

    def check_all(self) -> dict[str, list[dict]]:
        results = {}
        for name in self.list_macro_names():
            try:
                results[name] = self.check_macro(name)
            except Exception as e:  # noqa: BLE001
                results[name] = [{"action": "-", "params": {}, "status": "NG", "detail": str(e)}]
        return results
