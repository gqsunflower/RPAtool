"""
codegen: 登録済みマクロ(config/macros.json)を、同じ処理を行う単体の.py
スクリプトに変換する。main.py --run-macro と同等の処理を、JSONを都度
解釈する形ではなく、直接呼び出しのコードとして固定化するイメージ。

「動作確認が済んで問題なく動くと分かったマクロ」を、他の人に配布したり、
PyInstallerでexe化したりしやすい単体スクリプトに落とし込む用途を想定している。

対応範囲(v1):
- 基本は一直線(上から順に実行する)のマクロのみ対応する。
  For繰り返し/If分岐/Goto(handler="control"のうち
  label/goto/if_goto/for_start/for_end)を含むマクロは、まだ変換に対応
  していない(check_convertibleでエラーにする)。変数への値の設定
  (set_value)や型変換(to_str/to_int/to_float)は分岐・繰り返しを伴わない
  ため対応している。
- 各手順の「実行後の確認」(verify)は生成しない。動作確認済みのマクロを
  前提にしているため、確認なしでそのまま次の手順へ進む。リトライ回数は
  そのまま引き継ぐ。
- {{変数名}} 等のテンプレート解決は、生成したコードの中に埋め込まず、
  実行時に engine.executor._substitute をそのまま呼び出す形にしている
  (列文字の加減算等、複雑な書式を再現するコードを別途生成する必要が
  なく、本体と全く同じ挙動になる)。そのため生成したスクリプトは、
  このプロジェクトのフォルダ内(main.pyと同じ階層)に置いて実行する
  必要がある(コピーして別の場所だけに持ち出しても動かない)。
"""
from __future__ import annotations

from pathlib import Path

_UNSUPPORTED_CONTROL_ACTIONS = {"label", "goto", "if_goto", "for_start", "for_end"}

_UNSUPPORTED_ACTION_LABELS = {
    "label": "ラベルを置く",
    "goto": "指定したラベルへジャンプする(goto)",
    "if_goto": "条件を満たしたらジャンプする(IF文)",
    "for_start": "繰り返しを開始する(for)",
    "for_end": "繰り返しを終了する(next)",
}

# handler名 -> (importするモジュール, クラス名, 生成コード内での変数名)
_HANDLER_INFO = {
    "excel": ("handlers.excel_handler", "ExcelHandler", "excel"),
    "pdf": ("handlers.pdf_handler", "PdfHandler", "pdf"),
    "browser": ("handlers.browser_handler", "BrowserHandler", "browser"),
    "explorer": ("handlers.explorer_handler", "ExplorerHandler", "explorer"),
    "process": ("handlers.process_handler", "ProcessHandler", "process"),
    "desktop": ("handlers.desktop_handler", "DesktopHandler", "desktop"),
    "text": ("handlers.text_handler", "TextHandler", "text"),
    "list": ("handlers.list_handler", "ListHandler", "list_"),
}


class UnsupportedMacroError(Exception):
    pass


def check_convertible(macro_def: dict) -> None:
    """For/If/Gotoといった制御構文が含まれていないか確認する。
    含まれている場合は UnsupportedMacroError を送出する。
    """
    offending = []
    for i, step in enumerate(macro_def.get("steps", []), start=1):
        action = step.get("action")
        if step.get("handler") == "control" and action in _UNSUPPORTED_CONTROL_ACTIONS:
            offending.append(f"{i}番目({_UNSUPPORTED_ACTION_LABELS.get(action, action)})")
    if offending:
        raise UnsupportedMacroError(
            "このマクロにはFor繰り返し/If分岐/Gotoといった制御構文が含まれているため、"
            ".pyスクリプトへの変換にはまだ対応していません"
            f"({', '.join(offending)})。"
            "一直線(上から順に実行する)構成のマクロのみ変換できます。"
        )


def _cast_expr(action: str) -> str:
    if action == "to_str":
        return "str(_v)"
    if action == "to_int":
        return "int(float(_v))"
    return "float(_v)"


def generate_script(macro_name: str, macro_def: dict, output_path: Path) -> None:
    """macro_defの内容を、直接呼び出しの単体.pyスクリプトとして output_path に
    書き出す(check_convertibleは呼び出し側で先に確認しておくこと)。
    """
    check_convertible(macro_def)

    steps = macro_def.get("steps", [])
    used_handlers = sorted({s["handler"] for s in steps if s.get("handler") != "control"})
    required_slots = macro_def.get("required_slots", [])
    browser_choice = macro_def.get("browser") or "chrome"

    lines: list[str] = []
    lines.append('"""')
    lines.append(f"マクロ '{macro_name}' から自動生成されたスクリプト。")
    lines.append("RPAツールのフォルダ内(main.pyと同じ階層)に置いて実行してください。")
    lines.append("再生成すると上書きされるため、手直しする場合は別名でコピーしてから編集すること。")
    lines.append('"""')
    lines.append("from __future__ import annotations")
    lines.append("")
    lines.append("import json")
    lines.append("import sys")
    lines.append("import time")
    lines.append("from pathlib import Path")
    lines.append("")
    lines.append("PROJECT_DIR = Path(__file__).resolve().parent")
    lines.append("while not (PROJECT_DIR / \"handlers\").exists() and PROJECT_DIR != PROJECT_DIR.parent:")
    lines.append("    PROJECT_DIR = PROJECT_DIR.parent")
    lines.append("if str(PROJECT_DIR) not in sys.path:")
    lines.append("    sys.path.insert(0, str(PROJECT_DIR))")
    lines.append("")
    lines.append("from engine.executor import _substitute  # noqa: E402")
    for h in used_handlers:
        module, cls, _ = _HANDLER_INFO[h]
        lines.append(f"from {module} import {cls}  # noqa: E402")
    lines.append("")
    lines.append('CONFIG_DIR = PROJECT_DIR / "config"')
    lines.append("")
    lines.append("")
    lines.append(
        "def run_step(fn, params_template, slots, variables, retry_count=0, retry_interval=2.0):"
    )
    lines.append("    params = _substitute(params_template, {**slots, **variables})")
    lines.append("    last_err = None")
    lines.append("    for attempt in range(retry_count + 1):")
    lines.append("        try:")
    lines.append("            return fn(**params)")
    lines.append("        except Exception as e:  # noqa: BLE001")
    lines.append("            last_err = e")
    lines.append("            if attempt < retry_count:")
    lines.append("                time.sleep(retry_interval)")
    lines.append("    raise last_err")
    lines.append("")
    lines.append("")
    lines.append("def main() -> None:")
    lines.append("    slots = {}")
    for slot in required_slots:
        prompt = f"'{slot}' を入力してください(JSON形式で書けばdict/listも可。単純な文字列はそのままでOK): "
        lines.append(f"    _raw = input({prompt!r})")
        lines.append("    try:")
        lines.append(f"        slots[{slot!r}] = json.loads(_raw)")
        lines.append("    except json.JSONDecodeError:")
        lines.append(f"        slots[{slot!r}] = _raw")
    lines.append("    variables = {}")
    lines.append("")

    for h in used_handlers:
        _, _, var = _HANDLER_INFO[h]
        if h == "browser":
            lines.append(
                f'    {var} = BrowserHandler(CONFIG_DIR / "whitelist_urls.json", '
                f"headless=False, browser={browser_choice!r})"
            )
        elif h == "process":
            lines.append(f'    {var} = ProcessHandler(CONFIG_DIR / "exec_whitelist.json")')
        else:
            _, cls, _ = _HANDLER_INFO[h]
            lines.append(f"    {var} = {cls}()")
    lines.append("")

    for i, step in enumerate(steps, start=1):
        handler_name = step.get("handler")
        action_name = step.get("action")
        params = step.get("params", {})
        store_as = step.get("store_as")

        if handler_name == "control":
            if action_name == "set_value":
                lines.append(f"    # ステップ{i}: 変数に値を設定")
                lines.append(
                    f"    variables[{store_as!r}] = _substitute({params.get('value')!r}, "
                    f"{{**slots, **variables}})"
                )
            elif action_name in ("to_str", "to_int", "to_float"):
                lines.append(f"    # ステップ{i}: 型変換({action_name})")
                lines.append(
                    f"    _v = _substitute({params.get('value')!r}, {{**slots, **variables}})"
                )
                lines.append(f"    variables[{store_as!r}] = {_cast_expr(action_name)}")
            lines.append("")
            continue

        _, _, var = _HANDLER_INFO[handler_name]
        retry_cfg = step.get("retry") or {}
        retry_count = int(retry_cfg.get("count", 0))
        retry_interval = float(retry_cfg.get("interval_seconds", 2))
        lines.append(f"    # ステップ{i}: {handler_name}.{action_name}")
        lines.append(
            f"    result = run_step({var}.{action_name}, {params!r}, slots, variables, "
            f"retry_count={retry_count}, retry_interval={retry_interval})"
        )
        if store_as:
            lines.append(f"    variables[{store_as!r}] = result")
        lines.append("")

    if "browser" in used_handlers:
        lines.append("    browser.close()")
    lines.append('    print("完了しました。")')
    lines.append("")
    lines.append("")
    lines.append('if __name__ == "__main__":')
    lines.append("    main()")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
