"""
codegen: 登録済みマクロ(config/macros.json)を、同じ処理を行う単体の.py
スクリプトに変換する。main.py --run-macro と同等の処理を、JSONを都度
解釈する形ではなく、直接呼び出しのコードとして固定化するイメージ。

「動作確認が済んで問題なく動くと分かったマクロ」を、他の人に配布したり、
PyInstallerでexe化したりしやすい単体スクリプトに落とし込む用途を想定している。

対応範囲:
- 制御構文(handler="control")を含まない、一直線(上から順に実行する)の
  マクロは、手順をそのまま上から並べたコードに変換する。
- For繰り返し/If分岐/Goto/ラベル(label/goto/if_goto/for_start/for_end)を
  含むマクロは、手順番号(pc)をたどる形のコードに変換する。ジャンプや
  入れ子のforを含め、実行エンジン(engine.executor)と全く同じ順序で
  実行されるよう、繰り返しの状態管理(for_startで終了値を一度だけ評価、
  for_endでカウンタを+1して本体の先頭へ戻る、本体は最低1回実行される、
  等)も同じ手順でコードに埋め込んでいる。
  変数への値の設定(set_value)や型変換(to_str/to_int/to_float)もそのまま
  変換できる。
- 各手順の「実行後の確認」(verify)は生成しない。動作確認済みのマクロを
  前提にしているため、確認なしでそのまま次の手順へ進む。リトライ回数は
  そのまま引き継ぐ。
- {{変数名}} 等のテンプレート解決は、生成したコードの中に埋め込まず、
  実行時に engine.executor._substitute をそのまま呼び出す形にしている
  (列文字の加減算等、複雑な書式を再現するコードを別途生成する必要が
  なく、本体と全く同じ挙動になる)。{{clipboard}}(実行時のクリップボード)も
  同様に engine.executor.builtin_variables で解決する。そのため生成した
  スクリプトは、このプロジェクトのフォルダ内(main.pyと同じ階層)に置いて
  実行する必要がある(コピーして別の場所だけに持ち出しても動かない)。
"""
from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
from pathlib import Path

_PROJECT_DIR = Path(__file__).resolve().parent.parent

_FLOW_CONTROL_ACTIONS = {"label", "goto", "if_goto", "for_start", "for_end"}
_VALUE_CONTROL_ACTIONS = {"set_value", "to_str", "to_int", "to_float"}

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


def _has_flow_control(steps: list[dict]) -> bool:
    return any(
        s.get("handler") == "control" and s.get("action") in _FLOW_CONTROL_ACTIONS for s in steps
    )


def _scan_flow(steps: list[dict]) -> dict[str, int]:
    """制御構文の整合性を確認し、ラベル名 -> 手順番号(1始まり)を返す。
    実行エンジンが実行開始時に行う事前チェックと同じ内容を、変換時点で行う
    (壊れたマクロから、実行時にしかエラーにならないスクリプトを作らないため)。
    """
    problems: list[str] = []
    labels: dict[str, int] = {}
    for_stack: list[int] = []
    for i, step in enumerate(steps, start=1):
        if step.get("handler") != "control":
            continue
        action = step.get("action")
        params = step.get("params", {})
        if action == "label":
            name = params.get("name")
            if not name:
                problems.append(f"{i}番目: ラベルに名前がありません")
            elif name in labels:
                problems.append(f"{i}番目: ラベル名 '{name}' が重複しています")
            else:
                labels[name] = i
        elif action == "for_start":
            for_stack.append(i)
        elif action == "for_end":
            if not for_stack:
                problems.append(f"{i}番目: 対応する「繰り返しを開始する」がありません")
            else:
                for_stack.pop()
        elif action not in _FLOW_CONTROL_ACTIONS and action not in _VALUE_CONTROL_ACTIONS:
            problems.append(f"{i}番目: 未知の制御構文です({action})")
    if for_stack:
        problems.append(f"対応する「繰り返しを終了する」が無い「繰り返しを開始する」があります({for_stack}番目)")
    for i, step in enumerate(steps, start=1):
        if step.get("handler") == "control" and step.get("action") in ("goto", "if_goto"):
            label = step.get("params", {}).get("label")
            if label not in labels:
                problems.append(f"{i}番目: ジャンプ先のラベル '{label}' が見つかりません")
    if problems:
        raise UnsupportedMacroError(
            "このマクロは制御構文(繰り返し/分岐/ジャンプ)に問題があるため、"
            ".pyスクリプトに変換できません: " + " / ".join(problems)
        )
    return labels


def check_convertible(macro_def: dict) -> None:
    """変換できるマクロかを確認する。制御構文(for/if/goto/label)自体は
    変換できるが、ジャンプ先が無い・forの対応が合わない等、壊れている場合は
    UnsupportedMacroError を送出する。
    """
    _scan_flow(macro_def.get("steps", []))


def _cast_expr(action: str) -> str:
    if action == "to_str":
        return "str(_v)"
    if action == "to_int":
        return "int(float(_v))"
    return "float(_v)"


def _step_body(i: int, step: dict, ctx, labels: dict[str, int], uses_clipboard: bool) -> list[str]:
    """1手順分のコード行(インデントなし)を返す。ctx(value_repr)は、その値を
    {{}}解決するときに渡す辞書を表すコード片を返す関数。
    pc(手順番号)の更新は含まない(呼び出し側が、制御構文を含むマクロのときだけ付ける)。
    """
    handler_name = step.get("handler")
    action_name = step.get("action")
    params = step.get("params", {})
    store_as = step.get("store_as")
    out: list[str] = []

    if handler_name == "control":
        if action_name == "set_value":
            out.append(f"# ステップ{i}: 変数に値を設定")
            v = repr(params.get("value"))
            out.append(f"variables[{store_as!r}] = _substitute({v}, {ctx(v)})")
        elif action_name in ("to_str", "to_int", "to_float"):
            out.append(f"# ステップ{i}: 型変換({action_name})")
            v = repr(params.get("value"))
            out.append(f"_v = _substitute({v}, {ctx(v)})")
            out.append(f"variables[{store_as!r}] = {_cast_expr(action_name)}")
        return out

    _, _, var = _HANDLER_INFO[handler_name]
    retry_cfg = step.get("retry") or {}
    retry_count = int(retry_cfg.get("count", 0))
    retry_interval = float(retry_cfg.get("interval_seconds", 2))
    out.append(f"# ステップ{i}: {handler_name}.{action_name}")
    text_arg = ", text=_builtin_text" if uses_clipboard else ""
    out.append(
        f"result = run_step({var}.{action_name}, {params!r}, slots, variables, "
        f"retry_count={retry_count}, retry_interval={retry_interval}{text_arg})"
    )
    if store_as:
        out.append(f"variables[{store_as!r}] = result")
    return out


def _flow_step(i: int, step: dict, ctx, labels: dict[str, int]) -> list[str]:
    """制御構文(label/goto/if_goto/for_start/for_end)1手順分のコード行を返す。
    実行エンジン(engine.executor.MacroExecutor.run)の同名の処理と同じ動き。
    """
    action = step.get("action")
    params = step.get("params", {})
    out: list[str] = []
    if action == "label":
        out.append(f"# ステップ{i}: ラベル {params.get('name')!r}(目印のみ。何もしない)")
        out.append("pc += 1")
    elif action == "goto":
        label = params.get("label")
        out.append(f"# ステップ{i}: ラベル {label!r}(ステップ{labels[label]})へジャンプ")
        out.append(f"pc = {labels[label]}")
    elif action == "if_goto":
        label = params.get("label")
        left, right = repr(params.get("left")), repr(params.get("right"))
        out.append(f"# ステップ{i}: 条件を満たせばラベル {label!r}(ステップ{labels[label]})へジャンプ")
        out.append(f"_left = _substitute({left}, {ctx(left)})")
        out.append(f"_right = _substitute({right}, {ctx(right)})")
        out.append(f"pc = {labels[label]} if _evaluate_condition(_left, {params.get('op')!r}, _right) else pc + 1")
    elif action == "for_start":
        var_name = params.get("var", "i")
        start, end = repr(params.get("start", 0)), repr(params.get("end"))
        out.append(f"# ステップ{i}: 繰り返しを開始する(for {var_name})")
        out.append(f"if loop_stack and loop_stack[-1]['for_start_idx'] == {i}:")
        out.append("    pass  # gotoで戻ってきた2周目以降(設定済みなので初期化し直さない)")
        out.append("else:")
        out.append(f"    _start = _substitute({start}, {ctx(start)})")
        out.append(f"    _end = _substitute({end}, {ctx(end)})")
        out.append(
            f"    loop_stack.append({{'var': {var_name!r}, 'end': int(float(_end)), 'for_start_idx': {i}}})"
        )
        out.append(f"    variables[{var_name!r}] = int(float(_start))")
        out.append("pc += 1")
    elif action == "for_end":
        out.append(f"# ステップ{i}: 繰り返しを終了する(next)")
        out.append("_loop = loop_stack[-1]")
        out.append("variables[_loop['var']] += 1")
        out.append("if variables[_loop['var']] <= _loop['end']:")
        out.append("    pc = _loop['for_start_idx'] + 1  # 本体の先頭へ戻る")
        out.append("else:")
        out.append("    loop_stack.pop()")
        out.append("    pc += 1")
    return out


def generate_script(macro_name: str, macro_def: dict, output_path: Path) -> None:
    """macro_defの内容を、直接呼び出しの単体.pyスクリプトとして output_path に
    書き出す(check_convertibleは呼び出し側で先に確認しておくこと)。
    """
    steps = macro_def.get("steps", [])
    labels = _scan_flow(steps)
    flow_mode = _has_flow_control(steps)
    uses_clipboard = "clipboard" in json.dumps(steps, ensure_ascii=False)

    used_handlers = sorted({s["handler"] for s in steps if s.get("handler") != "control"})
    required_slots = macro_def.get("required_slots", [])
    browser_choice = macro_def.get("browser") or "chrome"

    def ctx(value_repr: str) -> str:
        if uses_clipboard:
            return f"{{**builtin_variables({value_repr}, _builtin_text), **slots, **variables}}"
        return "{**slots, **variables}"

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
    imports = ["_substitute"]
    if flow_mode:
        imports.append("_evaluate_condition")
    if uses_clipboard:
        imports.append("builtin_variables")
    lines.append(f"from engine.executor import {', '.join(imports)}  # noqa: E402")
    if uses_clipboard:
        lines.append("from handlers.text_handler import TextHandler as _BuiltinText  # noqa: E402")
    for h in used_handlers:
        module, cls, _ = _HANDLER_INFO[h]
        lines.append(f"from {module} import {cls}  # noqa: E402")
    lines.append("")
    lines.append('CONFIG_DIR = PROJECT_DIR / "config"')
    lines.append("")
    lines.append("")
    lines.append(
        "def run_step(fn, params_template, slots, variables, retry_count=0, retry_interval=2.0, text=None):"
    )
    if uses_clipboard:
        lines.append(
            "    params = _substitute(params_template, "
            "{**builtin_variables(params_template, text), **slots, **variables})"
        )
    else:
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
    if uses_clipboard:
        lines.append("    _builtin_text = _BuiltinText()")
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

    if not flow_mode:
        for i, step in enumerate(steps, start=1):
            for body_line in _step_body(i, step, ctx, labels, uses_clipboard):
                lines.append(f"    {body_line}")
            lines.append("")
    else:
        total = len(steps)
        lines.append("    # 制御構文(繰り返し/分岐/ジャンプ)を含むため、手順番号(pc)をたどって実行する。")
        lines.append("    # 実行エンジン(engine.executor)と同じ順序・同じ繰り返しの扱いになる。")
        lines.append("    loop_stack = []")
        lines.append("    pc = 1")
        lines.append(f"    while pc <= {total}:")
        for i, step in enumerate(steps, start=1):
            lines.append(f"        if pc == {i}:")
            if step.get("handler") == "control" and step.get("action") in _FLOW_CONTROL_ACTIONS:
                body = _flow_step(i, step, ctx, labels)
                for body_line in body:
                    lines.append(f"            {body_line}")
            else:
                for body_line in _step_body(i, step, ctx, labels, uses_clipboard):
                    lines.append(f"            {body_line}")
                lines.append("            pc += 1")
            lines.append("            continue")
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


def pyinstaller_available() -> bool:
    return importlib.util.find_spec("PyInstaller") is not None


def build_exe(script_path: Path) -> tuple[bool, str]:
    """PyInstallerで.pyファイルを単体exeに変換する(--onefile)。
    戻り値: (成功したか, 成功時はexeのパス・失敗時はエラーメッセージ)。
    呼び出し元のUI(CLI/GUI)に依存しないよう、printは行わない。
    """
    if not pyinstaller_available():
        return False, (
            "PyInstallerがインストールされていません。"
            "'pip install pyinstaller' を実行してから、もう一度お試しください。"
        )

    dist_dir = script_path.parent / "dist"
    build_dir = script_path.parent / "build"
    result = subprocess.run(
        [
            sys.executable, "-m", "PyInstaller", "--onefile",
            "--distpath", str(dist_dir), "--workpath", str(build_dir),
            "--specpath", str(script_path.parent), str(script_path),
        ],
        cwd=_PROJECT_DIR,
    )
    if result.returncode != 0:
        return False, "exe化に失敗しました(PyInstallerのログを確認してください)。"
    exe_path = dist_dir / f"{script_path.stem}.exe"
    return True, str(exe_path)
