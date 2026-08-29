"""
疑似ローカルAI (RPA特化) - CLIエントリーポイント

使い方:
    python main.py

対話的にコマンドを入力すると、
  1. IntentEngine が意図(実行すべきマクロ)を判定
  2. 必要なスロット(パラメータ)が未入力なら質問
  3. MacroExecutor が macros.json の手順どおりに実行

主なコマンド(指示欄にそのまま入力):
    操作を登録     ... Excel・PDF・Web・エクスプローラー・exe/py・デスクトップ・テキスト加工・リスト・制御構文の操作を組み合わせてQ&A形式で記録する
    GUIで操作を登録 ... 同じ内容をボタン操作中心のGUI(gui_recorder.py)で記録する
    確認設定       ... 各手順の「実行後の確認」を行う/省略するを切り替える
    リトライ設定   ... 各手順の失敗時リトライ回数を後から設定・変更する
    手順編集       ... 登録済みの手順を並び替え・削除・挿入する
    ステップ実行   ... F8のように1手順ずつ実行しながら動作確認する
    実行ログ出力   ... これまでの実行記録をExcelに書き出す
    パイプライン作成 / パイプライン実行 ... 登録済みマクロ同士を繋げて連続実行する
    ヘルスチェック ... 登録済みボタン等が今も見つかるか(非破壊)確認する

--dry-run を付けると実際には実行せず、実行予定のステップだけ表示します。
--healthcheck-all を付けると、全マクロのヘルスチェックだけを実行して終了します
(OSのタスクスケジューラ/cron等から定期実行する用途)。
--run-macro NAME [--slots JSON] を付けると、対話なしで指定マクロを1回だけ
実行して終了します(終了コード0=成功、1=失敗)。他のツール(Power Automate
Desktopから呼び出すPythonスクリプト等)からこのRPAツールを1コマンドで
起動したい場合に使います。例:
    python main.py --run-macro monthly_report --slots "{\"path\": \"C:\\\\data\\\\a.xlsx\"}"
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime
from pathlib import Path

from engine.backup import backup_file
from engine.executor import MacroEditRequested, MacroExecutor
from engine.health_check import HealthChecker
from engine.intent_engine import IntentEngine
from engine.pipeline import PipelineRunner
from engine.recorder import MacroRecorder
from engine.run_logger import RunLogger
from handlers.browser_handler import BrowserHandler
from handlers.desktop_handler import DesktopHandler
from handlers.excel_handler import ExcelHandler
from handlers.explorer_handler import ExplorerHandler
from handlers.list_handler import ListHandler
from handlers.pdf_handler import PdfHandler
from handlers.process_handler import ProcessHandler
from handlers.text_handler import TextHandler

BASE_DIR = Path(__file__).resolve().parent
CONFIG_DIR = BASE_DIR / "config"
LOG_DIR = BASE_DIR / "logs"
SCREENSHOT_DIR = LOG_DIR / "screenshots"
BACKUP_DIR = CONFIG_DIR / "backups"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(LOG_DIR / "rpa_local_ai.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger("rpa_local_ai.main")


def build_handlers(headless: bool) -> dict:
    return {
        "excel": ExcelHandler(),
        "pdf": PdfHandler(),
        "browser": BrowserHandler(CONFIG_DIR / "whitelist_urls.json", headless=headless),
        "explorer": ExplorerHandler(),
        "process": ProcessHandler(CONFIG_DIR / "exec_whitelist.json"),
        "desktop": DesktopHandler(),
        "text": TextHandler(),
        "list": ListHandler(),
    }


def run_macro_noninteractive(macro_name: str, slots_json: str | None, headless: bool) -> bool:
    """--run-macro 用: 対話なしで1つのマクロを実行して終了する。
    Power Automate Desktop等の外部ツールから、このRPAツールを1コマンドで
    呼び出すための入口。失敗時は3択メニューを出さず、そのまま失敗として
    終了コード1を返す(無人実行を想定しているため)。
    戻り値: 成功したか。
    """
    try:
        slots = json.loads(slots_json) if slots_json else {}
    except json.JSONDecodeError as e:
        print(f"⚠ --slots のJSONが不正です: {e}")
        return False

    executor = build_executor(headless)
    try:
        results = executor.run(macro_name, slots, dry_run=False)
    except Exception as e:  # noqa: BLE001
        logger.exception("非対話実行でエラーが発生しました: %s", macro_name)
        print(f"⚠ マクロ '{macro_name}' の実行に失敗しました: {e}")
        return False

    print(f"マクロ '{macro_name}' が完了しました。")
    for r in results:
        print(f"  - {r}")
    return True


def build_executor(headless: bool) -> MacroExecutor:
    run_logger = RunLogger(LOG_DIR / "execution_log.csv")
    return MacroExecutor(
        CONFIG_DIR / "macros.json",
        build_handlers(headless),
        run_logger=run_logger,
        screenshot_dir=SCREENSHOT_DIR,
    )


def prompt_for_slot(slot_name: str) -> str:
    raw = input(f"  '{slot_name}' を入力してください "
                f"(JSON形式で書けばdict/listも可。単純な文字列はそのままでOK): ").strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return raw


# ---------- 確認設定(後から確認の有無を切り替える) ----------

def _verify_status_label(step: dict) -> str:
    verify = step.get("verify") or {"type": "none"}
    vtype = verify.get("type", "none")
    if vtype == "none":
        return "確認なし(そもそも確認方法が未設定)"
    if step.get("verify_skip", False):
        return f"確認は省略中 (設定は残っています: {vtype})"
    return f"確認あり ({vtype})"


def manage_verification(config_dir: Path) -> None:
    """登録済みマクロの各手順について、実行後の確認動作を
    行う/省略する を後から個別に切り替えるためのメニュー。
    ほぼ失敗しないとわかっている手順だけ確認を省略する、という用途を想定。
    """
    macros_path = config_dir / "macros.json"

    while True:
        with open(macros_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        macros: dict = data.get("macros", {})
        names = list(macros.keys())
        if not names:
            print("  登録済みマクロがありません。\n")
            return

        print("\n確認動作を設定するマクロを選んでください:")
        for i, name in enumerate(names, start=1):
            print(f"  {i}. {name}")
        print("  0. 終了する")
        choice = input("番号> ").strip()
        if choice == "0":
            return
        try:
            macro_name = names[int(choice) - 1]
        except (ValueError, IndexError):
            print("  入力が正しくありません\n")
            continue

        steps = macros[macro_name].get("steps", [])
        while True:
            print(f"\n『{macro_name}』の手順:")
            for i, step in enumerate(steps, start=1):
                print(f"  {i}. {step['handler']}.{step['action']} — {_verify_status_label(step)}")
            print("  番号を入力するとその手順の確認の有無を切り替えます。0で戻ります。")
            s = input("番号> ").strip()
            if s == "0":
                break
            try:
                idx = int(s) - 1
                step = steps[idx]
            except (ValueError, IndexError):
                print("  入力が正しくありません")
                continue

            verify = step.get("verify") or {"type": "none"}
            if verify.get("type", "none") == "none":
                print("  この手順にはそもそも確認方法が登録されていないため、切り替えられません。")
                print("  (確認したい場合は '操作を登録' で登録し直してください)")
                continue

            step["verify_skip"] = not step.get("verify_skip", False)
            backup_file(macros_path, BACKUP_DIR)
            with open(macros_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            new_state = "省略する" if step["verify_skip"] else "確認する"
            print(f"  → この手順は今後『{new_state}』ようになりました")


def _retry_status_label(step: dict) -> str:
    retry = step.get("retry") or {}
    count = retry.get("count", 0)
    if not count:
        return "リトライなし"
    interval = retry.get("interval_seconds", 2)
    return f"最大{count}回再試行(間隔{interval}秒)"


def manage_retry(config_dir: Path) -> None:
    """登録済みマクロの各手順について、失敗時の自動リトライ回数を
    後から個別に設定・変更するためのメニュー。制御構文(control)の
    手順はハンドラを呼び出さないためリトライの対象外。
    """
    macros_path = config_dir / "macros.json"

    while True:
        with open(macros_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        macros: dict = data.get("macros", {})
        names = list(macros.keys())
        if not names:
            print("  登録済みマクロがありません。\n")
            return

        print("\nリトライ回数を設定するマクロを選んでください:")
        for i, name in enumerate(names, start=1):
            print(f"  {i}. {name}")
        print("  0. 終了する")
        choice = input("番号> ").strip()
        if choice == "0":
            return
        try:
            macro_name = names[int(choice) - 1]
        except (ValueError, IndexError):
            print("  入力が正しくありません\n")
            continue

        steps = macros[macro_name].get("steps", [])
        while True:
            print(f"\n『{macro_name}』の手順:")
            for i, step in enumerate(steps, start=1):
                if step.get("handler") == "control":
                    print(f"  {i}. {step['handler']}.{step['action']} — (制御構文のため対象外)")
                else:
                    print(f"  {i}. {step['handler']}.{step['action']} — {_retry_status_label(step)}")
            print("  番号を入力するとその手順のリトライ回数を設定します。0で戻ります。")
            s = input("番号> ").strip()
            if s == "0":
                break
            try:
                idx = int(s) - 1
                step = steps[idx]
            except (ValueError, IndexError):
                print("  入力が正しくありません")
                continue

            if step.get("handler") == "control":
                print("  制御構文の手順はハンドラを呼び出さないため、リトライの対象外です。")
                continue

            raw = input("  最大リトライ回数を入力してください(0でリトライなし): ").strip()
            try:
                count = max(int(raw), 0)
            except ValueError:
                print("  数字で入力してください")
                continue

            if count == 0:
                step.pop("retry", None)
            else:
                interval_raw = input("  再試行の間隔(秒、空Enterで既定の2秒): ").strip()
                try:
                    interval = float(interval_raw) if interval_raw else 2.0
                except ValueError:
                    interval = 2.0
                step["retry"] = {"count": count, "interval_seconds": interval}

            backup_file(macros_path, BACKUP_DIR)
            with open(macros_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"  → この手順のリトライ設定を更新しました: {_retry_status_label(step)}")


# ---------- 手順編集(並び替え・削除・挿入) ----------

def _check_control_flow_integrity(steps: list[dict]) -> list[str]:
    """制御構文(ラベル参照・for対応)の整合性を簡易チェックし、
    問題があれば警告メッセージの一覧を返す(空なら問題なし)。
    並び替え・削除・挿入の後に呼び、参考情報として表示する
    (自動修正はしない。実行してみないと壊れているか分からない複雑な
    ケースまでは検出できないため、あくまで簡易チェック)。
    """
    warnings: list[str] = []
    labels = {
        s["params"]["name"] for s in steps
        if s.get("handler") == "control" and s.get("action") == "label"
    }
    for s in steps:
        if s.get("handler") != "control":
            continue
        if s.get("action") in ("goto", "if_goto"):
            label = s.get("params", {}).get("label")
            if label not in labels:
                warnings.append(f"'{label}' へのジャンプがありますが、そのラベルが見つかりません")

    depth = 0
    for s in steps:
        if s.get("handler") != "control":
            continue
        if s.get("action") == "for_start":
            depth += 1
        elif s.get("action") == "for_end":
            depth -= 1
            if depth < 0:
                warnings.append("対応する「繰り返しを開始する」より前に「繰り返しを終了する」があります")
                depth = 0
    if depth > 0:
        warnings.append("対応する「繰り返しを終了する」が無い「繰り返しを開始する」があります")

    return warnings


def _print_control_flow_warnings(steps: list[dict]) -> None:
    warnings = _check_control_flow_integrity(steps)
    for w in warnings:
        print(f"  ⚠ 制御構文の整合性チェック: {w}")


def _insert_step_interactive(config_dir: Path, macro_name: str, insert_before: int) -> bool:
    """既存マクロの指定位置(0始まり)の直前に、新しい手順を1つ以上
    対話形式で挿入する。内部的にはMacroRecorderの各領域メニューを
    そのまま再利用し、追加された手順だけを本来の位置へ差し込む。
    戻り値: 実際に何か追加できたか。
    """
    from engine.recorder import MacroRecorder

    macros_path = config_dir / "macros.json"
    data = json.loads(macros_path.read_text(encoding="utf-8"))
    macro = data["macros"][macro_name]
    existing_steps = macro["steps"]

    rec = MacroRecorder(config_dir)
    rec.steps = list(existing_steps)  # 既存の手順をそのまま引き継ぐ(末尾に追加される)
    rec.required_slots = list(macro.get("required_slots", []))

    # 既にWebサイトを開く手順があれば、挿入作業中もその状態を再現しておく
    # (ベストエフォート: 最後に開かれていたサイトだけを開き直す)
    last_site_key = None
    for s in existing_steps[:insert_before]:
        if s.get("handler") == "browser" and s.get("action") == "open_registered_site":
            last_site_key = s.get("params", {}).get("site_key")
    if last_site_key:
        try:
            rec.browser.open_registered_site(last_site_key)
            rec._site_opened = True
            print(f"  (挿入作業のため、参考にサイト '{last_site_key}' を開き直しました)")
        except Exception as e:  # noqa: BLE001
            print(f"  ⚠ サイトの再現に失敗しました(手動でサイトを開く操作から始めてください): {e}")

    # 同様に、Excelのブックを開く/切り替える手順があればベストエフォートで再現する
    # (テンプレート({{ }})を含むパスは実際の値が分からないためスキップする)
    for s in existing_steps[:insert_before]:
        if s.get("handler") != "excel":
            continue
        params = s.get("params", {})
        try:
            if s.get("action") == "load_workbook":
                path = params.get("path", "")
                if "{{" in str(path):
                    continue
                rec.excel.load_workbook(path, alias=params.get("alias"))
            elif s.get("action") == "switch_workbook":
                alias = params.get("alias")
                if alias in rec.excel.list_open_workbooks():
                    rec.excel.switch_workbook(alias)
        except Exception:  # noqa: BLE001
            pass  # 再現できなくても致命的ではないため、そのまま先へ進む

    before_count = len(rec.steps)

    print("\n挿入する手順の領域を選んでください:")
    print("  1) Excel  2) PDF  3) Webサイト  4) エクスプローラー  5) 実行ファイル")
    print("  6) デスクトップ  7) テキスト加工  8) リスト  9) 制御構文")
    domain_choice = input("番号> ").strip()
    domain_map = {
        "1": rec._record_excel_menu, "2": rec._record_pdf_menu, "3": rec._record_web_menu,
        "4": rec._record_explorer_menu, "5": rec._record_process_menu, "6": rec._record_desktop_menu,
        "7": rec._record_text_menu, "8": rec._record_list_menu, "9": rec._record_control_menu,
    }
    domain_fn = domain_map.get(domain_choice)
    if domain_fn is None:
        print("  入力が正しくありません。挿入を中止しました。\n")
        if rec._site_opened:
            rec.browser.close()
        return False

    domain_fn()  # このメニューは「0) 戻る」で抜けるまでループする(複数手順の追加も可)

    if rec._site_opened:
        rec.browser.close()

    new_steps = rec.steps[before_count:]
    if not new_steps:
        print("  → 新しい手順は追加されませんでした。\n")
        return False

    final_steps = existing_steps[:insert_before] + new_steps + existing_steps[insert_before:]
    macro["steps"] = final_steps
    macro["required_slots"] = rec.required_slots
    backup_file(macros_path, BACKUP_DIR)
    with open(macros_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  → {len(new_steps)}件の手順を{insert_before + 1}番目の位置に挿入しました。\n")
    _print_control_flow_warnings(final_steps)
    return True


def manage_steps(config_dir: Path) -> None:
    """登録済みマクロの手順を、後から並び替え・削除・挿入するためのメニュー。"""
    macros_path = config_dir / "macros.json"

    while True:
        data = json.loads(macros_path.read_text(encoding="utf-8"))
        macros: dict = data.get("macros", {})
        names = list(macros.keys())
        if not names:
            print("  登録済みマクロがありません。\n")
            return

        print("\n手順を編集するマクロを選んでください:")
        for i, name in enumerate(names, start=1):
            print(f"  {i}. {name}")
        print("  0. 終了する")
        choice = input("番号> ").strip()
        if choice == "0":
            return
        try:
            macro_name = names[int(choice) - 1]
        except (ValueError, IndexError):
            print("  入力が正しくありません\n")
            continue

        while True:
            data = json.loads(macros_path.read_text(encoding="utf-8"))
            steps = data["macros"][macro_name]["steps"]

            print(f"\n『{macro_name}』の手順:")
            for i, step in enumerate(steps, start=1):
                print(f"  {i}. {step['handler']}.{step['action']} {step.get('params', {})}")
            print("  1) 上へ移動する")
            print("  2) 下へ移動する")
            print("  3) 削除する")
            print("  4) 指定した位置の前に新しい手順を挿入する")
            print("  0) 戻る")
            action = input("番号> ").strip()

            if action == "0":
                break

            if action not in ("1", "2", "3", "4"):
                print("  0〜4のいずれかを入力してください\n")
                continue

            if action == "4":
                pos_raw = input(f"  何番目の手順の前に挿入しますか?(末尾に追加するなら{len(steps) + 1}): ").strip()
                try:
                    pos = int(pos_raw) - 1
                    if not (0 <= pos <= len(steps)):
                        raise ValueError
                except ValueError:
                    print("  入力が正しくありません\n")
                    continue
                _insert_step_interactive(config_dir, macro_name, pos)
                continue

            s = input("  対象の手順番号> ").strip()
            try:
                idx = int(s) - 1
                if not (0 <= idx < len(steps)):
                    raise ValueError
            except ValueError:
                print("  入力が正しくありません\n")
                continue

            if action == "1":
                if idx == 0:
                    print("  ⚠ すでに先頭です\n")
                    continue
                steps[idx - 1], steps[idx] = steps[idx], steps[idx - 1]
                backup_file(macros_path, BACKUP_DIR)
                with open(macros_path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                print("  → 上へ移動しました。\n")
                _print_control_flow_warnings(steps)

            elif action == "2":
                if idx == len(steps) - 1:
                    print("  ⚠ すでに末尾です\n")
                    continue
                steps[idx + 1], steps[idx] = steps[idx], steps[idx + 1]
                backup_file(macros_path, BACKUP_DIR)
                with open(macros_path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                print("  → 下へ移動しました。\n")
                _print_control_flow_warnings(steps)

            elif action == "3":
                removed = steps[idx]
                if not messagebox_confirm_cli(
                    f"'{removed['handler']}.{removed['action']}' を削除します。よろしいですか?"
                ):
                    print("  → 削除を取り消しました。\n")
                    continue
                steps.pop(idx)
                backup_file(macros_path, BACKUP_DIR)
                with open(macros_path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                print(f"  → 削除しました: {removed['handler']}.{removed['action']}\n")
                _print_control_flow_warnings(steps)


def messagebox_confirm_cli(message: str) -> bool:
    return input(f"  {message} (y/N): ").strip().lower() == "y"


# ---------- 失敗時メニュー(即時終了 / 手動補完して再開 / 修正画面を開いて終了) ----------

def cli_failure_prompt(step_number: int, total: int, step: dict, error: Exception) -> str:
    print(f"\n⚠ ステップ {step_number}/{total} ({step['handler']}.{step['action']}) が失敗しました:")
    print(f"   {error}")
    print("  どうしますか?")
    print("    1) 今すぐ中断する")
    print("    2) この操作を手動で終わらせてから、続きを再開する")
    print("    3) 中断して、この手順の修正画面を開く")
    choice = input("  番号> ").strip()

    if choice == "2":
        input("  手動での対応が終わったら Enter キーを押してください... ")
        return "resume"
    if choice == "3":
        return "edit"
    return "abort"


def open_step_editor(config_dir: Path, macro_name: str, step_number: int) -> None:
    """失敗したステップのparamsをその場で修正できる簡易エディタ。
    (「コード修正画面を開く」の実装: 対話でparamsを直接書き換え、
    可能であればOS標準のテキストエディタでもmacros.jsonを開く)
    """
    path = config_dir / "macros.json"
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    step = data["macros"][macro_name]["steps"][step_number - 1]
    print(f"\n--- 手順 {step_number} の修正 ({step['handler']}.{step['action']}) ---")
    print(f"  現在の内容: {json.dumps(step.get('params', {}), ensure_ascii=False)}")

    params = step.get("params", {})
    changed = False
    for key, val in list(params.items()):
        if not isinstance(val, str):
            continue
        new_val = input(f"  '{key}' の新しい値(変更しないならそのままEnter) [{val}]: ").strip()
        if new_val:
            params[key] = new_val
            changed = True

    if changed:
        backup_file(path, BACKUP_DIR)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print("  → 修正内容を保存しました。")
    else:
        print("  → 変更はありませんでした。")

    # ベストエフォートでOS標準のエディタ/ファイラーでも開く(失敗しても無視する)
    try:
        import platform
        import subprocess

        system = platform.system()
        if system == "Windows":
            import os as _os
            _os.startfile(str(path))  # type: ignore[attr-defined]
        elif system == "Darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        print(f"  (自動でエディタを開けなかったため、必要なら手動で開いてください: {path})")


# ---------- ステップ実行(F8相当)・途中スキップ開始 ----------

def cli_step_prompt(step_number: int, total: int, step: dict) -> str:
    print(f"\n--- ステップ {step_number}/{total}: {step['handler']}.{step['action']} ---")
    print(f"  params: {step.get('params', {})}")
    print("  [Enter] 1ステップだけ実行して次で止まる(F8相当)   "
          "[c] 続きを自動実行   [s] このステップをスキップ   [a] 中断")
    choice = input("  > ").strip().lower()
    if choice == "c":
        return "run"
    if choice == "s":
        return "skip"
    if choice == "a":
        return "abort"
    return "step"


def cli_step_result(step_number: int, total: int, step: dict, result) -> None:
    print(f"  実行結果: {result}")


def run_macro_stepwise(executor: MacroExecutor, macro_name: str, slots: dict) -> None:
    """指定マクロをEXCELマクロのF8のように1手順ずつ実行する動作確認モード。
    途中の番号を指定して、そこから開始することもできる(前段は手動で
    準備済みであることが前提)。
    """
    total_steps = len(executor.get_macro(macro_name)["steps"])
    print(f"『{macro_name}』は全{total_steps}ステップです。")
    start_raw = input(
        f"何番目のステップから開始しますか?(先頭からなら1。1〜{total_steps}): "
    ).strip()
    try:
        start_step = int(start_raw) if start_raw else 1
    except ValueError:
        start_step = 1

    try:
        results = executor.run(
            macro_name, slots,
            dry_run=False,
            start_step=start_step,
            on_step=cli_step_prompt,
            on_result=cli_step_result,
            on_failure=cli_failure_prompt,
        )
        print("\n  ステップ実行が完了しました。実行結果:")
        for r in results:
            print(f"    - {r}")
    except MacroEditRequested as e:
        print(f"\n手順 {e.step_number} を修正します。")
        open_step_editor(CONFIG_DIR, macro_name, e.step_number)
        print("修正のため、ここで終了します。")
        raise SystemExit(0)
    except Exception as e:  # noqa: BLE001
        logger.exception("ステップ実行中にエラーが発生しました")
        print(f"  ⚠ エラー: {e}")


def launch_gui_recorder() -> None:
    """gui_recorder.py を別プロセスとして起動し、終了まで待つ。
    別プロセスにするのは、TkinterのGUIイベントループとこのCLIの入力待ちを
    同じプロセス内で共存させないため(お互いをブロックしてしまう)。
    """
    import subprocess

    gui_path = BASE_DIR / "gui_recorder.py"
    if not gui_path.exists():
        print(f"  ⚠ GUIレコーダーが見つかりません: {gui_path}\n")
        return

    print("  GUIレコーダーを起動します(ウィンドウを閉じるとここに戻ります)...")
    try:
        subprocess.run([sys.executable, str(gui_path)])
    except Exception as e:  # noqa: BLE001
        print(f"  ⚠ GUIレコーダーの起動に失敗しました: {e}")
    print("  → GUIレコーダーを終了しました。\n")


def choose_macro(executor: MacroExecutor) -> str | None:
    names = list(executor.macros.keys())
    if not names:
        print("  登録済みマクロがありません。\n")
        return None
    print("対象のマクロを選んでください:")
    for i, name in enumerate(names, start=1):
        print(f"  {i}. {name}")
    choice = input("番号> ").strip()
    try:
        return names[int(choice) - 1]
    except (ValueError, IndexError):
        print("  入力が正しくありません\n")
        return None


# ---------- 実行ログの出力 ----------

def export_run_log(executor: MacroExecutor) -> None:
    if executor.run_logger is None:
        print("  実行ログが有効になっていません。\n")
        return
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = LOG_DIR / f"execution_log_{ts}.xlsx"
    saved = executor.run_logger.export_to_excel(out_path)
    print(f"  → 実行ログをExcelに出力しました: {saved}\n")


# ---------- パイプライン(複数マクロの連続実行) ----------

def create_pipeline(config_dir: Path, executor: MacroExecutor) -> None:
    names = list(executor.macros.keys())
    if not names:
        print("  登録済みマクロがありません。\n")
        return

    print("パイプラインに含めるマクロを、実行したい順番でカンマ区切りの番号で選んでください:")
    for i, name in enumerate(names, start=1):
        print(f"  {i}. {name}")
    raw = input("番号(例: 1,3,2)> ").strip()
    try:
        indices = [int(x.strip()) - 1 for x in raw.split(",") if x.strip()]
        macro_names = [names[i] for i in indices]
        if not macro_names:
            raise ValueError
    except (ValueError, IndexError):
        print("  入力が正しくありません\n")
        return

    pipeline_name = input("このパイプラインの保存名(半角英数字)を入力してください: ").strip()
    description = input("説明を一言でお願いします: ").strip()

    runner = PipelineRunner(config_dir, executor)
    runner.save_pipeline(pipeline_name, description, macro_names)
    print(f"  → パイプライン '{pipeline_name}' を保存しました "
          f"({' → '.join(macro_names)})\n")


def run_pipeline(config_dir: Path, executor: MacroExecutor, dry_run: bool) -> None:
    runner = PipelineRunner(config_dir, executor)
    names = runner.list_names()
    if not names:
        print("  登録済みのパイプラインがありません。'パイプライン作成' で作成してください。\n")
        return

    print("実行するパイプラインを選んでください:")
    for i, name in enumerate(names, start=1):
        pipeline = runner.get_pipeline(name)
        macros_str = " → ".join(m["macro"] for m in pipeline.get("macros", []))
        print(f"  {i}. {name}: {pipeline.get('description', '')} [{macros_str}]")
    choice = input("番号> ").strip()
    try:
        pipeline_name = names[int(choice) - 1]
    except (ValueError, IndexError):
        print("  入力が正しくありません\n")
        return

    def slot_prompt_fn(macro_name: str, slot_name: str):
        print(f"  [{macro_name}]")
        return prompt_for_slot(slot_name)

    try:
        results = runner.run(
            pipeline_name, slot_prompt_fn, dry_run=dry_run, on_failure=cli_failure_prompt
        )
        print("\n  パイプラインの実行が完了しました:")
        for macro_name, macro_results in results.items():
            print(f"  ■ {macro_name}")
            for r in macro_results:
                print(f"      - {r}")
    except MacroEditRequested as e:
        print(f"\n手順 {e.step_number} を修正します。")
        open_step_editor(config_dir, pipeline_name, e.step_number)
        print("修正のため、プログラムを終了します。")
        raise SystemExit(0)
    except Exception as e:  # noqa: BLE001
        logger.exception("パイプライン実行中にエラーが発生しました")
        print(f"  ⚠ エラー: {e}")
    print()


# ---------- ヘルスチェック(登録済み要素が今も見つかるかの非破壊確認) ----------

def _print_health_report(macro_name: str, report: list[dict]) -> None:
    print(f"\n■ {macro_name}")
    if not report:
        print("  (確認対象のブラウザ手順がありません)")
        return
    for entry in report:
        mark = {"OK": "✅", "NG": "❌", "SKIP": "―"}.get(entry["status"], "?")
        detail = f" ({entry['detail']})" if entry["detail"] else ""
        print(f"  {mark} {entry['action']} {entry['params']}{detail}")


def run_health_check(config_dir: Path, headless: bool = True, target: str | None = None) -> bool:
    """全マクロ(またはtargetで指定した1つ)のヘルスチェックを行い、
    問題があれば標準出力に表示する。戻り値は「全てOKだったか」。
    """
    def factory():
        return BrowserHandler(config_dir / "whitelist_urls.json", headless=headless)

    checker = HealthChecker(config_dir, factory)
    all_ok = True

    names = [target] if target else checker.list_macro_names()
    if not names:
        print("  登録済みマクロがありません。\n")
        return True

    for macro_name in names:
        try:
            report = checker.check_macro(macro_name)
        except Exception as e:  # noqa: BLE001
            print(f"\n■ {macro_name}\n  ⚠ ヘルスチェックに失敗しました: {e}")
            all_ok = False
            continue
        _print_health_report(macro_name, report)
        if any(entry["status"] == "NG" for entry in report):
            all_ok = False

    print()
    print("この確認は同一画面内の要素検出に限られます。複数ページにまたがる"
          "手順の後半は正しく検証できない場合があるため、参考情報としてご利用ください。\n")
    return all_ok


def health_check_menu(executor: MacroExecutor) -> None:
    names = list(executor.macros.keys())
    if not names:
        print("  登録済みマクロがありません。\n")
        return
    print("ヘルスチェックの対象を選んでください:")
    print("  0. すべてのマクロ")
    for i, name in enumerate(names, start=1):
        print(f"  {i}. {name}")
    choice = input("番号> ").strip()
    if choice == "0":
        run_health_check(CONFIG_DIR, headless=True, target=None)
        return
    try:
        macro_name = names[int(choice) - 1]
    except (ValueError, IndexError):
        print("  入力が正しくありません\n")
        return
    run_health_check(CONFIG_DIR, headless=True, target=macro_name)


# ---------- メインREPL ----------

def run_repl(dry_run: bool, headless: bool) -> None:
    intent_engine = IntentEngine(CONFIG_DIR / "intents.json")
    executor = build_executor(headless)

    print("=== 疑似ローカルAI (RPA特化) ===")
    print("Excel / PDF / 登録済みWebサイト / エクスプローラー / exe・py実行 / デスクトップ操作 /")
    print("テキスト加工 に対応しています。")
    print("これらを組み合わせた新しい操作を覚えさせたいときは '操作を登録' と入力してください。")
    print("ボタン操作中心のGUIで登録したいときは 'GUIで操作を登録' と入力してください。")
    print("手順ごとの確認動作を後から見直したいときは '確認設定' と入力してください。")
    print("手順ごとのリトライ回数を後から見直したいときは 'リトライ設定' と入力してください。")
    print("手順の並び替え・削除・挿入をしたいときは '手順編集' と入力してください。")
    print("1手順ずつ動作確認したいときは 'ステップ実行' と入力してください。")
    print("実行記録をExcelで見たいときは '実行ログ出力' と入力してください。")
    print("複数マクロをまとめて実行したいときは 'パイプライン作成' / 'パイプライン実行' と入力してください。")
    print("登録済みボタン等が今も見つかるか確認したいときは 'ヘルスチェック' と入力してください。")
    print("終了するには 'exit' または 'quit' を入力してください。\n")

    RECORD_TRIGGERS = ("操作を登録", "マクロ登録", "レコーダー", "record")
    GUI_RECORD_TRIGGERS = ("GUIで操作を登録", "GUI操作登録", "gui", "GUIレコーダー")
    VERIFY_MENU_TRIGGERS = ("確認設定", "確認省略", "確認の設定")
    RETRY_MENU_TRIGGERS = ("リトライ設定", "再試行設定", "リトライ回数設定")
    STEPS_EDIT_TRIGGERS = ("手順編集", "手順の編集", "並び替え")
    STEP_TRIGGERS = ("ステップ実行", "動作確認", "step")
    LOG_EXPORT_TRIGGERS = ("実行ログ出力", "実行ログ", "ログ出力")
    PIPELINE_CREATE_TRIGGERS = ("パイプライン作成", "パイプライン登録")
    PIPELINE_RUN_TRIGGERS = ("パイプライン実行",)
    HEALTHCHECK_TRIGGERS = ("ヘルスチェック", "健全性チェック")

    while True:
        try:
            text = input("指示> ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break

        if not text:
            continue
        if text.lower() in ("exit", "quit"):
            break

        if text in RECORD_TRIGGERS:
            recorder = MacroRecorder(CONFIG_DIR)
            recorder.record()
            # 新しく登録されたマクロ/意図をすぐ使えるように再読込
            intent_engine.reload()
            executor.reload()
            continue

        if text in GUI_RECORD_TRIGGERS:
            launch_gui_recorder()
            # GUI側で保存された可能性があるので再読込
            intent_engine.reload()
            executor.reload()
            continue

        if text in VERIFY_MENU_TRIGGERS:
            manage_verification(CONFIG_DIR)
            executor.reload()
            continue

        if text in RETRY_MENU_TRIGGERS:
            manage_retry(CONFIG_DIR)
            executor.reload()
            continue

        if text in STEPS_EDIT_TRIGGERS:
            manage_steps(CONFIG_DIR)
            intent_engine.reload()
            executor.reload()
            continue

        if text in STEP_TRIGGERS:
            macro_name = choose_macro(executor)
            if macro_name is None:
                continue
            slots: dict = {}
            for slot_name in executor.required_slots(macro_name):
                slots[slot_name] = prompt_for_slot(slot_name)
            run_macro_stepwise(executor, macro_name, slots)
            continue

        if text in LOG_EXPORT_TRIGGERS:
            export_run_log(executor)
            continue

        if text in PIPELINE_CREATE_TRIGGERS:
            create_pipeline(CONFIG_DIR, executor)
            continue

        if text in PIPELINE_RUN_TRIGGERS:
            run_pipeline(CONFIG_DIR, executor, dry_run)
            continue

        if text in HEALTHCHECK_TRIGGERS:
            health_check_menu(executor)
            continue

        match = intent_engine.classify(text)
        if match is None:
            print("  → 対応するマクロが見つかりませんでした。'help' で一覧を確認できます。\n")
            continue

        if match.macro == "__list_macros__":
            for intent in intent_engine.list_intents():
                if intent["macro"] == "__list_macros__":
                    continue
                print(f"  - {intent['id']}: {intent['description']} (macro: {intent['macro']})")
            print()
            continue

        print(f"  → 意図: {match.intent_id} (score={match.score}) / マクロ: {match.macro}")

        slots = {}
        for slot_name in executor.required_slots(match.macro):
            slots[slot_name] = prompt_for_slot(slot_name)

        try:
            results = executor.run(match.macro, slots, dry_run=dry_run, on_failure=cli_failure_prompt)
            print("  実行結果:")
            for r in results:
                print(f"    - {r}")
        except MacroEditRequested as e:
            print(f"\n手順 {e.step_number} を修正します。")
            open_step_editor(CONFIG_DIR, match.macro, e.step_number)
            print("修正のため、プログラムを終了します。")
            return
        except Exception as e:  # noqa: BLE001
            logger.exception("マクロ実行中にエラーが発生しました")
            print(f"  ⚠ エラー: {e}")
        print()


def main() -> None:
    parser = argparse.ArgumentParser(description="疑似ローカルAI (RPA特化)")
    parser.add_argument("--dry-run", action="store_true", help="実際には実行せず手順のみ表示")
    parser.add_argument("--no-headless", action="store_true", help="ブラウザをheadlessにしない(画面表示)")
    parser.add_argument(
        "--healthcheck-all", action="store_true",
        help="全マクロのヘルスチェックだけを行って終了する(タスクスケジューラ/cronでの定期実行向け)",
    )
    parser.add_argument(
        "--run-macro", metavar="NAME",
        help="対話なしで指定マクロを1回だけ実行して終了する(他ツールからの呼び出し向け)",
    )
    parser.add_argument(
        "--slots", metavar="JSON",
        help="--run-macro と併用。マクロに渡すスロットをJSON文字列で指定する",
    )
    args = parser.parse_args()

    LOG_DIR.mkdir(exist_ok=True)

    if args.healthcheck_all:
        all_ok = run_health_check(CONFIG_DIR, headless=not args.no_headless)
        sys.exit(0 if all_ok else 1)

    if args.run_macro:
        ok = run_macro_noninteractive(args.run_macro, args.slots, headless=not args.no_headless)
        sys.exit(0 if ok else 1)

    run_repl(dry_run=args.dry_run, headless=not args.no_headless)


if __name__ == "__main__":
    main()
