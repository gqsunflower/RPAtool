"""
MacroExecutor: config/macros.json に登録された「操作手順(マクロ)」を
順番に実行するランナーです。

安全設計のポイント:
- 実行できるのは macros.json に事前登録されたステップのみ。
  自由なコード生成・動的なアクション追加は行いません。
- browser ハンドラは whitelist_urls.json にあるサイトにしか
  アクセスできない作りになっているため、このExecutor経由でも
  未登録サイトへは到達できません。
- 各ステップの params に "verify" (確認方法) が定義されていれば、
  実行直後にそれを検証し、失敗した場合は既定でマクロの残りの手順を
  一切実行せずに例外を送出して止まります(誤操作の連鎖を防ぐため)。
  ステップの "verify_skip" を true にすると、その手順だけ確認を
  省略できます(macros.jsonを直接編集するか、確認設定メニューから)。
  確認タイプは text_appears/text_disappears/url_changes(Web用)、
  image_appears/image_disappears(デスクトップ用)に対応しています。

リトライ("retry"):
  各ステップに {"retry": {"count": 2, "interval_seconds": 3}} のように
  設定しておくと、一時的な読み込み遅延などによる失敗をその場で
  count回まで自動的に再試行してから失敗扱いにします
  (ネットワーク遅延などほぼ確実に成功する操作向け)。

失敗時の挙動のカスタマイズ(on_failure):
  リトライを使い切っても失敗した場合、on_failure コールバックが
  渡されていればそれを呼び出し、"abort"(即時中断) / "resume"(手動で
  補完済みとみなし次のステップへ) / "edit"(中断してMacroEditRequestedを
  送出、呼び出し側で編集画面へ)のいずれかで応答してもらいます。
  渡されない場合は例外をそのまま送出します。

ステップ実行(F8相当)のカスタマイズ(on_step / on_result):
  on_step が渡されている場合、各ステップの実行前に呼び出され、
  "step"(1つだけ実行して次で止まる) / "run"(以降は止めずに自動実行)
  / "skip"(このステップを実行せず飛ばす) / "abort"(中断する)
  のいずれかで応答してもらいます。on_result は各ステップの実行結果を
  その場で確認したい場合に、成功のたびに呼び出されます。

start_step で指定した番号のステップから実行を開始できます。

前の手順の結果を後の手順で使う(store_as / 変数):
  ステップに {"store_as": "変数名"} を付けておくと、その手順の実行結果が
  変数として記録される。以降のどの手順の params の中でも "{{変数名}}" と
  書けば参照できる(ユーザー指定のスロットと全く同じ書き方)。
  値全体が "{{変数名}}" だけの場合は元の型(int/list等)のまま渡され、
  "A1:B{{last_row}}" や "report_{{today}}.xlsx" のように他の文字列に
  埋め込むこともできる(その場合は文字列として結合される)。
  さらに "{{last_row+1}}" / "{{last_row-1}}" のように変数名の後ろに
  +数値 / -数値 を付けると、その場で加減算した値を使える
  (「最終行の次の行」に貼り付けたい場合など)。
  例: Excelの最終行を取得するステップに store_as: "last_row" を付けておき、
  次のセルコピーの範囲に {"source_range": "A1:B{{last_row}}"} と書けば、
  実際に取得した最終行番号がそのままコピー範囲の終端に使われる。
  貼り付け先を最終行の次の行にしたい場合は {"dest_cell": "A{{last_row+1}}"}
  と書く。

実行ログ・スクリーンショット:
  コンストラクタに run_logger (RunLogger) を渡すと、各ステップの
  成功/失敗/スキップ/手動補完をCSVに記録します。screenshot_dir を
  渡すと、ブラウザ操作が失敗した瞬間の画面を自動で保存します。

制御構文("handler": "control"):
  Excel/PDF/Web等のハンドラを呼ばず、マクロの実行順序そのものを操作する
  特別な手順。VBAでいう For〜Next、If〜Then〜Else、Gotoに相当する。

  - {"handler":"control","action":"label","params":{"name":"Label1"}}
    ジャンプ先の目印。何もしない。
  - {"handler":"control","action":"goto","params":{"label":"Label1"}}
    指定したラベルへ無条件にジャンプする。
  - {"handler":"control","action":"if_goto",
     "params":{"left":"{{A}}","op":"!=","right":"{{B}}","label":"Label1"}}
    left/opここでopは "==","!=","<","<=",">",">=" のいずれか。
    条件が真ならラベルへジャンプし、偽ならそのまま次のステップへ進む
    (次のステップ以降を素直に実行することで「else」相当になる)。
    例: "IF A<>B then goto Label1 else A=A+1 endif" は、
    if_goto(A!=B, Label1) の直後に set_value(A+1, store_as="A") を
    置くだけで表現できる。
  - {"handler":"control","action":"set_value","params":{"value":"{{A+1}}"},
     "store_as":"A"}
    値を計算して変数へ再代入する(A=A+1 に相当)。store_asは既存のパイプ
    ライン機構と同じで、既存の変数名を指定すればその変数を上書きできる。
  - {"handler":"control","action":"to_str","params":{"value":"{{A}}"},"store_as":"A"}
    {"handler":"control","action":"to_int","params":{"value":"{{A}}"},"store_as":"A"}
    {"handler":"control","action":"to_float","params":{"value":"{{A}}"},"store_as":"A"}
    変数の型を強制的に変換する。Excelのセル値等は先頭が0の値(郵便番号等)を
    そのまま保持するため基本的に文字列として取得されるが、計算に使いたい
    ときはto_int/to_floatで明示的に数値へ変換できる(int変換は小数点以下を
    切り捨てるのではなく一度floatを経由するため、"3.0"のような文字列も
    to_intで3に変換できる。"3.5"のような値をto_intすると3になる)。
  - {"handler":"control","action":"for_start",
     "params":{"var":"i","start":0,"end":"{{mylist.length-1}}"}}
    〜
    {"handler":"control","action":"for_end","params":{"var":"i"}}
    VBAの「For i = 0 To N ... Next」に相当。for_startとfor_endの間の
    手順が、varで指定した変数(既定"i")を0から順にendまで動かしながら
    繰り返し実行される。ループ本体の中では {{i}} や {{mylist[i]}} の形で
    現在のカウンタ値・対応するリスト要素を参照できる。入れ子(ループの中に
    別のループ)にも対応。endにリストの要素数を使いたい場合は
    "{{mylist.length}}"(要素数そのまま。0始まりで最後の要素まで回したい
    場合は endを "{{mylist.length-1}}" のように1つ減らして指定する)。

  dry_run実行時は、前段の通常ステップが実際には実行されず変数も定義されて
  いないため、制御構文は条件評価・ジャンプを行わずプレビュー表示のみ行い、
  直線的に(登録順のまま)流れる。
"""
from __future__ import annotations

import json
import logging
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

logger = logging.getLogger("rpa_local_ai.executor")

# {{ ... }} の中身(波括弧を含まない文字列)を丸ごと拾い、中身の解釈は
# _parse_var_expr が行う({{name}} / {{name+1}} / {{name[0]}} / {{name[i]}} /
# {{name.length}} のいずれにも対応するため)。
_VAR_PATTERN = re.compile(r"\{\{([^{}]+)\}\}")

_RE_LENGTH = re.compile(r"^(\w+)\.length$")
_RE_LENGTH_ARITH = re.compile(r"^(\w+)\.length\s*([+\-])\s*(\d+(?:\.\d+)?)$")
_RE_INDEX = re.compile(r"^(\w+)\[(\w+)\]$")
_RE_ARITH = re.compile(r"^(\w+)\s*([+\-])\s*(\d+(?:\.\d+)?)$")
_RE_PLAIN = re.compile(r"^(\w+)$")
_RE_COLUMN_LETTERS = re.compile(r"^[A-Za-z]*$")  # 空文字列は「A列より前」を表す特殊値として許容


def _column_letters_to_index(letters: str) -> int:
    """Excelの列文字("A","B",...,"Z","AA",...)を1始まりの列番号に変換する。"""
    idx = 0
    for ch in letters.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def _column_index_to_letters(index: int) -> str:
    """1始まりの列番号をExcelの列文字に変換する(_column_letters_to_indexの逆変換)。"""
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(rem + ord("A")) + letters
    return letters


class MacroEditRequested(Exception):
    """失敗時メニューで「終了して修正画面を開く」が選ばれたときに送出する。"""

    def __init__(self, step_number: int, step: dict, original_error: Exception):
        super().__init__(str(original_error))
        self.step_number = step_number
        self.step = step
        self.original_error = original_error


def _parse_var_expr(expr: str, slots: dict) -> Any:
    """{{ }} の中身(波括弧を除いた文字列)を実際の値に解決する。対応する書き方:

    - "name"          そのまま値を返す(型を維持)
    - "name+N" / "name-N"  数値として加減算する
    - "name[0]" / "name[i]"  リストの要素をPythonと同じ0始まりで取得する
      (角括弧の中は数字そのままか、他の変数名のどちらでも良い。
      後者の場合はその変数の値を整数のインデックスとして使う。forループの
      カウンタ変数をそのままインデックスに使いたい場合等)
    - "name.length"   リストの要素数を取得する
    """
    expr = expr.strip()

    m = _RE_LENGTH_ARITH.match(expr)
    if m:
        key, op, operand = m.group(1), m.group(2), m.group(3)
        if key not in slots:
            raise KeyError(f"未解決のスロットです: {key}")
        value = slots[key]
        if not isinstance(value, list):
            raise ValueError(f"'{key}' はリストではないため .length は使えません")
        result = len(value) + float(operand) if op == "+" else len(value) - float(operand)
        return int(result) if result == int(result) else result

    m = _RE_LENGTH.match(expr)
    if m:
        key = m.group(1)
        if key not in slots:
            raise KeyError(f"未解決のスロットです: {key}")
        value = slots[key]
        if not isinstance(value, list):
            raise ValueError(f"'{key}' はリストではないため .length は使えません")
        return len(value)

    m = _RE_INDEX.match(expr)
    if m:
        key, idx_raw = m.group(1), m.group(2)
        if key not in slots:
            raise KeyError(f"未解決のスロットです: {key}")
        lst = slots[key]
        if not isinstance(lst, list):
            raise ValueError(f"'{key}' はリストではないため [インデックス] 指定はできません")
        if idx_raw.isdigit():
            idx = int(idx_raw)
        else:
            if idx_raw not in slots:
                raise KeyError(f"未解決のスロットです(インデックス用の変数): {idx_raw}")
            idx = int(slots[idx_raw])
        if idx < 0 or idx >= len(lst):
            raise IndexError(
                f"'{key}' の範囲外のインデックスです: {idx}(要素数{len(lst)}、0〜{len(lst) - 1}の範囲で指定)"
            )
        return lst[idx]

    m = _RE_ARITH.match(expr)
    if m:
        key, op, operand = m.group(1), m.group(2), m.group(3)
        if key not in slots:
            raise KeyError(f"未解決のスロットです: {key}")
        value = slots[key]
        if isinstance(value, str) and _RE_COLUMN_LETTERS.match(value):
            # Excelの列文字("A","B",...,"Z","AA",...)に対する+N/-N演算。
            # 例: get_last_columnで取得した列文字を{{last_col+1}}のように
            # 使うと、その1つ右の列文字が得られる(A1形式のセル参照組み立て用)。
            idx = _column_letters_to_index(value)
            delta = int(float(operand))
            new_idx = idx + delta if op == "+" else idx - delta
            if new_idx < 1:
                raise ValueError(f"'{key}'({value})から{op}{operand}すると列がA列より前になります")
            return _column_index_to_letters(new_idx)
        try:
            num = float(value)
        except (TypeError, ValueError) as e:
            raise ValueError(
                f"'{key}' の値 {value!r} は数値ではないため、{op}{operand} の計算ができません"
            ) from e
        result = num + float(operand) if op == "+" else num - float(operand)
        return int(result) if result == int(result) else result

    m = _RE_PLAIN.match(expr)
    if m:
        key = m.group(1)
        if key not in slots:
            raise KeyError(f"未解決のスロットです: {key}")
        return slots[key]

    raise ValueError(f"認識できない書き方です: {{{{{expr}}}}}")


def _substitute(value: Any, slots: dict) -> Any:
    """paramsの値に含まれる {{...}} をslotsの値に置換する。

    - 値全体が単一の "{{...}}" だけの場合は、元の型(int/list/dict等)を
      保ったまま差し替える。例: {"end_page": "{{last_row}}"} で last_row が
      整数42なら end_page には整数42がそのまま渡る。
    - "A1:B{{last_row}}" や "report_{{today}}.xlsx" のように、他の文字列の
      中に埋め込まれている場合は、その部分だけを文字列として結合する
      (埋め込む値は str() で文字列化される)。1つの文字列に複数の
      {{変数}} を含めることもできる。
    - "{{last_row+1}}" / "{{last_row-1}}" のように、変数名の後ろに +数値 /
      -数値 を付けると、その場で加減算した値を使える。
    - "{{mylist[0]}}" / "{{mylist[i]}}" のように、変数名の後ろに [ ] を
      付けると、リストの要素をPythonと同じ0始まりの番号で取得できる
      (角括弧の中は数字でも、iのような別の変数名でも良い)。
    - "{{mylist.length}}" でリストの要素数を取得できる。
    - "{{last_col+1}}" のように、変数の値がExcelの列文字("A","B",...)の
      場合は列文字としての+N/-N演算になる(1つ右/左の列文字を返す)。

    dictの値だけでなく**キー**も同じルールで置換する。これは
    Excelのcell_values(例: {"B{{last_row+1}}": "値"})のように、
    セル参照そのものを前の手順の結果から組み立てたい場合に使う。
    """
    if isinstance(value, str):
        matches = list(_VAR_PATTERN.finditer(value))
        if not matches:
            return value

        # 値全体が単一の {{...}} だけで構成されている場合は、元の型を保ったまま返す
        if len(matches) == 1 and matches[0].span() == (0, len(value)):
            return _parse_var_expr(matches[0].group(1), slots)

        # それ以外は文字列として埋め込む(見つかった {{...}} をすべて置換)
        def _replace(match: re.Match) -> str:
            return str(_parse_var_expr(match.group(1), slots))

        return _VAR_PATTERN.sub(_replace, value)
    if isinstance(value, dict):
        return {
            (_substitute(k, slots) if isinstance(k, str) else k): _substitute(v, slots)
            for k, v in value.items()
        }
    if isinstance(value, list):
        return [_substitute(v, slots) for v in value]
    return value


def _evaluate_condition(left: Any, op: str, right: Any) -> bool:
    """if_goto用の条件式を評価する。両辺とも数値として解釈できれば数値比較、
    できなければ文字列として比較する。
    """
    def _try_number(v: Any) -> float | None:
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    left_num, right_num = _try_number(left), _try_number(right)
    if left_num is not None and right_num is not None:
        left_cmp, right_cmp = left_num, right_num
    else:
        left_cmp, right_cmp = str(left), str(right)

    if op == "==":
        return left_cmp == right_cmp
    if op == "!=":
        return left_cmp != right_cmp
    if op == "<":
        return left_cmp < right_cmp
    if op == "<=":
        return left_cmp <= right_cmp
    if op == ">":
        return left_cmp > right_cmp
    if op == ">=":
        return left_cmp >= right_cmp
    raise ValueError(f"未知の比較演算子です: {op}(使えるもの: == != < <= > >=)")


class MacroExecutor:
    def __init__(
        self,
        macros_path: Path,
        handlers: dict[str, Any],
        run_logger: Any | None = None,
        screenshot_dir: Path | None = None,
    ):
        """
        handlers: {"excel": ExcelHandlerインスタンス, "pdf": ..., "browser": ...}
        各ハンドラは action名と同名のメソッドを持つ必要がある。
        run_logger: engine.run_logger.RunLogger (省略可)
        screenshot_dir: 失敗時のスクリーンショット保存先(省略可、browserハンドラのみ対象)
        """
        self.macros_path = Path(macros_path)
        self.handlers = handlers
        self.run_logger = run_logger
        self.screenshot_dir = Path(screenshot_dir) if screenshot_dir else None
        self._load()

    def _load(self) -> None:
        with open(self.macros_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.macros: dict[str, dict] = data.get("macros", {})

    def reload(self) -> None:
        self._load()

    def get_macro(self, name: str) -> dict:
        if name not in self.macros:
            raise KeyError(f"未登録のマクロです: {name}")
        return self.macros[name]

    def required_slots(self, macro_name: str) -> list[str]:
        return self.get_macro(macro_name).get("required_slots", [])

    def _capture_failure_screenshot(self, handler: Any, macro_name: str, step_number: int) -> str:
        if not self.screenshot_dir or not hasattr(handler, "take_screenshot"):
            return ""
        try:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_macro = "".join(c if c.isalnum() else "_" for c in macro_name)
            path = self.screenshot_dir / f"{safe_macro}_step{step_number}_{ts}.png"
            handler.take_screenshot(str(path))
            logger.info("失敗時のスクリーンショットを保存しました: %s", path)
            return str(path)
        except Exception:  # noqa: BLE001
            return ""

    def run(
        self,
        macro_name: str,
        slots: dict,
        dry_run: bool = False,
        start_step: int = 1,
        on_step: Callable[[int, int, dict], str] | None = None,
        on_result: Callable[[int, int, dict, Any], None] | None = None,
        on_failure: Callable[[int, int, dict, Exception], str] | None = None,
    ) -> list[Any]:
        macro = self.get_macro(macro_name)
        steps = macro["steps"]
        total = len(steps)

        missing = [s for s in macro.get("required_slots", []) if s not in slots]
        if missing:
            raise ValueError(f"不足しているスロット: {missing}")

        if not (1 <= start_step <= max(total, 1)):
            raise ValueError(f"start_stepが範囲外です(1〜{total}): {start_step}")

        # 制御構文(label/for)の事前スキャン: ラベル名->ステップ番号、
        # for_start<->for_endの対応関係をあらかじめ調べておく
        label_positions: dict[str, int] = {}
        for_pairs: dict[int, int] = {}  # for_startの番号 -> 対応するfor_endの番号(相互)
        for_open_stack: list[int] = []
        for idx, step in enumerate(steps, start=1):
            if step.get("handler") != "control":
                continue
            action = step.get("action")
            if action == "label":
                name = step.get("params", {}).get("name")
                if not name:
                    raise ValueError(f"ステップ{idx}: labelにはnameが必要です")
                if name in label_positions:
                    raise ValueError(f"ラベル名が重複しています: {name}")
                label_positions[name] = idx
            elif action == "for_start":
                for_open_stack.append(idx)
            elif action == "for_end":
                if not for_open_stack:
                    raise ValueError(f"ステップ{idx}: 対応するfor_startがありません")
                start_idx = for_open_stack.pop()
                for_pairs[start_idx] = idx
                for_pairs[idx] = start_idx
        if for_open_stack:
            raise ValueError(f"対応するfor_endがありません(for_start: ステップ{for_open_stack})")

        loop_stack: list[dict[str, Any]] = []  # 実行中のforループの状態

        results: list[Any] = []
        variables: dict[str, Any] = {}  # store_as で記録された、前の手順の結果を後の手順で使うための変数
        auto_run = on_step is None  # on_stepが無ければ最初から自動実行
        i = start_step

        while i <= total:
            step = steps[i - 1]
            handler_name = step["handler"]
            action_name = step["action"]

            if not dry_run and on_step is not None and not auto_run:
                decision = on_step(i, total, step)
                if decision == "abort":
                    logger.info("ユーザー操作によりステップ%d/%dで中断しました", i, total)
                    return results
                if decision == "skip":
                    logger.info("ステップ%d/%dをスキップしました", i, total)
                    if self.run_logger:
                        self.run_logger.log(macro_name, i, handler_name, action_name, "skip", "ユーザー操作でスキップ")
                    i += 1
                    continue
                if decision == "run":
                    auto_run = True
                # "step" の場合はこのままこのステップだけ実行する

            # ---------- 制御構文(label/goto/if_goto/set_value/for_start/for_end) ----------
            # これらはExcel/PDF/Web等のハンドラを呼び出さない「マクロの流れ」を
            # 操作する特別なステップなので、通常のハンドラ実行より前に処理する。
            if handler_name == "control":
                raw_params = step.get("params", {})

                if dry_run:
                    # dry_run中は前段の通常ステップも実際には実行されておらず
                    # 変数が定義されていないため、条件評価やジャンプは行わず
                    # プレビュー表示のみ行って次のステップへ進む(直線的に流す)。
                    results.append(f"[dry-run] control.{action_name}({raw_params})")
                    i += 1
                    continue

                combined = {**slots, **variables}

                if action_name == "label":
                    i += 1
                    continue

                if action_name == "goto":
                    label = raw_params.get("label")
                    if label not in label_positions:
                        raise ValueError(f"ステップ{i}: ラベルが見つかりません: {label}")
                    i = label_positions[label]
                    continue

                if action_name == "if_goto":
                    left = _substitute(raw_params.get("left"), combined)
                    op = raw_params.get("op")
                    right = _substitute(raw_params.get("right"), combined)
                    label = raw_params.get("label")
                    cond = _evaluate_condition(left, op, right)
                    logger.info(
                        "step %d/%d: if %r %s %r -> %s", i, total, left, op, right, cond
                    )
                    if cond:
                        if label not in label_positions:
                            raise ValueError(f"ステップ{i}: ラベルが見つかりません: {label}")
                        i = label_positions[label]
                    else:
                        i += 1
                    continue

                if action_name == "set_value":
                    value = _substitute(raw_params.get("value"), combined)
                    store_as = step.get("store_as")
                    if store_as:
                        variables[store_as] = value
                    results.append(value)
                    if self.run_logger:
                        self.run_logger.log(
                            macro_name, i, "control", "set_value", "success", str(value)[:200]
                        )
                    i += 1
                    continue

                if action_name in ("to_str", "to_int", "to_float"):
                    value = _substitute(raw_params.get("value"), combined)
                    try:
                        if action_name == "to_str":
                            converted: Any = str(value)
                        elif action_name == "to_int":
                            converted = int(float(value))
                        else:
                            converted = float(value)
                    except (TypeError, ValueError) as e:
                        raise ValueError(
                            f"ステップ{i}: {value!r} を{action_name}で変換できません: {e}"
                        ) from e
                    store_as = step.get("store_as")
                    if store_as:
                        variables[store_as] = converted
                    results.append(converted)
                    if self.run_logger:
                        self.run_logger.log(
                            macro_name, i, "control", action_name, "success", str(converted)[:200]
                        )
                    i += 1
                    continue

                if action_name == "for_start":
                    var_name = raw_params.get("var", "i")
                    if loop_stack and loop_stack[-1]["for_start_idx"] == i:
                        # for_endからのジャンプバックによる2周目以降の到達
                        # (すでにセットアップ済みなので初期化し直さない)
                        pass
                    else:
                        start_val = _substitute(raw_params.get("start", 0), combined)
                        end_val = _substitute(raw_params.get("end"), combined)
                        loop_stack.append({
                            "var": var_name,
                            "end": int(float(end_val)),
                            "for_start_idx": i,
                        })
                        variables[var_name] = int(float(start_val))
                        logger.info(
                            "step %d/%d: for %s = %s to %s", i, total, var_name,
                            variables[var_name], loop_stack[-1]["end"],
                        )
                    i += 1
                    continue

                if action_name == "for_end":
                    if not loop_stack:
                        raise RuntimeError(f"ステップ{i}: 対応するfor_startがありません")
                    loop = loop_stack[-1]
                    variables[loop["var"]] += 1
                    if variables[loop["var"]] <= loop["end"]:
                        i = loop["for_start_idx"] + 1  # ループ本体の先頭へ戻る
                    else:
                        loop_stack.pop()
                        i += 1
                    continue

                raise ValueError(f"ステップ{i}: 未知の制御構文です: {action_name}")

            raw_params = step.get("params", {})
            # ユーザー指定のスロットと、前の手順が store_as で記録した変数を
            # 同じ名前空間として扱う({{name}}の書き方はどちらも共通)
            try:
                params = _substitute(raw_params, {**slots, **variables})
            except (KeyError, ValueError, IndexError) as e:
                if dry_run:
                    # dry_run中は制御構文(for/if等)が実行されず変数が
                    # まだ定義されていないことがあるため、置換に失敗しても
                    # クラッシュさせずプレビュー表示だけ行って先へ進む
                    results.append(
                        f"[dry-run] {handler_name}.{action_name}({raw_params}) "
                        f"[変数が未確定のため実際の値は表示できません: {e}]"
                    )
                    i += 1
                    continue
                raise

            handler = self.handlers.get(handler_name)
            if handler is None:
                raise RuntimeError(f"未登録のハンドラです: {handler_name}")

            fn: Callable | None = getattr(handler, action_name, None)
            if fn is None:
                raise RuntimeError(
                    f"ハンドラ '{handler_name}' に action '{action_name}' がありません"
                )

            verify_cfg = step.get("verify")
            verify_skip = step.get("verify_skip", False)
            verify_active = (
                not dry_run
                and verify_cfg
                and verify_cfg.get("type") not in (None, "none")
                and not verify_skip
            )

            logger.info("step %d/%d: %s.%s(%s)", i, total, handler_name, action_name, params)

            if dry_run:
                note = ""
                if verify_cfg and verify_cfg.get("type") not in (None, "none"):
                    note = f" [確認あり: {verify_cfg.get('type')}]"
                if verify_skip:
                    note = " [確認は省略設定]"
                retry_cfg = step.get("retry") or {}
                if retry_cfg.get("count", 0):
                    note += f" [リトライ{retry_cfg['count']}回]"
                if step.get("store_as"):
                    note += f" [store_as: {step['store_as']}]"
                results.append(f"[dry-run] {handler_name}.{action_name}({params}){note}")
                i += 1
                continue

            before_url = None
            if verify_active and verify_cfg.get("type") == "url_changes" and handler_name == "browser":
                before_url = handler.get_current_url()

            retry_cfg = step.get("retry") or {}
            max_attempts = int(retry_cfg.get("count", 0)) + 1
            interval = float(retry_cfg.get("interval_seconds", 2))

            attempt = 0
            result = None
            last_err: Exception | None = None
            success = False

            while attempt < max_attempts:
                attempt += 1
                try:
                    result = fn(**params)
                    if verify_active:
                        self._run_verification(handler, verify_cfg, before_url)
                    success = True
                    break
                except Exception as err:  # noqa: BLE001
                    last_err = err
                    if attempt < max_attempts:
                        logger.warning(
                            "step %d/%d 失敗(試行%d/%d)。%.1f秒後に再試行します: %s",
                            i, total, attempt, max_attempts, interval, err,
                        )
                        time.sleep(interval)

            if not success:
                err = last_err
                logger.error("step %d/%d が失敗しました(リトライ後も失敗): %s", i, total, err)
                screenshot_path = self._capture_failure_screenshot(handler, macro_name, i)
                if self.run_logger:
                    self.run_logger.log(
                        macro_name, i, handler_name, action_name, "failure", str(err), screenshot_path
                    )

                if on_failure is None:
                    raise err
                decision = on_failure(i, total, step, err)
                if decision == "resume":
                    if self.run_logger:
                        self.run_logger.log(
                            macro_name, i, handler_name, action_name, "manual", "手動補完により継続"
                        )
                    results.append(f"[manual] step {i} は手動補完として扱い、次に進みます")
                    i += 1
                    continue
                if decision == "edit":
                    raise MacroEditRequested(i, step, err) from err
                raise err  # "abort" またはその他 → 元の例外をそのまま送出

            results.append(result)
            store_as = step.get("store_as")
            if store_as:
                variables[store_as] = result
                logger.info("変数に保存しました: %s = %r", store_as, result)
            if self.run_logger:
                self.run_logger.log(macro_name, i, handler_name, action_name, "success", str(result)[:200])
            if on_result is not None:
                on_result(i, total, step, result)
            i += 1

        return results

    def _run_verification(self, handler: Any, verify_cfg: dict, before_url: str | None) -> Any:
        vtype = verify_cfg.get("type")
        timeout = verify_cfg.get("timeout", 10)
        if vtype == "text_appears":
            return handler.verify_text_appears(verify_cfg["value"], timeout=timeout)
        if vtype == "text_disappears":
            return handler.verify_text_disappears(verify_cfg["value"], timeout=timeout)
        if vtype == "url_changes":
            return handler.verify_url_changed(before_url or "", timeout=timeout)
        if vtype == "image_appears":
            return handler.verify_image_appears(verify_cfg["value"], timeout=timeout)
        if vtype == "image_disappears":
            return handler.verify_image_disappears(verify_cfg["value"], timeout=timeout)
        raise ValueError(f"未知の確認タイプです: {vtype}")
