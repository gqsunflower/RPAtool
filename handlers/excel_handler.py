"""
ExcelHandler: 既定ではopenpyxlのみで完結するExcel操作ハンドラ(実際のExcel
アプリケーションを起動しない、非表示のバックグラウンド動作)。
.xlsm(マクロ有効ブック)はVBAを保持したまま読み書きできる(keep_vba)。

**可視化モード(visible=True)**: コンストラクタで visible=True を指定すると、
create_workbook/load_workbook以降、実際のExcelアプリケーションをCOM経由で
起動し、画面に表示しながら操作する(**Windows + Excelインストール環境専用**、
pywin32が必要)。動作を目で見て確認しながらマクロを組み立てたい場合(レコーダー
では既定でこちらを使う)に使う。ただし対応している操作は限られており
(create_workbook/load_workbook/switch_workbook/close_workbook/write_cells/
get_cell_value/get_last_row/get_last_column/get_sheet_names/create_sheet/
rename_sheet/save_workbook/save_workbook_as/set_print_area/set_page_layout/
set_print_area_from_file/lookup_range_for_date)、それ以外(copy_cell_range/
copy_sheet/get_range_as_list/find_cell/append_row/
insert_rows/delete_rows/clear_range/delete_sheet/read_sheet_to_records/
write_records_to_csv)は可視化モードのブックに対して呼ぶとエラーになる
(その手順だけ非可視化モードのブックで行うか、動作確認が済んだら
set_visible(False)で非可視化モードに切り替えてから使うこと)。
可視化モードで作成した新規ブックの既定シート名("Sheet1"等)は、
openpyxlでの新規作成時("Sheet")と異なる点に注意。

複数ブックの同時オープンとアクティブ切り替え:
  load_workbook() は毎回「エイリアス」を付けてメモリ上に保持し、最後に
  読み込んだ(または switch_workbook() で切り替えた)ブックが「アクティブ」
  になる。read_sheet_to_records / write_cells / save_workbook_as 等は
  常にアクティブなブックに対して行われる。これにより、2つ以上のExcelを
  開いたまま行き来しながら操作できる(例: A.xlsxを読み込む → B.xlsxを
  読み込む → switch_workbookでAに戻す → Aに書き込む → Bに切り替えて書き込む)。
  可視化/非可視化はブックごとに個別に記録されるため(load_workbook/
  create_workbook実行時点のvisible設定が、そのブックにはそのまま使われる)、
  同じセッション内で可視化ブックと非可視化ブックを混在させることもできる。

RPAでよく使うセル/シート操作:
  get_cell_value(特定セル1つの読み込み)、get_last_row(最終行取得)、
  copy_cell_range(セルのコピー&貼り付け。値のみ/数式のみ/すべて(書式含む)
  から貼り付け方式を選べる)、copy_sheet(シートのコピー)、
  create_sheet(新規シート追加)、rename_sheet(既存シート名の変更)、
  get_sheet_names(シート名一覧)、set_print_area(印刷範囲の指定) に対応。

印刷範囲・ページ設定とPDF出力:
  set_print_area(印刷範囲)/set_page_layout(用紙サイズ・向き・縦横の枚数)/
  set_print_area_from_file(別ファイルの「日付,範囲」の対応表から、実行日に
  合わせて印刷範囲を決める。lookup_range_for_dateで範囲だけを取り出すことも
  可)で設定した内容は、同じファイルを対象にした save_as_pdf のPDFに反映される。

そのほか以下の2つの操作は、openpyxlでは実現できない
(openpyxlはファイル形式としての読み書きのみで、実際にExcelアプリケーションを
描画・実行する機能を持たないため)常にCOM経由で実際のExcelを操作する
(**Windows + Excelインストール環境専用**、pywin32が必要)。これらは毎回
パスを直接指定するため、複数ブックの「アクティブ切り替え」とは独立して、
どのファイルに対しても実行できる:

- save_as_pdf: 印刷範囲・ページ設定(このツールで設定した内容を含む)を反映してPDFへ書き出す
- run_excel_macro: .xlsmに既に組み込まれているVBAマクロを実行する

いずれもネットワークアクセスは一切行いません。
"""
from __future__ import annotations

import calendar
import csv
import datetime as _dt
import logging
import re
from copy import copy as _copy_style
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

logger = logging.getLogger("rpa_local_ai.excel")

_MACRO_ENABLED_EXTENSIONS = {".xlsm", ".xltm"}
_EXCEL_FILE_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}

# Excelの定数(win32comを遅延・非gencacheで使っているためwin32com.client.constants
# が使えず、数値をそのまま指定している。値はExcelのバージョンに依らず安定)
_XL_UP = -4162
_XL_TO_LEFT = -4159
_XL_ORIENTATION = {"portrait": 1, "landscape": 2}  # xlPortrait / xlLandscape
# xlPaperSize(openpyxlのpage_setup.paperSizeも同じ番号)。B4/B5はJIS規格
_PAPER_SIZES = {"A3": 8, "A4": 9, "A5": 11, "B4": 12, "B5": 13, "LETTER": 1, "LEGAL": 5}
_MAX_FIT_PAGES = 100
# 用紙サイズに「印刷範囲そのまま」を選んだ印を、記録済み設定(paper)に入れる値
_FIT_PAPER = "FIT"
FIT_PAPER_LABEL = "印刷範囲そのまま"
_FIT_PAPER_ALIASES = {
    "FIT", "AREA", "範囲", "範囲に合わせる", "印刷範囲", "印刷範囲そのまま", "印刷範囲の大きさそのまま",
}
_A3_POINTS = (841.92, 1190.52)  # A3の幅×高さ(縦向き、pt)。範囲そのまま出力の作業用紙

_CELL_RANGE_RE = re.compile(r"[A-Z]{1,3}[0-9]+(:[A-Z]{1,3}[0-9]+)?")
_DATE_RE = re.compile(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})")


def _normalize_paper_size(paper_size: str) -> str:
    key = str(paper_size).strip().upper().replace("(JIS)", "").replace("JIS", "").replace(" ", "")
    if key in _FIT_PAPER_ALIASES:
        return _FIT_PAPER
    if key not in _PAPER_SIZES:
        raise ValueError(
            f"用紙サイズが未対応です: {paper_size}"
            f"(使えるもの: A3 / A4 / A5 / B4 / B5 / Letter / Legal / {FIT_PAPER_LABEL})"
        )
    return key


def _parse_margin(label: str, value: Any) -> float | None:
    """余白(cm)を数値にする。None/空欄は「変更しない」としてNone。"""
    if value is None or str(value).strip() == "":
        return None
    try:
        cm = float(value)
    except (TypeError, ValueError) as e:
        raise ValueError(f"{label}の余白は数字(cm)で指定してください: {value}") from e
    if not (0 <= cm <= 50):
        raise ValueError(f"{label}の余白は0〜50cmで指定してください: {value}")
    return cm


def _normalize_orientation(orientation: str) -> str:
    key = str(orientation).strip().lower()
    if key in ("portrait", "縦", "縦向き", "たて"):
        return "portrait"
    if key in ("landscape", "横", "横向き", "よこ"):
        return "landscape"
    raise ValueError(f"向きは portrait(縦向き)か landscape(横向き)で指定してください: {orientation}")


def _parse_date(value: Any) -> _dt.date:
    """"2026/10/20" "2026-10-20" "2026年10月20日" "20261020" やExcelの日付セルの値を
    dateにする。年が無い等で日付として読めない場合はValueError。
    """
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    text = str(value).strip()
    m = _DATE_RE.search(text)
    if m is None and re.fullmatch(r"\d{8}", text):
        m = re.match(r"(\d{4})(\d{2})(\d{2})", text)
    if m is None:
        raise ValueError(f"日付として読めません: '{value}'(例: 2026/10/20)")
    try:
        return _dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError as e:
        raise ValueError(f"存在しない日付です: '{value}'") from e


def _read_date_range_table(path: Path, source_sheet: str = "") -> list[tuple[_dt.date, str]]:
    """「日付, セル範囲」の対応表(Excelまたはテキスト)を [(日付, 範囲), ...] で返す。
    日付として読めない行(見出し・空行・コメント)は読み飛ばす。
    """
    if not path.exists():
        raise FileNotFoundError(f"日付と範囲の対応表が見つかりません: {path}")
    pairs: list[tuple[Any, Any]] = []
    if path.suffix.lower() in _EXCEL_FILE_EXTENSIONS:
        wb = load_workbook(filename=path, read_only=True, data_only=True)
        try:
            if source_sheet:
                if source_sheet not in wb.sheetnames:
                    raise KeyError(f"対応表のシートが見つかりません: {source_sheet}(存在: {wb.sheetnames})")
                ws = wb[source_sheet]
            else:
                ws = wb.worksheets[0]
            for row in ws.iter_rows(min_col=1, max_col=2, values_only=True):
                pairs.append((row[0], row[1] if len(row) > 1 else None))
        finally:
            wb.close()
    else:
        raw = path.read_bytes()
        for enc in ("utf-8-sig", "cp932"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError(f"対応表の文字コードを判別できませんでした(UTF-8かShift_JISで保存してください): {path}")
        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = re.split(r"[,\t]", line, maxsplit=1)
            if len(parts) < 2:
                parts = re.split(r"\s+", line, maxsplit=1)
            if len(parts) < 2:
                continue
            pairs.append((parts[0].strip().strip('"'), parts[1].strip().strip('"')))

    rows: list[tuple[_dt.date, str]] = []
    for raw_date, raw_range in pairs:
        if raw_date is None or raw_range is None or not str(raw_range).strip():
            continue
        try:
            rows.append((_parse_date(raw_date), str(raw_range).strip()))
        except ValueError:
            continue  # 見出し行など
    return rows


def _com_value_to_str(value: Any) -> str:
    """COM経由で読み取ったセルの値を文字列化する。ExcelのCOMは数値セルを
    常にfloat(浮動小数点数)として返すため(openpyxlのように元の型が
    int/floatかを区別しない)、小数点以下が無い場合はint相当の見た目
    ("123")に揃える(openpyxlバックエンドとの見た目の一貫性のため)。
    """
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _import_win32com():
    try:
        import win32com.client  # type: ignore
        return win32com.client
    except ImportError as e:
        raise RuntimeError(
            "この操作にはWindows上のExcel本体 + pywin32が必要です"
            "(pip install pywin32)。Excelがインストールされていない環境や"
            "Windows以外の環境では実行できません。"
        ) from e


_WEEKDAY_JA = "月火水木金土日"
_FORMAT_TOKEN_RE = re.compile(
    r'"[^"]*"|\\.|\[[^\]]*\]|yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|aaaa|aaa|.', re.IGNORECASE,
)
_TODAY_EXPR = r"TODAY\(\)\s*(?:([+-])\s*(\d+))?"


def _format_excel_date(d: _dt.date, fmt: str) -> str:
    """日付をExcelの表示形式の文字列(例: "m/d" → "10/20"、"m月d日(aaa)" → "10月20日(火)")に
    する。対応: yyyy yy mmmm mmm mm m dddd ddd dd d aaaa(月曜日) aaa(月)、
    "..."で囲んだ文字・\\付きの文字はそのまま、[...]は無視する。
    """
    out: list[str] = []
    for tok in _FORMAT_TOKEN_RE.findall(fmt):
        low = tok.lower()
        if len(tok) >= 2 and tok.startswith('"') and tok.endswith('"'):
            out.append(tok[1:-1])
        elif len(tok) == 2 and tok.startswith("\\"):
            out.append(tok[1])
        elif tok.startswith("["):
            continue
        elif low == "yyyy":
            out.append(f"{d.year:04d}")
        elif low == "yy":
            out.append(f"{d.year % 100:02d}")
        elif low == "mmmm":
            out.append(calendar.month_name[d.month])
        elif low == "mmm":
            out.append(calendar.month_abbr[d.month])
        elif low == "mm":
            out.append(f"{d.month:02d}")
        elif low == "m":
            out.append(str(d.month))
        elif low == "dddd":
            out.append(calendar.day_name[d.weekday()])
        elif low == "ddd":
            out.append(calendar.day_abbr[d.weekday()])
        elif low == "dd":
            out.append(f"{d.day:02d}")
        elif low == "d":
            out.append(str(d.day))
        elif low == "aaaa":
            out.append(_WEEKDAY_JA[d.weekday()] + "曜日")
        elif low == "aaa":
            out.append(_WEEKDAY_JA[d.weekday()])
        else:
            out.append(tok)
    return "".join(out)


def _as_date(value: Any) -> _dt.date:
    return value.date() if isinstance(value, _dt.datetime) else value


def _display_from_format(value: Any, number_format: str) -> str | None:
    """日付セルが、表示形式(number_format)によって画面でどう見えるかを再現する
    (日付以外・標準の表示形式のときはNone)。"""
    if isinstance(value, (_dt.date, _dt.datetime)) and number_format and number_format != "General":
        try:
            return _format_excel_date(_as_date(value), number_format)
        except Exception:  # noqa: BLE001
            return None
    return None


def _com_display(ws, row: int, col: int) -> str | None:
    """COMのセルの、画面に表示されている文字列。列幅が足りず"####"になっている
    場合は、値と表示形式から再現する。"""
    cell = ws.Cells(row, col)
    shown = cell.Text
    if shown and not str(shown).startswith("#"):
        return str(shown)
    return _display_from_format(cell.Value, cell.NumberFormat)


def _evaluate_search_target(target: str) -> tuple[str, Any]:
    """find_cell_in_line / set_print_area_by_search の探す対象を評価する。
    戻り値: ("date", 日付) または ("text", 文字列)。
    対応する式: =TODAY() / =TODAY()+N / =TODAY()-N(日付として比較)、
    =TEXT(TODAY(),"m/d") / =TEXT(TODAY()+1,"m月d日(aaa)")(Excelの書式で文字にしたものと比較)。
    """
    text = str(target).strip()
    if not text:
        raise ValueError("探す文字列(または =TODAY() のような式)を指定してください")
    if not text.startswith("="):
        return "text", text

    def shifted(sign: str | None, num: str | None) -> _dt.date:
        days = int(num) if num else 0
        return _dt.date.today() + _dt.timedelta(days=-days if sign == "-" else days)

    m = re.fullmatch(rf"=\s*{_TODAY_EXPR}", text, re.IGNORECASE)
    if m:
        return "date", shifted(m.group(1), m.group(2))
    m = re.fullmatch(rf'=\s*TEXT\s*\(\s*{_TODAY_EXPR}\s*,\s*"([^"]*)"\s*\)', text, re.IGNORECASE)
    if m:
        return "text", _format_excel_date(shifted(m.group(1), m.group(2)), m.group(3))
    raise ValueError(
        f"この式には対応していません: {text}"
        '(使える式: =TODAY() / =TODAY()+1 / =TEXT(TODAY(),"m/d") のように、'
        "今日の日付に日数を加減したもの)"
    )


def _cell_matches(value: Any, display, kind: str, wanted: Any, mode: str) -> bool:
    """セルの値(と表示)が、探す対象に一致するか。kind="date"は日付どうし、
    kind="text"は文字列(日付セルの場合は画面の表示や yyyy/mm/dd とも)を比べる。"""
    if value is None:
        return False
    is_date = isinstance(value, (_dt.date, _dt.datetime))
    if kind == "date":
        if is_date:
            return _as_date(value) == wanted
        if isinstance(value, str):
            try:
                return _parse_date(value) == wanted
            except ValueError:
                return False
        return False

    candidates: list[str] = []
    if is_date:
        shown = display()
        if shown:
            candidates.append(str(shown))
        candidates.append(f"{_as_date(value):%Y/%m/%d}")
    else:
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        candidates.append(str(value))
    candidates = [c.strip() for c in candidates]
    if mode == "contains":
        return any(wanted in c for c in candidates)
    return any(c == wanted for c in candidates)


class ExcelHandler:
    def __init__(self, visible: bool = False):
        # alias -> {"wb": Workbook(openpyxl) or COM Workbook, "path": Path,
        #           "last_records": list, "com": bool}
        self._workbooks: dict[str, dict[str, Any]] = {}
        self._active_alias: str | None = None
        self.visible = visible
        self._com_app = None  # 可視化モード用に使い回すExcel.Applicationインスタンス

    def set_visible(self, visible: bool) -> bool:
        """以降 create_workbook/load_workbook するブックの可視化モードを切り替える。
        既に開いているブックはそのまま(そのブックを開いた時点のモードのまま)。
        戻り値: 実際に変更されたか。
        """
        if visible == self.visible:
            return False
        self.visible = visible
        return True

    def _get_com_app(self):
        if self._com_app is None:
            win32com = _import_win32com()
            self._com_app = win32com.DispatchEx("Excel.Application")
            self._com_app.Visible = True
            self._com_app.DisplayAlerts = False
        return self._com_app

    def _active(self) -> dict[str, Any]:
        if self._active_alias is None or self._active_alias not in self._workbooks:
            raise RuntimeError("先に load_workbook を実行してください")
        return self._workbooks[self._active_alias]

    def _require_openpyxl(self, entry: dict[str, Any], action_name: str) -> None:
        if entry.get("com"):
            raise RuntimeError(
                f"'{action_name}' は可視化モード(COM)のブックには未対応です。"
                f"この手順だけ非可視化モードのブックで行うか、動作確認が済んだ後に"
                f"set_visibleで非可視化モードに切り替えてから登録し直してください。"
            )

    def _com_sheet(self, entry: dict[str, Any], sheet_name: str):
        wb = entry["wb"]
        names = [s.Name for s in wb.Sheets]
        if sheet_name not in names:
            raise KeyError(f"シートが見つかりません: {sheet_name} (存在: {names})")
        return wb.Sheets(sheet_name)

    def list_open_workbooks(self) -> list[str]:
        return list(self._workbooks.keys())

    def _unique_alias(self, base: str) -> str:
        """save_as_pdf/run_excel_macro等、既存のブック一覧とは独立にファイルを
        開く操作で使うエイリアスを、既存のものと衝突しないように発行する。
        """
        if base not in self._workbooks:
            return base
        i = 2
        while f"{base}_{i}" in self._workbooks:
            i += 1
        return f"{base}_{i}"

    def create_workbook(self, alias: str | None = None) -> str:
        """新しい空のExcelブックを作成し、アクティブにする。まだファイルとして
        保存されていないため、保存する際は先に save_workbook_as で保存先パスを
        指定すること(未保存の状態で save_workbook(上書き保存)を呼ぶとエラーになる)。
        コンストラクタで visible=True の場合、実際のExcelアプリケーションで
        画面に表示しながら作成する。
        """
        used_alias = alias or f"新規ブック{len(self._workbooks) + 1}"
        if self.visible:
            wb = self._get_com_app().Workbooks.Add()
            self._workbooks[used_alias] = {"wb": wb, "path": None, "last_records": [], "com": True}
        else:
            wb = Workbook()
            self._workbooks[used_alias] = {"wb": wb, "path": None, "last_records": [], "com": False}
        self._active_alias = used_alias
        logger.info("新しいExcelブックを作成しました(alias=%s, visible=%s)", used_alias, self.visible)
        return f"created: alias={used_alias}"

    def load_workbook(self, path: str, alias: str | None = None) -> str:
        """Excelファイルを開いてアクティブにする。alias を省略した場合は
        ファイル名(拡張子抜き)を自動的に使う。同じaliasで再度読み込むと、
        既存の内容を上書きして開き直す。コンストラクタで visible=True の場合、
        実際のExcelアプリケーションで画面に表示しながら開く。
        """
        p = Path(path)
        if not p.exists():
            raise FileNotFoundError(f"Excelファイルが見つかりません: {p}")

        used_alias = alias or p.stem
        if self.visible:
            wb = self._get_com_app().Workbooks.Open(str(p.resolve()))
            self._workbooks[used_alias] = {"wb": wb, "path": p, "last_records": [], "com": True}
            sheet_names = [s.Name for s in wb.Sheets]
        else:
            keep_vba = p.suffix.lower() in _MACRO_ENABLED_EXTENSIONS
            wb = load_workbook(filename=p, data_only=True, keep_vba=keep_vba)
            self._workbooks[used_alias] = {"wb": wb, "path": p, "last_records": [], "com": False}
            sheet_names = wb.sheetnames
        self._active_alias = used_alias

        logger.info(
            "Excelを読み込みました: %s (alias=%s, シート: %s, visible=%s)",
            p, used_alias, sheet_names, self.visible,
        )
        return f"loaded: {p} (alias={used_alias})"

    def switch_workbook(self, alias: str) -> str:
        """既に開いている複数のExcelの中から、以降の操作対象(アクティブ)を切り替える。"""
        if alias not in self._workbooks:
            raise KeyError(
                f"開いていないエイリアスです: {alias}"
                f"(現在開いているもの: {list(self._workbooks.keys())})"
            )
        self._active_alias = alias
        logger.info("アクティブなExcelを切り替えました: %s", alias)
        return f"switched to: {alias}"

    def close_workbook(self, alias: str | None = None) -> str:
        """開いているExcelを閉じる(可視化モード(COM)のブックは、保存していない
        変更を破棄して実際にExcel上のウィンドウも閉じる。非可視化モードの
        ブックはメモリ上から手放すだけで、ファイル自体には影響しない)。
        alias省略時はアクティブなものを閉じる。
        """
        target = alias or self._active_alias
        if target is None or target not in self._workbooks:
            raise KeyError(f"開いていないエイリアスです: {target}")
        entry = self._workbooks[target]
        if entry.get("com"):
            try:
                entry["wb"].Close(SaveChanges=False)
            except Exception:  # noqa: BLE001
                pass
        del self._workbooks[target]
        if self._active_alias == target:
            self._active_alias = next(iter(self._workbooks), None)
        if not any(e.get("com") for e in self._workbooks.values()) and self._com_app is not None:
            try:
                self._com_app.Quit()
            except Exception:  # noqa: BLE001
                pass
            self._com_app = None
        return f"closed: {target}"

    def close(self) -> str:
        """開いている全ブックを閉じ、可視化モードのExcelアプリケーションを
        終了する(セッション終了時のクリーンアップ用)。
        """
        for alias in list(self._workbooks.keys()):
            try:
                self.close_workbook(alias)
            except Exception:  # noqa: BLE001
                pass
        if self._com_app is not None:
            try:
                self._com_app.Quit()
            except Exception:  # noqa: BLE001
                pass
            self._com_app = None
        return "closed all workbooks"

    def read_sheet_to_records(self, sheet_name: str) -> list[dict[str, Any]]:
        entry = self._active()
        self._require_openpyxl(entry, "read_sheet_to_records")
        wb = entry["wb"]
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: {sheet_name} (存在: {wb.sheetnames})")
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            entry["last_records"] = []
            return []
        header, *body = rows
        records = [dict(zip(header, row)) for row in body]
        entry["last_records"] = records
        return records

    def write_records_to_csv(self, output_path: str) -> str:
        entry = self._active()
        records = entry["last_records"]
        if not records:
            raise RuntimeError("書き出す対象のレコードがありません(read_sheet_to_records未実行)")
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        fieldnames = list(records[0].keys())
        with open(out, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(records)
        logger.info("CSVを出力しました: %s (%d行)", out, len(records))
        return str(out)

    def write_cells(self, sheet_name: str, cell_values: dict[str, Any]) -> str:
        """cell_values 例: {"B2": "GAO", "C2": 12345}"""
        entry = self._active()
        if entry.get("com"):
            ws = self._com_sheet(entry, sheet_name)
            for cell_ref, value in cell_values.items():
                ws.Range(cell_ref).Value = value
            return f"{len(cell_values)} セルを書き込みました"
        wb = entry["wb"]
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: {sheet_name}")
        ws = wb[sheet_name]
        for cell_ref, value in cell_values.items():
            ws[cell_ref] = value
        return f"{len(cell_values)} セルを書き込みました"

    def get_cell_value(self, sheet_name: str, cell_ref: str) -> str:
        """特定シート・特定セル1つの値を文字列として読み込む。
        テキスト加工パイプライン(textハンドラ)に渡す入力としてもよく使う。
        """
        entry = self._active()
        if entry.get("com"):
            ws = self._com_sheet(entry, sheet_name)
            value = ws.Range(cell_ref).Value
            return "" if value is None else _com_value_to_str(value)
        wb = entry["wb"]
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: {sheet_name} (存在: {wb.sheetnames})")
        value = wb[sheet_name][cell_ref].value
        return "" if value is None else str(value)

    def get_last_row(self, sheet_name: str, column: str | None = None) -> int:
        """最終行番号を取得する。columnを指定すると、その列で値が入っている
        最後の行番号を返す(Excelで言う Ctrl+↓ に相当)。省略時はシート全体の
        使用範囲の最終行(ws.max_row)を返す。
        """
        entry = self._active()
        if entry.get("com"):
            ws = self._com_sheet(entry, sheet_name)
            if column is None:
                used = ws.UsedRange
                return used.Row + used.Rows.Count - 1
            col_idx = int(column) if str(column).isdigit() else column_index_from_string(column)
            last_cell = ws.Cells(ws.Rows.Count, col_idx).End(_XL_UP)
            return last_cell.Row if last_cell.Value is not None else 0
        wb = entry["wb"]
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: {sheet_name}")
        ws = wb[sheet_name]
        if column is None:
            return ws.max_row

        col_idx = int(column) if str(column).isdigit() else column_index_from_string(column)
        last_row = 0
        for row in range(1, ws.max_row + 1):
            if ws.cell(row=row, column=col_idx).value is not None:
                last_row = row
        return last_row

    def get_last_column(self, sheet_name: str, row: str | int | None = None) -> str:
        """最終列を列文字("A","B",...)で取得する。rowを指定すると、その行で
        値が入っている最後の列を返す(Excelで言う行方向のCtrl+→に相当)。
        省略時はシート全体の使用範囲の最終列(ws.max_column)を返す。

        戻り値は列文字のため、そのまま {{変数名}} でセル参照の一部として
        埋め込める。さらに {{変数名+1}} / {{変数名-1}} と書くと、その1つ右/左の
        列文字が得られる(表の右隣の列に続けて書き込みたい場合等に使う)。
        """
        entry = self._active()
        if entry.get("com"):
            ws = self._com_sheet(entry, sheet_name)
            if row is None:
                used = ws.UsedRange
                return get_column_letter(used.Column + used.Columns.Count - 1)
            row_idx = int(row)
            last_cell = ws.Cells(row_idx, ws.Columns.Count).End(_XL_TO_LEFT)
            return get_column_letter(last_cell.Column) if last_cell.Value is not None else ""
        wb = entry["wb"]
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: {sheet_name}")
        ws = wb[sheet_name]
        if row is None:
            return get_column_letter(ws.max_column)

        row_idx = int(row)
        last_col = 0
        for col in range(1, ws.max_column + 1):
            if ws.cell(row=row_idx, column=col).value is not None:
                last_col = col
        # 該当行に値が1つも無い場合は「A列より前」を表す空文字列を返す
        # (get_last_rowが値0を返すのと同じ考え方。{{last_col+1}}とすれば
        # 正しく"A"が得られる)
        return get_column_letter(last_col) if last_col else ""

    def get_sheet_names(self) -> list[str]:
        """アクティブなブックのシート名一覧を取得する。"""
        entry = self._active()
        if entry.get("com"):
            return [s.Name for s in entry["wb"].Sheets]
        return list(entry["wb"].sheetnames)

    def get_range_as_list(self, sheet_name: str, cell_range: str) -> list[Any]:
        """指定範囲(1列または1行)の値をリストとして取得する。
        例: "A1:A10"(1列10件) または "A1:J1"(1行10件)。単一セルの指定も可
        (要素数1のリストになる)。2次元(複数行かつ複数列)の範囲は指定できない。
        戻り値のリストは、他のRPAツールと同様Pythonのリストと同じ0始まりの
        番号でアクセスできる({{変数名[0]}} が先頭要素)。
        """
        entry = self._active()
        self._require_openpyxl(entry, "get_range_as_list")
        wb = entry["wb"]
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: {sheet_name}")
        ws = wb[sheet_name]

        cells = ws[cell_range]
        if not isinstance(cells, tuple):
            # 単一セル
            return [cells.value]

        if cells and isinstance(cells[0], tuple):
            rows = cells
            if len(rows[0]) == 1:
                # 1列(複数行 x 1列)
                return [row[0].value for row in rows]
            if len(rows) == 1:
                # 1行(1行 x 複数列)
                return [c.value for c in rows[0]]
            raise ValueError(
                f"1列または1行の範囲のみ指定できます"
                f"(現在: {len(rows)}行 x {len(rows[0])}列): {cell_range}"
            )

        # 1行だけの範囲(cellsがセルのタプル1段のみ)
        return [c.value for c in cells]

    def find_cell(self, sheet_name: str, value: Any, column: str | None = None) -> str:
        """指定した値と一致する最初のセルのアドレス(例: "B5")を検索する。
        column を指定すると、その列だけを上から順に検索する
        (省略時はシート全体を行優先で検索する)。値は文字列化して比較する
        (Excel側の見た目の値と一致させるため)。
        """
        entry = self._active()
        self._require_openpyxl(entry, "find_cell")
        wb = entry["wb"]
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: {sheet_name}")
        ws = wb[sheet_name]
        target = str(value)

        if column is not None:
            col_idx = int(column) if str(column).isdigit() else column_index_from_string(column)
            for row in range(1, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                if cell.value is not None and str(cell.value) == target:
                    return cell.coordinate
        else:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None and str(cell.value) == target:
                        return cell.coordinate

        raise ValueError(f"値 '{value}' に一致するセルが見つかりませんでした(シート: {sheet_name})")

    def append_row(self, sheet_name: str, values: list[Any], start_column: str = "A") -> str:
        """シートの最終行の次の行に、values を横並びで書き込む(Excelでよく使う
        「表の末尾に1行追記する」操作)。start_column から右方向に書き込む。
        """
        entry = self._active()
        self._require_openpyxl(entry, "append_row")
        wb = entry["wb"]
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: {sheet_name}")
        ws = wb[sheet_name]
        start_col_idx = column_index_from_string(start_column)
        target_row = ws.max_row + 1
        if ws.max_row == 1 and all(c.value is None for c in ws[1]):
            target_row = 1
        for offset, value in enumerate(values):
            ws.cell(row=target_row, column=start_col_idx + offset, value=value)
        return f"{target_row}行目に{len(values)}件追記しました"

    def insert_rows(self, sheet_name: str, row: int, count: int = 1) -> str:
        """row で指定した行の直前に、count行分の空行を挿入する(既存の行は下にずれる)。"""
        entry = self._active()
        self._require_openpyxl(entry, "insert_rows")
        wb = entry["wb"]
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: {sheet_name}")
        if count < 1:
            raise ValueError("countは1以上を指定してください")
        wb[sheet_name].insert_rows(int(row), amount=int(count))
        return f"{row}行目の前に{count}行挿入しました"

    def delete_rows(self, sheet_name: str, row: int, count: int = 1) -> str:
        """row で指定した行から count行分を削除する(以降の行は上に詰まる)。"""
        entry = self._active()
        self._require_openpyxl(entry, "delete_rows")
        wb = entry["wb"]
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: {sheet_name}")
        if count < 1:
            raise ValueError("countは1以上を指定してください")
        wb[sheet_name].delete_rows(int(row), amount=int(count))
        return f"{row}行目から{count}行削除しました"

    def clear_range(self, sheet_name: str, cell_range: str) -> str:
        """指定範囲のセルの値をすべて空にする(書式は変更しない)。"""
        entry = self._active()
        self._require_openpyxl(entry, "clear_range")
        wb = entry["wb"]
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: {sheet_name}")
        ws = wb[sheet_name]
        cells = ws[cell_range]
        if not isinstance(cells, tuple):
            cells = ((cells,),)
        elif cells and not isinstance(cells[0], tuple):
            cells = (cells,)
        count = 0
        for row in cells:
            for cell in row:
                cell.value = None
                count += 1
        return f"{count} セルを空にしました: {cell_range}"

    def delete_sheet(self, sheet_name: str) -> str:
        """既存シートを削除する(ブック内で唯一のシートは削除できない)。"""
        entry = self._active()
        self._require_openpyxl(entry, "delete_sheet")
        wb = entry["wb"]
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: {sheet_name} (存在: {wb.sheetnames})")
        if len(wb.sheetnames) <= 1:
            raise ValueError("ブック内で唯一のシートは削除できません")
        wb.remove(wb[sheet_name])
        logger.info("シートを削除しました: %s", sheet_name)
        return f"deleted sheet: {sheet_name}"

    def copy_cell_range(
        self, sheet_name: str, source_range: str, dest_cell: str, paste_type: str = "values"
    ) -> str:
        """source_range(例: "A1:B3")を、dest_cellを左上として貼り付ける。

        paste_type:
          "values"   既定。計算済みの値のみを貼り付ける(数式は貼り付けない)。
          "formulas" 数式をそのまま貼り付ける(セル参照は自動調整されない。
                     ファイルを読み直して数式文字列を取得するため、この
                     ブック上でまだ保存していない変更が元セルにある場合、
                     その変更は数式の取得には反映されない。読み直す先は
                     load_workbook時のパスだが、その後 save_workbook_as を
                     呼んでいれば、以後はその保存先を読み直す。新規作成した
                     ブック(create_workbook)でまだ一度も保存していない
                     場合はエラーになる)。
          "all"      値/数式に加えて、数値の表示形式・フォント・塗りつぶし・
                     罫線・配置もあわせて貼り付ける(列幅/行高・条件付き書式・
                     コメント・入力規則は対象外)。
        """
        entry = self._active()
        self._require_openpyxl(entry, "copy_cell_range")
        wb = entry["wb"]
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: {sheet_name}")
        ws = wb[sheet_name]

        if paste_type not in ("values", "formulas", "all"):
            raise ValueError(
                f"paste_typeは 'values' / 'formulas' / 'all' のいずれかを指定してください: {paste_type}"
            )
        if paste_type in ("formulas", "all") and entry["path"] is None:
            raise RuntimeError(
                "このブックはまだ一度も保存先が指定されていません(新規作成したブック等)。"
                "数式を貼り付けるには、先に save_workbook_as で保存先パスを指定してください"
                "(数式は、保存済みのファイルを読み直して取得するため)。"
            )

        cells = ws[source_range]
        # 単一セルの場合はタプルのタプルにならないため揃える
        if not isinstance(cells, tuple):
            cells = ((cells,),)
        elif cells and not isinstance(cells[0], tuple):
            cells = (cells,)

        dest_col_letters = "".join(c for c in dest_cell if c.isalpha())
        dest_row_digits = "".join(c for c in dest_cell if c.isdigit())
        if not dest_col_letters or not dest_row_digits:
            raise ValueError(f"貼り付け先セルの指定が不正です: {dest_cell}")
        dest_row = int(dest_row_digits)
        dest_col = column_index_from_string(dest_col_letters)

        formula_ws = None
        if paste_type in ("formulas", "all"):
            # 数式文字列は data_only=True のブックからは取得できないため、
            # 同じファイルを data_only=False で読み直して数式だけを参照する
            formula_wb = load_workbook(filename=entry["path"], data_only=False)
            if sheet_name not in formula_wb.sheetnames:
                raise KeyError(f"シートが見つかりません(数式読込用): {sheet_name}")
            formula_ws = formula_wb[sheet_name]

        count = 0
        for r_offset, row in enumerate(cells):
            for c_offset, cell in enumerate(row):
                target = ws.cell(row=dest_row + r_offset, column=dest_col + c_offset)

                if paste_type == "values":
                    target.value = cell.value
                else:
                    src_formula_cell = formula_ws.cell(row=cell.row, column=cell.column)
                    target.value = src_formula_cell.value

                if paste_type == "all":
                    target.number_format = cell.number_format
                    target.font = _copy_style(cell.font)
                    target.fill = _copy_style(cell.fill)
                    target.border = _copy_style(cell.border)
                    target.alignment = _copy_style(cell.alignment)

                count += 1
        return f"{count} セルをコピーしました({paste_type}): {source_range} -> {dest_cell}"

    def create_sheet(self, sheet_name: str, index: int | None = None) -> str:
        """新しいシートを追加する。index省略時は末尾に追加する(0始まりで
        位置を指定可能。例: 0で先頭に追加)。"""
        entry = self._active()
        if entry.get("com"):
            wb = entry["wb"]
            names = [s.Name for s in wb.Sheets]
            if sheet_name in names:
                raise ValueError(f"シート名 '{sheet_name}' は既に存在します")
            # 注意: win32com の遅延バインディングでは Sheets.Add(After=...) の
            # ような名前付き引数が正しく解決されない(Before側に渡ってしまう)
            # ため、Add(Before, After) の位置引数で明示的に呼び出す。
            if index is None or int(index) >= wb.Sheets.Count:
                new_ws = wb.Sheets.Add(None, wb.Sheets(wb.Sheets.Count))
            else:
                idx = max(0, int(index))
                new_ws = wb.Sheets.Add(wb.Sheets(idx + 1))
            new_ws.Name = sheet_name
            logger.info("シートを追加しました: %s (index=%s)", sheet_name, index)
            return sheet_name
        wb = entry["wb"]
        if sheet_name in wb.sheetnames:
            raise ValueError(f"シート名 '{sheet_name}' は既に存在します")
        wb.create_sheet(title=sheet_name, index=index)
        logger.info("シートを追加しました: %s (index=%s)", sheet_name, index)
        return sheet_name

    def rename_sheet(self, old_name: str, new_name: str) -> str:
        """既存シートの名前を変更する。"""
        entry = self._active()
        if entry.get("com"):
            ws = self._com_sheet(entry, old_name)
            names = [s.Name for s in entry["wb"].Sheets]
            if new_name in names:
                raise ValueError(f"シート名 '{new_name}' は既に存在します")
            ws.Name = new_name
            logger.info("シート名を変更しました: %s -> %s", old_name, new_name)
            return new_name
        wb = entry["wb"]
        if old_name not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: {old_name} (存在: {wb.sheetnames})")
        if new_name in wb.sheetnames:
            raise ValueError(f"シート名 '{new_name}' は既に存在します")
        ws = wb[old_name]
        ws.title = new_name
        logger.info("シート名を変更しました: %s -> %s", old_name, new_name)
        return new_name

    def copy_sheet(self, source_sheet_name: str, new_sheet_name: str) -> str:
        """シートを丸ごとコピーする(同一ブック内)。"""
        entry = self._active()
        self._require_openpyxl(entry, "copy_sheet")
        wb = entry["wb"]
        if source_sheet_name not in wb.sheetnames:
            raise KeyError(f"シートが見つかりません: {source_sheet_name} (存在: {wb.sheetnames})")
        if new_sheet_name in wb.sheetnames:
            raise ValueError(f"シート名 '{new_sheet_name}' は既に存在します")
        src = wb[source_sheet_name]
        new_ws = wb.copy_worksheet(src)
        new_ws.title = new_sheet_name
        return new_sheet_name

    # ---------- 印刷範囲・ページ設定(PDF出力に反映される) ----------
    # set_print_area / set_page_layout / set_print_area_from_file で設定した内容は、
    # ブックごと・シートごとに entry["print_settings"] に記録され、同じファイルを
    # 対象にした save_as_pdf のPDF出力にそのまま反映される(openpyxlでメモリ上に
    # 持っているだけの設定は、ディスク上のファイルには入っていないため、
    # PDF出力時に実際のExcel(COM)へ同じ設定を適用してから書き出す)。

    def _print_settings(self, entry: dict[str, Any], sheet_name: str) -> dict[str, Any]:
        return entry.setdefault("print_settings", {}).setdefault(sheet_name, {})

    def _sheet_names_of(self, entry: dict[str, Any]) -> list[str]:
        if entry.get("com"):
            return [s.Name for s in entry["wb"].Sheets]
        return list(entry["wb"].sheetnames)

    @staticmethod
    def _normalize_cell_range(cell_range: str) -> str:
        """印刷範囲の書式(例: A1:K19 / $A$1:$K$19 / A1)を確認し、$を除いた形に揃える。
        "G50:X:88" のような書き間違いは、後続の処理で分かりにくいエラーになる前に
        ここで分かりやすいエラーにする。
        """
        text = str(cell_range).strip().replace("$", "").upper()
        if not _CELL_RANGE_RE.fullmatch(text):
            raise ValueError(
                f"印刷範囲の書式が正しくありません: '{cell_range}'"
                "(例: A1:K19 のように「列+行:列+行」で指定してください)"
            )
        return text

    def _apply_page_settings_com(self, ws, settings: dict[str, Any]) -> None:
        """記録済みの印刷設定(範囲/用紙/向き/枚数)を、COMのワークシートへ適用する。"""
        try:
            ps = ws.PageSetup
            # 「印刷範囲そのまま」のときの用紙・向き・拡大率・余白は、書き出しの直前に
            # _export_pdf_fit が印刷範囲の実寸から決める(ここでは触らない)
            fit = settings.get("paper") == _FIT_PAPER
            if "orientation" in settings and not fit:
                ps.Orientation = _XL_ORIENTATION[settings["orientation"]]
            if "paper" in settings and not fit:
                ps.PaperSize = _PAPER_SIZES[settings["paper"]]
            for key, attr in (("top", "TopMargin"), ("bottom", "BottomMargin"),
                              ("left", "LeftMargin"), ("right", "RightMargin")):
                if key in settings.get("margins", {}) and not fit:
                    setattr(ps, attr, settings["margins"][key] / 2.54 * 72)  # cm → ポイント
            if "area" in settings:
                ps.PrintArea = settings["area"]
            if "title_cols" in settings and not fit:
                ps.PrintTitleColumns = settings["title_cols"]
            if "title_rows" in settings and not fit:
                ps.PrintTitleRows = settings["title_rows"]
            if "pages_tall" in settings and not fit:
                ps.Zoom = False  # これを先に外さないと、FitToPages指定が効かない
                ps.FitToPagesTall = settings["pages_tall"]
                ps.FitToPagesWide = settings["pages_wide"]
        except Exception as e:  # noqa: BLE001
            raise RuntimeError(
                f"印刷設定をExcelへ反映できませんでした(シート '{ws.Name}'): {e}"
                "(既定のプリンターが1台も無い環境では、Excelがページ設定を変更できず、"
                "この問題が起きることがあります)"
            ) from e

    def set_print_area(self, sheet_name: str, cell_range: str) -> str:
        """指定範囲(例: "A1:H30")をそのシートの印刷範囲に設定する。この設定は、
        同じファイルに対する save_as_pdf のPDF出力にも反映される。
        """
        cell_range = self._normalize_cell_range(cell_range)
        entry = self._active()
        if entry.get("com"):
            ws = self._com_sheet(entry, sheet_name)
            self._print_settings(entry, sheet_name)["area"] = cell_range
            self._apply_page_settings_com(ws, {"area": cell_range})
        else:
            wb = entry["wb"]
            if sheet_name not in wb.sheetnames:
                raise KeyError(f"シートが見つかりません: {sheet_name}")
            wb[sheet_name].print_area = cell_range
            self._print_settings(entry, sheet_name)["area"] = cell_range
        return f"印刷範囲を設定しました: {sheet_name}!{cell_range}"

    def set_page_layout(
        self, sheet_name: str, paper_size: str = "A4", orientation: str = "portrait",
        pages_tall: int = 1, pages_wide: int = 1,
        margin_top: float | None = None, margin_bottom: float | None = None,
        margin_left: float | None = None, margin_right: float | None = None,
    ) -> str:
        """印刷の用紙サイズ・向き・「縦に何枚×横に何枚に収めるか」・余白を設定する。
        例: A4・縦向き・縦1×横1 / B5・横向き・縦5×横3。印刷範囲(set_print_area)が
        指定した枚数に収まるよう自動で縮小される(Excelの「拡大縮小印刷」の
        「次のページ数に合わせて印刷」と同じで、縮小のみ。100%より大きくは
        拡大されない)。この設定は save_as_pdf のPDF出力にも反映される。
        paper_size: A3 / A4 / A5 / B4 / B5(JIS) / Letter / Legal / 印刷範囲そのまま
          「印刷範囲そのまま」は、PDFのページを印刷範囲ちょうどの大きさ・形にする
          (壁紙などで余計な余白を出したくないとき用)。この場合、向き・枚数・余白の
          指定は使われない(1ページに、余白なしで収まる)。大きさは最大でA3に収まる
          範囲まで(それより大きい範囲は、形はそのまま縮小して1ページにする)。
          必ず sheet_name を指定した save_as_pdf で書き出すこと(全シート書き出しは不可)。
        orientation: portrait(縦向き)/ landscape(横向き)。"縦"/"横"も可。
        margin_top/bottom/left/right: 余白(cm、0〜50。小数可。例: 1.5)。
          省略(None/空欄)した辺は、今の余白のまま変更しない。
          Excelの標準は上下1.9cm・左右1.8cm、「狭い」は上下1.9・左右0.6cm。
          余白を0にしても書き出しはできるが、プリンターによっては印刷できない端がある。
        """
        paper = _normalize_paper_size(paper_size)
        if paper == _FIT_PAPER:
            entry = self._active()
            if entry.get("com"):
                self._apply_page_settings_com(self._com_sheet(entry, sheet_name), {"paper": _FIT_PAPER})
            elif sheet_name not in entry["wb"].sheetnames:
                raise KeyError(f"シートが見つかりません: {sheet_name}")
            self._print_settings(entry, sheet_name)["paper"] = _FIT_PAPER
            return (
                f"ページ設定しました: {sheet_name} {FIT_PAPER_LABEL}"
                "(PDFのページを印刷範囲ちょうどの大きさにします。向き・枚数・余白の指定は使いません)"
            )
        orient = _normalize_orientation(orientation)
        try:
            tall, wide = int(pages_tall), int(pages_wide)
        except (TypeError, ValueError) as e:
            raise ValueError(f"枚数は整数で指定してください: 縦{pages_tall} × 横{pages_wide}") from e
        if not (1 <= tall <= _MAX_FIT_PAGES and 1 <= wide <= _MAX_FIT_PAGES):
            raise ValueError(f"枚数は1〜{_MAX_FIT_PAGES}で指定してください: 縦{tall} × 横{wide}")
        margins: dict[str, float] = {}
        for key, label, value in (
            ("top", "上", margin_top), ("bottom", "下", margin_bottom),
            ("left", "左", margin_left), ("right", "右", margin_right),
        ):
            parsed = _parse_margin(label, value)
            if parsed is not None:
                margins[key] = parsed

        entry = self._active()
        settings: dict[str, Any] = {
            "paper": paper, "orientation": orient, "pages_tall": tall, "pages_wide": wide,
        }
        if margins:
            settings["margins"] = margins
        if entry.get("com"):
            ws = self._com_sheet(entry, sheet_name)
            self._apply_page_settings_com(ws, settings)
        else:
            wb = entry["wb"]
            if sheet_name not in wb.sheetnames:
                raise KeyError(f"シートが見つかりません: {sheet_name}")
            ws = wb[sheet_name]
            ws.page_setup.paperSize = _PAPER_SIZES[paper]
            ws.page_setup.orientation = orient
            ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
            ws.page_setup.fitToHeight = tall
            ws.page_setup.fitToWidth = wide
            for key, cm in margins.items():
                setattr(ws.page_margins, key, cm / 2.54)  # openpyxlの余白の単位はインチ
        stored = self._print_settings(entry, sheet_name)
        stored.update({k: v for k, v in settings.items() if k != "margins"})
        if margins:
            stored.setdefault("margins", {}).update(margins)  # 指定しなかった辺の以前の設定は残す
        label = "縦向き" if orient == "portrait" else "横向き"
        names = {"top": "上", "bottom": "下", "left": "左", "right": "右"}
        margin_text = (
            "・余白 " + " ".join(f"{names[k]}{margins[k]:g}cm" for k in ("top", "bottom", "left", "right") if k in margins)
        ) if margins else ""
        return f"ページ設定しました: {sheet_name} {paper}・{label}・縦{tall}枚×横{wide}枚に収める{margin_text}"

    def lookup_range_for_date(
        self, source_path: str, date: str = "", source_sheet: str = "", default_range: str = "",
    ) -> str:
        """「日付 と セル範囲」の対応表を書いた別のファイルから、指定した日付の
        セル範囲(例: A2:K19)を取り出して返す。store_as で変数に保存して、
        {{変数名}} として set_print_area 等で使うこともできる。

        source_path: Excel(.xlsx/.xlsm)または テキスト(.txt/.csv)。
          Excelは先頭シート(source_sheetで指定も可)のA列=日付、B列=範囲。
          テキストは1行に「日付,範囲」(区切りはカンマ/タブ/空白のいずれか)。
          見出し行や、日付として読めない行は読み飛ばす。
          例)  2026/10/20,A2:K19
               2026/10/21,B5:C80
        date: 取り出す日付。空欄なら実行した日(今日)。"2026/10/20" "2026-10-20"
          "2026年10月20日" などで指定できる。
        default_range: その日付の行が無い場合に使う範囲。空欄ならエラーにする
          (範囲が決まらないまま誤った範囲を印刷しないため)。
        同じ日付の行が複数あれば、後ろの行を優先する。
        """
        target = _parse_date(date) if str(date).strip() else _dt.date.today()
        rows = _read_date_range_table(Path(source_path), source_sheet)
        found = None
        for row_date, cell_range in rows:
            if row_date == target:
                found = cell_range
        if found is None:
            if str(default_range).strip():
                return self._normalize_cell_range(default_range)
            raise KeyError(
                f"{target:%Y/%m/%d} の範囲が対応表にありません: {source_path}"
                f"(表にある日付: {len(rows)}件。既定の範囲を指定すれば、無い日はそれを使えます)"
            )
        return self._normalize_cell_range(found)

    def set_print_area_from_file(
        self, sheet_name: str, source_path: str, date: str = "", source_sheet: str = "",
        default_range: str = "",
    ) -> str:
        """lookup_range_for_date で別ファイルの「日付 と セル範囲」の対応表から
        指定日(省略時は今日)の範囲を取り出し、そのシートの印刷範囲に設定する。
        毎日印刷する範囲が変わる場合に、実行日に合わせて自動で切り替えられる。
        """
        cell_range = self.lookup_range_for_date(source_path, date, source_sheet, default_range)
        target = _parse_date(date) if str(date).strip() else _dt.date.today()
        result = self.set_print_area(sheet_name, cell_range)
        return f"{target:%Y/%m/%d} の範囲を対応表から取り出しました。{result}"

    def _search_line_cells(self, entry: dict[str, Any], sheet_name: str, search_in: str, line: Any):
        """検索する行(または列)のセルを、先頭から順に (1始まりの位置, 値, 表示文字列を返す関数)
        として返す。表示文字列は、日付セルが「10/20」のように表示されている場合の
        その見た目(COMはExcelの表示そのまま、openpyxlは表示形式から再現)。"""
        axis = "row" if str(search_in).strip().lower() in ("row", "行") else "column"
        if axis == "row":
            try:
                index = int(str(line).strip())
            except ValueError as e:
                raise ValueError(f"探す行は行番号(数字)で指定してください: {line}") from e
        else:
            text = str(line).strip().upper()
            index = int(text) if text.isdigit() else column_index_from_string(text)
        if index < 1:
            raise ValueError(f"探す{'行' if axis == 'row' else '列'}の指定が不正です: {line}")

        cells: list[tuple[int, Any, Any]] = []
        if entry.get("com"):
            ws = self._com_sheet(entry, sheet_name)
            used = ws.UsedRange
            last_row = used.Row + used.Rows.Count - 1
            last_col = used.Column + used.Columns.Count - 1
            if axis == "row":
                rng = ws.Range(ws.Cells(index, 1), ws.Cells(index, last_col))
                values = rng.Value[0] if last_col > 1 else (rng.Value,)
                for pos, value in enumerate(values, start=1):
                    cells.append((pos, value, lambda r=index, c=pos: _com_display(ws, r, c)))
            else:
                rng = ws.Range(ws.Cells(1, index), ws.Cells(last_row, index))
                values = [row[0] for row in rng.Value] if last_row > 1 else [rng.Value]
                for pos, value in enumerate(values, start=1):
                    cells.append((pos, value, lambda r=pos, c=index: _com_display(ws, r, c)))
        else:
            wb = entry["wb"]
            if sheet_name not in wb.sheetnames:
                raise KeyError(f"シートが見つかりません: {sheet_name}")
            ws = wb[sheet_name]
            count = ws.max_column if axis == "row" else ws.max_row
            for pos in range(1, count + 1):
                cell = ws.cell(row=index, column=pos) if axis == "row" else ws.cell(row=pos, column=index)
                cells.append((pos, cell.value, lambda cell=cell: _display_from_format(cell.value, cell.number_format)))
        return axis, index, cells

    def find_cell_in_line(
        self, sheet_name: str, search_in: str, line: Any, target: str, match: str = "exact",
    ) -> str:
        """指定した行(search_in="row"・lineは行番号)または列(search_in="column"・
        lineは列文字)の中から、targetに一致する最初のセルを探し、そのアドレス
        (例: "M3")を返す。store_as で変数に保存して使うこともできる。
        target: 探す文字列。先頭が"="なら式として評価する(対応: =TODAY() /
          =TODAY()+1 のような日数の加減 / =TEXT(TODAY(),"m/d") のように日付を
          Excelの書式文字列で文字にしたもの)。=TODAY()系は日付セルと日付として
          比べ、=TEXT(...)や普通の文字列は、セルの値と、日付セルの場合は画面の
          表示("10/20"等)の両方と比べる。
        match: "exact"(完全一致)/ "contains"(部分一致)。
        """
        kind, wanted = _evaluate_search_target(target)
        mode = str(match).strip().lower()
        if mode in ("部分一致", "含む"):
            mode = "contains"
        elif mode in ("完全一致", ""):
            mode = "exact"
        if mode not in ("exact", "contains"):
            raise ValueError(f"matchは exact(完全一致)か contains(部分一致)で指定してください: {match}")

        entry = self._active()
        axis, index, cells = self._search_line_cells(entry, sheet_name, search_in, line)
        for pos, value, display in cells:
            if _cell_matches(value, display, kind, wanted, mode):
                return f"{get_column_letter(pos)}{index}" if axis == "row" else f"{get_column_letter(index)}{pos}"
        where = f"{index}行目" if axis == "row" else f"{get_column_letter(index)}列"
        shown = wanted if kind == "text" else f"{wanted:%Y/%m/%d}"
        raise KeyError(
            f"シート '{sheet_name}' の{where}に、'{target}'(= {shown})と一致するセルが見つかりません"
        )

    @staticmethod
    def _normalize_title_spec(spec: str, kind: str) -> str:
        """印刷タイトルの指定(列: "A" / "A:B"、行: "3" / "1:3")をExcelの形式
        ($A:$A / $3:$3)にする。空欄なら空文字(指定なし)。"""
        text = str(spec or "").strip().replace("$", "").upper()
        if not text:
            return ""
        pattern = r"[A-Z]{1,3}(:[A-Z]{1,3})?" if kind == "cols" else r"[0-9]+(:[0-9]+)?"
        if not re.fullmatch(pattern, text):
            example = "A または A:B" if kind == "cols" else "3 または 1:3"
            raise ValueError(
                f"印刷タイトルの{'列' if kind == 'cols' else '行'}の書式が正しくありません: '{spec}'(例: {example})"
            )
        parts = text.split(":")
        if len(parts) == 1:
            parts = parts * 2
        return f"${parts[0]}:${parts[1]}"

    def set_print_area_by_search(
        self, sheet_name: str, search_in: str = "row", line: Any = 1, target: str = "",
        match: str = "exact", delta_cols: int = 0, delta_rows: int = 0,
        offset_cols: int = 0, offset_rows: int = 0,
        title_columns: str = "", title_rows: str = "", default_range: str = "",
    ) -> str:
        """開いているシートの特定の行(または列)から、targetに一致するセルを探し、
        そのセルを基準にした範囲を印刷範囲に設定する。毎日1つずつ右(または下)に
        ずれていくカレンダー等で、「今日の列から1週間分」を自動で印刷するための機能。

        一致したセルの位置を (行r, 列c) とすると、
          開始 = (r + offset_rows, c + offset_cols)
          終了 = 開始 + (delta_rows, delta_cols)
        の範囲を印刷範囲にする(delta=0なら開始と同じ位置まで。例: 見出しが3行目に
        日付で並び、今日の列から7日分・担当4行分なら、search_in="row", line=3,
        target="=TODAY()", delta_cols=6, delta_rows=4 → 例: M3:S7)。deltaや
        offsetに負の数を指定すれば、左・上へ広げることもできる。
        search_in/line/target/match は find_cell_in_line と同じ。
        title_columns(例: "A" / "A:B")・title_rows(例: "3" / "1:3"): 印刷範囲が
        右(下)にずれても、名前の列/見出しの行を毎回同じページの左端(上端)に
        付けて印刷する「印刷タイトル」。担当者名の列を一緒に出したいときに使う。
        default_range: 一致するセルが無かったときに使う範囲(空欄ならエラーで止まる)。
        この設定は save_as_pdf のPDF出力にも反映される。
        """
        try:
            d_cols, d_rows = int(delta_cols), int(delta_rows)
            o_cols, o_rows = int(offset_cols), int(offset_rows)
        except (TypeError, ValueError) as e:
            raise ValueError("広げる幅・開始位置のずらしは整数で指定してください") from e
        t_cols = self._normalize_title_spec(title_columns, "cols")
        t_rows = self._normalize_title_spec(title_rows, "rows")

        try:
            matched = self.find_cell_in_line(sheet_name, search_in, line, target, match)
        except KeyError:
            if not str(default_range).strip():
                raise
            matched = None

        if matched is None:
            cell_range = self._normalize_cell_range(default_range)
            note = "一致するセルが無かったため既定の範囲を使いました"
        else:
            m = re.fullmatch(r"([A-Z]+)(\d+)", matched)
            col, row = column_index_from_string(m.group(1)), int(m.group(2))
            start_col, start_row = col + o_cols, row + o_rows
            end_col, end_row = start_col + d_cols, start_row + d_rows
            c1, c2 = sorted((start_col, end_col))
            r1, r2 = sorted((start_row, end_row))
            if c1 < 1 or r1 < 1:
                raise ValueError(
                    f"一致したセル({matched})から計算した範囲がシートの外にはみ出します"
                    f"(開始 {get_column_letter(max(c1, 1))}{max(r1, 1)} 付近。ずらし/広げる幅を見直してください)"
                )
            cell_range = f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}"
            note = f"一致したセル: {matched}"

        result = self.set_print_area(sheet_name, cell_range)
        entry = self._active()
        settings = self._print_settings(entry, sheet_name)
        if t_cols or t_rows:
            if entry.get("com"):
                ws = self._com_sheet(entry, sheet_name)
                self._apply_page_settings_com(ws, {k: v for k, v in (("title_cols", t_cols), ("title_rows", t_rows)) if v})
            else:
                ws = entry["wb"][sheet_name]
                if t_cols:
                    ws.print_title_cols = t_cols.replace("$", "")
                if t_rows:
                    ws.print_title_rows = t_rows.replace("$", "")
            if t_cols:
                settings["title_cols"] = t_cols
            if t_rows:
                settings["title_rows"] = t_rows
        def short(spec: str) -> str:
            a, b = spec.replace("$", "").split(":")
            return a if a == b else f"{a}:{b}"

        titles = "".join([
            f"、印刷タイトル(列 {short(t_cols)})" if t_cols else "",
            f"、印刷タイトル(行 {short(t_rows)})" if t_rows else "",
        ])
        return f"{note}。{result}{titles}"

    def save_workbook_as(self, output_path: str) -> str:
        entry = self._active()
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        if entry.get("com"):
            entry["wb"].SaveAs(str(out.resolve()))
        else:
            # .xlsm/.xltmとして保存する場合、load_workbook時にkeep_vba=Trueで
            # 読み込んでいれば、既存のVBAはそのまま保持される
            entry["wb"].save(out)
        entry["path"] = out  # 以後 save_workbook(上書き保存)はこのパスに対して行う
        logger.info("Excelを保存しました: %s", out)
        return str(out)

    def save_workbook(self) -> str:
        """アクティブなブックを、読み込んだとき(または前回save_workbook_asで
        指定したとき)と同じパスへ、名前を変えずに上書き保存する。
        新規作成したブック(create_workbook)でまだ一度も保存先が決まっていない
        場合はエラーになるため、先に save_workbook_as で保存先パスを指定すること。
        """
        entry = self._active()
        if entry["path"] is None:
            raise RuntimeError(
                "このブックはまだ一度も保存先が指定されていません(新規作成したブック等)。"
                "先に save_workbook_as で保存先パスを指定してください。"
            )
        if entry.get("com"):
            entry["wb"].Save()
        else:
            entry["wb"].save(entry["path"])
        logger.info("Excelを上書き保存しました: %s", entry["path"])
        return str(entry["path"])

    def _find_entry_by_path(self, p: Path) -> dict[str, Any] | None:
        """今開いているブックのうち、ファイルのパスがpと同じもの(複数あれば
        後から開いたもの)を返す。無ければNone。"""
        found = None
        for entry in self._workbooks.values():
            ep = entry.get("path")
            if ep is not None and Path(ep).resolve() == p:
                found = entry
        return found

    def _apply_saved_print_settings(self, wb, settings_by_sheet: dict[str, dict[str, Any]]) -> None:
        """set_print_area / set_page_layout で記録済みの設定を、COMで開いた
        ブックの各シートへ適用する(ディスク上のファイルには保存しない)。"""
        for sheet, settings in settings_by_sheet.items():
            try:
                ws = wb.Sheets(sheet)
            except Exception:  # noqa: BLE001
                logger.warning("印刷設定のあるシート '%s' がファイルに無いため、反映をスキップします", sheet)
                continue
            self._apply_page_settings_com(ws, settings)

    @staticmethod
    def _pdf_content_box(pdf_path: Path) -> tuple[int, float, float, float, float]:
        """PDFの(ページ数, 1ページ目の塗りの外接長方形 x0, top, x1, bottom)を返す。"""
        import pdfplumber

        with pdfplumber.open(str(pdf_path)) as pdf:
            rects = pdf.pages[0].rects
            if not rects:
                raise RuntimeError("印刷範囲の大きさをPDFから測れませんでした(描画された領域がありません)")
            return (
                len(pdf.pages),
                min(r["x0"] for r in rects), min(r["top"] for r in rects),
                max(r["x1"] for r in rects), max(r["bottom"] for r in rects),
            )

    def _export_pdf(self, wb, sheet_name: str | None, out: Path, fit: bool) -> None:
        if fit:
            self._export_pdf_fit(wb, wb.Sheets(sheet_name), out)
            return
        target = wb.Sheets(sheet_name) if sheet_name else wb
        target.ExportAsFixedFormat(0, str(out))  # xlTypePDF = 0

    def _export_pdf_fit(self, wb, ws, out: Path) -> None:
        """印刷範囲ちょうどの大きさのPDF(1ページ)を書き出す。

        PDF上のセルの大きさは、Excelが画面上で報告する列幅・行の高さと2〜7%ずれる
        (印刷時はプリンターの解像度で丸められるため)。そのため計算では切り抜けないので、
        ①余白0・A3・縮小率を決めて、②保存しない一時コピーで印刷範囲を塗りつぶして
        書き出し、PDF上の実際の範囲(外接長方形)を測り、③本番のPDFを同じ設定で書き出して
        その長方形に切り抜く。
        """
        import shutil
        import tempfile

        from pypdf import PdfReader, PdfWriter
        from pypdf.generic import RectangleObject

        ps = ws.PageSetup
        area = (ps.PrintArea or "").replace("$", "")
        if not area:
            area = ws.UsedRange.Address.replace("$", "")
        if "," in area:
            raise ValueError(
                f"印刷範囲が離れた複数の範囲({area})になっているため、「{FIT_PAPER_LABEL}」にできません"
                "(1つの長方形の範囲にしてください)"
            )
        rng = ws.Range(area)
        width, height = float(rng.Width), float(rng.Height)
        if width <= 0 or height <= 0:
            raise ValueError(f"印刷範囲 {area} の大きさが0です(行・列がすべて非表示ではありませんか)")

        landscape = width >= height
        page_w, page_h = (_A3_POINTS[1], _A3_POINTS[0]) if landscape else _A3_POINTS
        # PDF上は画面上より最大約10%大きくなることがあるので、その分の余裕を見て縮小率を決める
        zoom = int(min(100.0, (page_w - 8) / (width * 1.1) * 100, (page_h - 8) / (height * 1.1) * 100))

        ps.PrintArea = area
        ps.PrintTitleColumns = ""
        ps.PrintTitleRows = ""
        ps.PaperSize = _PAPER_SIZES["A3"]
        ps.Orientation = _XL_ORIENTATION["landscape" if landscape else "portrait"]
        for attr in ("TopMargin", "BottomMargin", "LeftMargin", "RightMargin", "HeaderMargin", "FooterMargin"):
            setattr(ps, attr, 0)
        ps.CenterHorizontally = False
        ps.CenterVertically = False
        for attr in ("LeftHeader", "CenterHeader", "RightHeader", "LeftFooter", "CenterFooter", "RightFooter"):
            setattr(ps, attr, "")

        win32com = _import_win32com()
        tmpdir = Path(tempfile.mkdtemp(prefix="rpa_fit_"))
        try:
            probe_excel = win32com.DispatchEx("Excel.Application")
            try:
                probe_excel.Visible = False
                probe_excel.DisplayAlerts = False
                probe_excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable: コピーのマクロは動かさない
                suffix = Path(str(wb.FullName)).suffix or ".xlsx"
                while True:
                    if zoom < 10:
                        raise ValueError(
                            f"印刷範囲 {area} が大きすぎて、A3の1ページ(最小縮小10%)に収まりません"
                        )
                    ps.Zoom = zoom
                    copy, probe_pdf = tmpdir / f"probe{suffix}", tmpdir / "probe.pdf"
                    wb.SaveCopyAs(str(copy))
                    pb = probe_excel.Workbooks.Open(str(copy))
                    try:
                        pws = pb.Sheets(ws.Name)
                        pws.Range(area).Interior.Color = 0x0000FF  # 範囲の外接長方形を測る目印(コピー上のみ)
                        pws.ExportAsFixedFormat(0, str(probe_pdf))
                    finally:
                        pb.Close(SaveChanges=False)
                    pages, x0, top, x1, bottom = self._pdf_content_box(probe_pdf)
                    if pages == 1:
                        break
                    zoom = int(zoom * 0.9)  # 1ページに収まらなかったので、少し縮小してやり直す
            finally:
                probe_excel.Quit()

            full_pdf = tmpdir / "full.pdf"
            ws.ExportAsFixedFormat(0, str(full_pdf))
            page = PdfReader(str(full_pdf)).pages[0]
            box = page.mediabox
            left, base = float(box.left), float(box.bottom)
            page_height = float(box.height)
            pad = 0.5  # 外枠の罫線は、セルの境目に太さの半分ほどはみ出して描かれるため、少しだけ広げる
            new_box = RectangleObject([
                max(left, left + x0 - pad), max(base, base + page_height - bottom - pad),
                min(float(box.right), left + x1 + pad), min(float(box.top), base + page_height - top + pad),
            ])
            page.mediabox = new_box
            page.cropbox = new_box
            writer = PdfWriter()
            writer.add_page(page)
            with open(out, "wb") as f:
                writer.write(f)
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
        logger.info("印刷範囲 %s(縮小率%d%%)ちょうどの大きさでPDFを書き出しました: %s", area, zoom, out)

    def save_as_pdf(self, path: str, output_path: str, sheet_name: str | None = None) -> str:
        """Excelファイルを、印刷範囲・ページ設定を反映してPDFへ書き出す
        (**Windows + Excelインストール環境専用**)。
        sheet_name を指定しない場合はブック全体(全シート)を書き出す。

        反映される印刷設定:
        - このツールで同じファイルを開いて set_print_area(印刷範囲) /
          set_page_layout(用紙サイズ・向き・縦横の枚数)/ set_print_area_from_file
          (日付で切り替わる印刷範囲)を実行していれば、その内容を反映する
          (ファイルには書き込まず、PDF出力用に開いたExcelへ一時的に適用する)。
        - ファイル側にもともと設定されている印刷範囲・ページ設定は、上記で
          指定していない項目についてはそのまま使われる。
        アクティブブックの状態とは独立して、指定したpathに対して直接動作する
        (pathは、設定を行ったブックを開いたときと同じファイルを指定すること)。

        コンストラクタで visible=True の場合、実際のExcelアプリケーションで
        画面に表示しながら書き出す(結果を目で見て確認できるよう、書き出し後も
        ブックは開いたままにする。閉じるにはclose_workbook/closeを使う)。
        既に同じファイルを開いている場合は、そのブックをそのまま書き出す。
        visible=False(既定)の場合は従来通り、専用の非表示Excelプロセスで
        開いて書き出し、即座に閉じる。
        """
        p = Path(path).resolve()
        if not p.exists():
            raise FileNotFoundError(f"Excelファイルが見つかりません: {p}")
        out = Path(output_path).resolve()
        out.parent.mkdir(parents=True, exist_ok=True)

        matched = self._find_entry_by_path(p)
        settings_by_sheet = (matched or {}).get("print_settings", {})
        fit_sheets = [n for n, st in settings_by_sheet.items() if st.get("paper") == _FIT_PAPER]
        if sheet_name is None and fit_sheets:
            raise ValueError(
                f"用紙サイズを「{FIT_PAPER_LABEL}」にしたシート({', '.join(fit_sheets)})があるため、"
                "書き出すシート名を指定してください(全シート書き出しでは、ページごとの大きさを揃えられません)"
            )
        fit = sheet_name in fit_sheets

        if self.visible:
            if matched is not None and matched.get("com"):
                # 既に開いているブックには印刷設定が適用済みなので、そのまま書き出す
                self._export_pdf(matched["wb"], sheet_name, out, fit)
                logger.info("開いているExcelをPDFに書き出しました: %s -> %s", p, out)
                return str(out)
            wb = self._get_com_app().Workbooks.Open(str(p))
            self._apply_saved_print_settings(wb, settings_by_sheet)
            self._export_pdf(wb, sheet_name, out, fit)
            alias = self._unique_alias(p.stem)
            self._workbooks[alias] = {"wb": wb, "path": p, "last_records": [], "com": True}
            self._active_alias = alias
            logger.info(
                "ExcelをPDFに書き出しました(画面表示のまま開いています。alias=%s): %s -> %s",
                alias, p, out,
            )
            return str(out)

        win32com = _import_win32com()
        excel = win32com.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(str(p))
            try:
                self._apply_saved_print_settings(wb, settings_by_sheet)
                self._export_pdf(wb, sheet_name, out, fit)
            finally:
                wb.Close(SaveChanges=False)
        finally:
            excel.Quit()

        logger.info("ExcelをPDFに書き出しました: %s -> %s", p, out)
        return str(out)

    def run_excel_macro(self, path: str, macro_name: str, save_after: bool = True) -> str:
        """.xlsm/.xltmに既に組み込まれているVBAマクロを実行する
        (**Windows + Excelインストール環境専用**)。openpyxlはVBAコードの
        実行を一切行えないため、実際のExcelアプリケーションを操作する。
        アクティブブックの状態とは独立して、指定したpathに対して直接動作する。

        コンストラクタで visible=True の場合、実際のExcelアプリケーションで
        画面に表示しながら実行する(マクロの動きを目で見て確認できるよう、
        実行後もブックは開いたままにする。閉じるにはclose_workbook/closeを使う)。
        visible=False(既定)の場合は従来通り、専用の非表示Excelプロセスで
        開いて実行し、即座に閉じる。
        """
        p = Path(path).resolve()
        if not p.exists():
            raise FileNotFoundError(f"Excelファイルが見つかりません: {p}")

        if self.visible:
            excel = self._get_com_app()
            wb = excel.Workbooks.Open(str(p))
            try:
                excel.Run(macro_name)
            except Exception as e:  # noqa: BLE001
                wb.Close(SaveChanges=False)
                raise RuntimeError(f"マクロ '{macro_name}' の実行に失敗しました: {e}") from e
            if save_after:
                wb.Save()
            alias = self._unique_alias(p.stem)
            self._workbooks[alias] = {"wb": wb, "path": p, "last_records": [], "com": True}
            self._active_alias = alias
            logger.info(
                "Excelマクロを実行しました(画面表示のまま開いています。alias=%s): %s (%s)",
                alias, macro_name, p,
            )
            return f"macro executed: {macro_name}"

        win32com = _import_win32com()
        excel = win32com.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(str(p))
            try:
                try:
                    excel.Run(macro_name)
                except Exception as e:  # noqa: BLE001
                    raise RuntimeError(f"マクロ '{macro_name}' の実行に失敗しました: {e}") from e
                if save_after:
                    wb.Save()
            finally:
                wb.Close(SaveChanges=save_after)
        finally:
            excel.Quit()

        logger.info("Excelマクロを実行しました: %s (%s)", macro_name, p)
        return f"macro executed: {macro_name}"
