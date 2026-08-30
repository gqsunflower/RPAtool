"""
TextHandler: Excel/PDF/Webなどから取得した文字列を加工するためのハンドラ。

「文字を取得する機能がある場合はその後の文字の加工もできるように」という
要望に応えるための領域。単体では意味を持たず、Excelのセル読込や
PDFの範囲テキスト抽出・OCR結果などを受け取って、切り出し・置換・
日付時刻の付加といった後処理を行う。

前の手順の結果を後の手順で使う(パイプライン)には、MacroExecutorの
"store_as" 機構を使う。手順に "store_as": "変数名" を付けておくと、その
手順の実行結果が変数として記録され、以降の手順の params の中で
"{{変数名}}" と書けば参照できる(値全体でも、"A1:B{{last_row}}" のように
他の文字列に埋め込んでも良い)。

OSのクリップボードとの連携(copy_to_clipboard / get_from_clipboard)にも
対応しており、取得した値を他のアプリへ手動で貼り付けたい場合や、逆に
手動でコピーしておいた値をマクロの入力として取り込みたい場合に使える。
要 pyperclip(Linuxではさらにxclip/xselが必要)。
"""
from __future__ import annotations

import re
from datetime import datetime

_DATETIME_TOKEN_RE = re.compile(r"yyyy|YYYY|MM|dd|hh|mm|ss|M|d")


class TextHandler:
    def cut_from_marker(
        self,
        text: str,
        marker: str,
        include_marker: bool = False,
        length: int | None = None,
        end_marker: str | None = None,
    ) -> str:
        """textの中から marker を探し、そこ(を含む/含まない)から
        length で指定した文字数、または end_marker が見つかるまでを切り出す。
        length と end_marker の両方が指定された場合は end_marker を優先する。
        どちらも指定しない場合は marker 以降の文字列すべてを返す。
        """
        idx = text.find(marker)
        if idx == -1:
            raise ValueError(f"'{marker}' が見つかりませんでした")
        start = idx if include_marker else idx + len(marker)

        if end_marker:
            end_idx = text.find(end_marker, start)
            if end_idx == -1:
                raise ValueError(f"終わりの目印 '{end_marker}' が見つかりませんでした")
            return text[start:end_idx]

        if length is not None:
            return text[start:start + length]

        return text[start:]

    def replace_text(self, text: str, search: str, replace: str) -> str:
        """textの中の search を replace に置換する(すべて置換)。"""
        return text.replace(search, replace)

    def format_now(self, format_code: str, fiscal_year_start_month: int = 4) -> str:
        """現在の日付・時刻を、指定した書式コードの文字列に変換する
        (Excelのユーザー定義表示形式と似た書き方の、よく使うものだけに絞った書式)。

        使えるトークン(大文字・小文字を区別する):
          yyyy  西暦年(4桁。例: 2026)
          YYYY  年度(4桁。fiscal_year_start_monthで開始月を指定、既定4月始まり)
          MM    月(2桁ゼロ埋め。例: 09)
          M     月(ゼロ埋めなし。例: 9)
          dd    日(2桁ゼロ埋め)
          d     日(ゼロ埋めなし)
          hh    時(2桁ゼロ埋め、24時間表記)
          mm    分(2桁ゼロ埋め)
          ss    秒(2桁ゼロ埋め)
        トークン以外の文字(区切りの"_"や"年","月"等)はそのまま残る。

        よく使う書式の例:
          "yyyyMMdd_hhmmss" → "20260830_143022"(ファイル名の一意化によく使う)
          "yyyyMMdd"        → "20260830"
          "hhmmss"          → "143022"
          "yyyy"            → "2026"(西暦年)
          "YYYY"            → "2025"(年度。4月始まりなら1〜3月は前年扱い)

        保存ファイル名(例: "report_{{now}}.xlsx")やセルへの入力、Webフォームへの
        入力など、Excel/PDF/エクスプローラー/Webのどの手順の値欄でも
        store_asで変数に保存すれば {{変数名}} として埋め込んで使える。
        """
        now = datetime.now()
        fiscal_year = now.year if now.month >= fiscal_year_start_month else now.year - 1

        def _replace(match: re.Match) -> str:
            token = match.group(0)
            if token == "yyyy":
                return f"{now.year:04d}"
            if token == "YYYY":
                return f"{fiscal_year:04d}"
            if token == "MM":
                return f"{now.month:02d}"
            if token == "M":
                return str(now.month)
            if token == "dd":
                return f"{now.day:02d}"
            if token == "d":
                return str(now.day)
            if token == "hh":
                return f"{now.hour:02d}"
            if token == "mm":
                return f"{now.minute:02d}"
            return f"{now.second:02d}"  # ss

        return _DATETIME_TOKEN_RE.sub(_replace, format_code)

    def append_text(self, text: str, suffix: str, before: bool = False) -> str:
        """textの前(before=True)または後ろ(既定)に文字列を付け加える。"""
        return f"{suffix}{text}" if before else f"{text}{suffix}"

    def combine_text(self, parts: list[str], separator: str = "") -> str:
        """複数の文字列を区切り文字でつなげる。"""
        return separator.join(parts)

    def copy_to_clipboard(self, text: str) -> str:
        """文字列をOSのクリップボードにコピーする(手動で他のアプリに貼り付けたい場合等)。
        要 pyperclip(pip install pyperclip)。Linuxではさらに xclip または xsel が必要。
        """
        try:
            import pyperclip
        except ImportError as e:
            raise RuntimeError(
                "クリップボードへのコピーには pyperclip が必要です(pip install pyperclip)。"
                "Linuxではさらに xclip または xsel のインストールが必要です。"
            ) from e
        pyperclip.copy(text)
        return f"copied to clipboard: {text[:80]}"

    def get_from_clipboard(self) -> str:
        """今クリップボードに入っている文字列を取得する
        (手動でコピーしておいた値をマクロの入力として使いたい場合等)。
        要 pyperclip(pip install pyperclip)。Linuxではさらに xclip または xsel が必要。
        """
        try:
            import pyperclip
        except ImportError as e:
            raise RuntimeError(
                "クリップボードの取得には pyperclip が必要です(pip install pyperclip)。"
                "Linuxではさらに xclip または xsel のインストールが必要です。"
            ) from e
        return pyperclip.paste()
