"""
PdfHandler: pdfplumber(読み取り)と pypdf(結合・分割・回転・書き出し)を
基本とするハンドラ。OCRのみ、任意で pytesseract + pdf2image を使う
(OS側にTesseract OCRとPopplerのインストールが必要)。

範囲指定でのテキスト取得(extract_text_in_area):
  ページ幅・高さに対する割合(0〜100%)でx_left/x_right/y_upper/y_lowerを
  指定し、その矩形内の文字だけを取得できる。基準点はページ左上(0,0)、
  yは下方向に増加する。ocr=Trueにすると該当範囲を画像として切り出して
  OCRにかける(スキャンPDF向け)。改行は "ж" に置換した1行の文字列で返す。

ページ切り出し(split_pdf / extract_page_range):
  「何ページから何ページを1ファイルに抜き出す」(extract_page_range)、
  「全ページを1ページずつ、または指定ページ数ごとに分割する」(split_pdf)
  の両方に対応し、split_pdfは出力ファイル名のルールも指定できる。

ネットワークアクセスは一切行いません。
"""
from __future__ import annotations

import csv
import logging
from pathlib import Path

import pdfplumber
from pypdf import PdfReader, PdfWriter

logger = logging.getLogger("rpa_local_ai.pdf")


class PdfHandler:
    def get_page_count(self, input_path: str) -> int:
        """PDFの総ページ数を取得する。"""
        p = Path(input_path)
        if not p.exists():
            raise FileNotFoundError(f"PDFが見つかりません: {p}")
        return len(PdfReader(str(p)).pages)

    def extract_tables(self, input_path: str, output_path: str, page_number: int | None = None) -> str:
        """PDF内の表(罫線等から検出できる表組み)をCSVとして書き出す。
        page_number を指定するとそのページだけ、省略時は全ページを対象にする。
        1ページに複数の表がある場合や複数ページにまたがる場合は、表と表の間に
        空行を挟んで1つのCSVにまとめて書き出す。罫線の無い表(スペースだけで
        列を揃えている等)は検出できないことがある。
        """
        p = Path(input_path)
        if not p.exists():
            raise FileNotFoundError(f"PDFが見つかりません: {p}")

        all_rows: list[list[str]] = []
        with pdfplumber.open(p) as pdf:
            if page_number is not None and (page_number < 1 or page_number > len(pdf.pages)):
                raise ValueError(f"ページ{page_number}が見つかりません(全{len(pdf.pages)}ページ)")
            pages = [pdf.pages[page_number - 1]] if page_number else pdf.pages
            for page in pages:
                for table in page.extract_tables():
                    if all_rows:
                        all_rows.append([])
                    for row in table:
                        all_rows.append(["" if c is None else str(c) for c in row])

        if not all_rows:
            raise ValueError("表が検出できませんでした(罫線の無い表は検出できない場合があります)")

        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        with open(out, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerows(all_rows)
        logger.info("PDFから表を抽出しました: %s -> %s", p, out)
        return str(out)

    def extract_images(self, input_path: str, output_dir: str, page_number: int | None = None) -> list[str]:
        """PDFに埋め込まれている画像をファイルとして書き出す(写真・署名画像等の
        抽出に使う。文字を画像化しただけの「スキャンPDF全体」を画像として
        欲しい場合は用途が異なるので注意)。page_number省略時は全ページが対象。
        出力ファイル名は "{stem}_p{ページ番号}_{連番}.{拡張子}"。
        """
        p = Path(input_path)
        if not p.exists():
            raise FileNotFoundError(f"PDFが見つかりません: {p}")
        reader = PdfReader(str(p))
        total = len(reader.pages)
        if page_number is not None and (page_number < 1 or page_number > total):
            raise ValueError(f"ページ{page_number}が見つかりません(全{total}ページ)")

        out_dir = Path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        outputs: list[str] = []
        page_indices = [page_number - 1] if page_number else range(total)
        for page_idx in page_indices:
            for i, img in enumerate(reader.pages[page_idx].images, start=1):
                suffix = Path(img.name).suffix or ".png"
                out_path = out_dir / f"{p.stem}_p{page_idx + 1}_{i}{suffix}"
                out_path.write_bytes(img.data)
                outputs.append(str(out_path))

        if not outputs:
            raise ValueError("埋め込み画像が見つかりませんでした")
        logger.info("PDFから画像を抽出しました(%d件): %s -> %s", len(outputs), p, out_dir)
        return outputs

    def extract_text(self, input_path: str, output_path: str) -> str:
        """文字ベースのPDFからテキストを抽出する(スキャンPDF等には効かない。
        その場合は ocr_pdf_to_text を使う)。
        """
        p = Path(input_path)
        if not p.exists():
            raise FileNotFoundError(f"PDFが見つかりません: {p}")

        texts = []
        with pdfplumber.open(p) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                texts.append(f"--- page {i} ---\n{text}")

        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text("\n\n".join(texts), encoding="utf-8")
        logger.info("PDFテキストを抽出しました: %s -> %s", p, out)
        return str(out)

    def merge_pdfs(self, input_paths: list[str], output_path: str) -> str:
        writer = PdfWriter()
        for path in input_paths:
            p = Path(path)
            if not p.exists():
                raise FileNotFoundError(f"PDFが見つかりません: {p}")
            writer.append(str(p))

        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        with open(out, "wb") as f:
            writer.write(f)
        logger.info("PDFを結合しました: %s -> %s", input_paths, out)
        return str(out)

    def split_pdf(
        self,
        input_path: str,
        output_dir: str,
        pages_per_file: int = 1,
        start_page: int | None = None,
        end_page: int | None = None,
        filename_pattern: str | None = None,
    ) -> list[str]:
        """指定ページ範囲(start_page〜end_page。省略時は全ページ)を、
        pages_per_fileページごとに分割して保存する。pages_per_file=1なら
        「すべてのページを1ページずつ個別ファイルに抜き出す」動作になる。

        filename_pattern で出力ファイル名のルールを指定できる(拡張子.pdfは
        自動付与、省略時は "{stem}_part{part}")。使えるプレースホルダ:
          {stem}      元のファイル名(拡張子なし)
          {page}      そのファイルの先頭ページ番号
          {page_end}  そのファイルの末尾ページ番号
          {part}      通し番号(1始まり)
        例: "{stem}_p{page:03d}" のようにPythonの書式指定も使える。
        """
        p = Path(input_path)
        if not p.exists():
            raise FileNotFoundError(f"PDFが見つかりません: {p}")
        if pages_per_file < 1:
            raise ValueError("pages_per_fileは1以上を指定してください")

        reader = PdfReader(str(p))
        total = len(reader.pages)
        start = start_page or 1
        end = end_page or total
        if start < 1 or end > total or start > end:
            raise ValueError(f"ページ範囲が不正です(全{total}ページ): start_page={start}, end_page={end}")

        out_dir = Path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        pattern = filename_pattern or "{stem}_part{part}"

        outputs: list[str] = []
        part = 0
        page_idx = start - 1  # 0始まりに変換
        while page_idx < end:
            part += 1
            writer = PdfWriter()
            chunk_start_page = page_idx + 1
            chunk_end_page = min(page_idx + pages_per_file, end)
            for i in range(page_idx, chunk_end_page):
                writer.add_page(reader.pages[i])

            filename = pattern.format(
                stem=p.stem, page=chunk_start_page, page_end=chunk_end_page, part=part
            )
            if not filename.lower().endswith(".pdf"):
                filename += ".pdf"
            out_path = out_dir / filename
            with open(out_path, "wb") as f:
                writer.write(f)
            outputs.append(str(out_path))
            page_idx = chunk_end_page

        logger.info("PDFを%d件に分割しました: %s -> %s", len(outputs), p, out_dir)
        return outputs

    def extract_page_range(self, input_path: str, output_path: str, start_page: int, end_page: int) -> str:
        """指定したページ範囲(start_page〜end_page、両端含む)を1つのPDFとして抜き出す。
        「複数ファイルに分割する」split_pdfとは異なり、常に1ファイルにまとめる。
        """
        p = Path(input_path)
        if not p.exists():
            raise FileNotFoundError(f"PDFが見つかりません: {p}")
        reader = PdfReader(str(p))
        total = len(reader.pages)
        if start_page < 1 or end_page > total or start_page > end_page:
            raise ValueError(f"ページ範囲が不正です(全{total}ページ): {start_page}〜{end_page}")

        writer = PdfWriter()
        for i in range(start_page - 1, end_page):
            writer.add_page(reader.pages[i])

        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        with open(out, "wb") as f:
            writer.write(f)
        logger.info("ページ範囲を抜き出しました(%d〜%d): %s -> %s", start_page, end_page, p, out)
        return str(out)

    def rotate_pdf(
        self,
        input_path: str,
        output_path: str,
        degrees: int = 90,
        pages: list[int] | None = None,
    ) -> str:
        """PDFのページを回転する。degreesは90単位を推奨(90/180/270、負数も可)。
        pagesは1始まりのページ番号のリスト。Noneなら全ページを回転する。
        """
        p = Path(input_path)
        if not p.exists():
            raise FileNotFoundError(f"PDFが見つかりません: {p}")

        reader = PdfReader(str(p))
        writer = PdfWriter()
        target_pages = set(pages) if pages else None

        for i, page in enumerate(reader.pages, start=1):
            if target_pages is None or i in target_pages:
                page.rotate(degrees)
            writer.add_page(page)

        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        with open(out, "wb") as f:
            writer.write(f)
        logger.info("PDFを回転しました(%d度): %s -> %s", degrees, p, out)
        return str(out)

    def ocr_pdf_to_text(self, input_path: str, output_path: str, language: str = "jpn+eng") -> str:
        """スキャン画像ベースのPDF(文字が選択できないPDF)からOCRでテキストを
        抽出する。文字ベースのPDFには extract_text の方が高速・高精度。

        追加の依存関係が必要:
        - pip install pytesseract pdf2image
        - OS側に Tesseract OCR 本体と Poppler のインストールが必要
          (Windows: 別途インストーラでの導入が必要。pipだけでは入らない)
        """
        try:
            import pytesseract
            from pdf2image import convert_from_path
        except ImportError as e:
            raise RuntimeError(
                "OCRには pytesseract と pdf2image が必要です"
                "(pip install pytesseract pdf2image)。"
                "さらにOS側に Tesseract OCR と Poppler のインストールが必要です。"
            ) from e

        p = Path(input_path)
        if not p.exists():
            raise FileNotFoundError(f"PDFが見つかりません: {p}")

        try:
            images = convert_from_path(str(p))
        except Exception as e:  # noqa: BLE001
            raise RuntimeError(
                f"PDFを画像に変換できませんでした(Popplerが未インストールの可能性があります): {e}"
            ) from e

        texts = []
        for i, img in enumerate(images, start=1):
            text = pytesseract.image_to_string(img, lang=language)
            texts.append(f"--- page {i} (OCR) ---\n{text}")

        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text("\n\n".join(texts), encoding="utf-8")
        logger.info("OCRでテキストを抽出しました: %s -> %s", p, out)
        return str(out)

    @staticmethod
    def _validate_area(x_left: float, x_right: float, y_upper: float, y_lower: float) -> None:
        for name, val in (("x_left", x_left), ("x_right", x_right), ("y_upper", y_upper), ("y_lower", y_lower)):
            if not (0 <= val <= 100):
                raise ValueError(f"{name}は0〜100の範囲(ページに対する%)で指定してください: {val}")
        if x_left >= x_right:
            raise ValueError(f"x_left({x_left})はx_right({x_right})より小さくしてください")
        if y_upper >= y_lower:
            raise ValueError(f"y_upper({y_upper})はy_lower({y_lower})より小さくしてください")

    def extract_text_in_area(
        self,
        input_path: str,
        x_left: float,
        x_right: float,
        y_upper: float,
        y_lower: float,
        page_number: int = 1,
        ocr: bool = False,
        ocr_language: str = "jpn+eng",
    ) -> str:
        """PDFページの指定範囲内にある文字を取得する。

        座標系: x_left/x_right/y_upper/y_lower はいずれも「ページの幅・高さに
        対する割合(0〜100のパーセント)」で指定する相対値。基準点はページの
        **左上を(0,0)**とし、xは右方向、yは下方向に向かって増える
        (pdfplumberの座標系にそのまま準拠)。
        例: 右下1/4の範囲を指定したい場合は
            x_left=50, x_right=100, y_upper=50, y_lower=100

        ocr=False(既定)の場合は文字ベースのPDFからそのままテキストを抽出する
        (pdfplumberのcrop機能を使用)。ocr=Trueの場合は該当範囲を画像として
        切り出しOCRにかける(スキャンPDF等、文字を選択できないPDF向け。
        要pytesseract/pdf2image + Tesseract OCR/Poppler)。

        改行は文字列 "ж" に置換し、1行の長い文字列として返す
        (改行を含んだままだと後続のテキスト加工が扱いにくいための仕様)。
        """
        p = Path(input_path)
        if not p.exists():
            raise FileNotFoundError(f"PDFが見つかりません: {p}")
        self._validate_area(x_left, x_right, y_upper, y_lower)

        if ocr:
            try:
                import pytesseract
                from pdf2image import convert_from_path
            except ImportError as e:
                raise RuntimeError(
                    "OCRには pytesseract と pdf2image が必要です"
                    "(pip install pytesseract pdf2image)。"
                ) from e
            try:
                images = convert_from_path(str(p), first_page=page_number, last_page=page_number)
            except Exception as e:  # noqa: BLE001
                raise RuntimeError(
                    f"PDFを画像に変換できませんでした(Popplerが未インストールの可能性があります): {e}"
                ) from e
            if not images:
                raise ValueError(f"ページ{page_number}が見つかりませんでした")
            img = images[0]
            w, h = img.size
            box = (
                max(0, min(w, w * x_left / 100)),
                max(0, min(h, h * y_upper / 100)),
                max(0, min(w, w * x_right / 100)),
                max(0, min(h, h * y_lower / 100)),
            )
            cropped = img.crop(box)
            text = pytesseract.image_to_string(cropped, lang=ocr_language)
        else:
            with pdfplumber.open(p) as pdf:
                if page_number < 1 or page_number > len(pdf.pages):
                    raise ValueError(f"ページ{page_number}が見つかりません(全{len(pdf.pages)}ページ)")
                page = pdf.pages[page_number - 1]
                # 100%指定時に浮動小数点の丸め誤差でページ境界をわずかに超えることが
                # あるため、ページの範囲内にクランプする
                box = (
                    max(0.0, min(page.width, page.width * x_left / 100)),
                    max(0.0, min(page.height, page.height * y_upper / 100)),
                    max(0.0, min(page.width, page.width * x_right / 100)),
                    max(0.0, min(page.height, page.height * y_lower / 100)),
                )
                cropped_page = page.crop(box)
                text = cropped_page.extract_text() or ""

        single_line = text.replace("\r\n", "ж").replace("\n", "ж").replace("\r", "ж")
        logger.info(
            "範囲指定でテキストを抽出しました(ページ%d, x:%s-%s%%, y:%s-%s%%, ocr=%s): %s",
            page_number, x_left, x_right, y_upper, y_lower, ocr, p,
        )
        return single_line
