"""
DesktopHandler: 画面上の画像を探してマウス移動・クリックする、デスクトップ全体の
自動操作を扱うハンドラ(pyautoguiを使用)。ブラウザに限らず、任意のアプリケーション
のウィンドウを対象にできる。

必要な追加依存:
- pip install pyautogui
- あいまい一致(confidence指定)を使うには pip install opencv-python も必要
  (無い場合は confidence を指定しても無視され、完全一致でのみ動作する)
- ウィンドウのサイズ・位置指定・アクティブ化(*_by_title系)には
  pip install pygetwindow も必要(Excel/PDFビューア/エクスプローラー等、
  どのアプリケーションのウィンドウでもタイトルを手がかりに操作できる。
  ブラウザだけはSeleniumの機能(browser.set_window_size等)で同じことが
  より確実にできるため、そちらを優先すること)
- Linux: python3-xlib が必要になる場合がある(pyautogui内部で使用)
- macOS: 「システム設定 > プライバシーとセキュリティ > アクセシビリティ / 画面収録」で
  実行しているターミナル/Pythonにアクセスを許可する必要がある

安全上の注意:
- 画面上のどこでもクリック・入力できてしまうため、ブラウザのサイトホワイトリストの
  ような「実行してよい対象の制限」は原理的に効かせられない。対象画像は誤検出しにくい、
  十分に特徴的なものを用意すること。
- pyautogui標準のFAILSAFE機能を有効にしている。マウスカーソルを画面の左上隅
  (座標(0,0))へ動かすと例外が発生して緊急停止できる。
- 座標を直接指定するクリック(click_at)は画面解像度・ウィンドウ配置に強く依存する。
  実行環境が変わると位置がずれる可能性があるため、できるだけ画像検索(locate_and_click)
  を優先すること。
"""
from __future__ import annotations

import ctypes
import logging
import time
from pathlib import Path

logger = logging.getLogger("rpa_local_ai.desktop")


class ImageNotFoundError(Exception):
    pass


class WindowNotFoundError(Exception):
    pass


def _import_pyautogui():
    try:
        import pyautogui
    except ImportError as e:
        raise RuntimeError(
            "この操作には pyautogui が必要です(pip install pyautogui)。"
            "あいまい一致(confidence指定)を使う場合はさらに"
            "opencv-python も必要です(pip install opencv-python)。"
        ) from e
    return pyautogui


class DesktopHandler:
    def __init__(self):
        self._pyautogui = None

    def _gui(self):
        if self._pyautogui is None:
            gui = _import_pyautogui()
            gui.FAILSAFE = True  # マウスを画面左上隅へ動かすと緊急停止できる
            self._pyautogui = gui
        return self._pyautogui

    @staticmethod
    def _normalize_region(region: tuple[int, int, int, int] | list[int] | None) -> tuple[int, int, int, int] | None:
        """region指定を (left, top, width, height) のタプルに正規化する。
        macros.json経由(JSON)ではタプルはリストとして保存・復元されるため、
        リストで渡ってきても受け付ける。
        """
        if region is None:
            return None
        region = tuple(int(v) for v in region)
        if len(region) != 4:
            raise ValueError(f"regionは (left, top, width, height) の4つの数値で指定してください: {region}")
        return region

    def get_screen_size(self) -> dict:
        """OSの画面解像度(モニタ全体のピクセル数)を取得する。
        region(検索する矩形領域)の座標を計算する際の基準として使う。
        """
        gui = self._gui()
        width, height = gui.size()
        return {"width": int(width), "height": int(height)}

    # ---------- ウィンドウのサイズ・位置・アクティブ化(タイトル指定) ----------
    # ブラウザに限らず、Excel/PDFビューア/エクスプローラー等どのアプリケーションの
    # ウィンドウでも、タイトルバーの文字列を手がかりに操作できる(pygetwindow使用)。
    # pyautoguiは常に「今アクティブな(最前面の)ウィンドウ」に対してクリック・
    # キー入力するため、複数のアプリを行き来しながら操作する場合は、対象操作の
    # 前に activate_window_by_title で切り替えておく必要がある。

    def _pygetwindow_window(self, title_hint: str):
        gui = self._gui()
        try:
            windows = gui.getWindowsWithTitle(title_hint)
        except (AttributeError, NotImplementedError) as e:
            raise RuntimeError(
                "ウィンドウの位置・サイズ操作には pygetwindow が必要です"
                "(pip install pygetwindow)。"
            ) from e
        if not windows:
            raise WindowNotFoundError(
                f"'{title_hint}' というタイトルを含むウィンドウが見つかりませんでした"
                f"(list_window_titlesで今開いているタイトル一覧を確認できます)"
            )
        return windows[0]

    def list_window_titles(self) -> list[str]:
        """今開いているウィンドウのタイトル一覧を取得する。
        *_by_title系の操作でどの文字列を目印にすればよいか調べる時に使う。
        """
        gui = self._gui()
        try:
            titles = gui.getAllTitles()
        except (AttributeError, NotImplementedError) as e:
            raise RuntimeError(
                "この操作には pygetwindow が必要です(pip install pygetwindow)。"
            ) from e
        return [t for t in titles if t.strip()]

    def activate_window_by_title(self, title_hint: str) -> str:
        """タイトルにtitle_hintを含むウィンドウを最前面(アクティブ)にする。
        Excel/PDFビューア/エクスプローラー等を切り替えながらpyautoguiで操作したい
        場合、対象を切り替えるたびにこれを呼んでおくこと。
        """
        win = self._pygetwindow_window(title_hint)
        self._force_activate(win._hWnd)
        logger.info("ウィンドウをアクティブにしました: %s", title_hint)
        return f"activated: {title_hint}"

    @staticmethod
    def _force_activate(hwnd: int) -> None:
        """SetForegroundWindowは、直前にユーザー入力を受け取っていないプロセスからの
        呼び出しをWindows側の仕様で拒否することがある(GetLastErrorが0=成功のまま
        呼び出し自体は失敗する、という紛らわしい既知の挙動)。ダミーのAltキー押下を
        挟むことで、フォアグラウンド変更が許可されるようにする
        (広く知られたWindows APIの回避策)。
        """
        user32 = ctypes.windll.user32
        VK_MENU = 0x12
        KEYEVENTF_KEYUP = 0x0002
        SW_RESTORE = 9
        user32.keybd_event(VK_MENU, 0, 0, 0)
        try:
            user32.ShowWindow(hwnd, SW_RESTORE)
            if not user32.SetForegroundWindow(hwnd):
                raise WindowNotFoundError(
                    "ウィンドウをアクティブにできませんでした"
                    "(Windows側の制限により拒否された可能性があります。リトライ設定を"
                    "有効にしていれば自動で再試行されます)"
                )
        finally:
            user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)

    def set_window_size_by_title(
        self,
        title_hint: str,
        width: int | None = None,
        height: int | None = None,
        width_percent: float | None = None,
        height_percent: float | None = None,
    ) -> str:
        """タイトルにtitle_hintを含むウィンドウ(Excel/PDFビューア/エクスプローラー等、
        どのアプリケーションでもよい)のサイズを指定する。width/height はピクセル指定。
        width_percent/height_percent は画面全体(モニタの解像度)に対する
        割合(0〜100%)で指定する相対値で、指定するとピクセル指定より優先される。
        どちらも指定しなかった軸は変更しない。
        """
        win = self._pygetwindow_window(title_hint)
        if width_percent is not None or height_percent is not None:
            screen = self.get_screen_size()
            if width_percent is not None:
                width = round(screen["width"] * width_percent / 100)
            if height_percent is not None:
                height = round(screen["height"] * height_percent / 100)
        final_width = int(width) if width is not None else win.width
        final_height = int(height) if height is not None else win.height
        win.resizeTo(final_width, final_height)
        logger.info("ウィンドウサイズを設定しました: %s -> %dx%d", title_hint, final_width, final_height)
        return f"{final_width}x{final_height}"

    def set_window_position_by_title(self, title_hint: str, x: int, y: int) -> str:
        """タイトルにtitle_hintを含むウィンドウの位置を指定する。基準点は画面の
        左上(0,0)で、右方向がx増加、下方向がy増加。
        """
        win = self._pygetwindow_window(title_hint)
        win.moveTo(int(x), int(y))
        logger.info("ウィンドウ位置を設定しました: %s -> (%d, %d)", title_hint, x, y)
        return f"({x}, {y})"

    def _locate(
        self,
        image_path: str,
        confidence: float,
        timeout: float,
        region: tuple[int, int, int, int] | list[int] | None = None,
    ):
        gui = self._gui()
        p = Path(image_path)
        if not p.exists():
            raise FileNotFoundError(f"画像ファイルが見つかりません: {p}")
        region = self._normalize_region(region)

        deadline = time.monotonic() + timeout
        last_err: Exception | None = None
        while True:
            try:
                box = gui.locateOnScreen(str(p), confidence=confidence, region=region)
            except TypeError:
                # opencv-python未インストール時、confidence引数は受け付けられない
                box = gui.locateOnScreen(str(p), region=region)
            except Exception as e:  # noqa: BLE001
                last_err = e
                box = None
            if box is not None:
                return box
            if time.monotonic() >= deadline:
                break
            time.sleep(0.5)

        area = f"領域{region}内" if region else "画面全体"
        msg = f"{area}に画像が見つかりませんでした: {p}"
        if last_err:
            msg += f"(内部エラー: {last_err})"
        raise ImageNotFoundError(msg)

    def move_to_image(
        self,
        image_path: str,
        confidence: float = 0.8,
        timeout: float = 10,
        region: tuple[int, int, int, int] | list[int] | None = None,
    ) -> str:
        """画面上(regionを指定した場合はその矩形領域内だけ)から画像を探し、
        見つかった位置の中心へマウスを移動する。regionは (left, top, width, height)
        をピクセルで指定する(省略時は画面全体を検索する)。
        """
        gui = self._gui()
        box = self._locate(image_path, confidence, timeout, region=region)
        center = gui.center(box)
        gui.moveTo(center.x, center.y, duration=0.2)
        logger.info("画像へマウスを移動しました: %s -> (%d, %d)", image_path, center.x, center.y)
        return f"moved to: ({center.x}, {center.y})"

    def locate_and_click(
        self,
        image_path: str,
        confidence: float = 0.8,
        timeout: float = 10,
        button: str = "left",
        clicks: int = 1,
        region: tuple[int, int, int, int] | list[int] | None = None,
    ) -> str:
        """画面上(regionを指定した場合はその矩形領域内だけ)から image_path の
        画像を探し、見つかった位置の中心をクリックする。regionは
        (left, top, width, height) をピクセルで指定する(省略時は画面全体を
        検索する。似た画像が画面の他の場所にもある場合の誤検出を防ぎたい場合や、
        検索範囲を絞って高速化したい場合に使う)。
        """
        gui = self._gui()
        box = self._locate(image_path, confidence, timeout, region=region)
        center = gui.center(box)
        gui.moveTo(center.x, center.y, duration=0.2)
        gui.click(center.x, center.y, button=button, clicks=clicks)
        logger.info("画像を見つけてクリックしました: %s -> (%d, %d)", image_path, center.x, center.y)
        return f"clicked at: ({center.x}, {center.y})"

    def click_at(self, x: int, y: int, button: str = "left", clicks: int = 1) -> str:
        """座標を直接指定してクリックする(画像検索がうまくいかない場合の代替手段)。
        画面解像度・ウィンドウ配置に依存するため、可能な限りlocate_and_clickを優先すること。
        """
        gui = self._gui()
        gui.moveTo(x, y, duration=0.2)
        gui.click(x, y, button=button, clicks=clicks)
        logger.info("座標を直接クリックしました: (%d, %d)", x, y)
        return f"clicked at: ({x}, {y})"

    def type_text(self, text: str, interval: float = 0.02) -> str:
        """今フォーカスされている場所に文字列を入力する(OSのキーボード入力として送る)。"""
        gui = self._gui()
        gui.write(text, interval=interval)
        return f"typed: {text}"

    def press_key(self, key: str) -> str:
        """特殊キーを送信する。例: 'enter', 'tab', 'esc'。
        'ctrl+s' のように '+' 区切りで組み合わせキーも指定できる。
        """
        gui = self._gui()
        keys = [k.strip() for k in key.split("+") if k.strip()]
        if not keys:
            raise ValueError("keyが空です")
        if len(keys) > 1:
            gui.hotkey(*keys)
        else:
            gui.press(keys[0])
        return f"pressed: {key}"

    def take_screenshot(
        self, save_path: str, region: tuple[int, int, int, int] | list[int] | None = None
    ) -> str:
        """デスクトップのスクリーンショットを保存する。regionを指定するとその
        矩形領域だけを切り出して保存する(省略時は画面全体)。
        ここから対象ボタン/アイコンをトリミングし、locate_and_click等で使う
        画像素材を作るための下準備として使う。
        """
        gui = self._gui()
        region = self._normalize_region(region)
        p = Path(save_path)
        p.parent.mkdir(parents=True, exist_ok=True)
        img = gui.screenshot(region=region) if region else gui.screenshot()
        img.save(str(p))
        logger.info("スクリーンショットを保存しました: %s (region=%s)", p, region)
        return str(p)

    def wait_for_image(
        self,
        image_path: str,
        confidence: float = 0.8,
        timeout: float = 15,
        region: tuple[int, int, int, int] | list[int] | None = None,
    ) -> str:
        """画面上(regionを指定した場合はその矩形領域内だけ)に画像が表示される
        まで待つ。regionは (left, top, width, height) をピクセルで指定する。
        """
        self._locate(image_path, confidence, timeout, region=region)
        return f"appeared: {image_path}"

    # ---------- 操作後の成功確認(Executorのverify機構と連携) ----------
    # value は "画像パス" または "画像パス|confidence" の形式を受け付ける。

    @staticmethod
    def _parse_verify_value(value: str) -> tuple[str, float]:
        if "|" in value:
            path, conf_raw = value.rsplit("|", 1)
            try:
                return path, float(conf_raw)
            except ValueError:
                return value, 0.8
        return value, 0.8

    def verify_image_appears(self, value: str, timeout: int = 10) -> str:
        image_path, confidence = self._parse_verify_value(value)
        self._locate(image_path, confidence, timeout)
        return f"verified: 画像の表示を確認しました({image_path})"

    def verify_image_disappears(self, value: str, timeout: int = 10) -> str:
        image_path, confidence = self._parse_verify_value(value)
        gui = self._gui()
        deadline = time.monotonic() + timeout
        while True:
            try:
                box = gui.locateOnScreen(image_path, confidence=confidence)
            except TypeError:
                box = gui.locateOnScreen(image_path)
            except Exception:  # noqa: BLE001
                box = None
            if box is None:
                return f"verified: 画像が消えたことを確認しました({image_path})"
            if time.monotonic() >= deadline:
                break
            time.sleep(0.5)
        raise ImageNotFoundError(
            f"操作後に画像が消えるはずでしたが、{timeout}秒経っても表示されたままでした: {image_path}"
        )
