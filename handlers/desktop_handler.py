"""
DesktopHandler: 画面上の画像を探してマウス移動・クリックする、デスクトップ全体の
自動操作を扱うハンドラ(pyautoguiを使用)。ブラウザに限らず、任意のアプリケーション
のウィンドウを対象にできる。

必要な追加依存:
- pip install pyautogui
- あいまい一致(confidence指定)を使うには pip install opencv-python も必要
  (無い場合は confidence を指定しても無視され、完全一致でのみ動作する)
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

import logging
import time
from pathlib import Path

logger = logging.getLogger("rpa_local_ai.desktop")


class ImageNotFoundError(Exception):
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

    def _locate(self, image_path: str, confidence: float, timeout: float):
        gui = self._gui()
        p = Path(image_path)
        if not p.exists():
            raise FileNotFoundError(f"画像ファイルが見つかりません: {p}")

        deadline = time.monotonic() + timeout
        last_err: Exception | None = None
        while True:
            try:
                box = gui.locateOnScreen(str(p), confidence=confidence)
            except TypeError:
                # opencv-python未インストール時、confidence引数は受け付けられない
                box = gui.locateOnScreen(str(p))
            except Exception as e:  # noqa: BLE001
                last_err = e
                box = None
            if box is not None:
                return box
            if time.monotonic() >= deadline:
                break
            time.sleep(0.5)

        msg = f"画面上に画像が見つかりませんでした: {p}"
        if last_err:
            msg += f"(内部エラー: {last_err})"
        raise ImageNotFoundError(msg)

    def move_to_image(self, image_path: str, confidence: float = 0.8, timeout: float = 10) -> str:
        gui = self._gui()
        box = self._locate(image_path, confidence, timeout)
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
    ) -> str:
        """画面上から image_path の画像を探し、見つかった位置の中心をクリックする。"""
        gui = self._gui()
        box = self._locate(image_path, confidence, timeout)
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

    def take_screenshot(self, save_path: str) -> str:
        """デスクトップ全体のスクリーンショットを保存する。
        ここから対象ボタン/アイコンをトリミングし、locate_and_click等で使う
        画像素材を作るための下準備として使う。
        """
        gui = self._gui()
        p = Path(save_path)
        p.parent.mkdir(parents=True, exist_ok=True)
        img = gui.screenshot()
        img.save(str(p))
        logger.info("デスクトップのスクリーンショットを保存しました: %s", p)
        return str(p)

    def wait_for_image(self, image_path: str, confidence: float = 0.8, timeout: float = 15) -> str:
        self._locate(image_path, confidence, timeout)
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
