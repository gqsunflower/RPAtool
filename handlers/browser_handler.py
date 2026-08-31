"""
BrowserHandler: Selenium を使ったブラウザ操作ハンドラ。
既定はChromeだが、コンストラクタの browser="edge" でMicrosoft Edge
(Chromiumベース)にも対応する。両者ともselenium 1パッケージで動くため
追加のpipインストールは不要(Edge本体のインストールは別途必要)。

安全上の制約(重要):
- config/whitelist_urls.json に登録された site_key の URL にしか
  遷移しません。goto的な「任意URLを開く」actionはそもそも存在しません。
- 検索エンジンへのクエリ送信や、リンクを辿っての自由なクロールは行いません。
- 各アクション実行のたびに「今のドメインが開いたサイトのドメインと一致しているか」
  を再検証します(意図しないリダイレクト対策)。

要素の見つけ方(重要な設計方針):
- ボタン等はCSSセレクタやclass名でガチガチに指定しません(UI変更で壊れやすいため)。
- 代わりに「画面に表示されているであろう文字」を手がかりに、
  button/a/input[type=submit,button]/role=button/aria-label/title を横断的に
  検索する「緩いテキスト一致」(click_by_text等)を基本にしています。
  完全一致を優先し、無ければ部分一致にフォールバックします。
- 表示テキストでの検索がうまく拾えない場合に備え、F12の開発者ツールで
  調べたCSSセレクタ(class/id等)を直接指定できる click_selector /
  type_by_selector / select_by_selector も用意しています
  (UI変更に弱くなるため、まずは *_by_text 系を試すことを推奨)。
- 入力操作(type_by_text / type_by_selector)は press_enter=True で
  入力後にEnterキー(送信)を送れます。
- 画面をPDFとして保存する save_page_as_pdf も用意しています
  (Chromeの印刷機能をDevTools Protocol経由で使用。倍率(scale)指定可能)。
"""
from __future__ import annotations

import json
import logging
import os
import re
import time
from pathlib import Path
from urllib.parse import urlparse

logger = logging.getLogger("rpa_local_ai.browser")


class SiteNotWhitelistedError(Exception):
    pass


class NavigationBlockedError(Exception):
    pass


class ElementNotFoundError(Exception):
    pass


class VerificationFailedError(Exception):
    """操作後の確認(期待した文字が表示された/URLが変わった等)が取れなかった場合に送出する。
    これが出た場合、Executor側はマクロの残りの手順を実行せず即座に停止する。
    """
    pass


def _xpath_literal(text: str) -> str:
    """XPathの文字列リテラルとして安全な形にする(シングル/ダブルクォート混在対応)。"""
    if "'" not in text:
        return f"'{text}'"
    if '"' not in text:
        return f'"{text}"'
    parts = text.split("'")
    return "concat(" + ", \"'\", ".join(f"'{p}'" for p in parts) + ")"


_SUPPORTED_BROWSERS = ("chrome", "edge")

# ドライバ起動失敗時のエラーメッセージから「対応バージョン」と「実際のブラウザバージョン」を
# 拾うための正規表現。Chrome/Edge両方のSessionNotCreatedExceptionメッセージが
# 概ねこの形式("This version of XxxDriver only supports Yyy version N" +
# "Current browser version is M...")のため、共通で使える。
_VERSION_MISMATCH_RE = re.compile(
    r"only supports .*? version (\d+).*?Current browser version is ([\d.]+)",
    re.IGNORECASE | re.DOTALL,
)


def _build_launch_error(browser_label: str, driver_name: str, env_var: str, exc: Exception) -> RuntimeError:
    """ブラウザ起動失敗時に、原因ごとに具体的な対処方法を含んだエラーにして返す。"""
    from selenium.common.exceptions import SessionNotCreatedException

    msg = str(exc)
    if isinstance(exc, SessionNotCreatedException):
        match = _VERSION_MISMATCH_RE.search(msg)
        if match:
            driver_supports, browser_version = match.group(1), match.group(2)
            return RuntimeError(
                f"{browser_label}のバージョンと{driver_name}のバージョンが合っていません"
                f"({driver_name}が対応しているのは{browser_label} {driver_supports}系ですが、"
                f"実際にインストールされている{browser_label}は{browser_version}です)。"
                f"{browser_label}のバージョンに合った{driver_name}を再ダウンロードし、"
                f"環境変数 {env_var} を新しいファイルのパスに更新してください"
                f"({browser_label}は自動更新されることが多いため、{driver_name}を手動で指定している場合は"
                f"ブラウザが更新されるたびにこのズレが起きやすくなります): {msg}"
            )
        return RuntimeError(
            f"{browser_label}のバージョンと{driver_name}のバージョンが合っていない可能性があります。"
            f"{browser_label}のバージョンに合った{driver_name}を再ダウンロードし、"
            f"環境変数 {env_var} を新しいファイルのパスに更新してください: {msg}"
        )
    return RuntimeError(
        f"{browser_label}の起動に失敗しました。Seleniumが{driver_name}を自動取得できなかった"
        f"可能性があります(社内ネットワークが配布サーバーをブロックしている場合によく起きます)。"
        f"{driver_name}を手動でダウンロードし、環境変数 {env_var} にそのファイルのフルパスを"
        f"設定してから再実行してください"
        f"(詳細はREADMEの「ブラウザのドライバを手動で用意する」を参照): {msg}"
    )


class BrowserHandler:
    def __init__(self, whitelist_path: Path, headless: bool = True, browser: str = "chrome"):
        """browser: "chrome"(既定)または"edge"。EdgeはChromiumベースのため、
        表示テキストでのクリック・CSSセレクタ・JavaScript実行・画面のPDF保存
        (DevTools Protocol)等、ここで使っている機能はすべて同様に動作する。
        追加のpipパッケージは不要(seleniumが両対応)だが、Edge本体のインストール
        は別途必要(ドライバ(msedgedriver)はSelenium Managerが自動取得する)。
        """
        browser = browser.lower()
        if browser not in _SUPPORTED_BROWSERS:
            raise ValueError(
                f"未対応のbrowserです: {browser}(使えるもの: {', '.join(_SUPPORTED_BROWSERS)})"
            )
        self.whitelist_path = Path(whitelist_path)
        self.headless = headless
        self.browser = browser
        self._driver = None
        self._sites = self._load_whitelist()
        self._current_site_key: str | None = None
        # 複数タブ対応: エイリアス名でタブ(ウィンドウハンドル)を管理する。
        # tab_aliasを指定しない既存のマクロは既定の"main"タブとして扱われ、
        # 従来どおり単一タブのまま動作する(後方互換)。
        self._tab_handles: dict[str, str] = {}       # alias -> window handle
        self._tab_site_urls: dict[str, str] = {}      # alias -> 現在そのタブで開いているサイトのURL
        self._current_tab_alias: str | None = None

    # ---------- ホワイトリスト管理 ----------

    def _load_whitelist(self) -> dict:
        with open(self.whitelist_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("sites", {})

    def _save_whitelist(self) -> None:
        with open(self.whitelist_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        data["sites"] = self._sites
        with open(self.whitelist_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def register_site(self, site_key: str, url: str) -> str:
        """新しいサイトをホワイトリストに登録する(レコーダーから使用)。"""
        if site_key in self._sites:
            raise ValueError(f"site_key '{site_key}' は既に登録されています")
        self._sites[site_key] = {"url": url}
        self._save_whitelist()
        logger.info("サイトを登録しました: %s -> %s", site_key, url)
        return f"registered: {site_key}"

    # ---------- 内部ヘルパー ----------

    def _get_driver(self):
        if self._driver is None:
            from selenium import webdriver

            if self.browser == "edge":
                from selenium.webdriver.edge.options import Options as EdgeOptions
                from selenium.webdriver.edge.service import Service as EdgeService

                options = EdgeOptions()
                if self.headless:
                    options.add_argument("--headless=new")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-dev-shm-usage")
                driver_path = os.environ.get("RPA_EDGE_DRIVER_PATH")
                service = EdgeService(executable_path=driver_path) if driver_path else None
                try:
                    self._driver = webdriver.Edge(options=options, service=service)
                except Exception as e:  # noqa: BLE001
                    raise _build_launch_error("Edge", "msedgedriver", "RPA_EDGE_DRIVER_PATH", e) from e
            else:
                from selenium.webdriver.chrome.options import Options as ChromeOptions
                from selenium.webdriver.chrome.service import Service as ChromeService

                options = ChromeOptions()
                if self.headless:
                    options.add_argument("--headless=new")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-dev-shm-usage")
                driver_path = os.environ.get("RPA_CHROME_DRIVER_PATH")
                service = ChromeService(executable_path=driver_path) if driver_path else None
                try:
                    self._driver = webdriver.Chrome(options=options, service=service)
                except Exception as e:  # noqa: BLE001
                    raise _build_launch_error("Chrome", "chromedriver", "RPA_CHROME_DRIVER_PATH", e) from e
        return self._driver

    def _assert_domain_allowed(self, url: str, expected_url: str) -> None:
        got_domain = urlparse(url).netloc
        expected_domain = urlparse(expected_url).netloc
        if got_domain != expected_domain:
            raise NavigationBlockedError(
                f"ホワイトリスト外のドメインへの遷移を検知したため停止しました: {got_domain}"
            )

    def _assert_still_on_site(self) -> None:
        """クリックや入力の結果、想定外のドメインに飛んでいないかを毎回確認する。
        現在アクティブなタブに登録されているサイトのドメインと比較する。
        """
        if self._current_tab_alias is None:
            return
        expected_url = self._tab_site_urls.get(self._current_tab_alias)
        if expected_url is None:
            return
        driver = self._get_driver()
        self._assert_domain_allowed(driver.current_url, expected_url)

    # ---------- サイトを開く / タブ切り替え / ログイン ----------

    def open_registered_site(self, site_key: str, tab_alias: str = "main") -> str:
        """登録済みサイトを開く。tab_alias省略時は既定タブ("main")で開く
        (=既存の単一タブ運用と同じ挙動)。既に他のタブが開いている状態で
        未使用のtab_aliasを指定すると、新しいタブを開いてそこにサイトを表示する。
        既に存在するtab_aliasを指定した場合は、そのタブへ切り替えたうえで
        指定サイトへ遷移する(同じタブで別サイトに開き直す用途にも使える)。
        """
        if site_key not in self._sites:
            raise SiteNotWhitelistedError(
                f"'{site_key}' はホワイトリストに登録されていません。"
                f"config/whitelist_urls.json に追加してください。"
            )
        site = self._sites[site_key]
        driver = self._get_driver()

        if tab_alias in self._tab_handles:
            driver.switch_to.window(self._tab_handles[tab_alias])
        elif self._tab_handles:
            # 既に他のタブが開いている状態で、新しいタブ名を指定された場合は新規タブを開く
            driver.switch_to.new_window("tab")
            self._tab_handles[tab_alias] = driver.current_window_handle
        else:
            # これがまだ最初のタブ(ドライバ起動直後の既定ウィンドウをそのまま使う)
            self._tab_handles[tab_alias] = driver.current_window_handle

        driver.get(site["url"])
        self._assert_domain_allowed(driver.current_url, site["url"])
        self._tab_site_urls[tab_alias] = site["url"]
        self._current_tab_alias = tab_alias
        self._current_site_key = site_key
        logger.info("登録サイトを開きました: %s (タブ=%s, %s)", site_key, tab_alias, driver.current_url)
        return driver.current_url

    # ---------- ウィンドウサイズ・位置・表示倍率(pyautogui併用時の座標合わせ用) ----------
    # pyautoguiは画面上の絶対座標で操作するため、ブラウザのウィンドウサイズ・位置・
    # 表示倍率が実行のたびに変わると、記録した画像/座標がずれて動かなくなる。
    # ここでこれらを固定しておくことで、pyautoguiとの併用を安定させる。

    def get_screen_size(self) -> dict:
        """OSの画面解像度(モニタ全体のピクセル数)を取得する。
        set_window_sizeで相対値(%)を指定する際の基準にもなる。
        """
        driver = self._get_driver()
        size = driver.execute_script("return {width: window.screen.width, height: window.screen.height};")
        return {"width": int(size["width"]), "height": int(size["height"])}

    def set_window_size(
        self,
        width: int | None = None,
        height: int | None = None,
        width_percent: float | None = None,
        height_percent: float | None = None,
    ) -> str:
        """ブラウザのウィンドウサイズを指定する。width/height はピクセル指定。
        width_percent/height_percent は画面全体(モニタの解像度)に対する
        割合(0〜100%)で指定する相対値で、指定するとピクセル指定より優先される。
        どちらも指定しなかった軸は変更しない。
        """
        driver = self._get_driver()
        if width_percent is not None or height_percent is not None:
            screen = self.get_screen_size()
            if width_percent is not None:
                width = round(screen["width"] * width_percent / 100)
            if height_percent is not None:
                height = round(screen["height"] * height_percent / 100)
        current = driver.get_window_size()
        final_width = int(width) if width is not None else current["width"]
        final_height = int(height) if height is not None else current["height"]
        driver.set_window_size(final_width, final_height)
        logger.info("ウィンドウサイズを設定しました: %dx%d", final_width, final_height)
        return f"{final_width}x{final_height}"

    def set_window_position(self, x: int, y: int) -> str:
        """ブラウザのウィンドウ位置を指定する。基準点は画面の左上(0,0)で、
        右方向がx増加、下方向がy増加(Windowsの画面座標と同じ)。
        """
        driver = self._get_driver()
        driver.set_window_position(int(x), int(y))
        logger.info("ウィンドウ位置を設定しました: (%d, %d)", x, y)
        return f"({x}, {y})"

    def set_zoom(self, percent: float = 100) -> str:
        """ページの表示倍率を変更する(100=等倍)。pyautoguiの画像検索の精度を
        上げたい(要素を大きく表示したい)場合等に使う。ブラウザ本体の
        Ctrl+スクロールズームと違い、CSSの`zoom`プロパティで実現しているため、
        別ページに遷移すると設定はリセットされる(維持したい場合はページ遷移の
        たびに再度このステップを実行すること)。
        """
        driver = self._get_driver()
        driver.execute_script("document.body.style.zoom = arguments[0] + '%';", percent)
        logger.info("表示倍率を設定しました: %s%%", percent)
        return f"{percent}%"

    def switch_to_tab(self, tab_alias: str) -> str:
        """既に開いている複数のタブの中から、以降の操作対象(アクティブタブ)を切り替える。"""
        if tab_alias not in self._tab_handles:
            raise ValueError(
                f"タブ '{tab_alias}' はまだ開かれていません"
                f"(開いているタブ: {list(self._tab_handles.keys())})"
            )
        driver = self._get_driver()
        driver.switch_to.window(self._tab_handles[tab_alias])
        self._current_tab_alias = tab_alias
        logger.info("タブを切り替えました: %s", tab_alias)
        return f"switched to tab: {tab_alias}"

    def close_tab(self, tab_alias: str) -> str:
        """指定したタブを閉じる。閉じたのがアクティブタブだった場合、
        残っている別のタブがあれば自動的にそちらへ切り替える。
        """
        if tab_alias not in self._tab_handles:
            raise ValueError(f"タブ '{tab_alias}' はまだ開かれていません")
        driver = self._get_driver()
        driver.switch_to.window(self._tab_handles[tab_alias])
        driver.close()
        del self._tab_handles[tab_alias]
        self._tab_site_urls.pop(tab_alias, None)

        if self._current_tab_alias == tab_alias:
            self._current_tab_alias = None
            remaining = list(self._tab_handles.keys())
            if remaining:
                self.switch_to_tab(remaining[0])
        logger.info("タブを閉じました: %s", tab_alias)
        return f"closed tab: {tab_alias}"

    def list_open_tabs(self) -> list[str]:
        return list(self._tab_handles.keys())

    def login(self, site_key: str) -> str:
        if site_key not in self._sites:
            raise SiteNotWhitelistedError(f"'{site_key}' はホワイトリストに登録されていません。")
        site = self._sites[site_key]
        login_cfg = site.get("login")
        if not login_cfg:
            raise ValueError(f"'{site_key}' にはlogin設定がありません")

        username = os.environ.get(login_cfg["username_env"])
        password = os.environ.get(login_cfg["password_env"])
        if not username or not password:
            raise EnvironmentError(
                f"環境変数 {login_cfg['username_env']} / {login_cfg['password_env']} "
                f"に認証情報を設定してください(パスワード等をコードや設定に直書きしない)"
            )

        driver = self._get_driver()
        driver.find_element("css selector", login_cfg["username_selector"]).send_keys(username)
        driver.find_element("css selector", login_cfg["password_selector"]).send_keys(password)
        driver.find_element("css selector", login_cfg["submit_selector"]).click()

        self._assert_domain_allowed(driver.current_url, site["url"])
        logger.info("ログイン操作を実行しました: %s", site_key)
        return "login submitted"

    # ---------- 表示テキストによる「緩い」要素検索(標準RPA操作) ----------

    def _build_clickable_xpath(self, text: str) -> str:
        lit = _xpath_literal(text)
        conditions = [
            f"(self::button and contains(normalize-space(.), {lit}))",
            f"(self::a and contains(normalize-space(.), {lit}))",
            f"((self::input) and (@type='submit' or @type='button') and contains(@value, {lit}))",
            f"(@role='button' and contains(normalize-space(.), {lit}))",
            f"(@aria-label and contains(@aria-label, {lit}))",
            f"(@title and contains(@title, {lit}))",
        ]
        return "//*[" + " or ".join(conditions) + "]"

    def _find_visible(self, xpath: str, exact_text: str | None = None) -> list:
        from selenium.webdriver.common.by import By

        driver = self._get_driver()
        elements = [el for el in driver.find_elements(By.XPATH, xpath) if el.is_displayed()]

        def label_of(el) -> str:
            return (el.text or el.get_attribute("value") or el.get_attribute("aria-label")
                    or el.get_attribute("title") or "").strip()

        if exact_text is not None:
            # 完全一致を優先し、無ければ部分一致(元の順序)にフォールバック
            elements.sort(key=lambda el: 0 if label_of(el) == exact_text else 1)
        return elements

    def click_by_text(
        self, text_hint: str, timeout: int = 10, obstruction_wait_seconds: float = 0
    ) -> str:
        """画面に表示されているであろう文字(ボタン/リンクのラベル)を手がかりにクリックする。
        CSSクラス名やid指定に依存しないため、多少のUI変更(見た目の微修正等)に強い。

        obstruction_wait_seconds: 広告のポップアップ等、別の要素に覆われてクリックが
        妨害された場合の対応。0(既定)なら妨害を検知した時点ですぐにエラーにする。
        0より大きい値を指定すると、その秒数の間は1秒おきに自動で再試行しながら
        待機する(いつ閉じられるか予測できないため、手動で閉じてもらうのを想定した
        待機)。待機しても解消しなければエラーで停止する。
        """
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.common.exceptions import TimeoutException

        self._assert_still_on_site()
        driver = self._get_driver()
        xpath = self._build_clickable_xpath(text_hint)

        try:
            WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))
        except TimeoutException:
            raise ElementNotFoundError(
                f"'{text_hint}' という表示のボタン/リンクが見つかりませんでした"
            )

        candidates = self._click_with_obstruction_wait(
            driver,
            lambda: self._find_visible(xpath, exact_text=text_hint),
            obstruction_wait_seconds,
            f"'{text_hint}' に一致する、現在表示されている要素",
        )
        self._assert_still_on_site()
        logger.info("テキスト一致でクリックしました: '%s' (候補%d件中1件目)", text_hint, len(candidates))
        return f"clicked by text: {text_hint}"

    def _click_with_obstruction_wait(
        self,
        driver,
        locate_candidates,
        obstruction_wait_seconds: float,
        not_found_label: str,
    ) -> list:
        """候補要素を探して先頭をクリックする。広告等の別要素にクリックを
        妨害された場合、obstruction_wait_seconds が0より大きければその秒数の間、
        1秒おきに候補を再取得してクリックを再試行する(DOMが変わっても対応
        できるよう、毎回locate_candidates()で探し直す)。要素が見つからない場合は
        待機せずすぐにエラーにする(妨害待ちは「見えているが押せない」場合のみ)。
        """
        from selenium.common.exceptions import (
            ElementClickInterceptedException,
            ElementNotInteractableException,
            StaleElementReferenceException,
        )

        deadline = time.monotonic() + obstruction_wait_seconds
        while True:
            candidates = locate_candidates()
            if not candidates:
                raise ElementNotFoundError(f"{not_found_label}が見つかりませんでした")
            target = candidates[0]
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target)
            try:
                target.click()
                return candidates
            except (
                ElementClickInterceptedException,
                ElementNotInteractableException,
                StaleElementReferenceException,
            ) as e:
                if obstruction_wait_seconds <= 0 or time.monotonic() >= deadline:
                    suffix = (
                        f"(手動で閉じるのを{obstruction_wait_seconds}秒待ちましたが解消しませんでした)"
                        if obstruction_wait_seconds > 0 else ""
                    )
                    raise ElementNotFoundError(
                        f"{not_found_label}は見つかりましたが、広告等の別の要素にクリックを"
                        f"妨害されています{suffix}: {e}"
                    ) from e
                time.sleep(1.0)

    def _build_input_xpath(self, label_hint: str, tag: str) -> str:
        lit = _xpath_literal(label_hint)
        conditions = [
            f"(@placeholder and contains(@placeholder, {lit}))",
            f"(@aria-label and contains(@aria-label, {lit}))",
            f"(@name and contains(@name, {lit}))",
            f"(@id and contains(@id, {lit}))",
        ]
        return f"//{tag}[" + " or ".join(conditions) + "]"

    def _find_input_like(self, label_hint: str, tags: tuple[str, ...]):
        from selenium.webdriver.common.by import By
        from selenium.common.exceptions import NoSuchElementException

        driver = self._get_driver()

        for tag in tags:
            candidates = self._find_visible(self._build_input_xpath(label_hint, tag))
            if candidates:
                return candidates[0]

        # placeholder等で見つからない場合、<label>の文言から辿る
        lit = _xpath_literal(label_hint)
        labels = driver.find_elements(By.XPATH, f"//label[contains(normalize-space(.), {lit})]")
        for lbl in labels:
            for_id = lbl.get_attribute("for")
            if for_id:
                try:
                    el = driver.find_element(By.ID, for_id)
                    if el.is_displayed():
                        return el
                except NoSuchElementException:
                    pass
            try:
                tag_xpath = " | ".join(f".//following::{t}[1]" for t in tags)
                el = lbl.find_element(By.XPATH, tag_xpath)
                if el.is_displayed():
                    return el
            except NoSuchElementException:
                continue
        return None

    def type_by_text(
        self, label_hint: str, value: str, clear_first: bool = True, press_enter: bool = False
    ) -> str:
        """ラベル・プレースホルダー・name/id等を手がかりに入力欄を探して入力する。
        press_enter=True の場合、入力後にEnterキー(送信)を送る。
        """
        self._assert_still_on_site()
        driver = self._get_driver()
        target = self._find_input_like(label_hint, tags=("input", "textarea"))
        if target is None:
            raise ElementNotFoundError(f"'{label_hint}' に一致する入力欄が見つかりませんでした")

        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target)
        if clear_first:
            target.clear()
        target.send_keys(value)
        if press_enter:
            from selenium.webdriver.common.keys import Keys
            target.send_keys(Keys.RETURN)
        self._assert_still_on_site()
        suffix = " (Enterで送信)" if press_enter else ""
        return f"typed into: {label_hint}{suffix}"

    def select_by_text(self, label_hint: str, option_text: str) -> str:
        """<select>要素をラベル手がかりで探し、表示文字が一致する選択肢を選ぶ。"""
        from selenium.webdriver.support.ui import Select

        self._assert_still_on_site()
        target = self._find_input_like(label_hint, tags=("select",))
        if target is None:
            raise ElementNotFoundError(f"'{label_hint}' に一致するドロップダウンが見つかりませんでした")

        select = Select(target)
        options = [o.text.strip() for o in select.options]
        exact = [o for o in options if o == option_text]
        if exact:
            select.select_by_visible_text(exact[0])
        else:
            partial = [o for o in options if option_text in o]
            if not partial:
                raise ElementNotFoundError(
                    f"'{label_hint}' の中に '{option_text}' という選択肢が見つかりませんでした"
                    f"(選択肢: {options})"
                )
            select.select_by_visible_text(partial[0])
        self._assert_still_on_site()
        return f"selected: {label_hint} = {option_text}"

    def wait_seconds(self, seconds: float) -> str:
        time.sleep(seconds)
        return f"waited {seconds}s"

    def wait_for_text(self, text: str, timeout: int = 15) -> str:
        """指定した文字が画面に表示されるまで待つ(画面遷移・非同期読み込み待ち用)。"""
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.common.exceptions import TimeoutException

        driver = self._get_driver()
        lit = _xpath_literal(text)
        xpath = f"//*[contains(normalize-space(.), {lit})]"
        try:
            WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))
        except TimeoutException:
            raise ElementNotFoundError(f"'{text}' が{timeout}秒待っても表示されませんでした")
        return f"appeared: {text}"

    # ---------- 操作後の成功確認(検証) ----------
    # 「押せたつもりが実は押せていなかった」を防ぐため、click_by_text等の実行直後に
    # ここのメソッドで結果を確認する。失敗時は VerificationFailedError を送出し、
    # Executor側でマクロの実行をその場で止める。

    def get_current_url(self) -> str:
        if self._driver is None:
            return ""
        return self._driver.current_url

    def verify_text_appears(self, text: str, timeout: int = 10) -> str:
        try:
            self.wait_for_text(text, timeout=timeout)
        except ElementNotFoundError as e:
            raise VerificationFailedError(
                f"操作後に '{text}' が表示されるはずでしたが、確認できませんでした: {e}"
            )
        return f"verified: '{text}' の表示を確認しました"

    def verify_text_disappears(self, text: str, timeout: int = 10) -> str:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.common.exceptions import TimeoutException

        driver = self._get_driver()
        lit = _xpath_literal(text)
        xpath = f"//*[contains(normalize-space(.), {lit})]"
        try:
            WebDriverWait(driver, timeout).until(
                lambda d: len(d.find_elements(By.XPATH, xpath)) == 0
            )
        except TimeoutException:
            raise VerificationFailedError(
                f"操作後に '{text}' が消えるはずでしたが、{timeout}秒経っても表示されたままでした"
            )
        return f"verified: '{text}' が消えたことを確認しました"

    def verify_url_changed(self, before_url: str, timeout: int = 10) -> str:
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.common.exceptions import TimeoutException

        driver = self._get_driver()
        try:
            WebDriverWait(driver, timeout).until(lambda d: d.current_url != before_url)
        except TimeoutException:
            raise VerificationFailedError(
                f"操作後にURLが変わるはずでしたが、{timeout}秒経っても "
                f"'{before_url}' のままでした(操作が反映されていない可能性があります)"
            )
        return "verified: URLの変化を確認しました"

    def take_screenshot(self, save_path: str) -> str:
        """今の画面をPNGで保存する(主に失敗時の自動記録用)。"""
        driver = self._get_driver()
        Path(save_path).parent.mkdir(parents=True, exist_ok=True)
        driver.save_screenshot(save_path)
        return save_path

    # ---------- ヘルスチェック用(非破壊: クリック/入力せず存在だけ確認する) ----------

    def check_clickable_exists(self, text_hint: str) -> bool:
        xpath = self._build_clickable_xpath(text_hint)
        return len(self._find_visible(xpath)) > 0

    def check_input_exists(self, label_hint: str) -> bool:
        return self._find_input_like(label_hint, tags=("input", "textarea")) is not None

    def check_select_exists(self, label_hint: str) -> bool:
        return self._find_input_like(label_hint, tags=("select",)) is not None

    def list_interactive_texts(self, limit: int = 30) -> list[str]:
        """今の画面で押せそうな要素の表示文字一覧を返す(レコーダーでの候補提示用)。"""
        from selenium.webdriver.common.by import By

        self._assert_still_on_site()
        driver = self._get_driver()
        xpath = "//button | //a | //input[@type='submit' or @type='button'] | //*[@role='button']"
        texts: list[str] = []
        for el in driver.find_elements(By.XPATH, xpath):
            if not el.is_displayed():
                continue
            t = (el.text or el.get_attribute("value") or el.get_attribute("aria-label") or "").strip()
            if t and t not in texts:
                texts.append(t)
            if len(texts) >= limit:
                break
        return texts

    # ---------- CSSセレクタ直接指定版(F12の開発者ツールで調べたclass/id向け) ----------
    # click_by_text等の表示テキスト検索でうまく見つからない場合の代替手段。
    # UI変更に弱くなるため、まずは *_by_text 系を試し、それでも見つからない
    # ときだけこちらを使うことを推奨する。

    def click_selector(self, selector: str, obstruction_wait_seconds: float = 0) -> str:
        """obstruction_wait_seconds: click_by_textと同じ考え方で、広告等の別要素に
        クリックを妨害された場合の最大待機秒数(0なら待機せずすぐにエラー)。
        """
        from selenium.common.exceptions import NoSuchElementException

        self._assert_still_on_site()
        driver = self._get_driver()

        def _locate():
            try:
                return [driver.find_element("css selector", selector)]
            except NoSuchElementException:
                return []

        self._click_with_obstruction_wait(
            driver, _locate, obstruction_wait_seconds, f"セレクタ '{selector}' の要素"
        )
        self._assert_still_on_site()
        return f"clicked: {selector}"

    def type_by_selector(
        self, selector: str, value: str, clear_first: bool = True, press_enter: bool = False
    ) -> str:
        """F12の開発者ツールで調べたCSSセレクタ(class/id等)を直接指定して入力する。
        press_enter=True の場合、入力後にEnterキー(送信)を送る。
        """
        self._assert_still_on_site()
        driver = self._get_driver()
        el = driver.find_element("css selector", selector)
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
        if clear_first:
            el.clear()
        el.send_keys(value)
        if press_enter:
            from selenium.webdriver.common.keys import Keys
            el.send_keys(Keys.RETURN)
        self._assert_still_on_site()
        suffix = " (Enterで送信)" if press_enter else ""
        return f"typed into selector: {selector}{suffix}"

    def select_by_selector(self, selector: str, option_text: str) -> str:
        """F12の開発者ツールで調べたCSSセレクタを直接指定してドロップダウンから選ぶ。"""
        from selenium.webdriver.support.ui import Select

        self._assert_still_on_site()
        driver = self._get_driver()
        el = driver.find_element("css selector", selector)
        select = Select(el)
        options = [o.text.strip() for o in select.options]
        exact = [o for o in options if o == option_text]
        if exact:
            select.select_by_visible_text(exact[0])
        else:
            partial = [o for o in options if option_text in o]
            if not partial:
                raise ElementNotFoundError(
                    f"selector '{selector}' の中に '{option_text}' が見つかりませんでした"
                    f"(選択肢: {options})"
                )
            select.select_by_visible_text(partial[0])
        self._assert_still_on_site()
        return f"selected via selector: {selector} = {option_text}"

    def get_text_by_selector(self, selector: str) -> str:
        """CSSセレクタで指定した要素の表示文字を取得する(画面から値を読み取って
        Excelへの転記やメール本文への埋め込みに使う。store_asと組み合わせる)。
        """
        self._assert_still_on_site()
        driver = self._get_driver()
        el = driver.find_element("css selector", selector)
        return el.text.strip()

    def get_attribute_by_selector(self, selector: str, attribute: str) -> str:
        """CSSセレクタで指定した要素の属性値(href/value/data-*等)を取得する。"""
        self._assert_still_on_site()
        driver = self._get_driver()
        el = driver.find_element("css selector", selector)
        value = el.get_attribute(attribute)
        return "" if value is None else value

    def get_text_list_by_selector(self, selector: str) -> list[str]:
        """CSSセレクタに一致するすべての要素の表示文字をリストとして取得する
        (一覧・表の1列を丸ごと読み取りたい場合等に使う。例: "table tr td:nth-child(2)")。
        取得したリストは store_as で変数に保存すれば、for繰り返し構文と組み合わせて
        1件ずつ処理できる。
        """
        self._assert_still_on_site()
        driver = self._get_driver()
        elements = driver.find_elements("css selector", selector)
        return [el.text.strip() for el in elements]

    def _find_checkbox_like(self, label_hint: str):
        """ラベルの表示文字を手がかりにチェックボックス本体(<input type=checkbox>)を探す。

        デザイン上、本体のinputを `opacity:0` や `width/height:0` で見た目上
        完全に隠し、`::before`/`::after`等の疑似要素で「四角にチェックが入る」
        見た目だけを描く実装がよくある(Bootstrapのカスタムチェックボックス等)。
        この場合Seleniumの is_displayed() は False を返す(要素は実在するが
        見た目のサイズが無いため)。ここでは意図的に is_displayed() では
        絞り込まず、DOM上に実在するかどうかだけで判定する(疑似要素自体は
        DOMノードではないためSeleniumから直接操作できないが、紐づく本体の
        inputは実在するのでそちらを操作対象にする)。
        戻り値: (checkbox_input, clickable_element)。clickable_elementは
        実際にクリックを試す対象(<label>があればそちらを優先。無ければinput自身)。
        """
        from selenium.webdriver.common.by import By
        from selenium.common.exceptions import NoSuchElementException

        driver = self._get_driver()

        def _label_for(checkbox):
            cb_id = checkbox.get_attribute("id")
            if cb_id:
                try:
                    return driver.find_element(By.XPATH, f"//label[@for={_xpath_literal(cb_id)}]")
                except NoSuchElementException:
                    pass
            try:
                return checkbox.find_element(By.XPATH, "ancestor::label[1]")
            except NoSuchElementException:
                return None

        all_inputs = driver.find_elements(By.XPATH, self._build_input_xpath(label_hint, "input"))
        checkboxes = [el for el in all_inputs if el.get_attribute("type") == "checkbox"]
        if checkboxes:
            cb = checkboxes[0]
            return cb, (_label_for(cb) or cb)

        lit = _xpath_literal(label_hint)
        labels = driver.find_elements(By.XPATH, f"//label[contains(normalize-space(.), {lit})]")
        for lbl in labels:
            cb = None
            for_id = lbl.get_attribute("for")
            if for_id:
                try:
                    candidate = driver.find_element(By.ID, for_id)
                    if candidate.get_attribute("type") == "checkbox":
                        cb = candidate
                except NoSuchElementException:
                    pass
            if cb is None:
                try:
                    cb = lbl.find_element(By.XPATH, ".//input[@type='checkbox']")
                except NoSuchElementException:
                    continue
            return cb, lbl
        return None

    def check_checkbox_by_text(self, label_hint: str, checked: bool = True) -> str:
        """ラベルの表示文字を手がかりにチェックボックスを探し、指定した状態
        (checked)に合わせる(既に希望の状態ならクリックしない)。

        本体のinputが見た目上隠されているカスタムデザインのチェックボックス
        (::before等で描画する類)にも対応するため、まず紐づく<label>への
        通常クリックを試し、それが失敗する場合はJavaScriptで直接
        input側の.click()を発火させる(座標ベースの操作ではないため、
        見た目のサイズ・位置に関わらず確実に状態を切り替えられる)。
        """
        self._assert_still_on_site()
        found = self._find_checkbox_like(label_hint)
        if found is None:
            raise ElementNotFoundError(f"'{label_hint}' に一致するチェックボックスが見つかりませんでした")
        checkbox, clickable = found

        if checkbox.is_selected() != checked:
            driver = self._get_driver()
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", clickable)
                clickable.click()
            except Exception:  # noqa: BLE001
                pass
            if checkbox.is_selected() != checked:
                driver.execute_script("arguments[0].click();", checkbox)

        if checkbox.is_selected() != checked:
            raise VerificationFailedError(
                f"'{label_hint}' のチェック状態を {checked} に変更できませんでした"
            )
        self._assert_still_on_site()
        return f"checkbox '{label_hint}' -> {checked}"

    def fill_form(self, form_values: dict[str, str]) -> str:
        """form_values 例: {"#name": "GAO", "#memo": "設備点検依頼"}"""
        self._assert_still_on_site()
        driver = self._get_driver()
        for selector, value in form_values.items():
            el = driver.find_element("css selector", selector)
            el.clear()
            el.send_keys(value)
        self._assert_still_on_site()
        return f"{len(form_values)} 項目を入力しました"

    # ---------- 番号指定(インデックス)による操作・プレビュー ----------
    # 表のセル、名前の無いボタン、カスタムデザインのトグル等、表示テキストでの
    # 目印が取りづらい要素向けに、画面に今表示されている要素を種類ごとに
    # 1番目から順に番号付けし、その番号を指定して操作する方式。
    # list_*系で「今何番に何があるか」を事前にプレビューしてから、対応する
    # *_by_index系(表はget_table_cell_text)で実際に操作・登録する、という
    # 2段階の使い方を想定している。番号は呼び出すたびにその場で数え直すため、
    # ページの内容が変わると対応がずれる可能性がある(プレビューした直後に
    # 登録することを推奨)。

    _CLICKABLE_XPATH = (
        "//button | //a | //input[@type='submit' or @type='button'] | //*[@role='button']"
    )
    _INPUT_XPATH = (
        "//input[not(@type='submit') and not(@type='button') and not(@type='checkbox') "
        "and not(@type='radio') and not(@type='hidden')] | //textarea"
    )
    _CHECKBOX_XPATH = "//input[@type='checkbox']"
    _TOGGLE_XPATH = "//*[@role='switch'] | //*[@aria-pressed]"
    _DROPDOWN_XPATH = "//select"
    _TABLE_XPATH = "//table"

    @staticmethod
    def _element_label(el) -> str:
        """要素自身の「表示されているであろう文字」を、テキスト→value→aria-label→
        titleの優先順で1つ取得する(見つからなければ空文字)。"""
        return (el.text or el.get_attribute("value") or el.get_attribute("aria-label")
                or el.get_attribute("title") or "").strip()

    def _find_label_for(self, el) -> str:
        """入力欄・チェックボックス等、自分自身に表示文字を持たない要素向けに、
        aria-label/placeholder/title、または紐づく<label>から目印になりそうな
        文字列を探す(見つからなければname属性、それも無ければ空文字)。"""
        from selenium.webdriver.common.by import By
        from selenium.common.exceptions import NoSuchElementException

        for attr in ("aria-label", "placeholder", "title"):
            value = el.get_attribute(attr)
            if value and value.strip():
                return value.strip()
        driver = self._get_driver()
        el_id = el.get_attribute("id")
        if el_id:
            try:
                lbl = driver.find_element(By.XPATH, f"//label[@for={_xpath_literal(el_id)}]")
                if lbl.text.strip():
                    return lbl.text.strip()
            except NoSuchElementException:
                pass
        try:
            lbl = el.find_element(By.XPATH, "ancestor::label[1]")
            if lbl.text.strip():
                return lbl.text.strip()
        except NoSuchElementException:
            pass
        return (el.get_attribute("name") or "").strip()

    def _enumerate_visible(self, xpath: str) -> list:
        """xpathに一致する、今画面に表示されている要素を出現順に取得する
        (1番目からの番号付けは呼び出し側で行う)。"""
        from selenium.webdriver.common.by import By

        self._assert_still_on_site()
        driver = self._get_driver()
        return [el for el in driver.find_elements(By.XPATH, xpath) if el.is_displayed()]

    def _nth_visible(self, xpath: str, index: int, kind_label: str):
        """xpathに一致する表示中の要素のうち、index番目(1始まり)を返す。
        範囲外ならエラーにする。"""
        elements = self._enumerate_visible(xpath)
        if not (1 <= index <= len(elements)):
            raise ElementNotFoundError(
                f"{kind_label}の{index}番目は存在しません"
                f"(今表示されているのは{len(elements)}件です。一覧で番号を確認し直してください)"
            )
        return elements[index - 1]

    # ---- ボタン/リンク(通常ボタン) ----

    def list_clickable_elements(self) -> list[dict]:
        """今画面に表示されているボタン/リンクを1番目から順に列挙する
        (登録前のプレビュー用)。"""
        elements = self._enumerate_visible(self._CLICKABLE_XPATH)
        return [
            {"index": i, "text": self._element_label(el), "tag": el.tag_name}
            for i, el in enumerate(elements, start=1)
        ]

    def click_by_index(self, index: int, obstruction_wait_seconds: float = 0) -> str:
        """list_clickable_elementsで確認した番号(1始まり)のボタン/リンクをクリックする。
        obstruction_wait_secondsはclick_by_textと同じ(広告等に妨害された場合の
        手動close待ちの最大秒数。0なら待機せずすぐにエラー)。
        """
        def _locate():
            elements = self._enumerate_visible(self._CLICKABLE_XPATH)
            if not (1 <= index <= len(elements)):
                return []
            return [elements[index - 1]]

        driver = self._get_driver()
        self._click_with_obstruction_wait(
            driver, _locate, obstruction_wait_seconds, f"{index}番目のボタン/リンク"
        )
        self._assert_still_on_site()
        return f"clicked by index: {index}"

    # ---- 入力欄 ----

    def list_input_elements(self) -> list[dict]:
        """今画面に表示されている入力欄(input/textarea)を1番目から順に列挙する。"""
        elements = self._enumerate_visible(self._INPUT_XPATH)
        return [
            {
                "index": i,
                "label": self._find_label_for(el),
                "current_value": el.get_attribute("value") or "",
                "type": el.get_attribute("type") or el.tag_name,
            }
            for i, el in enumerate(elements, start=1)
        ]

    def type_by_index(
        self, index: int, value: str, clear_first: bool = True, press_enter: bool = False
    ) -> str:
        """list_input_elementsで確認した番号(1始まり)の入力欄に入力する。"""
        self._assert_still_on_site()
        target = self._nth_visible(self._INPUT_XPATH, index, "入力欄")
        driver = self._get_driver()
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target)
        if clear_first:
            target.clear()
        target.send_keys(value)
        if press_enter:
            from selenium.webdriver.common.keys import Keys
            target.send_keys(Keys.RETURN)
        self._assert_still_on_site()
        suffix = " (Enterで送信)" if press_enter else ""
        return f"typed into index {index}{suffix}"

    # ---- チェックボックス ----

    def list_checkbox_elements(self) -> list[dict]:
        """今画面に表示されているチェックボックスを1番目から順に列挙する。
        opacity:0等で見た目上隠されたカスタムデザインのチェックボックスは
        「表示中」の判定から外れるため対象にならない(その場合は表示テキストでの
        操作`check_checkbox_by_text`を使うこと)。"""
        elements = self._enumerate_visible(self._CHECKBOX_XPATH)
        return [
            {"index": i, "label": self._find_label_for(el), "checked": el.is_selected()}
            for i, el in enumerate(elements, start=1)
        ]

    def check_checkbox_by_index(self, index: int, checked: bool = True) -> str:
        """list_checkbox_elementsで確認した番号(1始まり)のチェックボックスを
        指定した状態にする。"""
        self._assert_still_on_site()
        target = self._nth_visible(self._CHECKBOX_XPATH, index, "チェックボックス")
        driver = self._get_driver()
        if target.is_selected() != checked:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target)
                target.click()
            except Exception:  # noqa: BLE001
                pass
            if target.is_selected() != checked:
                driver.execute_script("arguments[0].click();", target)
        if target.is_selected() != checked:
            raise VerificationFailedError(f"{index}番目のチェック状態を{checked}に変更できませんでした")
        self._assert_still_on_site()
        return f"checkbox index {index} -> {checked}"

    # ---- トグルボタン(role=switch / aria-pressed) ----

    def list_toggle_elements(self) -> list[dict]:
        """今画面に表示されているトグルボタン(role=switch または aria-pressed属性を
        持つ要素)を1番目から順に列挙する。"""
        elements = self._enumerate_visible(self._TOGGLE_XPATH)
        result = []
        for i, el in enumerate(elements, start=1):
            state = el.get_attribute("aria-checked") or el.get_attribute("aria-pressed") or ""
            result.append({
                "index": i,
                "label": self._find_label_for(el) or self._element_label(el),
                "on": state.lower() == "true",
            })
        return result

    def toggle_by_index(self, index: int, on: bool = True) -> str:
        """list_toggle_elementsで確認した番号(1始まり)のトグルボタンを
        指定した状態(on/off)にする。"""
        self._assert_still_on_site()
        target = self._nth_visible(self._TOGGLE_XPATH, index, "トグルボタン")
        driver = self._get_driver()

        def _current_state() -> bool:
            state = target.get_attribute("aria-checked") or target.get_attribute("aria-pressed") or ""
            return state.lower() == "true"

        if _current_state() != on:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target)
            target.click()
            if _current_state() != on:
                raise VerificationFailedError(f"{index}番目のトグル状態を{on}に変更できませんでした")
        self._assert_still_on_site()
        return f"toggle index {index} -> {on}"

    # ---- プルダウン(<select>) ----

    def list_dropdown_elements(self) -> list[dict]:
        """今画面に表示されているプルダウン(<select>)を1番目から順に列挙する
        (選択肢の一覧と、現在選ばれている項目も併せて返す)。"""
        from selenium.webdriver.support.ui import Select

        elements = self._enumerate_visible(self._DROPDOWN_XPATH)
        result = []
        for i, el in enumerate(elements, start=1):
            select = Select(el)
            options = [o.text.strip() for o in select.options]
            selected = [o.text.strip() for o in select.all_selected_options]
            result.append({
                "index": i,
                "label": self._find_label_for(el),
                "options": options,
                "selected": selected[0] if selected else "",
            })
        return result

    def select_by_index(self, index: int, option_text: str) -> str:
        """list_dropdown_elementsで確認した番号(1始まり)のプルダウンから、
        表示文字が一致する選択肢を選ぶ。"""
        from selenium.webdriver.support.ui import Select

        self._assert_still_on_site()
        target = self._nth_visible(self._DROPDOWN_XPATH, index, "プルダウン")
        select = Select(target)
        options = [o.text.strip() for o in select.options]
        exact = [o for o in options if o == option_text]
        if exact:
            select.select_by_visible_text(exact[0])
        else:
            partial = [o for o in options if option_text in o]
            if not partial:
                raise ElementNotFoundError(
                    f"{index}番目のプルダウンの中に '{option_text}' が見つかりませんでした"
                    f"(選択肢: {options})"
                )
            select.select_by_visible_text(partial[0])
        self._assert_still_on_site()
        return f"selected index {index} = {option_text}"

    # ---- 表(テーブル) ----

    def list_tables(self) -> list[dict]:
        """今画面に表示されている表(<table>)を1番目から順に列挙し、行数・列数と
        先頭行のセルのプレビューを返す(どの表か見分けるための参考情報)。"""
        elements = self._enumerate_visible(self._TABLE_XPATH)
        result = []
        for i, table_el in enumerate(elements, start=1):
            rows = table_el.find_elements("xpath", ".//tr")
            row_count = len(rows)
            col_count = 0
            preview_cells: list[str] = []
            if rows:
                cells = rows[0].find_elements("xpath", "./th | ./td")
                col_count = len(cells)
                preview_cells = [c.text.strip() for c in cells[:4]]
            result.append({"index": i, "rows": row_count, "columns": col_count, "preview": preview_cells})
        return result

    def get_table_cell_text(self, table_index: int, row: int, column: int) -> str:
        """table_index番目(1始まり)の表の、row行目・column列目(いずれも1始まり)の
        セルの文字を取得する。行ごとにセル数が異なる場合がある(colspan等)ため、
        事前にlist_tablesで大まかな行数・列数を確認しておくこと。"""
        table_el = self._nth_visible(self._TABLE_XPATH, table_index, "表")
        rows = table_el.find_elements("xpath", ".//tr")
        if not (1 <= row <= len(rows)):
            raise ElementNotFoundError(
                f"{table_index}番目の表の{row}行目は存在しません(全{len(rows)}行)"
            )
        cells = rows[row - 1].find_elements("xpath", "./th | ./td")
        if not (1 <= column <= len(cells)):
            raise ElementNotFoundError(
                f"{table_index}番目の表の{row}行目に{column}列目は存在しません"
                f"(この行は{len(cells)}列です)"
            )
        return cells[column - 1].text.strip()

    # ---------- 画面をPDFとして保存(印刷) ----------

    def save_page_as_pdf(
        self,
        save_path: str,
        scale: float = 1.0,
        landscape: bool = False,
        paper_width: float = 8.27,
        paper_height: float = 11.69,
        margin: float = 0.4,
        print_background: bool = True,
    ) -> str:
        """今表示している画面を、Chromeの印刷機能(DevTools Protocol)でPDFとして
        保存する。scale は縮尺(1.0=100%)で、内容を1ページに収めたい場合は
        小さい値(例: 0.7)を指定する。用紙サイズは既定でA4(単位: インチ)。
        """
        self._assert_still_on_site()
        driver = self._get_driver()

        result = driver.execute_cdp_cmd("Page.printToPDF", {
            "landscape": landscape,
            "printBackground": print_background,
            "scale": scale,
            "paperWidth": paper_width,
            "paperHeight": paper_height,
            "marginTop": margin,
            "marginBottom": margin,
            "marginLeft": margin,
            "marginRight": margin,
            "preferCSSPageSize": False,
        })

        import base64
        out = Path(save_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        with open(out, "wb") as f:
            f.write(base64.b64decode(result["data"]))
        logger.info("画面をPDFとして保存しました: %s (scale=%s)", out, scale)
        return str(out)

    def close(self) -> str:
        if self._driver is not None:
            self._driver.quit()
            self._driver = None
        self._current_site_key = None
        self._tab_handles = {}
        self._tab_site_urls = {}
        self._current_tab_alias = None
        return "browser closed"
