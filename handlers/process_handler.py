"""
ProcessHandler: 自分で作成したexe/pyファイルを実行するハンドラ。

安全設計:
- 実行できるのは config/exec_whitelist.json に事前登録された
  スクリプト/実行ファイルのみ(run_key経由でのみ実行)。
  browser_handlerのサイトホワイトリストと同じ考え方で、スロットの値の
  混入などにより意図しないファイルを実行してしまうことを防ぐ。
- shell=True は使わない(引数のシェルインジェクションを避けるため)。
- 標準出力・標準エラー・終了コードを取得し、終了コードが0以外の場合は
  例外を送出する(Executorのリトライ/失敗時メニューと連携できる)。
"""
from __future__ import annotations

import json
import logging
import subprocess
import sys
from pathlib import Path

logger = logging.getLogger("rpa_local_ai.process")


class ScriptNotWhitelistedError(Exception):
    pass


class ProcessExecutionError(Exception):
    pass


class ProcessHandler:
    def __init__(self, whitelist_path: Path):
        self.whitelist_path = Path(whitelist_path)
        self._scripts = self._load_whitelist()

    def _load_whitelist(self) -> dict:
        with open(self.whitelist_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("scripts", {})

    def _save_whitelist(self) -> None:
        with open(self.whitelist_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        data["scripts"] = self._scripts
        with open(self.whitelist_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def list_registered(self) -> dict:
        return dict(self._scripts)

    def register_script(self, run_key: str, path: str, kind: str) -> str:
        """kind: 'python'(pyファイルをPythonで実行) または 'exe'(実行ファイルを直接実行)"""
        if run_key in self._scripts:
            raise ValueError(f"run_key '{run_key}' は既に登録されています")
        if kind not in ("python", "exe"):
            raise ValueError("kindは 'python' か 'exe' を指定してください")
        p = Path(path)
        if not p.exists():
            raise FileNotFoundError(f"ファイルが見つかりません: {p}")
        self._scripts[run_key] = {"path": str(p), "kind": kind}
        self._save_whitelist()
        logger.info("スクリプトを登録しました: %s -> %s (%s)", run_key, p, kind)
        return f"registered: {run_key}"

    def run_registered(
        self,
        run_key: str,
        args: list[str] | None = None,
        timeout: float | None = None,
        cwd: str | None = None,
    ) -> str:
        if run_key not in self._scripts:
            raise ScriptNotWhitelistedError(
                f"'{run_key}' はホワイトリストに登録されていません。"
                f"config/exec_whitelist.json に追加してください。"
            )
        entry = self._scripts[run_key]
        path = entry["path"]
        kind = entry["kind"]
        args = args or []

        if not Path(path).exists():
            raise FileNotFoundError(f"登録されたファイルが見つかりません: {path}")

        if kind == "python":
            cmd = [sys.executable, path, *args]
        else:
            cmd = [path, *args]

        logger.info("プロセスを実行します: %s", cmd)
        try:
            result = subprocess.run(
                cmd, cwd=cwd, timeout=timeout,
                capture_output=True, text=True,
            )
        except subprocess.TimeoutExpired as e:
            raise ProcessExecutionError(f"'{run_key}' の実行がタイムアウトしました: {e}") from e
        except OSError as e:
            raise ProcessExecutionError(f"'{run_key}' を起動できませんでした: {e}") from e

        if result.returncode != 0:
            raise ProcessExecutionError(
                f"'{run_key}' が終了コード{result.returncode}で終了しました。"
                f"stderr: {(result.stderr or '')[:500]}"
            )

        logger.info("プロセスが正常終了しました: %s (returncode=0)", run_key)
        return result.stdout[:2000] if result.stdout else "(標準出力なし・正常終了)"
