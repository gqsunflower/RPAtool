"""
ExplorerHandler: パスを開く・フォルダ作成・ファイル/フォルダの移動・コピー・
名前変更を行うハンドラ。ネットワークアクセスは一切行わない。

安全設計:
- 移動・コピー・名前変更の先に同名のファイル/フォルダが既に存在する場合、
  既定ではエラーにして処理を止める(意図しない上書き・データ消失を防ぐため)。
  上書きしたい場合のみ明示的に overwrite=True を指定する。
- 移動先が既存のフォルダの場合は「その中に移動/コピーする」挙動
  (Windowsエクスプローラーのドラッグ&ドロップと同じ考え方)。
- 名前変更(rename_file/rename_folder)は new_name に「名前だけ」を要求し、
  パス区切り文字が含まれる場合はエラーにする(誤って別の場所へ移動する
  ことを防ぐため。場所を変えたい場合はmove_file/move_folderを使う)。
"""
from __future__ import annotations

import logging
import os
import platform
import shutil
import subprocess
from pathlib import Path

logger = logging.getLogger("rpa_local_ai.explorer")


class PathConflictError(Exception):
    """移動/コピー先が既に存在し、上書きが許可されていない場合に送出する。"""


def _resolve_destination(src: Path, destination: str) -> Path:
    """destinationが既存のフォルダなら「その中に src と同じ名前で」置く。
    それ以外は destination をそのまま最終的なパスとして扱う。
    """
    dest = Path(destination)
    if dest.exists() and dest.is_dir():
        return dest / src.name
    return dest


class ExplorerHandler:
    def open_path(self, path: str) -> str:
        """指定したパス(フォルダ or ファイル)をエクスプローラー(既定のファイラー)で開く。"""
        p = Path(path)
        if not p.exists():
            raise FileNotFoundError(f"パスが見つかりません: {p}")

        system = platform.system()
        if system == "Windows":
            os.startfile(str(p))  # type: ignore[attr-defined]
        elif system == "Darwin":
            subprocess.Popen(["open", str(p)])
        else:
            subprocess.Popen(["xdg-open", str(p)])
        logger.info("パスを開きました: %s", p)
        return f"opened: {p}"

    def create_folder(self, path: str, exist_ok: bool = False) -> str:
        p = Path(path)
        if p.exists() and not exist_ok:
            raise FileExistsError(
                f"既にこのパスが存在します: {p}(既存のまま使う場合はexist_okを指定)"
            )
        p.mkdir(parents=True, exist_ok=exist_ok)
        logger.info("フォルダを作成しました: %s", p)
        return str(p)

    def _check_conflict(self, dest: Path, overwrite: bool) -> None:
        if dest.exists() and not overwrite:
            raise PathConflictError(
                f"移動/コピー先が既に存在します: {dest}"
                f"(上書きする場合は overwrite=true を指定してください)"
            )

    def move_file(self, source: str, destination: str, overwrite: bool = False) -> str:
        src = Path(source)
        if not src.exists():
            raise FileNotFoundError(f"移動元のファイルが見つかりません: {src}")
        if not src.is_file():
            raise ValueError(f"ファイルではありません(フォルダの場合はmove_folderを使用): {src}")

        dest = _resolve_destination(src, destination)
        self._check_conflict(dest, overwrite)
        dest.parent.mkdir(parents=True, exist_ok=True)
        if dest.exists() and overwrite:
            dest.unlink()
        shutil.move(str(src), str(dest))
        logger.info("ファイルを移動しました: %s -> %s", src, dest)
        return str(dest)

    def copy_file(self, source: str, destination: str, overwrite: bool = False) -> str:
        src = Path(source)
        if not src.exists():
            raise FileNotFoundError(f"コピー元のファイルが見つかりません: {src}")
        if not src.is_file():
            raise ValueError(f"ファイルではありません(フォルダの場合はcopy_folderを使用): {src}")

        dest = _resolve_destination(src, destination)
        self._check_conflict(dest, overwrite)
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(src), str(dest))
        logger.info("ファイルをコピーしました: %s -> %s", src, dest)
        return str(dest)

    def move_folder(self, source: str, destination: str, overwrite: bool = False) -> str:
        src = Path(source)
        if not src.exists() or not src.is_dir():
            raise NotADirectoryError(f"移動元のフォルダが見つかりません: {src}")

        dest = _resolve_destination(src, destination)
        self._check_conflict(dest, overwrite)
        if dest.exists() and overwrite:
            shutil.rmtree(dest)
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(src), str(dest))
        logger.info("フォルダを移動しました: %s -> %s", src, dest)
        return str(dest)

    def copy_folder(self, source: str, destination: str, overwrite: bool = False) -> str:
        src = Path(source)
        if not src.exists() or not src.is_dir():
            raise NotADirectoryError(f"コピー元のフォルダが見つかりません: {src}")

        dest = _resolve_destination(src, destination)
        self._check_conflict(dest, overwrite)
        if dest.exists() and overwrite:
            shutil.rmtree(dest)
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copytree(str(src), str(dest))
        logger.info("フォルダをコピーしました: %s -> %s", src, dest)
        return str(dest)

    @staticmethod
    def _validate_new_name(new_name: str) -> None:
        if not new_name or "/" in new_name or "\\" in new_name:
            raise ValueError(
                f"new_nameには『名前だけ』を指定してください(パス区切り文字は不可): {new_name!r}"
            )

    def rename_file(self, path: str, new_name: str, overwrite: bool = False) -> str:
        """ファイル名だけを変更する(同じフォルダ内で名前だけ差し替える)。"""
        src = Path(path)
        if not src.exists():
            raise FileNotFoundError(f"ファイルが見つかりません: {src}")
        if not src.is_file():
            raise ValueError(f"ファイルではありません(フォルダの場合はrename_folderを使用): {src}")
        self._validate_new_name(new_name)

        dest = src.parent / new_name
        self._check_conflict(dest, overwrite)
        if dest.exists() and overwrite:
            dest.unlink()
        src.rename(dest)
        logger.info("ファイル名を変更しました: %s -> %s", src, dest)
        return str(dest)

    def rename_folder(self, path: str, new_name: str, overwrite: bool = False) -> str:
        """フォルダ名だけを変更する(同じ場所で名前だけ差し替える)。"""
        src = Path(path)
        if not src.exists() or not src.is_dir():
            raise NotADirectoryError(f"フォルダが見つかりません: {src}")
        self._validate_new_name(new_name)

        dest = src.parent / new_name
        self._check_conflict(dest, overwrite)
        if dest.exists() and overwrite:
            shutil.rmtree(dest)
        src.rename(dest)
        logger.info("フォルダ名を変更しました: %s -> %s", src, dest)
        return str(dest)

    def path_exists(self, path: str) -> bool:
        """指定したパス(ファイル/フォルダいずれも)が存在するかどうかを返す。
        制御構文(if_goto)と組み合わせて「無ければ作る」等の分岐に使う。
        """
        return Path(path).exists()

    def list_files_in_folder(self, path: str, pattern: str = "*", include_folders: bool = False) -> list[str]:
        """フォルダ内のファイル(既定)またはファイル+フォルダの一覧を、パスの
        文字列リストとして取得する。pattern はワイルドカード指定(例: "*.xlsx"
        でExcelファイルだけ)。サブフォルダの中までは辿らない。取得したリストは
        store_as で変数に保存し、for繰り返し構文と組み合わせて1件ずつ処理できる。
        """
        p = Path(path)
        if not p.exists() or not p.is_dir():
            raise NotADirectoryError(f"フォルダが見つかりません: {p}")
        entries = sorted(p.glob(pattern))
        if not include_folders:
            entries = [e for e in entries if e.is_file()]
        return [str(e) for e in entries]

    def delete_file(self, path: str) -> str:
        """ファイルを完全に削除する(ゴミ箱には移動しない。元に戻せないため、
        レコーダーでは実行前にもう一段階の確認を挟む)。
        """
        p = Path(path)
        if not p.exists():
            raise FileNotFoundError(f"ファイルが見つかりません: {p}")
        if not p.is_file():
            raise ValueError(f"ファイルではありません(フォルダの場合はdelete_folderを使用): {p}")
        p.unlink()
        logger.info("ファイルを削除しました: %s", p)
        return f"deleted: {p}"

    def delete_folder(self, path: str, recursive: bool = False) -> str:
        """フォルダを削除する。recursive=False(既定)では中身が空の場合のみ削除でき、
        中にファイル/フォルダが残っている場合はエラーで停止する(誤って中身ごと
        消してしまうことを防ぐ)。中身ごと削除したい場合のみ明示的に
        recursive=True を指定する。
        """
        p = Path(path)
        if not p.exists() or not p.is_dir():
            raise NotADirectoryError(f"フォルダが見つかりません: {p}")
        if recursive:
            shutil.rmtree(p)
        else:
            try:
                p.rmdir()
            except OSError as e:
                raise OSError(
                    f"フォルダの中身が空ではないため削除できません: {p}"
                    f"(中身ごと削除する場合はrecursive=trueを指定してください)"
                ) from e
        logger.info("フォルダを削除しました: %s (recursive=%s)", p, recursive)
        return f"deleted: {p}"
