"""
共通のバックアップユーティリティ。

macros.json / pipelines.json など、動作定義を上書きする前に必ずこれを通し、
タイムスタンプ付きのコピーを config/backups/ に保存しておく。
誤って登録内容を書き換えてしまった/壊してしまった場合に、人の手で
戻せるようにするための仕組み(自動復元は行わない。復元は該当ファイルを
バックアップからコピーし直すだけでよい)。
"""
from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path


def backup_file(path: Path, backup_dir: Path, keep: int = 30) -> Path | None:
    path = Path(path)
    if not path.exists():
        return None

    backup_dir = Path(backup_dir)
    backup_dir.mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    dest = backup_dir / f"{path.stem}_{ts}{path.suffix}"
    shutil.copy2(path, dest)

    # 直近 keep 件だけ残し、古いものは削除して肥大化を防ぐ
    pattern = f"{path.stem}_*{path.suffix}"
    backups = sorted(backup_dir.glob(pattern))
    if len(backups) > keep:
        for old in backups[: len(backups) - keep]:
            old.unlink(missing_ok=True)

    return dest
