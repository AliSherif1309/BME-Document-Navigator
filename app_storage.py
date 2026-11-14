"""Stable per-user storage, with one-time non-destructive legacy migration."""
from pathlib import Path
from contextlib import closing
import os
import shutil
import sqlite3
import sys

def data_file(app, name, legacy_dirs=()):
    override = os.environ.get(app.upper().replace('-', '_') + '_DATA_DIR')
    if override:
        folder = Path(override).expanduser()
    elif sys.platform == 'win32':
        folder = Path(os.environ.get('LOCALAPPDATA', Path.home() / 'AppData' / 'Local')) / app
    elif sys.platform == 'darwin':
        folder = Path.home() / 'Library' / 'Application Support' / app
    else:
        folder = Path(os.environ.get('XDG_DATA_HOME', Path.home() / '.local' / 'share')) / app
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / name
    if not target.exists():
        for legacy_dir in legacy_dirs:
            source = Path(legacy_dir) / name
            if source.is_file() and source.resolve() != target.resolve():
                temporary = target.with_name(target.name + '.migrating')
                try:
                    if name.endswith('.db'):
                        with closing(sqlite3.connect(source.resolve().as_uri() + '?mode=ro', uri=True)) as old:
                            with closing(sqlite3.connect(temporary)) as new:
                                old.backup(new)
                    else:
                        shutil.copy2(source, temporary)
                    temporary.replace(target)
                finally:
                    temporary.unlink(missing_ok=True)
                break
    return str(target)
