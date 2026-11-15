"""Build on the target OS: python build_app.py. Requires PyInstaller."""
from pathlib import Path
import subprocess
import sys


def main():
    root = Path(__file__).resolve().parent
    subprocess.run([sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean',
        '--name', 'BME-Navigator', '--onedir', '--console',
        '--add-data', f'{root / "web"}:web',
        '--collect-all', 'pymupdf', '--collect-all', 'docx',
        str(root / 'navigator_web.py')], cwd=root, check=True)
    print('Ready in dist/BME-Navigator. Distribute the entire folder.')


if __name__ == '__main__': main()
