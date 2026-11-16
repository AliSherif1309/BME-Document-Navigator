"""Build on the target OS: python build_app.py. Requires PyInstaller."""
from pathlib import Path
import argparse
import subprocess
import sys


def main():
    root = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description='Build a portable BME Navigator bundle.')
    parser.add_argument('--output', type=Path, default=root / 'dist', help='Output directory; use a different one while an older bundle is running')
    parser.add_argument('--replace', action='store_true', help='Replace an existing bundle after closing it')
    args = parser.parse_args()
    if (args.output.resolve() / 'BME-Navigator').exists() and not args.replace:
        parser.error('An application bundle already exists here. Use a new --output folder, or close it and pass --replace.')
    subprocess.run([sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean',
        '--name', 'BME-Navigator', '--onedir', '--console',
        '--distpath', str(args.output.resolve()),
        '--add-data', f'{root / "web"}:web',
        '--collect-all', 'pymupdf', '--collect-all', 'docx',
        str(root / 'navigator_web.py')], cwd=root, check=True)
    print(f'Ready in {args.output.resolve() / "BME-Navigator"}. Distribute the entire folder.')


if __name__ == '__main__': main()
