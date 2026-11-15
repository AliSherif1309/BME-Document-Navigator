"""Verify the real portable bundle on its target OS: python smoke_app.py."""
from pathlib import Path
import json
import re
import subprocess
import sys
import tempfile
import time
from urllib.request import Request, urlopen


def main():
    import pymupdf
    from docx import Document
    root = Path(__file__).resolve().parent
    executable = root / 'dist' / 'BME-Navigator' / (
        'BME-Navigator.exe' if sys.platform == 'win32' else 'BME-Navigator')
    with tempfile.TemporaryDirectory() as folder:
        work = Path(folder)
        docs = work / 'documents'
        docs.mkdir()
        with pymupdf.open() as pdf:
            pdf.new_page().insert_text((72, 72), 'PackagePDFfault123')
            pdf.save(docs / 'reference.pdf')
        word = Document()
        word.add_paragraph('PackageWordfault456')
        word.save(docs / 'reference.docx')
        with (work / 'startup.log').open('w', encoding='utf-8') as log:
            process = subprocess.Popen([str(executable), '--no-browser',
                '--scan-interval', '0', '--database', str(work / 'library.db')],
                stdout=log, stderr=subprocess.STDOUT, cwd=work)
            try:
                deadline = time.monotonic() + 30
                match = None
                while time.monotonic() < deadline:
                    text = (work / 'startup.log').read_text(encoding='utf-8', errors='replace')
                    match = re.search(r'http://127\.0\.0\.1:\d+/', text)
                    if match: break
                    if process.poll() is not None: raise RuntimeError('Bundle failed to start: ' + text)
                    time.sleep(.1)
                if not match: raise RuntimeError('Bundle startup timed out: ' + text)
                url = match.group()
                with urlopen(url, timeout=10) as response: html = response.read().decode()
                token = re.search(r'name="navigator-token" content="([^"]+)"', html).group(1)
                for asset in ('app.js', 'styles.css'):
                    with urlopen(url + asset, timeout=10) as response:
                        assert response.read(), f'Missing asset: {asset}'
                def api(path, body=None):
                    request = Request(url + path, data=json.dumps(body).encode() if body is not None else None,
                                      headers={'X-Navigator-Token': token})
                    with urlopen(request, timeout=10) as response: return json.load(response)
                api('api/paths/add', {'path': str(docs)})
                api('api/scan', {})
                deadline = time.monotonic() + 30
                while time.monotonic() < deadline:
                    state = api('api/scan')
                    if not state['running']: break
                    time.sleep(.1)
                assert not state['running'], 'Packaged scan timed out'
                assert state['result']['added'] == 2, state
                for query in ('PackagePDFfault123', 'PackageWordfault456'):
                    assert api('api/documents?q=' + query)['total'] == 1, query
                print('Portable bundle passed: startup, HTML/CSS/JS, PDF/DOCX indexing, and search.')
            finally:
                process.terminate()
                process.wait(timeout=15)


if __name__ == '__main__': main()
