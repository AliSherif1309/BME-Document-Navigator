"""Verify the real portable bundle on its target OS: python smoke_app.py."""
from pathlib import Path
import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import time
from urllib.request import Request, urlopen


def main():
    import pymupdf
    from docx import Document
    root = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description='Verify a portable BME Navigator bundle.')
    parser.add_argument('--output', type=Path, default=root / 'dist')
    args = parser.parse_args()
    executable = args.output.resolve() / 'BME-Navigator' / (
        'BME-Navigator.exe' if sys.platform == 'win32' else 'BME-Navigator')
    with tempfile.TemporaryDirectory() as folder:
        work = Path(folder)
        docs = work / 'documents'
        docs.mkdir()
        with pymupdf.open() as pdf:
            pdf.new_page().insert_text((72, 72), 'PackagePDFfault123')
            pdf.new_page().insert_text((72, 72), 'PackageReaderPageTwo')
            pdf.set_toc([[1,'First reference',1],[1,'Second reference',2]])
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
                for asset in ('app.js', 'reader.js', 'styles.css'):
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
                ids = [doc['id'] for doc in api('api/documents')['documents']]
                pdf_id = api('api/documents?q=PackagePDFfault123')['documents'][0]['id']
                info = api('api/reader?id='+str(pdf_id))
                assert info['page_count'] == 2 and info['outline'][1]['page'] == 2
                request = Request(url+f'api/reader/page?id={pdf_id}&page=2&q=PackageReaderPageTwo',headers={'X-Navigator-Token':token})
                with urlopen(request,timeout=10) as response:
                    assert response.headers['Content-Type'] == 'image/png' and response.headers['X-Match-Count'] == '1'
                    assert response.read().startswith(b'\x89PNG')
                assert 'PackageReaderPageTwo' in api(f'api/reader/text?id={pdf_id}&page=2')['text']
                api('api/reader/position',{'id':pdf_id,'page':2,'zoom':1.5})
                api('api/bookmarks',{'id':pdf_id,'page':2,'name':'Packaged page reference'})
                api('api/links',{'id':ids[0],'target_id':ids[1],'description':'Packaged related document'})
                assert api('api/links?id='+str(ids[1]))['links'][0]['doc_id'] == ids[0]
                api('api/searches', {'name':'Packaged references','filters':{'q':'PackagePDFfault123'}})
                assert api('api/searches')['searches'][0]['name'] == 'Packaged references'
                preview = api('api/batch/preview', {'ids':ids,'values':{'status':'Bundle test'}})
                assert api('api/batch/apply', {'preview_id':preview['preview_id']})['updated'] == 2
                api('api/notes', {'id':ids[0], 'text':'Keep this note', 'page':1})
                api('api/favorite', {'id':ids[0]})
                relocated = work / 'relocated-documents'
                shutil.copytree(docs, relocated)
                preview = api('api/relocation/preview', {'old_folder':str(docs),'new_folder':str(relocated)})
                assert preview['blocked'] == 0, preview
                assert api('api/relocation/apply', {'preview_id':preview['preview_id']})['updated'] == 2
                doc = api('api/document?id=' + str(ids[0]))
                assert doc['notes'][0]['note_text'] == 'Keep this note' and doc['favorite']
                assert doc['status'] == 'Bundle test' and doc['available']
                assert api('api/documents?q=PackagePDFfault123')['total'] == 1
                reader = api('api/reader?id='+str(pdf_id))
                assert reader['page'] == 2 and reader['bookmarks'][0]['page_number'] == 1
                print('Portable bundle passed: library workflows, PDF reader, page bookmarks, reading position, and related documents.')
            finally:
                process.terminate()
                process.wait(timeout=15)


if __name__ == '__main__': main()
