"""Regression checks for indexing, retained history, search, and local HTTP boundaries."""
from pathlib import Path
import json
import os
import tempfile
import threading
import unittest
from unittest.mock import patch
from urllib.error import HTTPError
from urllib.request import Request, urlopen
from http.server import ThreadingHTTPServer

from navigator_core import Library
from navigator_web import App, handler_for


class LibraryFixture(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.docs = self.folder / 'Philips'
        self.docs.mkdir()
        self.library = Library(self.folder / 'library.db')
        self.library.add_path(str(self.docs))

    def create(self, name='pump_service.txt', text='E042 pressure sensor calibration'):
        path = self.docs / name
        path.write_text(text, encoding='utf-8')
        return path

    def first(self):
        return self.library.search()['documents'][0]['id']


class LibraryTests(LibraryFixture):
    def test_incremental_full_text_and_metadata_filters(self):
        path = self.create()
        result = self.library.scan()
        self.assertEqual((result['added'], result['reindexed']), (1, 1))
        hit = self.library.search('E042 pressure', 'Philips', 'Service manual')
        self.assertEqual(hit['total'], 1)
        self.assertEqual(hit['documents'][0]['page_number'], 0)
        self.assertEqual(self.library.scan()['updated'], 0)
        doc_id = self.first()
        self.library.update_metadata(doc_id, {'manufacturer': 'Custom', 'device_model': 'ABC-7', 'status': 'Approved'})
        path.write_text('replacement fault code E999', encoding='utf-8')
        stamp = path.stat().st_mtime + 5
        os.utime(path, (stamp, stamp))
        self.assertEqual(self.library.scan()['updated'], 1)
        self.assertEqual(self.library.search('E042')['total'], 0)
        self.assertEqual(self.library.search('E999')['total'], 1)
        self.assertEqual(self.library.search('ABC-7', 'Custom')['total'], 1)
        self.assertEqual(self.library.details(doc_id)['status'], 'Approved')

    def test_offline_folder_and_removed_file_preserve_history(self):
        path = self.create()
        self.library.scan()
        doc_id = self.first()
        self.library.add_note(doc_id, 'Keep this reference', 3)
        self.library.toggle_favorite(doc_id)
        path.unlink()
        self.docs.rmdir()
        result = self.library.scan()
        self.assertTrue(result['warnings'])
        doc = self.library.details(doc_id)
        self.assertFalse(doc['available'])
        self.assertEqual(doc['notes'][0]['page_number'], 2)
        self.assertTrue(doc['favorite'])
        self.assertEqual(self.library.search(favorites=True)['total'], 1)

    def test_extraction_failure_preserves_previous_index_and_retries(self):
        path = self.create()
        self.library.scan()
        stamp = path.stat().st_mtime + 10
        os.utime(path, (stamp, stamp))
        with patch('navigator_core.extract_pages', side_effect=OSError('File locked')):
            self.assertEqual(self.library.scan()['errors'], 1)
        self.assertEqual(self.library.search('E042')['total'], 1)
        self.assertEqual(self.library.scan()['updated'], 1)

    def test_html_entities_hidden_content_and_literal_search(self):
        self.create('guide.html', '<style>hidden_style</style><script>hidden_script</script><p>oxygen &amp; flow</p>')
        self.library.scan()
        self.assertEqual(self.library.search('oxygen " & flow:*')['total'], 1)
        self.assertEqual(self.library.search('hidden_script')['total'], 0)
        self.assertEqual(self.library.search('oxygen', kind='SOP')['total'], 0)
        self.assertEqual(self.library.search('nonexistent OR oxygen')['total'], 0)

    def test_overlapping_roots_do_not_duplicate_and_empty_query_paginates(self):
        self.create()
        self.library.add_path(str(self.folder))
        result = self.library.scan()
        self.assertEqual(result['added'], 1)
        self.assertEqual(self.library.search(limit=1)['total'], 1)
        self.assertEqual(self.library.search(offset=1)['documents'], [])

    def test_docx_tables_and_pdf_page_numbers(self):
        try:
            import pymupdf as fitz
            from docx import Document
        except ImportError:
            self.skipTest('Install requirements.txt to verify PDF and DOCX extraction')
        document = Document()
        table = document.add_table(rows=1, cols=1)
        table.cell(0, 0).text = 'Tablefault987'
        document.save(self.docs / 'table.docx')
        with fitz.open() as pdf:
            pdf.new_page().insert_text((72,72), 'First page')
            pdf.new_page().insert_text((72,72), 'Secondfault654')
            pdf.save(self.docs / 'pages.pdf')
        self.assertEqual(self.library.scan()['errors'], 0)
        self.assertEqual(self.library.search('Tablefault987')['total'], 1)
        self.assertEqual(self.library.search('Secondfault654')['documents'][0]['page_number'], 1)


class HTTPFixture(LibraryFixture):
    def setUp(self):
        super().setUp()
        self.create('=formula.html', '<h1>oxygen</h1>')
        self.library.scan()
        self.app = App(self.library)
        self.server = ThreadingHTTPServer(('127.0.0.1', 0), handler_for(self.app))
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.base = f'http://127.0.0.1:{self.server.server_port}'
        self.addCleanup(self.close_server)

    def close_server(self):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join()

    def request(self, path, body=None, token=True, headers=None):
        request = Request(self.base + path,
            data=json.dumps(body).encode() if body is not None else None,
            headers={**({'X-Navigator-Token': self.app.token} if token else {}), **(headers or {})})
        return urlopen(request, timeout=5)


class HTTPTests(HTTPFixture):
    def test_ui_bootstrap_and_api_authorization(self):
        with self.request('/', token=False) as response:
            html = response.read().decode()
            self.assertIn(self.app.token, html)
            self.assertNotIn('__SESSION_TOKEN__', html)
            self.assertIn("script-src 'self'", response.headers['Content-Security-Policy'])
        with self.assertRaises(HTTPError) as error: self.request('/api/stats', token=False)
        self.assertEqual(error.exception.code, 403)
        with self.assertRaises(HTTPError) as error:
            self.request('/api/favorite', {'id': self.first()}, headers={'Origin':'https://example.com'})
        self.assertEqual(error.exception.code, 403)
        with self.assertRaises(HTTPError) as error:
            self.request('/api/stats', headers={'Host':'attacker.example'})
        self.assertEqual(error.exception.code, 403)

    def test_mutations_export_and_safe_html_download(self):
        doc_id = self.first()
        with self.request('/api/notes', {'id':doc_id,'text':'<script>inert</script>','page':2}): pass
        with self.request('/api/favorite', {'id':doc_id}): pass
        with self.request('/api/document?id=' + str(doc_id)) as response:
            doc = json.load(response)
            self.assertEqual(doc['notes'][0]['page_number'], 1)
            self.assertTrue(doc['favorite'])
        with self.request('/api/export') as response:
            self.assertIn("'=formula.html", response.read().decode('utf-8-sig'))
        with self.request('/api/file?id=' + str(doc_id)) as response:
            self.assertEqual(response.headers['Content-Type'], 'application/octet-stream')
            self.assertIn('attachment', response.headers['Content-Disposition'])
        with self.request('/api/backup') as response:
            self.assertTrue(response.read().startswith(b'SQLite format 3'))

    def test_invalid_ids_payloads_and_static_traversal(self):
        for path in ['/api/document?id=99999', '/api/documents?offset=nope']:
            with self.assertRaises(HTTPError) as error: self.request(path)
            self.assertEqual(error.exception.code, 400)
        with self.assertRaises(HTTPError) as error: self.request('/../navigator_web.py')
        self.assertEqual(error.exception.code, 404)
        with self.assertRaises(HTTPError) as error: self.request('/api/notes', {'id': self.first(),'text':'','page':0})
        self.assertEqual(error.exception.code, 400)


if __name__ == '__main__': unittest.main()
