"""Local browser application: python navigator_web.py (no web framework required)."""
import argparse
import csv
import io
import json
import secrets
import sqlite3
import sys
import tempfile
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

from app_storage import data_file
from navigator_core import FIELDS, Library
from document_reader import DocumentReader
from ocr_service import OCRService


def resource_dir():
    return Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent)) / 'web'


class App:
    def __init__(self, library):
        self.library = library
        self.reader = DocumentReader(library)
        self.token = secrets.token_urlsafe(32)
        self.lock = threading.Lock()
        self.ocr = OCRService(library,self.lock)
        self.state_lock = threading.Lock()
        self.state = {'running': False, 'message': 'Ready to index your library.', 'result': None}
        self.previews = {}

    def preview(self, kind, body):
        if not self.lock.acquire(blocking=False):
            raise ValueError('Wait for the current scan or library operation to finish.')
        try:
            if kind == 'batch': plan = self.library.preview_batch(body.get('ids'), body.get('values'))
            else: plan = self.library.preview_relocation(body.get('old_folder', ''), body.get('new_folder', ''))
            now = time.monotonic()
            self.previews = {key: value for key, value in self.previews.items() if value['expires'] > now}
            if len(self.previews) >= 16: self.previews.pop(next(iter(self.previews)))
            key = secrets.token_urlsafe(24)
            self.previews[key] = {'plan': plan, 'expires': now + 600}
            return {**plan, 'preview_id': key}
        finally:
            self.lock.release()

    def apply_preview(self, kind, key):
        if not self.lock.acquire(blocking=False):
            raise ValueError('Wait for the current scan or library operation to finish.')
        try:
            preview = self.previews.get(key)
            if not preview or preview['expires'] <= time.monotonic() or preview['plan']['kind'] != kind:
                raise ValueError('This preview expired or was already used. Preview the changes again.')
            plan = preview['plan']
            result = self.library.apply_batch(plan) if kind == 'batch' else self.library.apply_relocation(plan)
            self.previews.pop(key)
            return result
        finally:
            self.lock.release()

    def scan(self):
        if not self.lock.acquire(blocking=False): return False
        with self.state_lock:
            self.state = {'running': True, 'message': 'Starting scan…', 'result': None}
        def run():
            try:
                def notify(message):
                    with self.state_lock: self.state['message'] = message['message']
                result = self.library.scan(notify)
                with self.state_lock:
                    self.state.update(message='Scan complete.', result=result)
            except Exception as error:
                with self.state_lock: self.state['message'] = 'Scan failed: ' + str(error)
            finally:
                with self.state_lock: self.state['running'] = False
                self.lock.release()
        threading.Thread(target=run, daemon=True).start()
        return True


def handler_for(app):
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, format, *args):
            pass

        def send_data(self, data, content_type='application/json; charset=utf-8', status=200, attachment=None, headers=None):
            if not isinstance(data, bytes): data = json.dumps(data, ensure_ascii=False).encode('utf-8')
            self.send_response(status)
            self.send_header('Content-Type', content_type)
            self.send_header('Content-Length', str(len(data)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.send_header('Referrer-Policy', 'no-referrer')
            self.send_header('X-Frame-Options', 'SAMEORIGIN')
            self.send_header('Content-Security-Policy', "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data: blob:; frame-src 'self'; object-src 'none'; base-uri 'none'; form-action 'self'; frame-ancestors 'self'")
            for key, value in (headers or {}).items(): self.send_header(key, str(value))
            if attachment: self.send_header('Content-Disposition', f'attachment; filename="{attachment}"')
            self.end_headers()
            self.wfile.write(data)

        def route(self, mutate=False):
            try:
                expected = f'127.0.0.1:{self.server.server_port}'
                if self.headers.get('Host') != expected:
                    self.send_data({'error': 'Invalid host.'}, status=403); return
                parts = urlsplit(self.path)
                query = parse_qs(parts.query)
                path = parts.path
                if not mutate and path in ('/', '/app.js', '/reader.js', '/ocr.js', '/styles.css'):
                    name = {'/': 'index.html', '/app.js': 'app.js', '/reader.js': 'reader.js', '/ocr.js': 'ocr.js', '/styles.css': 'styles.css'}[path]
                    content = (resource_dir() / name).read_bytes().replace(b'__SESSION_TOKEN__', app.token.encode())
                    mime = {'/': 'text/html', '/app.js': 'text/javascript', '/reader.js': 'text/javascript', '/ocr.js': 'text/javascript', '/styles.css': 'text/css'}[path]
                    self.send_data(content, mime + '; charset=utf-8'); return
                token = self.headers.get('X-Navigator-Token', query.get('token', [''])[0])
                if not secrets.compare_digest(token, app.token):
                    self.send_data({'error': 'Reload the application to reconnect.'}, status=403); return
                if mutate:
                    origin = self.headers.get('Origin')
                    if origin and origin != f'http://{expected}':
                        self.send_data({'error': 'Invalid origin.'}, status=403); return
                    size = int(self.headers.get('Content-Length', '0'))
                    if not 0 < size <= 65536: raise ValueError('Invalid request size.')
                    body = json.loads(self.rfile.read(size))
                    if not isinstance(body, dict): raise ValueError('Expected an object.')
                    if path == '/api/paths/add': app.library.add_path(str(body.get('path', '')))
                    elif path == '/api/paths/remove':
                        with app.library.connect() as db:
                            db.execute('DELETE FROM scan_paths WHERE path=?', (body['path'],))
                    elif path == '/api/scan':
                        self.send_data({'started': app.scan()}); return
                    elif path == '/api/metadata': app.library.update_metadata(int(body['id']), body)
                    elif path == '/api/ocr/setup':
                        self.send_data(app.ocr.configure(body.get('tessdata',''))); return
                    elif path == '/api/ocr':
                        self.send_data(app.ocr.start(int(body['id']),body.get('languages'),int(body.get('dpi',200)),body.get('mode','missing'))); return
                    elif path == '/api/ocr/cancel': app.ocr.cancel()
                    elif path == '/api/reader/position': app.reader.save_position(int(body['id']), int(body['page']), float(body['zoom']))
                    elif path == '/api/bookmarks':
                        self.send_data({'id': app.reader.add_bookmark(int(body['id']), int(body['page']), body.get('name',''))}); return
                    elif path == '/api/bookmarks/delete': app.library.remove_bookmark(int(body['bookmark_id']))
                    elif path == '/api/links': app.library.add_link(int(body['id']), int(body['target_id']), body.get('description',''))
                    elif path == '/api/links/delete': app.library.remove_link(int(body['id']), int(body['link_id']))
                    elif path in ('/api/batch/preview', '/api/relocation/preview'):
                        self.send_data(app.preview('batch' if '/batch/' in path else 'relocation', body)); return
                    elif path in ('/api/batch/apply', '/api/relocation/apply'):
                        self.send_data(app.apply_preview('batch' if '/batch/' in path else 'relocation', str(body.get('preview_id', '')))); return
                    elif path == '/api/searches':
                        self.send_data({'id': app.library.save_search(body.get('name', ''), body.get('filters'))}); return
                    elif path == '/api/searches/delete': app.library.delete_search(int(body['id']))
                    elif path == '/api/notes': app.library.add_note(int(body['id']), str(body.get('text', '')), int(body.get('page', 1)))
                    elif path == '/api/notes/delete':
                        with app.library.connect() as db:
                            db.execute('DELETE FROM notes WHERE note_id=?', (int(body['note_id']),))
                    elif path == '/api/favorite': app.library.toggle_favorite(int(body['id']))
                    else:
                        self.send_data({'error': 'Route not found.'}, status=404); return
                    self.send_data({'ok': True}); return
                def arg(key, default=''): return query.get(key, [default])[0]
                if path == '/api/stats': self.send_data(app.library.stats())
                elif path == '/api/searches': self.send_data({'searches': app.library.saved_searches()})
                elif path == '/api/ocr/setup': self.send_data(app.ocr.setup())
                elif path == '/api/ocr': self.send_data(app.ocr.status())
                elif path == '/api/bookmarks': self.send_data({'bookmarks': app.library.bookmarks(int(arg('id')) if arg('id') else None)})
                elif path == '/api/links': self.send_data({'links': app.library.links(int(arg('id')))})
                elif path == '/api/reader': self.send_data(app.reader.info(int(arg('id'))))
                elif path == '/api/reader/text': self.send_data({'text': app.reader.text(int(arg('id')), int(arg('page','1')))})
                elif path == '/api/reader/page':
                    png, matches = app.reader.render(int(arg('id')), int(arg('page','1')), float(arg('zoom','1.25')), arg('q'))
                    self.send_data(png, 'image/png', headers={'X-Match-Count': matches})
                elif path == '/api/scan':
                    with app.state_lock: state = dict(app.state)
                    self.send_data(state)
                elif path == '/api/documents':
                    self.send_data(app.library.search(arg('q'), arg('manufacturer'), arg('kind'),
                        arg('favorites') == '1', max(0, int(arg('offset', '0')))))
                elif path == '/api/document': self.send_data(app.library.details(int(arg('id'))))
                elif path == '/api/file':
                    doc = app.library.details(int(arg('id')))
                    file = Path(doc['filepath'])
                    if not file.is_file(): raise ValueError('Document is unavailable. Reconnect its drive and try again.')
                    # Only PDFs render inline; HTML and all other formats download inertly.
                    is_pdf = file.suffix.lower() == '.pdf'
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/pdf' if is_pdf else 'application/octet-stream')
                    self.send_header('Content-Length', str(file.stat().st_size))
                    self.send_header('Content-Disposition', 'inline' if is_pdf else 'attachment; filename="document' + file.suffix.lower() + '"')
                    self.send_header('Cache-Control', 'no-store')
                    self.send_header('X-Content-Type-Options', 'nosniff')
                    self.send_header('Content-Security-Policy', 'sandbox')
                    self.end_headers()
                    with file.open('rb') as stream:
                        while chunk := stream.read(1024 * 1024): self.wfile.write(chunk)
                elif path == '/api/export':
                    output = io.StringIO(newline='')
                    writer = csv.writer(output)
                    fields = ('filename', 'filepath') + FIELDS
                    writer.writerow(fields)
                    with app.library.connect() as db:
                        for row in db.execute('SELECT * FROM documents ORDER BY filename'):
                            # Prevent spreadsheet formula execution when opening the export.
                            writer.writerow([("'" + str(row[f]) if str(row[f] or '').lstrip().startswith(('=', '+', '-', '@')) else row[f]) for f in fields])
                    self.send_data(output.getvalue().encode('utf-8-sig'), 'text/csv; charset=utf-8', attachment='bme-library.csv')
                elif path == '/api/backup':
                    with tempfile.TemporaryDirectory() as folder:
                        target = Path(folder) / 'library.db'
                        with app.library.connect() as source:
                            destination = sqlite3.connect(target)
                            try: source.backup(destination)
                            finally: destination.close()
                        self.send_data(target.read_bytes(), 'application/octet-stream', attachment='bme-library-backup.db')
                else: self.send_data({'error': 'Route not found.'}, status=404)
            except (ValueError, KeyError, TypeError) as error:
                self.send_data({'error': str(error)}, status=400)
            except (BrokenPipeError, ConnectionResetError):
                pass
            except Exception as error:
                self.send_data({'error': str(error)}, status=500)

        def do_GET(self): self.route()
        def do_POST(self): self.route(mutate=True)
    return Handler


def main():
    parser = argparse.ArgumentParser(description='BME Document Navigator — private local library')
    parser.add_argument('--port', type=int, default=0, help='Local port (default: choose a free port)')
    parser.add_argument('--no-browser', action='store_true')
    parser.add_argument('--scan-interval', type=int, default=300, help='Rescan interval in seconds; 0 disables automatic scans')
    parser.add_argument('--database', type=Path, help='Use a specific library database')
    args = parser.parse_args()
    if args.scan_interval < 0: parser.error('Scan interval must be 0 or greater.')
    script = Path(__file__).resolve().parent
    database = args.database or data_file('BMEDocumentNavigator', 'bme_doc_index.db', [script, Path.cwd()])
    app = App(Library(database))
    server = ThreadingHTTPServer(('127.0.0.1', args.port), handler_for(app))
    stop = threading.Event()
    def auto_scan():
        app.scan()
        while not stop.wait(args.scan_interval): app.scan()
    if args.scan_interval: threading.Thread(target=auto_scan, daemon=True).start()
    url = f'http://127.0.0.1:{server.server_port}/'
    print(f'BME Document Navigator: {url}\nKeep this window open. Press Ctrl+C to stop.', flush=True)
    if not args.no_browser: webbrowser.open(url)
    try: server.serve_forever()
    except KeyboardInterrupt: pass
    finally:
        stop.set()
        app.ocr.cancel()
        server.server_close()


if __name__ == '__main__': main()
