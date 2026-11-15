"""Shared, GUI-independent library services. No document files are modified."""
from contextlib import contextmanager
from html.parser import HTMLParser
from pathlib import Path
import hashlib
import json
import os
import re
import secrets
import sqlite3
import time

TEXT_EXTENSIONS = {'.pdf', '.docx', '.txt', '.md', '.csv', '.html', '.htm'}
SUPPORTED = TEXT_EXTENSIONS | {'.doc', '.rtf', '.xlsx', '.xls', '.pptx', '.ppt',
    '.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff', '.zip', '.7z', '.rar'}
FIELDS = ('manufacturer', 'device_model', 'document_type', 'keywords',
          'revision_number', 'revision_date', 'status', 'applicable_models',
          'associated_test_equipment')


def fingerprint(path):
    """Read in bounded chunks and reject files that change while being hashed."""
    before = path.stat()
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    after = path.stat()
    if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
        raise ValueError(f'File changed while reading: {path.name}. Try again.')
    return {'sha256': digest.hexdigest(), 'size': after.st_size,
            'mtime_ns': after.st_mtime_ns}


class HTMLText(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.hidden = 0
        self.parts = []

    def handle_starttag(self, tag, attrs):
        if tag in ('script', 'style'): self.hidden += 1

    def handle_endtag(self, tag):
        if tag in ('script', 'style'): self.hidden = max(0, self.hidden - 1)

    def handle_data(self, data):
        if not self.hidden: self.parts.append(data)


def extract_pages(path):
    ext = path.suffix.lower()
    if ext == '.pdf':
        import pymupdf as fitz
        with fitz.open(path) as document:
            return [(i, page.get_text()) for i, page in enumerate(document)]
    if ext == '.docx':
        from docx import Document
        document = Document(path)
        text = '\n'.join([p.text for p in document.paragraphs] +
                         [' | '.join(c.text for c in row.cells)
                          for table in document.tables for row in table.rows])
        return [(0, text)]
    if ext in TEXT_EXTENSIONS:
        raw = path.read_bytes()
        for encoding in ('utf-8-sig', 'utf-16', 'cp1252'):
            try:
                text = raw.decode(encoding)
                break
            except UnicodeError:
                continue
        else:
            text = raw.decode('utf-8', errors='replace')
        if ext in ('.html', '.htm'):
            parser = HTMLText()
            parser.feed(text)
            text = ' '.join(parser.parts)
        return [(0, text)]
    return []


def infer_metadata(path):
    words = re.sub(r'[_\W]+', ' ', str(path).lower())
    manufacturer = next((name for name in ('Philips', 'Siemens', 'GE', 'Dräger',
        'Draeger', 'Medtronic', 'Mindray', 'Baxter', 'B Braun', 'Stryker',
        'Getinge', 'Olympus', 'Fresenius')
        if re.search(r'\b' + re.escape(name.lower()) + r'\b', words)), '')
    kind = next((label for key, label in (
        ('service', 'Service manual'), ('calibration', 'Calibration'),
        ('sop', 'SOP'), ('pm', 'Preventive maintenance'), ('user', 'User manual'),
        ('datasheet', 'Datasheet'), ('manual', 'Manual'))
        if re.search(r'\b' + key + r'\b', words)), '')
    return manufacturer, kind


class Library:
    def __init__(self, database):
        self.database = str(database)
        with self.connect() as db:
            db.executescript('''
                CREATE TABLE IF NOT EXISTS scan_paths (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, path TEXT NOT NULL UNIQUE);
                CREATE TABLE IF NOT EXISTS documents (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT NOT NULL,
                    filepath TEXT NOT NULL UNIQUE, manufacturer TEXT, device_model TEXT,
                    document_type TEXT, keywords TEXT, last_modified REAL NOT NULL,
                    revision_number TEXT, revision_date TEXT, status TEXT,
                    applicable_models TEXT, associated_test_equipment TEXT);
                CREATE VIRTUAL TABLE IF NOT EXISTS documents_fts USING fts5(
                    doc_id UNINDEXED, page_number UNINDEXED, content);
                CREATE TRIGGER IF NOT EXISTS documents_ad_trigger AFTER DELETE ON documents
                    BEGIN DELETE FROM documents_fts WHERE doc_id=old.id; END;
                CREATE TABLE IF NOT EXISTS notes (
                    note_id INTEGER PRIMARY KEY AUTOINCREMENT, doc_id INTEGER NOT NULL,
                    page_number INTEGER, note_text TEXT NOT NULL, created_timestamp REAL NOT NULL,
                    FOREIGN KEY(doc_id) REFERENCES documents(id) ON DELETE CASCADE);
                CREATE TABLE IF NOT EXISTS favorites (
                    fav_id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE,
                    doc_id INTEGER NOT NULL, page_number INTEGER NOT NULL,
                    FOREIGN KEY(doc_id) REFERENCES documents(id) ON DELETE CASCADE);
                CREATE TABLE IF NOT EXISTS links (
                    link_id INTEGER PRIMARY KEY AUTOINCREMENT, source_doc_id INTEGER NOT NULL,
                    target_doc_id INTEGER NOT NULL, description TEXT,
                    FOREIGN KEY(source_doc_id) REFERENCES documents(id) ON DELETE CASCADE,
                    FOREIGN KEY(target_doc_id) REFERENCES documents(id) ON DELETE CASCADE,
                    UNIQUE(source_doc_id, target_doc_id));
                CREATE INDEX IF NOT EXISTS idx_note_doc ON notes(doc_id);
                CREATE INDEX IF NOT EXISTS idx_fav_doc ON favorites(doc_id);
                CREATE TABLE IF NOT EXISTS document_files (
                    doc_id INTEGER PRIMARY KEY REFERENCES documents(id) ON DELETE CASCADE,
                    sha256 TEXT NOT NULL, size INTEGER NOT NULL, mtime_ns INTEGER NOT NULL);
                CREATE TABLE IF NOT EXISTS saved_searches (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE COLLATE NOCASE,
                    filters TEXT NOT NULL, created_timestamp REAL NOT NULL);
            ''')

    @contextmanager
    def connect(self):
        db = sqlite3.connect(self.database, timeout=30)
        db.row_factory = sqlite3.Row
        db.execute('PRAGMA foreign_keys=ON')
        try:
            with db: yield db
        finally:
            db.close()

    def paths(self):
        with self.connect() as db:
            return [row[0] for row in db.execute('SELECT path FROM scan_paths ORDER BY path')]

    def add_path(self, path):
        if not str(path).strip(): raise ValueError('Enter a full folder path.')
        folder = Path(path).expanduser().resolve()
        if not folder.is_dir(): raise ValueError('Choose an existing, accessible folder.')
        with self.connect() as db:
            db.execute('INSERT OR IGNORE INTO scan_paths(path) VALUES (?)', (str(folder),))

    def scan(self, notify=lambda message: None):
        """Incremental indexing, per-file rollback, and retention of offline records.

        Missing documents are deliberately retained with their notes and bookmarks.
        Removing a scan path also never deletes library history.
        """
        started = time.monotonic()
        result = dict(added=0, updated=0, reindexed=0, removed=0, errors=0,
                      warnings=[], processed=0)
        seen = set()
        for folder in self.paths():
            if not Path(folder).is_dir():
                result['warnings'].append('Folder unavailable: ' + folder)
                continue
            def walk_error(error):
                result['warnings'].append(str(error))
            for root, dirs, files in os.walk(folder, onerror=walk_error):
                dirs[:] = sorted(d for d in dirs if d not in ('.git', '.venv', '__pycache__'))
                for name in sorted(files):
                    path = Path(root, name).resolve()
                    if path.suffix.lower() not in SUPPORTED or name.startswith('~$') or path in seen:
                        continue
                    seen.add(path)
                    result['processed'] += 1
                    notify({'type': 'status', 'message': 'Indexing: ' + name})
                    try:
                        stamp = path.stat().st_mtime
                        with self.connect() as db:
                            old = db.execute('''SELECT d.id,d.last_modified,f.sha256,f.size,f.mtime_ns
                                FROM documents d LEFT JOIN document_files f ON f.doc_id=d.id WHERE filepath=?''',
                                             (str(path),)).fetchone()
                            current = path.stat()
                            unchanged = old and stamp == old['last_modified'] and (
                                old['sha256'] is None or
                                (old['size'], old['mtime_ns']) == (current.st_size, current.st_mtime_ns))
                            if unchanged and old['sha256']: continue
                            identity = fingerprint(path)
                            # Extract first: a failed read must not destroy the previous index.
                            pages = extract_pages(path)
                            after = path.stat()
                            if (after.st_size, after.st_mtime_ns) != (identity['size'], identity['mtime_ns']):
                                raise ValueError('File changed during extraction; retry the scan.')
                            if old:
                                doc_id = old['id']
                                db.execute('UPDATE documents SET last_modified=?,filename=? WHERE id=?',
                                           (stamp, name, doc_id))
                                db.execute('DELETE FROM documents_fts WHERE doc_id=?', (doc_id,))
                            else:
                                manufacturer, kind = infer_metadata(path)
                                doc_id = db.execute('''INSERT INTO documents
                                    (filename,filepath,last_modified,manufacturer,document_type)
                                    VALUES (?,?,?,?,?)''',
                                    (name, str(path), stamp, manufacturer, kind)).lastrowid
                            db.executemany('INSERT INTO documents_fts(doc_id,page_number,content) VALUES (?,?,?)',
                                [(doc_id, page, text) for page, text in pages if text.strip()])
                            db.execute('INSERT OR REPLACE INTO document_files VALUES (?,?,?,?)',
                                (doc_id, identity['sha256'], identity['size'], identity['mtime_ns']))
                        result['updated' if old else 'added'] += 1
                        if any(text.strip() for _, text in pages): result['reindexed'] += 1
                    except Exception as error:
                        result['errors'] += 1
                        result['warnings'].append(f'{name}: {error}')
        result['duration'] = round(time.monotonic() - started, 2)
        return result

    def search(self, query='', manufacturer='', kind='', favorites=False, offset=0, limit=100):
        # Literal tokens prevent user punctuation being interpreted as FTS syntax.
        tokens = re.findall(r'\w+', query, flags=re.UNICODE)
        clauses, params = [], []
        with self.connect() as db:
            if tokens:
                match = ' AND '.join('"' + token + '"' for token in tokens)
                clauses.append('''(d.id IN (SELECT doc_id FROM documents_fts WHERE documents_fts MATCH ?)
                    OR (''' + ' AND '.join("lower(coalesce(d.filename,'') || ' ' || " +
                    " || ' ' || ".join(f"coalesce(d.{field},'')" for field in FIELDS) + ') LIKE ? ESCAPE \'\\\''
                    for _ in tokens) + '))')
                params = [match] + ['%' + t.lower().replace('_', '\\_') + '%' for t in tokens]
            if manufacturer:
                clauses.append('d.manufacturer=?'); params.append(manufacturer)
            if kind:
                clauses.append('d.document_type=?'); params.append(kind)
            if favorites:
                clauses.append('EXISTS(SELECT 1 FROM favorites f WHERE f.doc_id=d.id)')
            where = ' WHERE ' + ' AND '.join(clauses) if clauses else ''
            total = db.execute('SELECT count(*) FROM documents d' + where, params).fetchone()[0]
            rows = db.execute('''SELECT d.*, EXISTS(SELECT 1 FROM favorites f WHERE f.doc_id=d.id) AS favorite
                FROM documents d''' + where + ' ORDER BY d.filename COLLATE NOCASE,d.id LIMIT ? OFFSET ?',
                params + [limit, offset]).fetchall()
            documents = [dict(row) for row in rows]
            if tokens:
                for doc in documents:
                    hit = db.execute('''SELECT page_number, snippet(documents_fts,2,'[',']',' … ',24) AS snippet
                        FROM documents_fts WHERE documents_fts MATCH ? AND doc_id=? ORDER BY rank LIMIT 1''',
                        (match, doc['id'])).fetchone()
                    if hit: doc.update(dict(hit))
            return {'documents': documents, 'total': total}

    def details(self, doc_id):
        with self.connect() as db:
            row = db.execute('SELECT * FROM documents WHERE id=?', (doc_id,)).fetchone()
            if row is None: raise ValueError('Document not found.')
            doc = dict(row)
            doc['available'] = Path(doc['filepath']).is_file()
            doc['notes'] = [dict(r) for r in db.execute(
                'SELECT * FROM notes WHERE doc_id=? ORDER BY created_timestamp DESC', (doc_id,))]
            doc['favorite'] = bool(db.execute('SELECT 1 FROM favorites WHERE doc_id=?', (doc_id,)).fetchone())
            doc['indexed_pages'] = db.execute('SELECT count(*) FROM documents_fts WHERE doc_id=?', (doc_id,)).fetchone()[0]
            return doc

    def update_metadata(self, doc_id, values):
        self.details(doc_id)
        with self.connect() as db:
            db.execute('UPDATE documents SET ' + ','.join(f'{f}=?' for f in FIELDS) + ' WHERE id=?',
                       [str(values.get(f, '')).strip() for f in FIELDS] + [doc_id])

    def saved_searches(self):
        with self.connect() as db:
            return [{**dict(row), 'filters': json.loads(row['filters'])}
                    for row in db.execute('SELECT * FROM saved_searches ORDER BY name COLLATE NOCASE')]

    def save_search(self, name, filters):
        name = str(name).strip()
        if not name or len(name) > 100:
            raise ValueError('Enter a search name of 1–100 characters.')
        if not isinstance(filters, dict): raise ValueError('Invalid search filters.')
        clean = {key: str(filters.get(key, '')).strip() for key in ('q', 'manufacturer', 'kind')}
        if any(len(value) > 4000 for value in clean.values()):
            raise ValueError('Search filters are too long.')
        if not isinstance(filters.get('favorites', False), bool):
            raise ValueError('Invalid favorites filter.')
        clean['favorites'] = filters.get('favorites', False)
        with self.connect() as db:
            try:
                return db.execute('INSERT INTO saved_searches(name,filters,created_timestamp) VALUES (?,?,?)',
                    (name, json.dumps(clean), time.time())).lastrowid
            except sqlite3.IntegrityError:
                raise ValueError('A saved search already uses that name. Choose another name.') from None

    def delete_search(self, search_id):
        with self.connect() as db:
            db.execute('DELETE FROM saved_searches WHERE id=?', (search_id,))

    def preview_batch(self, ids, values):
        if not isinstance(ids, list) or not 1 <= len(ids) <= 500:
            raise ValueError('Select 1–500 documents.')
        if any(type(item) is not int or item < 1 for item in ids):
            raise ValueError('Invalid document selection.')
        if not isinstance(values, dict) or not values or set(values) - set(FIELDS):
            raise ValueError('Select the metadata fields to change.')
        if any(not isinstance(value, str) or len(value) > 4000 for value in values.values()):
            raise ValueError('Metadata values must be text of at most 4000 characters.')
        clean = {field: value.strip() for field, value in values.items()}
        changes = []
        with self.connect() as db:
            for doc_id in dict.fromkeys(ids):
                row = db.execute('SELECT * FROM documents WHERE id=?', (doc_id,)).fetchone()
                if row is None: raise ValueError('A selected document no longer exists. Refresh the library.')
                before = {field: row[field] for field in clean}
                if all((before[field] or '') == value for field, value in clean.items()): continue
                changes.append({'id': doc_id, 'filename': row['filename'], 'before': before, 'after': clean})
        return {'kind': 'batch', 'changes': changes, 'selected': len(set(ids))}

    def backup(self):
        folder = Path(self.database).resolve().parent / 'backups'
        folder.mkdir(exist_ok=True)
        path = folder / f'library-{time.strftime("%Y%m%d-%H%M%S")}-{secrets.token_hex(4)}.db'
        with self.connect() as source:
            destination = sqlite3.connect(path)
            try: source.backup(destination)
            finally: destination.close()
        return str(path)

    def apply_batch(self, plan):
        if plan.get('kind') != 'batch' or not plan['changes']:
            raise ValueError('There are no changes to apply.')
        backup = self.backup()
        with self.connect() as db:
            db.execute('BEGIN IMMEDIATE')
            for change in plan['changes']:
                row = db.execute('SELECT * FROM documents WHERE id=?', (change['id'],)).fetchone()
                if row is None or any(row[field] != value for field, value in change['before'].items()):
                    raise ValueError('Metadata changed since the preview. Preview again; nothing was applied.')
                values = change['after']
                db.execute('UPDATE documents SET ' + ','.join(f'{field}=?' for field in values) + ' WHERE id=?',
                           [*values.values(), change['id']])
        return {'updated': len(plan['changes']), 'backup': backup}

    def preview_relocation(self, old_folder, new_folder):
        if not str(old_folder).strip() or not str(new_folder).strip():
            raise ValueError('Enter both the previous and new folder paths.')
        old_root, new_root = Path(old_folder).expanduser().resolve(), Path(new_folder).expanduser().resolve()
        if old_root == new_root: raise ValueError('Choose a different destination folder.')
        if not new_root.is_dir(): raise ValueError('The new folder must be accessible.')
        if old_root in new_root.parents or new_root in old_root.parents:
            raise ValueError('The old and new folders must not contain one another.')
        changes = []
        with self.connect() as db:
            rows = db.execute('''SELECT d.id,d.filename,d.filepath,f.sha256,f.size FROM documents d
                LEFT JOIN document_files f ON f.doc_id=d.id ORDER BY d.filepath''').fetchall()
            occupied = {Path(row['filepath']).resolve(): row['id'] for row in rows}
            for row in rows:
                original = Path(row['filepath']).resolve()
                if not original.is_relative_to(old_root): continue
                destination = (new_root / original.relative_to(old_root)).resolve()
                change = {'id': row['id'], 'filename': row['filename'], 'old_path': row['filepath'],
                          'new_path': str(destination), 'status': 'ready', 'sha256': row['sha256']}
                try:
                    if not destination.is_relative_to(new_root):
                        raise ValueError('Destination leaves the selected folder through a symbolic link.')
                    if destination in occupied and occupied[destination] != row['id']:
                        raise ValueError('Destination is already indexed. Resolve duplicate records first.')
                    if not destination.is_file(): raise ValueError('Matching file is missing in the new folder.')
                    expected = row['sha256']
                    if not expected and original.is_file(): expected = fingerprint(original)['sha256']
                    if not expected:
                        raise ValueError('No recorded fingerprint. Reconnect and scan the original folder first.')
                    actual = fingerprint(destination)
                    if expected != actual['sha256']: raise ValueError('File content differs from the indexed original.')
                    change.update(sha256=expected, size=actual['size'])
                except (OSError, ValueError) as error:
                    change.update(status='blocked', reason=str(error))
                changes.append(change)
        if not changes: raise ValueError('No indexed documents belong to the previous folder.')
        return {'kind': 'relocation', 'old_root': str(old_root), 'new_root': str(new_root),
                'changes': changes, 'ready': sum(c['status'] == 'ready' for c in changes),
                'blocked': sum(c['status'] != 'ready' for c in changes)}

    def apply_relocation(self, plan):
        if plan.get('kind') != 'relocation' or not plan['changes'] or plan['blocked']:
            raise ValueError('Every document must be verified before reconnecting the folder.')
        backup = self.backup()
        with self.connect() as db:
            db.execute('BEGIN IMMEDIATE')
            actual_ids = {row['id'] for row in db.execute('SELECT id,filepath FROM documents')
                          if Path(row['filepath']).resolve().is_relative_to(Path(plan['old_root']))}
            if actual_ids != {change['id'] for change in plan['changes']}:
                raise ValueError('The folder contents changed in the library. Preview again; nothing was applied.')
            for change in plan['changes']:
                row = db.execute('''SELECT d.filepath,f.sha256 FROM documents d
                    LEFT JOIN document_files f ON f.doc_id=d.id WHERE d.id=?''', (change['id'],)).fetchone()
                if (row is None or row['filepath'] != change['old_path'] or
                        (row['sha256'] and row['sha256'] != change['sha256'])):
                    raise ValueError('Library changed since the preview. Preview again; nothing was applied.')
                destination = Path(change['new_path'])
                if not destination.resolve().is_relative_to(Path(plan['new_root'])):
                    raise ValueError('Destination changed since the preview. Preview again.')
                identity = fingerprint(destination)
                if identity['sha256'] != change['sha256']:
                    raise ValueError('File content changed since the preview. Nothing was applied.')
                if db.execute('SELECT 1 FROM documents WHERE filepath=? AND id!=?',
                              (str(destination), change['id'])).fetchone():
                    raise ValueError('A destination was indexed after the preview. Nothing was applied.')
                db.execute('UPDATE documents SET filepath=?,last_modified=? WHERE id=?',
                           (str(destination), destination.stat().st_mtime, change['id']))
                db.execute('INSERT OR REPLACE INTO document_files VALUES (?,?,?,?)',
                           (change['id'], identity['sha256'], identity['size'], identity['mtime_ns']))
            # Replace registered nested roots too, so an automatic scan cannot recreate old paths.
            for row in db.execute('SELECT id,path FROM scan_paths').fetchall():
                previous = Path(row['path']).resolve()
                if previous.is_relative_to(Path(plan['old_root'])):
                    target = str(Path(plan['new_root']) / previous.relative_to(Path(plan['old_root'])))
                    db.execute('DELETE FROM scan_paths WHERE id=?', (row['id'],))
                    db.execute('INSERT OR IGNORE INTO scan_paths(path) VALUES (?)', (target,))
            db.execute('INSERT OR IGNORE INTO scan_paths(path) VALUES (?)', (plan['new_root'],))
        return {'updated': len(plan['changes']), 'backup': backup}

    def add_note(self, doc_id, text, page=1):
        self.details(doc_id)
        if not text.strip(): raise ValueError('Write a note first.')
        if page < 1: raise ValueError('Page must be 1 or greater.')
        with self.connect() as db:
            db.execute('INSERT INTO notes(doc_id,page_number,note_text,created_timestamp) VALUES (?,?,?,?)',
                       (doc_id, page - 1, text.strip(), time.time()))

    def toggle_favorite(self, doc_id):
        doc = self.details(doc_id)
        with self.connect() as db:
            if doc['favorite']: db.execute('DELETE FROM favorites WHERE doc_id=?', (doc_id,))
            else: db.execute('INSERT INTO favorites(name,doc_id,page_number) VALUES (?,?,0)',
                             (f"{doc['filename']} [#{doc_id}]", doc_id))

    def stats(self):
        with self.connect() as db:
            return {
                'documents': db.execute('SELECT count(*) FROM documents').fetchone()[0],
                'indexed': db.execute('SELECT count(DISTINCT doc_id) FROM documents_fts').fetchone()[0],
                'notes': db.execute('SELECT count(*) FROM notes').fetchone()[0],
                'favorites': db.execute('SELECT count(DISTINCT doc_id) FROM favorites').fetchone()[0],
                'manufacturers': [r[0] for r in db.execute("SELECT DISTINCT manufacturer FROM documents WHERE coalesce(manufacturer,'')!='' ORDER BY manufacturer")],
                'types': [r[0] for r in db.execute("SELECT DISTINCT document_type FROM documents WHERE coalesce(document_type,'')!='' ORDER BY document_type")],
                'paths': self.paths()}
