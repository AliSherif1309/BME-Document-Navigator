"""Shared, GUI-independent library services. No document files are modified."""
from contextlib import contextmanager
from html.parser import HTMLParser
from pathlib import Path
import os
import re
import sqlite3
import time

TEXT_EXTENSIONS = {'.pdf', '.docx', '.txt', '.md', '.csv', '.html', '.htm'}
SUPPORTED = TEXT_EXTENSIONS | {'.doc', '.rtf', '.xlsx', '.xls', '.pptx', '.ppt',
    '.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff', '.zip', '.7z', '.rar'}
FIELDS = ('manufacturer', 'device_model', 'document_type', 'keywords',
          'revision_number', 'revision_date', 'status', 'applicable_models',
          'associated_test_equipment')


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
                            old = db.execute('SELECT id,last_modified FROM documents WHERE filepath=?',
                                             (str(path),)).fetchone()
                            if old and stamp == old['last_modified']: continue
                            # Extract first: a failed read must not destroy the previous index.
                            pages = extract_pages(path)
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
