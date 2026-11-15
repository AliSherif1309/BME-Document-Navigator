"""Read-only PDF rendering and page references for the local workspace."""
from contextlib import contextmanager
import math
from pathlib import Path

from pdf_support import open_pdf, pdf_task


class DocumentReader:
    def __init__(self, library):
        self.library = library

    @contextmanager
    def document(self, doc_id):
        details = self.library.details(doc_id)
        path = Path(details['filepath'])
        if path.suffix.lower() != '.pdf': raise ValueError('The in-app reader supports PDF files.')
        if not path.is_file(): raise ValueError('This PDF is offline. Reconnect its folder first.')
        with open_pdf(path) as pdf:
            if pdf.needs_pass: raise ValueError('This PDF is password-protected. Open it in your PDF application.')
            if not pdf.page_count: raise ValueError('This PDF has no pages.')
            yield pdf

    @staticmethod
    def check_page(pdf, page):
        if not 1 <= page <= pdf.page_count:
            raise ValueError(f'Choose a page between 1 and {pdf.page_count}.')

    @staticmethod
    def check_zoom(zoom):
        if not math.isfinite(zoom) or not .5 <= zoom <= 3:
            raise ValueError('Zoom must be between 50% and 300%.')

    @pdf_task
    def info(self, doc_id):
        with self.document(doc_id) as pdf:
            count = pdf.page_count
            outline = [{'level': level, 'title': title, 'page': page}
                       for level, title, page in pdf.get_toc() if 1 <= page <= count][:2000]
        with self.library.connect() as db:
            position = db.execute('SELECT * FROM reading_positions WHERE doc_id=?', (doc_id,)).fetchone()
        return {'page_count': count, 'outline': outline,
                'page': min(count, max(1, position['page_number'] + 1)) if position else 1,
                'zoom': position['zoom'] if position else 1.25,
                'bookmarks': self.library.bookmarks(doc_id)}

    @pdf_task
    def render(self, doc_id, page, zoom=1.25, query=''):
        self.check_zoom(zoom)
        if len(query) > 200: raise ValueError('Use a phrase of 200 characters or fewer.')
        import pymupdf
        with self.document(doc_id) as pdf:
            self.check_page(pdf, page)
            source = pdf[page - 1]
            matches = source.search_for(query.strip())[:200] if query.strip() else []
            # Draw only in this in-memory document; never save changes to the source.
            for rectangle in matches:
                source.draw_rect(rectangle, color=None, fill=(1, .8, .15), fill_opacity=.3, overlay=True)
            area = max(1, source.rect.width * source.rect.height)
            scale = min(zoom, 6000 / max(1, source.rect.width, source.rect.height), math.sqrt(12000000 / area))
            pixmap = source.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
            return pixmap.tobytes('png'), len(matches)

    @pdf_task
    def text(self, doc_id, page):
        with self.document(doc_id) as pdf:
            self.check_page(pdf, page)
            return pdf[page - 1].get_text(sort=True)

    @pdf_task
    def save_position(self, doc_id, page, zoom):
        self.check_zoom(zoom)
        with self.document(doc_id) as pdf: self.check_page(pdf, page)
        with self.library.connect() as db:
            db.execute('''INSERT INTO reading_positions(doc_id,page_number,zoom) VALUES (?,?,?)
                ON CONFLICT(doc_id) DO UPDATE SET page_number=excluded.page_number,zoom=excluded.zoom''',
                (doc_id, page - 1, zoom))

    @pdf_task
    def add_bookmark(self, doc_id, page, name):
        with self.document(doc_id) as pdf: self.check_page(pdf, page)
        return self.library.add_bookmark(doc_id, page, name)
