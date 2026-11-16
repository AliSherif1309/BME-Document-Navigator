"""Serialize PDF engine access shared by indexing and the local reader."""
from contextlib import contextmanager
from concurrent.futures import ThreadPoolExecutor
from functools import wraps
import threading

PDF_LOCK = threading.RLock()
_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix='bme-pdf')
_worker = threading.local()


def pdf_task(function):
    """Keep all new reader/index PDF work on one thread, with reentrant calls."""
    @wraps(function)
    def call(*args, **kwargs):
        if getattr(_worker, 'active', False): return function(*args, **kwargs)
        def run():
            _worker.active = True
            try: return function(*args, **kwargs)
            finally: _worker.active = False
        return _executor.submit(run).result()
    return call


@contextmanager
def open_pdf(path):
    import pymupdf
    # PyMuPDF document operations must not overlap across HTTP/indexing threads.
    with PDF_LOCK:
        try:
            document = pymupdf.open(path)
        except (RuntimeError, OSError) as error:
            raise ValueError(f'Could not open this PDF: {error}') from None
        try: yield document
        finally: document.close()


@pdf_task
def extract_pdf_pages(path):
    with open_pdf(path) as document:
        return [(i, page.get_text()) for i, page in enumerate(document)]
