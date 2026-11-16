"""Local, cancellable PDF OCR with atomic index replacement and provenance."""
import json
import os
from pathlib import Path
import re
import threading
import time

from navigator_core import fingerprint
from pdf_support import open_pdf, pdf_task


@pdf_task
def pdf_count(path):
    with open_pdf(path) as pdf:
        if pdf.needs_pass: raise ValueError('Unlock this PDF in your PDF application before OCR.')
        return pdf.page_count


@pdf_task
def recognize_page(path, page, language, dpi, tessdata, mode):
    with open_pdf(path) as pdf:
        source = pdf[page]
        native = source.get_text(sort=True)
        if mode == 'missing' and native.strip(): return {'text':native,'ocr':False,'words':[]}
        pixels = source.rect.width * source.rect.height * (dpi / 72) ** 2
        if pixels > 20000000: raise ValueError(f'Page {page+1} is too large at {dpi} DPI. Choose a lower resolution.')
        textpage = source.get_textpage_ocr(language=language,dpi=dpi,full=True,tessdata=tessdata)
        text = source.get_text(sort=True,textpage=textpage)
        words = source.get_text('words',textpage=textpage)
        if len(text)>2000000 or len(words)>100000:
            raise ValueError(f'Page {page+1} exceeds the OCR text limit.')
        return {'text':text,'ocr':True,'words':[list(word[:5]) for word in words]}


class OCRService:
    def __init__(self, library, operation_lock):
        self.library = library
        self.operation_lock = operation_lock
        self.lock = threading.Lock()
        self.cancel_event = threading.Event()
        self.state = dict(running=False,message='No OCR job has run yet.',page=0,total=0,result=None)

    def setup(self):
        with self.library.connect() as db:
            row = db.execute("SELECT value FROM app_settings WHERE key='ocr_tessdata'").fetchone()
        explicit = row['value'] if row else (os.environ.get('BMEDOCUMENTNAVIGATOR_TESSDATA') or os.environ.get('TESSDATA_PREFIX'))
        candidates = [explicit] if explicit else [
            str(Path(os.environ.get('PROGRAMFILES','C:/Program Files'))/'Tesseract-OCR'/'tessdata'),
            '/opt/homebrew/share/tessdata','/usr/local/share/tessdata',
            '/usr/share/tesseract-ocr/5/tessdata','/usr/share/tesseract-ocr/4.00/tessdata','/usr/share/tessdata']
        for candidate in candidates:
            folder = Path(candidate).expanduser()
            if not folder.is_dir(): continue
            languages = sorted(p.stem for p in folder.glob('*.traineddata')
                               if p.stem not in ('osd','equ') and re.fullmatch(r'[A-Za-z0-9_]+',p.stem))
            if languages: return dict(available=True,tessdata=str(folder.resolve()),languages=languages,
                                     message='Local language data is ready.')
        return dict(available=False,tessdata=explicit or '',languages=[],
                    message='Choose a tessdata folder containing Tesseract .traineddata language files.')

    def configure(self, folder):
        folder = str(folder).strip()
        if folder:
            path = Path(folder).expanduser().resolve()
            if not path.is_dir() or not any(path.glob('*.traineddata')):
                raise ValueError('Choose a folder containing Tesseract .traineddata language files.')
            folder = str(path)
        with self.library.connect() as db:
            if folder: db.execute("INSERT OR REPLACE INTO app_settings VALUES ('ocr_tessdata',?)",(folder,))
            else: db.execute("DELETE FROM app_settings WHERE key='ocr_tessdata'")
        return self.setup()

    def status(self):
        with self.lock: return dict(self.state)

    def cancel(self):
        with self.lock:
            if self.state['running']:
                self.cancel_event.set()
                self.state['message']='Cancelling after the current page; the previous index will be retained.'

    def start(self, doc_id, languages, dpi=200, mode='missing'):
        setup = self.setup()
        if not setup['available']: raise ValueError(setup['message'])
        if (not isinstance(languages,list) or not languages or len(languages)>4 or
                any(not isinstance(item,str) or item not in setup['languages'] for item in languages)):
            raise ValueError('Select 1–4 installed OCR languages.')
        if dpi not in (150,200,300): raise ValueError('Choose 150, 200, or 300 DPI.')
        if mode not in ('missing','all'): raise ValueError('Choose scanned pages or all pages.')
        doc = self.library.details(doc_id)
        if Path(doc['filepath']).suffix.lower()!='.pdf' or not doc['available']:
            raise ValueError('OCR requires an accessible PDF document.')
        if not self.operation_lock.acquire(blocking=False):
            raise ValueError('Wait for the current scan or library operation to finish.')
        self.cancel_event.clear()
        language = '+'.join(dict.fromkeys(languages))
        with self.lock:
            self.state = dict(running=True,message='Preparing local OCR…',doc_id=doc_id,
                filename=doc['filename'],page=0,total=0,result=None)
        threading.Thread(target=self.run,args=(doc,language,dpi,setup['tessdata'],mode),daemon=True).start()
        return {'started':True}

    def run(self, doc, language, dpi, tessdata, mode):
        try:
            path = Path(doc['filepath'])
            identity = fingerprint(path)
            total = pdf_count(path)
            if not total: raise ValueError('This PDF has no pages.')
            pages = []
            staged_bytes = 0
            with self.lock: self.state['total']=total
            for index in range(total):
                if self.cancel_event.is_set(): break
                with self.lock:
                    self.state.update(page=index+1,message=f'Recognizing page {index+1} of {total}…')
                page = recognize_page(path,index,language,dpi,tessdata,mode)
                page['words'] = json.dumps(page['words'])
                staged_bytes += len(page['text'].encode('utf-8')) + len(page['words'].encode('utf-8'))
                if staged_bytes>64000000: raise ValueError('This document exceeds the 64 MB OCR staging limit. Split it into smaller PDFs.')
                pages.append(page)
                stat = path.stat()
                if (stat.st_size,stat.st_mtime_ns)!=(identity['size'],identity['mtime_ns']):
                    raise ValueError('The PDF changed during OCR. The previous index was retained.')
            if self.cancel_event.is_set():
                with self.lock: self.state.update(running=False,message='OCR cancelled. The previous index was retained.',result={'cancelled':True})
                return
            if fingerprint(path)['sha256']!=identity['sha256']:
                raise ValueError('The PDF changed during OCR. The previous index was retained.')
            now = time.time()
            # Serialize cancellation with the final transaction; either all pages commit or none do.
            with self.lock:
                if self.cancel_event.is_set():
                    self.state.update(running=False,message='OCR cancelled. The previous index was retained.',result={'cancelled':True})
                    return
                with self.library.connect() as db:
                    db.execute('BEGIN IMMEDIATE')
                    current = db.execute('SELECT filepath FROM documents WHERE id=?',(doc['id'],)).fetchone()
                    if not current or current['filepath']!=doc['filepath']:
                        raise ValueError('The document moved during OCR. The previous index was retained.')
                    stat = path.stat()
                    if (stat.st_size,stat.st_mtime_ns)!=(identity['size'],identity['mtime_ns']):
                        raise ValueError('The PDF changed before OCR finished. The previous index was retained.')
                    db.execute('DELETE FROM documents_fts WHERE doc_id=?',(doc['id'],))
                    db.execute('DELETE FROM ocr_pages WHERE doc_id=?',(doc['id'],))
                    for index,page in enumerate(pages):
                        if page['text'].strip(): db.execute('INSERT INTO documents_fts VALUES (?,?,?)',(doc['id'],index,page['text']))
                        if page['ocr']:
                            db.execute('INSERT INTO ocr_pages VALUES (?,?,?,?,?,?,?,?)',
                                (doc['id'],index,page['text'],page['words'],language,dpi,identity['sha256'],now))
                    db.execute('UPDATE documents SET last_modified=? WHERE id=?',(stat.st_mtime,doc['id']))
                    db.execute('INSERT OR REPLACE INTO document_files VALUES (?,?,?,?)',
                        (doc['id'],identity['sha256'],identity['size'],identity['mtime_ns']))
                converted = sum(page['ocr'] for page in pages)
                self.state.update(running=False,message='OCR complete. Search and page text are ready.',
                    result=dict(ocr_pages=converted,skipped=total-converted,language=language,dpi=dpi))
        except Exception as error:
            with self.lock: self.state.update(running=False,message='OCR failed: '+str(error),result={'error':str(error)})
        finally:
            with self.lock: self.state['running']=False
            self.operation_lock.release()
