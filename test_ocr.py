"""OCR commit/cancellation/provenance checks plus optional real recognition."""
import hashlib
import json
import os
import shutil
from pathlib import Path
import threading
import time
from unittest.mock import patch
from urllib.error import HTTPError

import pymupdf
from document_reader import DocumentReader
from ocr_service import OCRService
from test_navigator import LibraryFixture, HTTPFixture


def make_scan(path):
    with pymupdf.open() as original:
        page = original.new_page()
        page.insert_text((60,100),'BIOMEDICAL OCR REFERENCE',fontsize=24)
        page.insert_text((60,160),'Pressure sensor calibration',fontsize=24)
        image = page.get_pixmap(dpi=150).tobytes('png')
    with pymupdf.open() as scanned:
        page = scanned.new_page();page.insert_image(page.rect,stream=image)
        scanned.save(path)


def wait_job(service):
    deadline=time.monotonic()+60
    while time.monotonic()<deadline:
        if not service.status()['running'] and not service.operation_lock.locked(): return service.status()
        time.sleep(.02)
    raise AssertionError('OCR job did not finish in 60 seconds')


class OCRTests(LibraryFixture):
    def setUp(self):
        super().setUp()
        self.pdf=self.docs/'scan.pdf';make_scan(self.pdf);self.library.scan()
        self.doc_id=self.first();self.service=OCRService(self.library,threading.Lock())
        self.models=self.folder/'models';self.models.mkdir();(self.models/'eng.traineddata').write_bytes(b'fixture')
        self.service.configure(str(self.models))

    def recognized(self,*args):
        return dict(text='Pressure sensor calibration',ocr=True,words=[[60,140,180,170,'Pressure'],[190,140,300,170,'sensor']])

    def test_commit_provenance_reader_and_unchanged_scan(self):
        before=hashlib.sha256(self.pdf.read_bytes()).hexdigest()
        self.library.add_note(self.doc_id,'Keep note');self.library.toggle_favorite(self.doc_id)
        with patch('ocr_service.recognize_page',side_effect=self.recognized):
            self.service.start(self.doc_id,['eng']);state=wait_job(self.service)
        self.assertEqual(state['result']['ocr_pages'],1)
        self.assertEqual(self.library.search('Pressure sensor')['total'],1)
        self.assertEqual(self.library.scan()['updated'],0)
        cached=self.library.ocr_page(self.doc_id,1)
        self.assertEqual(cached['language'],'eng');self.assertEqual(cached['source_sha256'],before)
        reader=DocumentReader(self.library)
        self.assertIn('Pressure',reader.text(self.doc_id,1))
        self.assertEqual(reader.render(self.doc_id,1,query='Pressure sensor')[1],1)
        self.assertEqual(hashlib.sha256(self.pdf.read_bytes()).hexdigest(),before)
        self.assertTrue(self.library.details(self.doc_id)['favorite'])
        self.assertEqual(len(self.library.details(self.doc_id)['notes']),1)
        moved=self.folder/'moved-documents';shutil.copytree(self.docs,moved)
        self.library.apply_relocation(self.library.preview_relocation(str(self.docs),str(moved)))
        self.assertEqual(self.library.ocr_page(self.doc_id,1)['source_sha256'],before)
        self.assertIn('Pressure',reader.text(self.doc_id,1))

    def test_cancel_and_failure_preserve_previous_ocr(self):
        with patch('ocr_service.recognize_page',side_effect=self.recognized):
            self.service.start(self.doc_id,['eng']);wait_job(self.service)
        entered=threading.Event();release=threading.Event()
        def paused(*args): entered.set();release.wait(5);return dict(text='Wrong replacement',ocr=True,words=[])
        with patch('ocr_service.recognize_page',side_effect=paused):
            self.service.start(self.doc_id,['eng']);self.assertTrue(entered.wait(5));self.service.cancel();release.set()
            self.assertTrue(wait_job(self.service)['result']['cancelled'])
        self.assertEqual(self.library.search('Pressure sensor')['total'],1)
        with patch('ocr_service.recognize_page',side_effect=RuntimeError('Engine failure')):
            self.service.start(self.doc_id,['eng']);self.assertIn('Engine failure',wait_job(self.service)['result']['error'])
        self.assertEqual(self.library.search('Pressure sensor')['total'],1)

    def test_changed_source_prevents_commit_and_invalidates_cached_text(self):
        def changed(*args):
            with self.pdf.open('ab') as stream: stream.write(b'\n% changed\n')
            return self.recognized()
        with patch('ocr_service.recognize_page',side_effect=changed):
            self.service.start(self.doc_id,['eng']);self.assertIn('changed',wait_job(self.service)['result']['error'])
        self.assertEqual(self.library.search('Pressure sensor')['total'],0)
        with patch('ocr_service.recognize_page',side_effect=self.recognized):
            self.service.start(self.doc_id,['eng']);wait_job(self.service)
        with self.pdf.open('ab') as stream: stream.write(b'\n% new revision\n')
        self.assertIsNone(self.library.ocr_page(self.doc_id,1))
        self.library.scan();self.assertEqual(self.library.details(self.doc_id)['ocr']['pages'],0)

    def test_validation_and_operation_lock(self):
        for languages,dpi,mode in [([],200,'missing'),(['ara'],200,'missing'),(['eng'],900,'missing'),(['eng'],200,'bad')]:
            with self.assertRaises(ValueError):self.service.start(self.doc_id,languages,dpi,mode)
        self.service.operation_lock.acquire()
        try:
            with self.assertRaises(ValueError):self.service.start(self.doc_id,['eng'])
        finally:self.service.operation_lock.release()
        with self.assertRaises(ValueError):self.service.configure(str(self.folder/'absent'))


class RealOCRTests(LibraryFixture):
    def test_real_scanned_pdf_becomes_searchable_without_source_changes(self):
        folder=os.environ.get('BMEDOCUMENTNAVIGATOR_TESSDATA','')
        if not folder or not (Path(folder)/'eng.traineddata').is_file():
            if os.environ.get('BME_OCR_REQUIRE_TEST')=='1':self.fail('CI requires real English OCR test data')
            self.skipTest('Set BMEDOCUMENTNAVIGATOR_TESSDATA to run real OCR')
        path=self.docs/'real.pdf';make_scan(path);self.library.scan();doc_id=self.first()
        self.assertEqual(self.library.search('pressure sensor')['total'],0)
        before=path.read_bytes();service=OCRService(self.library,threading.Lock());service.configure(folder)
        service.start(doc_id,['eng']);state=wait_job(service)
        self.assertNotIn('error',state['result'])
        self.assertEqual(self.library.search('pressure sensor')['total'],1)
        reader=DocumentReader(self.library)
        self.assertIn('calibration',reader.text(doc_id,1).lower())
        self.assertEqual(reader.render(doc_id,1,query='pressure sensor')[1],1)
        self.assertEqual(path.read_bytes(),before)


class HTTPOCRTests(HTTPFixture):
    def call(self,path,body=None):
        with self.request(path,body) as response:return json.load(response)

    def test_setup_and_endpoints_require_authorization(self):
        for path in ('/api/ocr','/api/ocr/setup'):
            with self.assertRaises(HTTPError) as error:self.request(path,token=False)
            self.assertEqual(error.exception.code,403)
        self.assertFalse(self.call('/api/ocr')['running'])
        with self.assertRaises(HTTPError):self.call('/api/ocr/setup',{'tessdata':str(self.folder/'absent')})

    def test_ocr_start_progress_and_cancel_api(self):
        path=self.docs/'scan.pdf';make_scan(path);self.library.scan()
        doc_id=self.library.search('scan.pdf')['documents'][0]['id']
        models=self.folder/'models';models.mkdir();(models/'eng.traineddata').write_bytes(b'fixture')
        self.assertTrue(self.call('/api/ocr/setup',{'tessdata':str(models)})['available'])
        entered=threading.Event();release=threading.Event()
        def paused(*args):entered.set();release.wait(5);return dict(text='Recognized',ocr=True,words=[])
        with patch('ocr_service.recognize_page',side_effect=paused):
            self.call('/api/ocr',{'id':doc_id,'languages':['eng']});self.assertTrue(entered.wait(5))
            self.assertTrue(self.call('/api/ocr')['running']);self.call('/api/ocr/cancel',{});release.set()
            self.assertTrue(wait_job(self.app.ocr)['result']['cancelled'])
