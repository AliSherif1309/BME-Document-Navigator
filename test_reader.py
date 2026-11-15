"""PDF reader, page references, and related-document regressions."""
from concurrent.futures import ThreadPoolExecutor
import hashlib
import json
import shutil
import struct
from urllib.error import HTTPError

import pymupdf

from document_reader import DocumentReader
from navigator_core import Library
from test_navigator import LibraryFixture, HTTPFixture


def make_pdf(path, pages=3, password=None, large=False):
    with pymupdf.open() as pdf:
        for index in range(pages):
            page = pdf.new_page(width=4000 if large else 595, height=4000 if large else 842)
            page.insert_text((72,72), f'SAMPLE reference page {index+1}', fontsize=18)
            page.insert_text((72,110), 'E042 pressure sensor calibration', fontsize=12)
        pdf.set_toc([[1, 'Introduction', 1]] + ([[1,'Calibration',2]] if pages>1 else []))
        options = dict(encryption=pymupdf.PDF_ENCRYPT_AES_256, owner_pw='owner-secret', user_pw=password) if password else {}
        pdf.save(path, **options)


class ReaderTests(LibraryFixture):
    def setUp(self):
        super().setUp()
        self.pdf = self.docs / 'service.pdf'
        make_pdf(self.pdf)
        self.library.scan()
        self.doc_id = self.first()
        self.reader = DocumentReader(self.library)

    def test_outline_highlight_text_and_source_are_read_only(self):
        before = hashlib.sha256(self.pdf.read_bytes()).hexdigest()
        info = self.reader.info(self.doc_id)
        self.assertEqual(info['page_count'],3)
        self.assertEqual(info['outline'][1]['page'],2)
        plain, zero = self.reader.render(self.doc_id,2)
        highlighted, count = self.reader.render(self.doc_id,2,query='pressure sensor')
        self.assertEqual((zero,count),(0,1))
        self.assertTrue(plain.startswith(b'\x89PNG\r\n\x1a\n'))
        self.assertNotEqual(plain,highlighted)
        self.assertIn('reference page 2',self.reader.text(self.doc_id,2))
        self.assertEqual(hashlib.sha256(self.pdf.read_bytes()).hexdigest(),before)

    def test_saved_position_and_bookmarks_survive_reopening_and_relocation(self):
        self.reader.save_position(self.doc_id,3,2)
        bookmark_id = self.reader.add_bookmark(self.doc_id,2,'Calibration reference')
        reopened = DocumentReader(Library(self.library.database))
        self.assertEqual((reopened.info(self.doc_id)['page'],reopened.info(self.doc_id)['zoom']),(3,2))
        new = self.folder / 'moved-documents'
        shutil.copytree(self.docs,new)
        self.library.apply_relocation(self.library.preview_relocation(str(self.docs),str(new)))
        info = reopened.info(self.doc_id)
        self.assertEqual(info['page'],3)
        self.assertEqual(info['bookmarks'][0]['fav_id'],bookmark_id)
        self.assertEqual(info['bookmarks'][0]['page_number'],1)

    def test_document_star_never_deletes_page_bookmarks(self):
        bookmark_id = self.reader.add_bookmark(self.doc_id,2,'Keep this page')
        self.assertFalse(self.library.details(self.doc_id)['favorite'])
        self.library.toggle_favorite(self.doc_id)
        with self.library.connect() as db:
            favorite_id = db.execute('SELECT fav_id FROM favorites WHERE page_number=0').fetchone()[0]
        self.library.remove_bookmark(favorite_id)
        self.assertTrue(self.library.details(self.doc_id)['favorite'])
        self.library.toggle_favorite(self.doc_id)
        self.assertEqual(self.library.bookmarks()[0]['fav_id'],bookmark_id)
        self.assertEqual(self.library.search(favorites=True)['total'],1)
        self.assertFalse(self.library.search()['documents'][0]['favorite'])
        self.library.remove_bookmark(bookmark_id)
        self.assertEqual(self.library.bookmarks(),[])

    def test_invalid_pages_zoom_formats_and_offline_files(self):
        for page in (0,-1,4):
            with self.assertRaises(ValueError): self.reader.render(self.doc_id,page)
            with self.assertRaises(ValueError): self.reader.save_position(self.doc_id,page,1)
            with self.assertRaises(ValueError): self.reader.add_bookmark(self.doc_id,page,'Invalid page')
        for zoom in (0,4,float('nan'),float('inf')):
            with self.assertRaises(ValueError): self.reader.render(self.doc_id,1,zoom)
        with self.assertRaises(ValueError): self.reader.render(self.doc_id,1,query='x'*201)
        self.create('text.txt'); self.library.scan()
        text_id = self.library.search('text.txt')['documents'][0]['id']
        with self.assertRaisesRegex(ValueError,'supports PDF'): self.reader.info(text_id)
        self.pdf.unlink()
        with self.assertRaisesRegex(ValueError,'offline'): self.reader.info(self.doc_id)

    def test_protected_pdf_has_actionable_error(self):
        protected = self.docs / 'protected.pdf'
        make_pdf(protected,password='secret')
        with self.library.connect() as db:
            doc_id = db.execute('INSERT INTO documents(filename,filepath,last_modified) VALUES (?,?,?)',
                (protected.name,str(protected),protected.stat().st_mtime)).lastrowid
        with self.assertRaisesRegex(ValueError,'password-protected'): self.reader.info(doc_id)

    def test_rendering_limits_and_concurrent_reads(self):
        large = self.docs / 'large.pdf'; make_pdf(large,pages=1,large=True)
        self.library.scan()
        doc_id = self.library.search('large.pdf')['documents'][0]['id']
        png,_ = self.reader.render(doc_id,1,3)
        width,height = struct.unpack('>II',png[16:24])
        self.assertLessEqual(width,6000); self.assertLessEqual(height,6000)
        self.assertLessEqual(width*height,12010000)
        with ThreadPoolExecutor(max_workers=3) as pool:
            futures = [pool.submit(self.reader.render,self.doc_id,page) for page in (1,2,3)]
            self.assertTrue(all(f.result()[0].startswith(b'\x89PNG') for f in futures))

    def test_links_are_bidirectional_and_reject_duplicates(self):
        self.create('related_sop.txt'); self.create('unrelated.txt'); self.library.scan()
        target = self.library.search('related_sop')['documents'][0]['id']
        unrelated = self.library.search('unrelated')['documents'][0]['id']
        self.library.add_link(self.doc_id,target,'Calibration SOP')
        forward = self.library.links(self.doc_id)[0]
        reverse = self.library.links(target)[0]
        self.assertEqual(forward['doc_id'],target)
        self.assertEqual(reverse['doc_id'],self.doc_id)
        self.assertEqual(forward['description'],'Calibration SOP')
        for source,destination in [(target,self.doc_id),(self.doc_id,self.doc_id),(self.doc_id,999999)]:
            with self.assertRaises(ValueError): self.library.add_link(source,destination)
        self.library.remove_link(unrelated,forward['link_id'])
        self.assertEqual(len(self.library.links(target)),1)
        self.library.remove_link(target,forward['link_id'])
        self.assertEqual(self.library.links(self.doc_id),[])


class HTTPReaderTests(HTTPFixture):
    def setUp(self):
        super().setUp()
        make_pdf(self.docs / 'reader.pdf')
        self.library.scan()
        self.doc_id = self.library.search('reader.pdf')['documents'][0]['id']

    def call(self,path,body=None):
        with self.request(path,body) as response: return json.load(response)

    def test_render_headers_text_position_and_reader_asset(self):
        with self.request('/api/reader/page?id='+str(self.doc_id)+'&page=2&q=pressure') as response:
            self.assertEqual(response.headers['Content-Type'],'image/png')
            self.assertEqual(response.headers['X-Match-Count'],'1')
            self.assertTrue(response.read().startswith(b'\x89PNG'))
        self.assertIn('page 2',self.call('/api/reader/text?id='+str(self.doc_id)+'&page=2')['text'])
        self.call('/api/reader/position',{'id':self.doc_id,'page':2,'zoom':1.5})
        self.assertEqual(self.call('/api/reader?id='+str(self.doc_id))['page'],2)
        with self.request('/reader.js',token=False) as response:
            self.assertIn(b'openReader',response.read())

    def test_bookmark_and_link_round_trips(self):
        bookmark = self.call('/api/bookmarks',{'id':self.doc_id,'page':3,'name':'HTTP page reference'})
        self.assertEqual(self.call('/api/bookmarks?id='+str(self.doc_id))['bookmarks'][0]['page_number'],2)
        self.call('/api/bookmarks/delete',{'bookmark_id':bookmark['id']})
        self.assertEqual(self.call('/api/bookmarks')['bookmarks'],[])
        target = self.library.search('oxygen')['documents'][0]['id']
        self.call('/api/links',{'id':self.doc_id,'target_id':target,'description':'Related reference'})
        link = self.call('/api/links?id='+str(target))['links'][0]
        self.assertEqual(link['doc_id'],self.doc_id)
        self.call('/api/links/delete',{'id':self.doc_id,'link_id':link['link_id']})
        self.assertEqual(self.call('/api/links?id='+str(target))['links'],[])

    def test_reader_boundaries_require_token_and_validate_input(self):
        for path in ('/api/reader','/api/reader/page','/api/reader/text'):
            with self.assertRaises(HTTPError) as error: self.request(path+'?id='+str(self.doc_id),token=False)
            self.assertEqual(error.exception.code,403)
        for tail in ('&page=0','&page=999','&zoom=nan','&zoom=50'):
            with self.assertRaises(HTTPError) as error: self.request('/api/reader/page?id='+str(self.doc_id)+tail)
            self.assertEqual(error.exception.code,400)
