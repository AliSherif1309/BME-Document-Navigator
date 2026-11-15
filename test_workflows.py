"""Workflow regressions: preview/apply, stale plans, identity, and saved searches."""
import json
from contextlib import closing
from pathlib import Path
import shutil
import sqlite3
from urllib.error import HTTPError

from navigator_core import Library
from test_navigator import LibraryFixture, HTTPFixture


class WorkflowTests(LibraryFixture):
    def copy_folder(self):
        destination = self.folder / 'new-location'
        # Both roots are named temporary test folders, never the user's library.
        assert destination.resolve().is_relative_to(self.folder.resolve())
        shutil.copytree(self.docs, destination)
        return destination

    def test_saved_searches_persist_with_filters_and_unique_names(self):
        filters = dict(q='E042 pressure', manufacturer='Philips', kind='Service manual', favorites=True)
        search_id = self.library.save_search('Ventilator references', filters)
        reopened = Library(self.library.database)
        saved = reopened.saved_searches()
        self.assertEqual(saved[0]['filters'], filters)
        with self.assertRaises(ValueError): self.library.save_search('ventilator REFERENCES', filters)
        reopened.delete_search(search_id)
        self.assertEqual(self.library.saved_searches(), [])
        with self.assertRaises(ValueError): self.library.save_search('', filters)
        with self.assertRaises(ValueError): self.library.save_search('Invalid', {'favorites':'false'})

    def test_batch_preview_is_read_only_and_applies_only_selected_fields(self):
        self.create()
        self.library.scan()
        doc_id = self.first()
        self.library.update_metadata(doc_id, {'manufacturer':'Philips','device_model':'ABC','keywords':'keep'})
        plan = self.library.preview_batch([doc_id, doc_id], {'manufacturer':'Mindray','device_model':''})
        self.assertEqual(plan['selected'], 1)
        self.assertEqual(self.library.details(doc_id)['manufacturer'], 'Philips')
        result = self.library.apply_batch(plan)
        doc = self.library.details(doc_id)
        self.assertEqual((doc['manufacturer'], doc['device_model'], doc['keywords']), ('Mindray','','keep'))
        with closing(sqlite3.connect(result['backup'])) as backup:
            self.assertEqual(backup.execute('SELECT manufacturer FROM documents').fetchone()[0], 'Philips')
        self.assertEqual(self.library.preview_batch([doc_id], {'manufacturer':'Mindray'})['changes'], [])
        for ids, values in [([], {'status':'Approved'}), ([doc_id], {}), ([doc_id], {'filepath':'bad'}), ([9999], {'status':'Approved'})]:
            with self.assertRaises(ValueError): self.library.preview_batch(ids, values)

    def test_stale_batch_rolls_back_every_document(self):
        self.create('a.txt'); self.create('b.txt')
        self.library.scan()
        ids = [doc['id'] for doc in self.library.search()['documents']]
        plan = self.library.preview_batch(ids, {'status':'Approved'})
        with self.library.connect() as db:
            db.execute('UPDATE documents SET status=? WHERE id=?', ('Needs review',ids[1]))
        with self.assertRaisesRegex(ValueError, 'Metadata changed'):
            self.library.apply_batch(plan)
        self.assertIsNone(self.library.details(ids[0])['status'])
        self.assertEqual(self.library.details(ids[1])['status'], 'Needs review')

    def test_relocation_preserves_ids_notes_favorites_links_and_search(self):
        self.create('a.txt'); self.create('b.txt')
        self.library.scan()
        ids = [doc['id'] for doc in self.library.search()['documents']]
        self.library.add_note(ids[0], 'Preserve my note', 2)
        self.library.toggle_favorite(ids[0])
        with self.library.connect() as db:
            db.execute('INSERT INTO links(source_doc_id,target_doc_id,description) VALUES (?,?,?)', (*ids,'Related SOP'))
        new_root = self.copy_folder()
        self.docs.rename(self.folder / 'original-offline')
        plan = self.library.preview_relocation(str(self.docs),str(new_root))
        self.assertEqual((plan['ready'], plan['blocked']), (2,0))
        self.assertFalse(self.library.details(ids[0])['available'])
        result = self.library.apply_relocation(plan)
        doc = self.library.details(ids[0])
        self.assertTrue(doc['available'])
        self.assertTrue(doc['favorite'])
        self.assertEqual(doc['notes'][0]['page_number'], 1)
        self.assertEqual(self.library.search('E042')['total'], 2)
        self.assertEqual(self.library.scan()['added'], 0)
        self.assertEqual(self.library.paths(), [str(new_root)])
        with self.library.connect() as db:
            self.assertEqual(db.execute('SELECT count(*) FROM links').fetchone()[0], 1)
        self.assertTrue(Path(result['backup']).is_file())

    def test_relocation_blocks_wrong_content_missing_files_and_duplicates(self):
        self.create('a.txt'); self.create('b.txt')
        self.library.scan()
        new_root = self.copy_folder()
        (new_root / 'a.txt').write_text('Different document',encoding='utf-8')
        (new_root / 'b.txt').unlink()
        plan = self.library.preview_relocation(str(self.docs), str(new_root))
        self.assertEqual(plan['blocked'], 2)
        with self.assertRaises(ValueError): self.library.apply_relocation(plan)
        for file in self.docs.iterdir(): shutil.copy2(file, new_root / file.name)
        self.library.add_path(str(new_root)); self.library.scan()
        plan = self.library.preview_relocation(str(self.docs), str(new_root))
        self.assertEqual(plan['blocked'], 2)
        self.assertIn('already indexed', plan['changes'][0]['reason'])

    def test_relocation_detects_post_preview_content_change_and_rolls_back(self):
        self.create('a.txt'); self.create('b.txt')
        self.library.scan()
        new_root = self.copy_folder()
        plan = self.library.preview_relocation(str(self.docs), str(new_root))
        (new_root / 'b.txt').write_text('Changed after preview', encoding='utf-8')
        with self.assertRaisesRegex(ValueError, 'changed since'):
            self.library.apply_relocation(plan)
        self.assertTrue(all(Path(doc['filepath']).parent == self.docs for doc in self.library.search()['documents']))
        self.assertEqual(self.library.paths(), [str(self.docs)])

    def test_relocation_detects_new_library_records_after_preview(self):
        self.create()
        self.library.scan()
        new_root = self.copy_folder()
        plan = self.library.preview_relocation(str(self.docs), str(new_root))
        self.create('new.txt'); self.library.scan()
        with self.assertRaisesRegex(ValueError, 'folder contents changed'):
            self.library.apply_relocation(plan)

    def test_legacy_offline_files_require_verified_identity(self):
        self.create()
        self.library.scan()
        with self.library.connect() as db: db.execute('DELETE FROM document_files')
        new_root = self.copy_folder()
        self.docs.rename(self.folder / 'old-offline')
        plan = self.library.preview_relocation(str(self.docs), str(new_root))
        self.assertEqual(plan['blocked'],1)
        self.assertIn('No recorded fingerprint', plan['changes'][0]['reason'])
        (self.folder / 'old-offline').rename(self.docs)
        self.assertEqual(self.library.scan()['reindexed'],1)
        self.assertEqual(self.library.preview_relocation(str(self.docs), str(new_root))['ready'],1)

    def test_relocation_rejects_nested_roots_and_similar_sibling_names(self):
        self.create(); self.library.scan()
        nested = self.docs / 'nested'; nested.mkdir()
        with self.assertRaises(ValueError): self.library.preview_relocation(str(self.docs), str(nested))
        sibling = self.folder / 'Philips-other'; sibling.mkdir()
        self.library.add_path(str(sibling))
        (sibling / 'keep.txt').write_text('Keep separate', encoding='utf-8')
        self.library.scan()
        plan = self.library.preview_relocation(str(self.docs),str(self.copy_folder()))
        self.assertEqual(len(plan['changes']),1)


class HTTPWorkflowTests(HTTPFixture):
    def call(self, path, body=None):
        with self.request(path, body) as response: return json.load(response)

    def test_saved_search_round_trip_and_delete(self):
        filters = dict(q='oxygen', manufacturer='', kind='', favorites=False)
        saved = self.call('/api/searches', {'name':'Oxygen', 'filters':filters})
        self.assertEqual(self.call('/api/searches')['searches'][0]['filters'], filters)
        self.call('/api/searches/delete',saved)
        self.assertEqual(self.call('/api/searches')['searches'], [])

    def test_batch_requires_server_preview_and_is_single_use(self):
        payload = {'ids':[self.first()],'values':{'status':'Approved'}}
        preview = self.call('/api/batch/preview', payload)
        with self.assertRaises(HTTPError): self.call('/api/batch/apply',payload)
        with self.assertRaises(HTTPError): self.call('/api/relocation/apply',{'preview_id':preview['preview_id']})
        result = self.call('/api/batch/apply',{'preview_id':preview['preview_id'], 'values':{'status':'Forged'}})
        self.assertEqual(result['updated'],1)
        self.assertEqual(self.library.details(self.first())['status'],'Approved')
        with self.assertRaises(HTTPError): self.call('/api/batch/apply',{'preview_id':preview['preview_id']})

    def test_preview_expiration_scan_lock_and_origin(self):
        payload = {'ids':[self.first()],'values':{'status':'Approved'}}
        preview = self.call('/api/batch/preview',payload)
        self.app.previews[preview['preview_id']]['expires']=0
        with self.assertRaises(HTTPError): self.call('/api/batch/apply',{'preview_id':preview['preview_id']})
        self.app.lock.acquire()
        try:
            with self.assertRaises(HTTPError) as error: self.call('/api/batch/preview',payload)
            self.assertEqual(error.exception.code,400)
        finally: self.app.lock.release()
        with self.assertRaises(HTTPError) as error:
            self.request('/api/searches',{'name':'Bad','filters':{}},headers={'Origin':'https://example.com'})
        self.assertEqual(error.exception.code,403)

    def test_relocation_http_round_trip(self):
        destination = self.folder / 'new-location'
        shutil.copytree(self.docs, destination)
        preview = self.call('/api/relocation/preview',{'old_folder':str(self.docs),'new_folder':str(destination)})
        self.assertEqual(preview['ready'],1)
        self.assertEqual(self.call('/api/relocation/apply',{'preview_id':preview['preview_id']})['updated'],1)
        self.assertEqual(Path(self.library.details(self.first())['filepath']).parent,destination)
