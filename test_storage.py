import os
from contextlib import closing
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import patch
from app_storage import data_file


class StorageTests(unittest.TestCase):
    def test_database_migration_keeps_original_and_does_not_overwrite(self):
        with tempfile.TemporaryDirectory() as folder:
            legacy = Path(folder) / 'legacy'
            legacy.mkdir()
            source = legacy / 'history.db'
            with closing(sqlite3.connect(source)) as connection, connection:
                connection.execute('create table notes (text)')
                connection.execute("insert into notes values ('keep me')")
            with patch.dict(os.environ, {'TESTAPP_DATA_DIR': str(Path(folder) / 'new')}):
                target = data_file('TestApp', 'history.db', [legacy])
                with closing(sqlite3.connect(target)) as connection, connection:
                    self.assertEqual(connection.execute('select text from notes').fetchone()[0], 'keep me')
                    connection.execute("insert into notes values ('new note')")
                self.assertEqual(data_file('TestApp', 'history.db', [legacy]), target)
                with closing(sqlite3.connect(target)) as connection:
                    self.assertEqual(connection.execute('select count(*) from notes').fetchone()[0], 2)
                self.assertTrue(source.exists())

    def test_launch_folder_does_not_change_storage(self):
        with tempfile.TemporaryDirectory() as folder:
            with patch.dict(os.environ, {'TESTAPP_DATA_DIR': folder}):
                with patch('os.getcwd', return_value='/somewhere'):
                    first = data_file('TestApp', 'settings.ini')
                with patch('os.getcwd', return_value='/elsewhere'):
                    second = data_file('TestApp', 'settings.ini')
                self.assertEqual(first, second)


if __name__ == '__main__':
    unittest.main()
