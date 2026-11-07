"""Firefox import tests use synthetic sessions, never the user's browser data."""
import csv
import contextlib
import io
import json
import os
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import lz4.block

import import_firefox_tabs as importer

A = 'https://www.flipkart.com/shirt/p/itm123?pid=RED'
B = 'https://www.flipkart.com/jeans/p/itm456'


class FirefoxImportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.csv = self.root / 'urls.csv'

    def tearDown(self):
        self.temp.cleanup()

    def test_compressed_and_plain_session(self):
        session = {'windows': [{'tabs': [{'entries': [{'url': A}], 'index': 1}]}]}
        raw = json.dumps(session).encode()
        for filename, data in [('session.json', raw),
                               ('recovery.jsonlz4', b'mozLz40\0' + lz4.block.compress(raw))]:
            path = self.root / filename
            path.write_bytes(data)
            self.assertEqual(importer.read_session(path), session)
        path.write_bytes(b'mozLz40\0broken')
        with self.assertRaises(lz4.block.LZ4BlockError):
            importer.read_session(path)

    def test_only_selected_open_tabs_across_windows(self):
        session = {'windows': [
            {'tabs': [{'entries': [{'url': A}, {'url': 'https://example.com'}], 'index': 1},
                      {'entries': [{'url': 'https://example.com'}, {'url': B}], 'index': 2},
                      {'entries': [{'url': 'private'}], 'index': 1, 'isPrivate': True},
                      {'entries': [{'url': 'bad-index'}], 'index': 0}],
             '_closedTabs': [{'entries': [{'url': 'closed'}]}]},
            {'tabs': [{'entries': [{'url': 'other-window'}]}]},
            {'isPrivate': True, 'tabs': [{'entries': [{'url': 'private-window'}]}]},
        ], '_closedWindows': [{'tabs': [{'entries': [{'url': 'closed-window'}]}]}]}
        self.assertEqual(list(importer.tab_urls(session)), [A, B, 'other-window'])

    def test_append_keeps_original_bytes_columns_and_backup(self):
        original = ('\ufeff Product_URL ,note\r\n' + A + ',keep me').encode('utf-8')
        self.csv.write_bytes(original)
        stats = importer.append_product_urls(self.csv, [A + '&otracker=search', B, B, 'https://example.com'])
        self.assertEqual(stats, {'added': 1, 'duplicates': 2, 'skipped': 1})
        self.assertTrue(self.csv.read_bytes().startswith(original + b'\r\n'))
        self.assertEqual(self.csv.with_name('urls.csv.bak').read_bytes(), original)
        with self.csv.open(encoding='utf-8-sig', newline='') as stream:
            rows = list(csv.DictReader(stream))
        self.assertEqual(rows, [{' Product_URL ': A, 'note': 'keep me'},
                                {' Product_URL ': B, 'note': ''}])
        before = self.csv.read_bytes()
        self.assertEqual(importer.append_product_urls(self.csv, [A, B])['added'], 0)
        self.assertEqual(self.csv.read_bytes(), before)

    def test_dry_run_never_writes_csv_or_backup(self):
        self.assertEqual(importer.append_product_urls(self.csv, [A], dry_run=True)['added'], 1)
        self.assertFalse(self.csv.exists())
        self.csv.write_text('url\n' + A + '\n', encoding='utf-8')
        before = self.csv.read_bytes()
        importer.append_product_urls(self.csv, [B], dry_run=True)
        self.assertEqual(self.csv.read_bytes(), before)
        self.assertFalse(self.csv.with_name('urls.csv.bak').exists())

    def test_new_csv_and_product_variants(self):
        blue = A.replace('RED', 'BLUE')
        stats = importer.append_product_urls(self.csv, [A, blue, 'about:blank'])
        self.assertEqual(stats, {'added': 2, 'duplicates': 0, 'skipped': 1})
        with self.csv.open(newline='', encoding='utf-8') as stream:
            self.assertEqual(list(csv.DictReader(stream)), [{'url': A, 'already_done': ''},
                                                          {'url': blue, 'already_done': ''}])

    def test_invalid_existing_csv_and_failed_replace_preserve_file(self):
        self.csv.write_text('wrong\nkeep\n', encoding='utf-8')
        before = self.csv.read_bytes()
        with self.assertRaises(ValueError):
            importer.append_product_urls(self.csv, [A])
        self.assertEqual(self.csv.read_bytes(), before)
        self.csv.write_text('url\n' + A + '\n', encoding='utf-8')
        before = self.csv.read_bytes()
        with patch.object(importer.os, 'replace', side_effect=PermissionError('File locked')):
            with self.assertRaises(PermissionError):
                importer.append_product_urls(self.csv, [B])
        self.assertEqual(self.csv.read_bytes(), before)
        self.assertFalse(list(self.root.glob('.firefox-import-*')))

    def test_profile_discovery_selection_and_latest_snapshot(self):
        self.root.joinpath('profiles.ini').write_text(
            '[Profile0]\nName=Main\nIsRelative=1\nPath=Profiles/main\n'
            '[InstallABC]\nDefault=Profiles/main\n'
            '[Profile1]\nName=Work\nIsRelative=1\nPath=Profiles/work\n', encoding='utf-8')
        profiles = importer.firefox_profiles(self.root)
        self.assertEqual(len(profiles), 2)
        for i, profile in enumerate(profiles):
            path = profile / 'sessionstore-backups/recovery.jsonlz4'
            path.parent.mkdir(parents=True)
            path.write_bytes(b'fixture')
            os.utime(path, (100 + i, 100 + i))
        self.assertEqual(importer.choose_session(profiles)[1], 'Work')
        self.assertEqual(importer.choose_session(profiles, 'Main')[1], 'Main')
        with self.assertRaises(ValueError):
            importer.choose_session(profiles, 'missing-profile')
        with self.assertRaises(ValueError):
            importer.choose_session({})

    def test_profile_prompt_retries_invalid_input_and_uses_chosen_number(self):
        main, work = self.root / 'main', self.root / 'work'
        output = io.StringIO()
        with patch('builtins.input', side_effect=['', '0', '3', 'text', '2']), contextlib.redirect_stdout(output):
            self.assertEqual(importer.prompt_profile({main: 'Main', work: 'Work'}), work)
        self.assertIn('1. Main', output.getvalue())
        self.assertIn('2. Work', output.getvalue())
        self.assertEqual(output.getvalue().count('Enter a number'), 4)

    def test_selectable_profiles_within_group_are_named_and_imported_separately(self):
        self.root.joinpath('profiles.ini').write_text(
            '[Profile0]\nName=ESHAAN MAIN\nIsRelative=1\nPath=Profiles/es\nStoreID=abc123\n'
            '[InstallABC]\nDefault=Profiles/es\n', encoding='utf-8')
        folder = self.root / 'Profile Groups'
        folder.mkdir()
        database = folder / 'abc123.sqlite'
        with sqlite3.connect(database) as connection:
            connection.execute('CREATE TABLE Profiles (id INTEGER PRIMARY KEY, path TEXT, name TEXT)')
            connection.executemany('INSERT INTO Profiles VALUES (?, ?, ?)', [
                (1, 'Profiles/pr', 'Prabhu'), (2, 'Profiles/es', 'Eshaan'), (3, 'Profiles/se', 'Seema')])
        connection.close()
        before = database.read_bytes()
        profiles = importer.firefox_profiles(self.root)
        self.assertEqual(list(profiles.values()), [
            'ESHAAN MAIN / Prabhu', 'ESHAAN MAIN / Eshaan', 'ESHAAN MAIN / Seema'])
        self.assertEqual(database.read_bytes(), before)
        for path in profiles:
            path.mkdir(parents=True)
            (path / 'sessionstore.jsonlz4').write_bytes(b'fixture')
        self.assertEqual(importer.choose_session(profiles, 'Prabhu')[0].parent, self.root / 'Profiles/pr')
        self.assertEqual(importer.choose_session(profiles, 'ESHAAN MAIN / Seema')[0].parent,
                         self.root / 'Profiles/se')

    def test_unreadable_group_keeps_legacy_profile_and_warns(self):
        self.root.joinpath('profiles.ini').write_text(
            '[Profile0]\nName=Main\nPath=Profiles/main\nStoreID=abc123\n', encoding='utf-8')
        output = io.StringIO()
        with patch.object(importer, 'grouped_profiles', side_effect=sqlite3.OperationalError('locked')), \
             contextlib.redirect_stdout(output):
            profiles = importer.firefox_profiles(self.root)
        self.assertEqual(profiles, {(self.root / 'Profiles/main').resolve(): 'Main'})
        self.assertIn('Warning:', output.getvalue())

    def test_duplicate_person_names_require_full_group_name(self):
        profiles = {self.root / 'a': 'Main / Eshaan', self.root / 'b': 'Other / Eshaan'}
        with self.assertRaisesRegex(ValueError, 'ambiguous'):
            importer.choose_session(profiles, 'Eshaan')

    def test_profile_prompt_can_cancel_and_handles_closed_input(self):
        for response in ('q', EOFError(), KeyboardInterrupt()):
            with self.subTest(response=response), patch('builtins.input', side_effect=[response]), contextlib.redirect_stdout(io.StringIO()):
                self.assertIsNone(importer.prompt_profile({self.root: 'Main'}))
        with self.assertRaises(ValueError):
            importer.prompt_profile({})

    def test_main_imports_selected_profile_without_using_newest_profile(self):
        main, work = self.root / 'main', self.root / 'work'
        source = main / 'sessionstore.jsonlz4'
        source.parent.mkdir()
        source.write_bytes(b'fixture')
        with patch('sys.argv', ['import_firefox_tabs.py', '--csv', str(self.csv)]), \
             patch.object(importer, 'firefox_profiles', return_value={main: 'Main', work: 'Work'}), \
             patch('builtins.input', return_value='1'), \
             patch.object(importer, 'choose_session', return_value=(source, 'Main')) as choose, \
             patch.object(importer, 'read_session', return_value={'windows': []}), \
             patch.object(importer, 'append_product_urls', return_value={'added': 0, 'duplicates': 0, 'skipped': 0}), \
             contextlib.redirect_stdout(io.StringIO()):
            importer.main()
        choose.assert_called_once_with({main: 'Main', work: 'Work'}, str(main))

    def test_explicit_profile_bypasses_prompt(self):
        source = self.root / 'sessionstore.jsonlz4'
        source.write_bytes(b'fixture')
        with patch('sys.argv', ['import_firefox_tabs.py', '--profile', 'Main', '--dry-run']), \
             patch.object(importer, 'firefox_profiles', return_value={self.root: 'Main'}), \
             patch.object(importer, 'prompt_profile') as prompt, \
             patch.object(importer, 'choose_session', return_value=(source, 'Main')), \
             patch.object(importer, 'read_session', return_value={'windows': []}), \
             patch.object(importer, 'append_product_urls', return_value={'added': 0, 'duplicates': 0, 'skipped': 0}), \
             contextlib.redirect_stdout(io.StringIO()):
            importer.main()
        prompt.assert_not_called()


if __name__ == '__main__':
    unittest.main()
