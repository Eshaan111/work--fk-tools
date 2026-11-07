"""Completion checkpoints and retry behavior, without browser/network access."""
import contextlib
import csv
import io
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import downloader as d
import import_firefox_tabs as importer

A = 'https://www.flipkart.com/shirt/p/itm123?pid=RED'
B = 'https://www.flipkart.com/jeans/p/itm456'


class CompletionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        self.csv = self.folder / 'urls.csv'

    def tearDown(self):
        self.temp.cleanup()

    def rows(self):
        with self.csv.open(encoding='utf-8-sig', newline='') as stream:
            return list(csv.DictReader(stream))

    def pending(self):
        with contextlib.redirect_stdout(io.StringIO()):
            return d.read_urls_csv(self.csv)

    def test_migrate_existing_csv_preserves_cells_and_backs_up(self):
        original = ('\ufeffurl,note\n' + A + ',"keep, this"\n').encode('utf-8')
        self.csv.write_bytes(original)
        d.update_csv(self.csv)
        self.assertEqual(self.rows(), [{'url': A, 'note': 'keep, this', 'already_done': ''}])
        self.assertEqual(self.csv.with_name('urls.csv.bak').read_bytes(), original)
        self.assertTrue(self.csv.read_bytes().startswith(b'\xef\xbb\xbf'))

    def test_any_nonblank_status_skips_and_blank_status_runs(self):
        self.csv.write_text('url,already_done\n' + A + ',anything\n' + B + ',  \n')
        self.assertEqual(self.pending(), [B])

    def test_marking_updates_duplicate_variants_but_not_other_products(self):
        self.csv.write_text('url,note,already_done\n' + A + ',one,\n' + A + '&otracker=share,two,\n' + B + ',three,\n')
        d.mark_url_done(self.csv, A)
        self.assertEqual([row['already_done'] for row in self.rows()], ['yes', 'yes', ''])
        self.assertEqual([row['note'] for row in self.rows()], ['one', 'two', 'three'])
        self.assertEqual(self.pending(), [B])

    def test_completed_duplicate_skips_earlier_blank_row_and_clear_all_resets(self):
        self.csv.write_text('url,already_done\n' + A + ',\n' + A + '&otracker=share,yes\n')
        self.assertEqual(self.pending(), [])
        self.csv.write_text('url,already_done\n' + A + ',\n')
        self.assertEqual(self.pending(), [A])

    def test_imported_links_start_blank_and_existing_status_is_preserved(self):
        self.csv.write_text('url,already_done\n' + A + ',yes\n')
        importer.append_product_urls(self.csv, [A, B])
        self.assertEqual(self.rows(), [{'url': A, 'already_done': 'yes'}, {'url': B, 'already_done': ''}])
        self.assertEqual(self.pending(), [B])

    def test_locked_csv_remains_unchanged(self):
        self.csv.write_text('url,already_done\n' + A + ',\n')
        before = self.csv.read_bytes()
        with patch.object(d.os, 'replace', side_effect=PermissionError('locked')):
            with self.assertRaises(PermissionError):
                d.mark_url_done(self.csv, A)
        self.assertEqual(self.csv.read_bytes(), before)
        self.assertFalse(list(self.folder.glob('.csv-status-*')))

    def test_main_checkpoints_success_and_leaves_failure_pending(self):
        self.csv.write_text('url\n' + A + '\n' + B + '\n')
        playwright = MagicMock()
        with patch('sys.argv', ['downloader.py', '--csv', str(self.csv), '--output', str(self.folder), '--retries', '0']), \
             patch.object(d, 'sync_playwright', return_value=playwright), \
             patch.object(d, 'process_one', side_effect=[3, RuntimeError('partial download')]), \
             contextlib.redirect_stdout(io.StringIO()):
            d.main()
        self.assertEqual(self.rows(), [{'url': A, 'already_done': 'yes'}, {'url': B, 'already_done': ''}])
        self.assertEqual(self.pending(), [B])

    def test_all_done_does_not_launch_browser(self):
        self.csv.write_text('url,already_done\n' + A + ',yes\n')
        with patch('sys.argv', ['downloader.py', '--csv', str(self.csv), '--output', str(self.folder)]), \
             patch.object(d, 'sync_playwright') as browser, contextlib.redirect_stdout(io.StringIO()):
            d.main()
        browser.assert_not_called()

    def test_failures_are_retried_after_other_urls_and_then_marked_done(self):
        self.csv.write_text('url,already_done\n' + A + ',\n' + B + ',\n')
        args = SimpleNamespace(manual=False, retries=2, csv=self.csv, output=self.folder)
        with patch.object(d, 'process_one', side_effect=[RuntimeError('delayed image'), 2, 3]) as process, \
             contextlib.redirect_stdout(io.StringIO()):
            total, pending = d.run_downloads(MagicMock(), MagicMock(), [A, B], args)
        self.assertEqual([call.args[2] for call in process.call_args_list], [A, B, A])
        self.assertEqual([call.args[3] for call in process.call_args_list], [1, 2, 1])
        self.assertEqual((total, pending), (5, 0))
        self.assertEqual([row['already_done'] for row in self.rows()], ['yes', 'yes'])

    def test_retry_limit_leaves_failed_rows_blank(self):
        self.csv.write_text('url,already_done\n' + A + ',\n')
        args = SimpleNamespace(manual=False, retries=1, csv=self.csv, output=self.folder)
        with patch.object(d, 'process_one', side_effect=RuntimeError('still failing')) as process, \
             contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(d.run_downloads(MagicMock(), MagicMock(), [A], args), (0, 1))
        self.assertEqual(process.call_count, 2)
        self.assertEqual(self.rows()[0]['already_done'], '')

    def test_headless_launch_and_manual_conflict(self):
        playwright = MagicMock()
        d.launch_browser(playwright, headless=True)
        self.assertTrue(playwright.chromium.launch.call_args.kwargs['headless'])
        with patch('sys.argv', ['downloader.py', '--headless', '--manual']), \
             contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit) as error:
                d.main()
        self.assertEqual(error.exception.code, 2)

    def test_partial_product_download_raises_after_writing_manifest(self):
        page = MagicMock(url=A)
        args = SimpleNamespace(wait=0, steps=18, output=self.folder, format='jpg')
        with patch.object(d, 'product_name', return_value='product'), \
             patch.object(d, 'open_gallery', return_value=True), \
             patch.object(d, 'collect_gallery', return_value=['image1', 'image2']), \
             patch.object(d, 'download_image', side_effect=['01.jpg', RuntimeError('failed')]), \
             contextlib.redirect_stdout(io.StringIO()):
            with self.assertRaisesRegex(RuntimeError, 'CSV stays pending'):
                d.process_one(MagicMock(), page, A, 1, args)
        self.assertTrue((self.folder / 'product/manifest.json').exists())


if __name__ == '__main__':
    unittest.main()
