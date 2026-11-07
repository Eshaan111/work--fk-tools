"""CSV URL validation tests; no browser or network required."""
import contextlib
import csv
import io
import tempfile
import unittest
from pathlib import Path

import downloader as d


class CsvUrlTests(unittest.TestCase):
    def read(self, urls):
        output = io.StringIO()
        with tempfile.TemporaryDirectory() as folder:
            file = Path(folder) / 'urls.csv'
            with file.open('w', encoding='utf-8-sig', newline='') as stream:
                writer = csv.writer(stream)
                writer.writerow(['url'])
                writer.writerows([url] for url in urls)
            with contextlib.redirect_stdout(output):
                result = d.read_urls_csv(file)
        return result, output.getvalue()

    def test_duplicates_and_tracking_variants_keep_first_url_and_order(self):
        first = 'https://www.flipkart.com/shirt/p/itm123?pid=ABC&otracker=search'
        second = 'https://www.flipkart.com/jeans/p/itm456'
        urls, log = self.read([
            first, first, second,
            'http://m.flipkart.com/renamed-shirt/p/itm123/?pid=ABC&lid=other#photo',
            'https://flipkart.com/jeans/p/itm456?utm_source=share',
        ])
        self.assertEqual(urls, [first, second])
        self.assertEqual(log.count('Skipping duplicate URL'), 3)
        self.assertIn('CSV row 3', log)

    def test_different_product_variants_are_preserved(self):
        variants = ['https://www.flipkart.com/shirt/p/itm123?pid=RED',
                    'https://www.flipkart.com/shirt/p/itm123?pid=BLUE']
        self.assertEqual(self.read(variants)[0], variants)

    def test_invalid_urls_are_skipped_without_stopping_csv(self):
        invalid = [
            'not-a-url', 'https://www.flipkart.com/search?q=shirt',
            'https://flipkart.com.evil.test/shirt/p/itm123',
            'https://www.flipkart.com/p/', 'https://www.flipkart.com/p/itm 123',
            'https://www.flipkart.com:bad/p/itm123',
            'https://www.flipkart.com:99999/p/itm123',
            'https://www.flipkart.com:8080/p/itm123',
            'https://user:pass@www.flipkart.com/p/itm123',
            'https://[invalid/p/itm123', 'ftp://www.flipkart.com/p/itm123',
            'https://www.flipkart.com\\evil.test/p/itm123',
        ]
        valid = 'https://www.flipkart.com/shirt/p/itm123'
        urls, log = self.read([*invalid, '', '  '+valid+'  '])
        self.assertEqual(urls, [valid])
        self.assertEqual(log.count('Skipping CSV row'), len(invalid))

    def test_pidless_links_ignore_fragment_and_default_port(self):
        urls = ['https://www.flipkart.com/shirt/p/itm123',
                'https://flipkart.com:443/shirt/p/itm123/#images']
        self.assertEqual(self.read(urls)[0], urls[:1])


if __name__ == '__main__':
    unittest.main()
