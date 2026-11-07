"""Offline regression tests using excerpts of the user's two saved pages."""
import io
import json
import tempfile
import threading
import time
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from PIL import Image
from playwright.sync_api import sync_playwright
import downloader as d

FIXTURES = Path(__file__).parent / 'fixtures'
CDN = 'https://rukminim2.flixcart.com/image/'


def image_bytes(size=(240, 320)):
    buffer = io.BytesIO()
    Image.new('RGBA', size, (40, 80, 120, 128)).save(buffer, 'PNG')
    return buffer.getvalue()


class Handler(BaseHTTPRequestHandler):
    requests = []

    def do_GET(self):
        self.requests.append((self.path, self.headers.get('Referer')))
        if self.path.startswith('/image/1200/1200/fallback'):
            self.send_error(404)
            return
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b'<html>blocked</html>' if 'bad' in self.path else
                         image_bytes((40, 40) if 'small' in self.path else (240, 320)))

    def log_message(self, *args):
        pass


class DownloaderTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.playwright = sync_playwright().start()
        cls.browser = cls.playwright.chromium.launch(headless=True)
        cls.server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.base = f'http://127.0.0.1:{cls.server.server_port}'

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join()
        cls.browser.close()
        cls.playwright.stop()

    def setUp(self):
        self.context = self.browser.new_context(viewport={'width': 1536, 'height': 850})
        self.context.route('**/*', lambda route: route.fulfill(
            status=200, content_type='image/png', body=image_bytes()))
        self.page = self.context.new_page()
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)

    def tearDown(self):
        self.context.close()
        self.temp.cleanup()

    def load(self, listing=False, lazy=False):
        gallery = (FIXTURES / 'gallery.html').read_text(encoding='utf-8')
        css = '''<style>
          .TT5KcC{position:fixed;inset:0;background:white;z-index:99}
          .S0f9kM{position:relative}.ThdfW9{overflow:hidden}
          .hKE1H7{display:flex}.P56aSK{flex-shrink:0}
          .y7Pyz6{display:flex;align-items:center;justify-content:center}
          .Xz2QEH{width:400px;height:600px}
          .Vi7peE,.HbzLrW{position:fixed;top:300px;width:32px;height:62px;z-index:100;background:gray}
          .Vi7peE{right:0}.HbzLrW{left:0}
          .OfydJ4{aspect-ratio:3/4;width:100%}
          .ARoD00{position:relative}.eWJFnV{position:absolute;inset:0}
        </style>'''
        content = (FIXTURES / 'listing.html').read_text(encoding='utf-8') if listing else gallery
        self.page.set_content(css + content, wait_until='domcontentloaded')
        self.page.evaluate('''({gallery, listing, lazy, cdn}) => {
          window.videoClicks=0;
          function wire() {
            const root=document.querySelector('.TT5KcC');
            const imgs=[...root.querySelectorAll('img.Xz2QEH')];
            window.slideIndex=4;
            function render() {
              root.querySelector('.hKE1H7').style.transform=`translate3d(${-1536*slideIndex}px,0,0)`;
              root.querySelectorAll('.PylJiO').forEach((el,i)=>el.classList.toggle('Y1JNr7', i===slideIndex));
              imgs.forEach((el,i)=>{
                if (!lazy || i===slideIndex) el.src=cdn+'832/832/asset'+i+'.png';
                else el.removeAttribute('src');
              });
              root.querySelector('.HbzLrW').classList.toggle('B_lGg3',slideIndex===0);
              root.querySelector('.Vi7peE').classList.toggle('K5mlx2',slideIndex===5);
            }
            root.querySelector('.HbzLrW').onclick=()=>{slideIndex=Math.max(0,slideIndex-1);render()};
            root.querySelector('.Vi7peE').onclick=()=>{slideIndex=Math.min(5,slideIndex+1);render()};
            render();
          }
          if (listing) {
            document.querySelector('.ARoD00').onclick=()=>videoClicks++;
            document.querySelector('.OfydJ4').parentElement.onclick=()=>setTimeout(()=>{
              document.body.insertAdjacentHTML('beforeend',gallery);wire();
            },250);
          } else wire();
        }''', {'gallery': gallery, 'listing': listing, 'lazy': lazy, 'cdn': CDN})

    def test_actual_photo_card_opens_delayed_gallery_without_clicking_video(self):
        self.load(listing=True)
        self.assertFalse(d.gallery_is_open(self.page))
        self.assertEqual(d.urls_from_page(self.page), [])
        self.assertTrue(d.open_gallery(self.page))
        self.assertEqual(self.page.evaluate('videoClicks'), 0)
        self.assertEqual(len(d.urls_from_page(self.page)), 5)

    def test_lazy_gallery_navigates_both_directions_from_middle(self):
        self.load(lazy=True)
        urls = d.collect_gallery(self.page, 18)
        self.assertEqual({d.cdn_identity(u) for u in urls},
                         {f'/image/asset{i}.png' for i in (0, 2, 3, 4, 5)})
        self.assertEqual(self.page.evaluate('slideIndex'), 5)

    def test_extract_all_srcset_sizes_and_exclude_recommendations_and_video(self):
        self.load()
        self.page.locator(d.GALLERY).first.evaluate('''el => {
          el.src='https://rukminim2.flixcart.com/image/128/128/asset0.png?q=20';
          el.srcset='https://rukminim2.flixcart.com/image/1600/1600/asset0.png 2x, https://rukminim2.flixcart.com/image/400/400/asset0.png 1x';
          document.body.insertAdjacentHTML('beforeend','<img alt="Image" src="https://rukminim2.flixcart.com/image/832/832/recommendation.png">');
        }''')
        urls = d.collect_gallery(self.page, 1)
        self.assertEqual(len(urls), 5)
        self.assertFalse(any('asset1' in u or 'recommendation' in u for u in urls))
        self.assertIn(CDN+'1600/1600/asset0.png', urls)

    def test_hidden_first_gallery_image_does_not_hide_viewer(self):
        self.load()
        self.page.locator(d.GALLERY).first.evaluate("el => el.style.display='none'")
        self.assertTrue(d.gallery_is_open(self.page))
        self.assertTrue(d.open_gallery(self.page))

    def test_gallery_is_detected_before_images_have_layout(self):
        self.load()
        self.page.locator(d.GALLERY).evaluate_all("els => els.forEach(el => {el.style.width='0px';el.style.height='0px'})")
        self.assertEqual(self.page.locator(f'{d.GALLERY}:visible').count(), 0)
        self.assertTrue(d.gallery_is_open(self.page))
        self.assertEqual(len(d.wait_for_image_urls(self.page, timeout=500)), 5)

    def test_closed_hidden_viewer_is_not_gallery_success(self):
        self.load()
        self.page.locator(d.VIEWER).evaluate("el => el.style.display='none'")
        self.assertFalse(d.gallery_is_open(self.page))

    def test_gallery_open_works_when_animation_frames_stop(self):
        self.load(listing=True)
        self.page.evaluate("""() => {
          Object.defineProperty(document, 'hidden', {value: true, configurable: true});
          window.requestAnimationFrame = () => 0;
        }""")
        self.assertTrue(d.open_gallery(self.page, timeout=2000))
        self.assertEqual(self.page.evaluate('videoClicks'), 0)

    def test_image_source_wait_has_timeout_and_accepts_delayed_urls(self):
        self.load()
        self.page.locator(d.GALLERY).evaluate_all("els => els.forEach(el => el.removeAttribute('src'))")
        with self.assertRaisesRegex(RuntimeError, 'no photo URLs'):
            d.wait_for_image_urls(self.page, timeout=150)
        self.page.evaluate("""cdn => setTimeout(() => {
          document.querySelector('img.Xz2QEH').src = cdn + '832/832/late.png';
        }, 200)""", CDN)
        self.assertEqual(d.wait_for_image_urls(self.page, timeout=1000), [CDN+'832/832/late.png'])

    def test_open_gallery_stops_at_deadline(self):
        self.load(listing=True)
        self.page.locator('.OfydJ4').evaluate('(el) => el.parentElement.onclick = null')
        start = time.monotonic()
        self.assertFalse(d.open_gallery(self.page, timeout=350))
        self.assertLess(time.monotonic() - start, 2)

    def test_real_minimized_window_opens_and_navigates_gallery(self):
        browser = d.launch_browser(self.playwright, headless=False)
        context = browser.new_context(viewport={'width': 1536, 'height': 850})
        context.route('**/*', lambda route: route.fulfill(status=200, content_type='image/png', body=image_bytes()))
        original = self.page
        try:
            self.page = context.new_page()
            self.load(listing=True, lazy=True)
            cdp = context.new_cdp_session(self.page)
            window = cdp.send('Browser.getWindowForTarget')
            cdp.send('Browser.setWindowBounds', {'windowId': window['windowId'], 'bounds': {'windowState': 'minimized'}})
            self.page.wait_for_timeout(200)
            self.assertEqual(cdp.send('Browser.getWindowBounds', {'windowId': window['windowId']})['bounds']['windowState'], 'minimized')
            self.assertTrue(d.open_gallery(self.page, timeout=5000))
            urls = d.collect_gallery(self.page, 18)
            self.assertEqual({d.cdn_identity(url) for url in urls},
                             {f'/image/asset{i}.png' for i in (0, 2, 3, 4, 5)})
        finally:
            self.page = original
            context.close()
            browser.close()

    def test_cdn_identity_and_quality(self):
        self.assertEqual(d.cdn_identity(CDN+'128/128/xif0q/a.jpg?q=20'),
                         d.cdn_identity(CDN+'1600/1600/xif0q/a.jpg?q=90'))
        self.assertEqual(d.cdn_identity(CDN+'original/xif0q/a.jpg'), '/image/xif0q/a.jpg')
        self.assertEqual(d.high_res(CDN+'1600/1800/a.jpg'), CDN+'1600/1800/a.jpg')
        self.assertGreater(d.image_quality(CDN+'1600/1600/a.jpg'), d.image_quality(CDN+'128/128/a.jpg?long=query'))

    def test_real_http_download_fallback_and_jpeg_conversion(self):
        Handler.requests.clear()
        path = d.download_image(self.context, self.base+'/image/832/832/fallback.png', self.folder, 1, 'jpg', 'https://www.flipkart.com/product')
        with Image.open(path) as img:
            self.assertEqual((img.format, img.mode, img.size), ('JPEG', 'RGB', (240, 320)))
        self.assertEqual([x[0] for x in Handler.requests],
                         ['/image/1200/1200/fallback.png', '/image/832/832/fallback.png'])
        self.assertTrue(all(x[1]=='https://www.flipkart.com/product' for x in Handler.requests))

    def test_invalid_and_small_downloads_leave_no_jpg(self):
        for name in ('bad', 'small'):
            with self.assertRaises(RuntimeError):
                d.download_image(self.context, self.base+'/'+name, self.folder, 1, 'jpg', '')
        self.assertEqual(list(self.folder.iterdir()), [])

    def test_original_format_preserves_png(self):
        path = d.download_image(self.context, self.base+'/original', self.folder, 1, 'original', '')
        with Image.open(path) as img:
            self.assertEqual(img.format, 'PNG')

    def test_csv_input_preserves_valid_links_and_ignores_invalid(self):
        file = self.folder/'urls.csv'
        file.write_text('\ufeff Product_URL ,note\nhttps://www.flipkart.com/a/p/itm1,a\nhttps://flipkart.com.evil.test/no,b\n,blank\nhttps://flipkart.com/b/p/itm2,c\n',encoding='utf-8')
        self.assertEqual(d.read_urls_csv(file), ['https://www.flipkart.com/a/p/itm1','https://flipkart.com/b/p/itm2'])

    def test_product_flow_writes_manifest_and_five_jpegs_without_manual_reload(self):
        self.load(listing=True)
        args=SimpleNamespace(wait=0, steps=18, output=self.folder, format='jpg')
        with patch.object(d, 'fetch', side_effect=lambda context,url,referer:
                          context.request.get(self.base+'/image').body()), \
             patch.object(self.page, 'goto', side_effect=AssertionError('Manual page reloaded')):
            count=d.process_one(self.context,self.page,'https://www.flipkart.com/a/p/itm1',1,args,already_loaded=True)
        self.assertEqual(count,5)
        manifest=json.loads(next(self.folder.rglob('manifest.json')).read_text())
        self.assertEqual(len(manifest['saved']),5)
        self.assertEqual(manifest['errors'],[])
        for path in manifest['saved']:
            with Image.open(path) as img:
                self.assertEqual(img.format,'JPEG')


if __name__ == '__main__':
    unittest.main()
