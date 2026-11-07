"""Flipkart product gallery downloader. Python 3.10+, Playwright, Pillow."""
import argparse
import csv
import io
import json
import os
import re
import shutil
import tempfile
import time
from pathlib import Path
from urllib.parse import parse_qs, urlparse
from PIL import Image, ImageOps
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

GALLERY = 'img.Xz2QEH.m6YU3e'
PHOTO = '.OfydJ4:has(img[alt="Image"])'
VIEWER = '.TT5KcC'
DONE_COLUMN = 'already_done'
GALLERY_READY = """selector => [...document.querySelectorAll(selector)].some(image => {
  const viewer = image.closest('.TT5KcC');
  if (!viewer || !image.closest('.P56aSK') || !viewer.querySelector('.ThdfW9')) return false;
  // Image layout/paint may be delayed in a minimized window. The expanded
  // viewer's DOM is enough; a hidden, closed viewer must still be rejected.
  for (let el = image.parentElement; el; el = el.parentElement) {
    if (el.hidden || getComputedStyle(el).display === 'none' ||
        getComputedStyle(el).visibility === 'hidden') return false;
  }
  return true;
})"""


def urls_from_page(page):
    """Read gallery slides only; never mix in recommendations or listing photos."""
    return page.evaluate('''(selector) => {
      const nodes = document.querySelectorAll(selector);
      const found = [];
      for (const el of nodes) {
        // This slide is a playable video poster, not a product photograph.
        if (el.closest('.P56aSK')?.querySelector('.WTuzrs, video')) continue;
        const sets = [el.getAttribute('srcset'),
          ...[...(el.closest('picture')?.querySelectorAll('source') || [])]
            .map(source => source.getAttribute('srcset'))];
        const candidates = [el.currentSrc, el.src, el.getAttribute('data-src'),
          ...sets.flatMap(set => (set || '').split(',').map(v => v.trim().split(/\\s+/)[0]))];
        for (const value of candidates) {
          if (!value) continue;
          const url = new URL(value, document.baseURI);
          if (/^https?:$/.test(url.protocol) &&
              /^rukminim[0-9]*\\.flixcart\\.com$/i.test(url.hostname)) found.push(url.href);
        }
      }
      return [...new Set(found)];
    }''', GALLERY)


def cdn_identity(url):
    """Ignore dimensions, query and optional transformation prefixes for de-duplication."""
    parsed = urlparse(url)
    path = parsed.path
    path = re.sub(r'^/image/\d+/\d+/', '/image/', path)
    path = re.sub(r'^/image/(?:\d+|original)/', '/image/', path)
    return path


def high_res(url):
    # Flipkart CDN paths often look like /image/128/128/x.webp.
    # Request at least 1200px without shrinking an already larger source.
    return re.sub(r'/image/(\d+)/(\d+)/',
                  lambda m: f'/image/{max(1200, int(m[1]))}/{max(1200, int(m[2]))}/',
                  url, count=1)


def image_quality(url):
    match = re.search(r'/image/(\d+)/(\d+)/', urlparse(url).path)
    if '/image/original/' in urlparse(url).path:
        return float('inf')
    return int(match[1]) * int(match[2]) if match else 0


def fetch(context, url, referer):
    options = {'timeout': 25000, 'headers': {'Referer': referer}}
    response = context.request.get(url, **options)
    try:
        if response.ok:
            return response.body()
        raise RuntimeError(f'HTTP {response.status}')
    finally:
        response.dispose()


def download_image(context, url, folder, idx, fmt, referer):
    candidates = list(dict.fromkeys([high_res(url), url]))
    errors = []
    for candidate in candidates:
        try:
            raw = fetch(context, candidate, referer)
            img = Image.open(io.BytesIO(raw))
            img.load()
            if img.width < 120 or img.height < 120:
                raise ValueError(f'Only {img.width}x{img.height}px')
            if fmt == 'jpg':
                img = ImageOps.exif_transpose(img).convert('RGB')
                path = folder / f'{idx:02d}.jpg'
                img.save(path, 'JPEG', quality=94, optimize=True)
            else:
                ext = 'webp' if img.format == 'WEBP' else ('png' if img.format == 'PNG' else 'jpg')
                path = folder / f'{idx:02d}.{ext}'
                path.write_bytes(raw)
            return str(path)
        except Exception as exc:
            errors.append(str(exc))
    raise RuntimeError('; '.join(errors))


def gallery_is_open(page):
    """Detect the expanded viewer without waiting for image layout or paint."""
    return page.evaluate(GALLERY_READY, GALLERY)


def wait_for_gallery(page, timeout=8000):
    try:
        # Default requestAnimationFrame polling can pause in minimized windows.
        handle = page.wait_for_function(GALLERY_READY, arg=GALLERY, polling=200, timeout=timeout)
        handle.dispose()
        return True
    except PlaywrightTimeoutError:
        return False


def open_gallery(page, timeout=8000):
    deadline = time.monotonic() + timeout / 1000

    def remaining():
        return max(1, int((deadline - time.monotonic()) * 1000))

    def opened():
        return wait_for_gallery(page, timeout=min(3000, remaining()))

    if gallery_is_open(page):
        return True

    # In the supplied listing .ARoD00/.eWJFnV is a VIDEO poster.
    # The actual product photos are .OfydJ4 pictures. Click their parent card
    # so the sibling "Key Highlights" overlay receives a real pointer click.
    try:
        page.locator(PHOTO).first.wait_for(state='attached', timeout=remaining())
    except PlaywrightTimeoutError:
        return False
    photos = page.locator(PHOTO)
    for index in range(min(photos.count(), 6)):
        if gallery_is_open(page):
            return True
        if time.monotonic() >= deadline:
            break
        photo = photos.nth(index)
        if not photo.is_visible():
            continue
        rect = photo.bounding_box()
        if not rect or min(rect['width'], rect['height']) < 180:
            continue
        card = photo.locator('..')
        if page.evaluate('document.hidden'):
            print(f'  Clicking product photo card {index + 1} in background', flush=True)
            card.evaluate('(el) => el.click()')
            if opened():
                return True
        try:
            print(f'  Clicking product photo card {index + 1}', flush=True)
            card.click(timeout=min(4000, remaining()))
            if opened():
                return True
        except PlaywrightTimeoutError as exc:
            print(f'  Photo click timed out: {str(exc)[:100]}', flush=True)

        # A container may have a different center from its photo. Scroll first,
        # then click the photo coordinates; the browser hit-tests its overlay.
        if gallery_is_open(page):
            return True
        if time.monotonic() >= deadline:
            break
        try:
            photo.scroll_into_view_if_needed(timeout=min(3000, remaining()))
            rect = photo.bounding_box()
            if rect:
                page.mouse.click(rect['x'] + rect['width'] / 2,
                                 rect['y'] + rect['height'] / 2)
                if opened():
                    return True
            if time.monotonic() >= deadline:
                break
            # Bounded DOM fallback on this photo's card only, never unrelated
            # product-page ancestors or the video player.
            card.evaluate('(el) => el.click()')
            if opened():
                return True
        except PlaywrightTimeoutError:
            pass
    return gallery_is_open(page)


def gallery_state(page):
    # The supplied viewer mounts every slide up front and moves the strip.
    # URLs alone therefore cannot tell whether navigation actually worked.
    return page.evaluate("""selector => JSON.stringify({
      strips: [...document.querySelectorAll('.TT5KcC .hKE1H7')]
        .map(el => el.getAttribute('style')),
      selected: [...document.querySelectorAll('.TT5KcC .PylJiO')]
        .map(el => el.className),
      images: [...document.querySelectorAll(selector)]
        .map(el => [el.currentSrc, el.src])
    })""", GALLERY)


def wait_for_gallery_change(page, previous):
    deadline = time.monotonic() + 2
    while time.monotonic() < deadline:
        page.wait_for_timeout(100)
        if gallery_state(page) != previous:
            # The saved CSS has a 250ms carousel animation.
            page.wait_for_timeout(300)
            return True
    return False


def move_gallery(page, direction):
    previous = gallery_state(page)
    control = '.Vi7peE' if direction == 'next' else '.HbzLrW'
    disabled = '.K5mlx2' if direction == 'next' else '.B_lGg3'
    buttons = page.locator(f'{VIEWER} .S0f9kM {control}:visible')
    if buttons.count():
        button = buttons.first
        if button.evaluate('(el, disabled) => el.matches(disabled) || el.getAttribute("aria-disabled") === "true"', disabled):
            return False
        try:
            if page.evaluate('document.hidden'):
                button.evaluate('(el) => el.click()')
            else:
                button.click(timeout=3000)
        except PlaywrightTimeoutError:
            # This specific gallery control can be activated without waiting
            # for animation frames to make its position stable.
            button.evaluate('(el) => el.click()')
        if wait_for_gallery_change(page, previous):
            return True
    page.keyboard.press('ArrowRight' if direction == 'next' else 'ArrowLeft')
    return wait_for_gallery_change(page, previous)


def launch_browser(playwright, headless=False):
    return playwright.chromium.launch(headless=headless, args=[
        '--disable-background-timer-throttling',
        '--disable-backgrounding-occluded-windows',
        '--disable-renderer-backgrounding',
    ])


def wait_for_image_urls(page, timeout=3000):
    """Wait for usable gallery sources without requiring browser image paint."""
    deadline = time.monotonic() + timeout / 1000
    while True:
        urls = urls_from_page(page)
        if urls:
            return urls
        if time.monotonic() >= deadline:
            raise RuntimeError(f'Gallery structure found but no photo URLs loaded within {timeout / 1000:g}s')
        page.wait_for_timeout(100)


def collect_gallery(page, max_steps, image_timeout=3000):
    if max_steps < 1:
        raise ValueError('steps must be at least 1')
    if not gallery_is_open(page):
        raise RuntimeError('Cannot collect images without an open gallery')
    wait_for_image_urls(page, image_timeout)
    seen = {}

    def gather():
        for url in urls_from_page(page):
            key = cdn_identity(url)
            if key not in seen or image_quality(url) > image_quality(seen[key]):
                seen[key] = url

    gather()
    # A listing can open on any photo. Visit earlier positions as well as later
    # ones so lazy galleries do not lose images before the clicked photo.
    for direction in ('previous', 'next'):
        states = {gallery_state(page)}
        for _ in range(max_steps):
            if not move_gallery(page, direction):
                break
            gather()
            state = gallery_state(page)
            if state in states:
                break
            states.add(state)
    return list(seen.values())


def product_name(page, ordinal):
    try:
        title = page.locator('h1').first.inner_text(timeout=2500).strip()
    except Exception:
        title = page.title().split('|')[0]
    title = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', title).strip()[:75]
    return f'{ordinal:03d} - {title or "Product"}'


def process_one(context, page, url, ordinal, args, already_loaded=False):
    print(f'\n[{ordinal}] {url}', flush=True)
    if not already_loaded:
        page.goto(url, wait_until='domcontentloaded', timeout=60000)
    page.wait_for_timeout(args.wait * 1000)
    name = product_name(page, ordinal)
    gallery_open = open_gallery(page, timeout=int(getattr(args, 'gallery_timeout', 8) * 1000))
    print('Gallery detected:', gallery_open, flush=True)
    if not gallery_open:
        raise RuntimeError('Gallery structure was not detected before timeout. See debug_###.png/html.')
    images = collect_gallery(page, args.steps, image_timeout=int(getattr(args, 'image_timeout', 3) * 1000))
    if not images:
        raise RuntimeError('Gallery opened but no CDN images were found. Inspect debug screenshot and HTML.')
    folder = args.output / name
    folder.mkdir(parents=True, exist_ok=True)
    saved, errors = [], []
    for idx, image_url in enumerate(images, 1):
        try:
            path = download_image(context, image_url, folder, idx, args.format, page.url)
            saved.append(path)
            print(f'  saved {idx}/{len(images)}: {Path(path).name}', flush=True)
        except Exception as exc:
            errors.append({'url': image_url, 'error': str(exc)})
            print(f'  FAILED image {idx}: {exc}', flush=True)
    (folder / 'manifest.json').write_text(json.dumps({'source': url, 'gallery_detected': gallery_open, 'image_urls': images, 'saved': saved, 'errors': errors}, indent=2), encoding='utf-8')
    print(f'Completed: {len(saved)} images, {len(errors)} image errors', flush=True)
    if not saved:
        raise RuntimeError('All image downloads failed; see manifest.json for details.')
    if errors:
        raise RuntimeError(f'{len(errors)} image downloads failed; CSV stays pending. See manifest.json for details.')
    return len(saved)


def product_url_identity(url):
    """Validate a product URL and ignore tracking, fragments and host aliases.

    PID identifies a product variant, so different PIDs remain separate.
    The original URL is retained for navigation.
    """
    try:
        if re.search(r'\s|[\\\x00-\x1f\x7f]', url):
            raise ValueError('URL contains whitespace or invalid characters')
        parsed = urlparse(url)
        host = (parsed.hostname or '').lower()
        if parsed.scheme not in ('http', 'https'):
            raise ValueError('expected an http or https URL')
        if not (host == 'flipkart.com' or host.endswith('.flipkart.com')):
            raise ValueError('expected a flipkart.com host')
        if parsed.username is not None or parsed.password is not None:
            raise ValueError('URLs with login credentials are not supported')
        expected_port = 443 if parsed.scheme == 'https' else 80
        if parsed.port is not None and parsed.port != expected_port:
            raise ValueError('unexpected URL port')
        product = re.fullmatch(r'(?:/[^/]+)*/p/([A-Za-z0-9_-]+)/?', parsed.path)
        if not product or parsed.params:
            raise ValueError('expected a product link containing /p/<product-id>')
        query = parse_qs(parsed.query)
        pids = tuple(sorted({pid.upper() for pid in query.get('pid', []) if pid}))
        return product[1].lower(), pids
    except ValueError as exc:
        raise ValueError(f'Invalid Flipkart product URL: {exc}') from exc


def done_column(fields):
    return next((name for name in fields if name and name.strip().lower().replace(' ', '_')
                 in (DONE_COLUMN, 'done')), None)


def update_csv(path, mark_identity=None):
    """Add the status column or checkpoint a completed product, preserving other cells."""
    path = Path(path)
    original = path.read_bytes()
    reader = csv.DictReader(io.StringIO(original.decode('utf-8-sig'), newline=''))
    fields = reader.fieldnames
    if not fields:
        raise ValueError('CSV is empty or missing its header row.')
    url_column = next((name for name in fields if name and name.strip().lower()
                       in ('url', 'product_url', 'product link', 'link')), None)
    if not url_column:
        raise ValueError('CSV must contain a url column.')
    rows = list(reader)
    if any(None in row for row in rows):
        raise ValueError('CSV has rows with extra cells. Fix them before updating completion status.')
    status = done_column(fields)
    changed = status is None
    if status is None:
        status = DONE_COLUMN
        fields.append(status)
    matched = False
    if mark_identity is not None:
        for row in rows:
            try:
                identity = product_url_identity((row.get(url_column) or '').strip())
            except ValueError:
                continue
            if identity == mark_identity:
                matched = True
                if not (row.get(status) or '').strip():
                    row[status] = 'yes'
                    changed = True
        if not matched:
            raise ValueError('Completed URL is no longer in the CSV; completion was not recorded.')
    if not changed:
        return
    output = io.StringIO(newline='')
    writer = csv.DictWriter(output, fieldnames=fields)
    writer.writeheader()
    writer.writerows(rows)
    encoding = 'utf-8-sig' if original.startswith(b'\xef\xbb\xbf') else 'utf-8'
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(dir=path.parent, prefix='.csv-status-', delete=False) as stream:
            temporary = Path(stream.name)
            stream.write(output.getvalue().encode(encoding))
        if path.read_bytes() != original:
            raise ValueError('CSV changed during status update. Please retry.')
        shutil.copy2(path, path.with_name(path.name + '.bak'))
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def mark_url_done(path, url):
    update_csv(path, product_url_identity(url))


def read_urls_csv(path):
    """Read product links from a CSV containing a url column."""
    with Path(path).open('r', encoding='utf-8-sig', newline='') as file:
        reader = csv.DictReader(file)
        if not reader.fieldnames:
            raise ValueError('CSV is empty or missing its header row.')
        url_column = next((name for name in reader.fieldnames if name and name.strip().lower() in ('url', 'product_url', 'product link', 'link')), None)
        if not url_column:
            raise ValueError('CSV must contain a column named url (or product_url / link).')
        status = done_column(reader.fieldnames)
        rows = [(reader.line_num, row) for row in reader]
        completed = set()
        for _, row in rows:
            if status and (row.get(status) or '').strip():
                try:
                    completed.add(product_url_identity((row.get(url_column) or '').strip()))
                except ValueError:
                    pass
        urls = []
        seen = set()
        for line_num, row in rows:
            url = (row.get(url_column) or '').strip()
            if not url:
                continue
            try:
                identity = product_url_identity(url)
            except ValueError as exc:
                print(f'Skipping CSV row {line_num}: {exc} ({url})', flush=True)
                continue
            if identity in completed:
                print(f'Skipping already done URL at CSV row {line_num}: {url}', flush=True)
                continue
            if identity in seen:
                print(f'Skipping duplicate URL at CSV row {line_num}: {url}', flush=True)
                continue
            seen.add(identity)
            urls.append(url)
        return urls


def run_downloads(context, page, urls, args):
    """Finish a pass before retrying failures; keep stable product ordinals."""
    pending = list(enumerate(urls, 1))
    total = 0
    retries = getattr(args, 'retries', 2)
    for attempt in range(retries + 1):
        if not pending:
            break
        if attempt:
            print(f'\nRETRY PASS {attempt}/{retries}: {len(pending)} pending products', flush=True)
        retry_queue = []
        for ordinal, url in pending:
            try:
                manual_first = args.manual and ordinal == 1 and attempt == 0
                if manual_first:
                    page.goto(url, wait_until='domcontentloaded', timeout=60000)
                    input('Complete any login/challenge in browser, then press Enter here: ')
                count = process_one(context, page, url, ordinal, args, already_loaded=manual_first)
                mark_url_done(args.csv, url)
                total += count
                print('  Marked already_done=yes in CSV', flush=True)
            except Exception as exc:
                print(f'ERROR product {ordinal}, attempt {attempt + 1}/{retries + 1}: {exc}', flush=True)
                retry_queue.append((ordinal, url))
                if attempt < retries:
                    print('  Queued for retry after the other URLs', flush=True)
                debug = args.output / f'debug_{ordinal:03d}'
                # Save the DOM first: taking a screenshot can itself trigger
                # rendering, which would hide evidence of a paint delay.
                try:
                    debug.with_suffix('.html').write_text(page.content(), encoding='utf-8')
                    debug.with_suffix('.json').write_text(json.dumps({
                        'url': url, 'attempt': attempt + 1, 'error': str(exc),
                    }, indent=2), encoding='utf-8')
                except Exception:
                    pass
                try:
                    page.screenshot(path=str(debug.with_suffix('.png')), full_page=True, timeout=7000)
                except Exception:
                    pass
        pending = retry_queue
    return total, len(pending)


def main():
    parser = argparse.ArgumentParser(description='Download Flipkart gallery images from product URLs.')
    parser.add_argument('--csv', '--urls', dest='csv', default='urls.csv', help='CSV file with a url column')
    parser.add_argument('--output', type=Path, default=Path('downloads'))
    parser.add_argument('--format', choices=['jpg', 'original'], default='jpg')
    parser.add_argument('--wait', type=int, default=4, help='Seconds after loading page')
    parser.add_argument('--steps', type=int, default=18, help='Maximum gallery moves in each direction')
    parser.add_argument('--manual', action='store_true', help='Pause for manual login/challenge after first page')
    parser.add_argument('--headless', action='store_true', help='Run in the background without a browser window')
    parser.add_argument('--gallery-timeout', type=float, default=8, help='Seconds allowed to open the gallery')
    parser.add_argument('--image-timeout', type=float, default=3, help='Seconds to wait for gallery photo URLs')
    parser.add_argument('--retries', type=int, default=2, help='Extra passes for failed URLs, after processing the others')
    args = parser.parse_args()
    if args.steps < 1 or args.wait < 0:
        parser.error('--steps must be at least 1 and --wait must be nonnegative')
    if args.manual and args.headless:
        parser.error('--manual requires a visible browser; omit --headless')
    if args.retries < 0 or args.gallery_timeout <= 0 or args.image_timeout <= 0:
        parser.error('--retries must be nonnegative and gallery/image timeouts must be positive')
    args.output.mkdir(parents=True, exist_ok=True)
    update_csv(args.csv)
    urls = read_urls_csv(args.csv)
    if not urls:
        print('No pending valid Flipkart URLs. All completed rows are skipped.')
        return
    with sync_playwright() as p:
        browser = launch_browser(p, headless=args.headless)
        context = browser.new_context(viewport={'width': 1365, 'height': 850}, accept_downloads=True, locale='en-IN')
        page = context.new_page()
        try:
            total, failed = run_downloads(context, page, urls, args)
        finally:
            context.close()
            browser.close()
    print(f'\nFINISHED: {len(urls) - failed}/{len(urls)} products completed, {failed} still pending, '
          f'{total} images saved. Output: {args.output.resolve()}')

if __name__ == '__main__':
    main()
