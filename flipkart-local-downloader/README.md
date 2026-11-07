# Flipkart Local Product Gallery Downloader

A local Python tool; **not a browser extension**. Opens a visible Chromium browser using Playwright, navigates each Flipkart product page, clicks a product image to reveal the full-screen gallery, collects the gallery images, and downloads them into a separate folder for each product. Converts actual image data to JPG using Pillow.

## Windows quick start

1. Install Python 3.10+ (enable **Add Python to PATH**).
2. Open `urls.csv` in this folder and paste your Flipkart product links into the `url` column, one per row, then save it. Replace the example URL.
3. Double-click `START_WINDOWS.bat`. It installs dependencies and starts the scraper in a visible browser.
4. Find the downloaded images under `downloads/` in this folder.

For background runs, double-click `START_BACKGROUND.bat`. It runs Chromium headless, so there is no browser window to keep open or minimize. Use `START_WINDOWS.bat` for a visible browser, especially when using `--manual` for login/challenges. The visible launcher also accepts command-line options, e.g. `START_WINDOWS.bat --retries 1`.

## CSV format

```csv
url,already_done
https://www.flipkart.com/example-one/p/itm123,
https://www.flipkart.com/example-two/p/itm456,yes
```

The column name can also be `product_url` or `link`. Additional CSV columns are ignored.

Before opening the browser, the tool skips invalid URLs, completed products, and duplicate products within the CSV, printing the row number and reason. Links must use HTTP/HTTPS, a Flipkart host, and a product path containing `/p/<product-id>`. Duplicate detection ignores tracking parameters, fragments, HTTP/HTTPS differences, and host aliases such as `www` and `m`; the first URL is kept in its original position. Different `pid` values are treated as separate product variants. A link without a `pid` is also kept separate from one with a `pid`.

The downloader adds `already_done` automatically to older CSV files. Only rows where this cell is blank are processed; any nonblank value means skip. After all collected product photos download successfully and the manifest is saved, it immediately writes `yes` for that product (including duplicate URL rows). Failures and partial downloads remain blank for retry. Previously downloaded folders do not automatically mark rows complete. To download a completed product again, clear its `already_done` cell on every duplicate row. Existing columns are preserved, updates replace the file atomically, and the prior CSV is saved as `urls.csv.bak`. Keep the CSV closed in Excel while downloading so status updates can be saved. Firefox imports preserve existing statuses and leave newly added URLs blank.

## Import Firefox tabs

Open your product pages in normal Firefox windows, wait a few seconds, then double-click `IMPORT_FIREFOX_TABS.bat`. Choose your Firefox profile by entering its number (or enter `Q` to cancel). It adds new Flipkart product URLs to this folder's `urls.csv`. Then run `START_WINDOWS.bat` to download their images.

The importer reads Firefox's local saved session across all windows in the selected profile. It uses each tab's selected history entry, ignores closed/private tabs and other websites, and skips products already in the CSV. It preserves existing rows and additional columns. Before an update it saves the previous CSV as `urls.csv.bak` (replaced on the next successful import). Firefox profile files are only read; tabs stay open. No browser extension is required.

By default it asks which profile to use, even if only one profile is configured. It then reads the most recently saved session in that profile. The profile and snapshot time are printed. Saved sessions may lag behind your current tabs or represent the last session when Firefox is closed. Private tabs are not available through this method. If tabs are missing, wait and retry. To select a profile directly without the prompt:

```powershell
.\.venv\Scripts\python.exe import_firefox_tabs.py --list-profiles
.\.venv\Scripts\python.exe import_firefox_tabs.py --profile "ESHAAN MAIN / Eshaan" --dry-run
.\.venv\Scripts\python.exe import_firefox_tabs.py --profile "ESHAAN MAIN / Eshaan"
```

`--dry-run` shows counts without changing the CSV. You can also pass `--csv another.csv` or `--session path/to/recovery.jsonlz4`. The importer only updates the CSV; it does not start downloading.

The menu includes Firefox's selectable profiles inside profile groups, displayed as `ESHAAN MAIN / Prabhu`, `ESHAAN MAIN / Eshaan`, and `ESHAAN MAIN / Seema`. Select the specific person to import only that profile's tabs. These names come from Firefox's local `Profile Groups` database, opened read-only. A short name such as `--profile Prabhu` also works when it is unique.

## Manual command

```powershell
pip install -r requirements.txt
python -m playwright install chromium
python downloader.py --csv urls.csv --format jpg
```

Optional: `--format original` preserves WebP/PNG/JPEG, `--manual` lets you handle login or challenge in the browser, `--wait 8` waits longer, `--steps 25` allows up to 25 moves in each direction. Gallery collection excludes playable video posters and listing recommendations.

Gallery detection checks the expanded viewer's structure rather than requiring a painted image. It uses timer polling so a paused animation frame in a minimized window does not prevent detection. Chromium launches with background throttling disabled, and photo/control DOM clicks are used when the document is hidden or normal navigation clicks time out.

The default limits are 8 seconds to open the gallery (`--gallery-timeout 8`) and 3 seconds for usable photo URLs (`--image-timeout 3`). Actual image requests are separately downloaded and decoded before success is recorded. Failed URLs are queued while the remaining URLs continue, then retried in up to two additional passes (`--retries 2`). `--retries 0` disables retry passes. Successful products are not repeated in those passes; failures remain blank in `already_done` after retries are exhausted. Debug HTML, PNG, and JSON files record the latest failed attempt for each original product number.

## Troubleshooting

If a product fails, review `downloads/debug_001.png` and `debug_001.html`. Per-product `manifest.json` files list the selected image URLs, saved files, and download errors.

## Gallery click fix (v3)

Inspection of the two supplied saved pages identified the original click bug: `.ARoD00 .eWJFnV` is the **video poster**, not the product photo. The product photos are inside `.OfydJ4` containers, with a sibling Key Highlights overlay. The downloader now clicks the photo's parent card, then waits for a visible expanded-gallery image. Coordinate and bounded card DOM clicks are fallbacks.

The expanded viewer uses `img.Xz2QEH.m6YU3e`, next `.Vi7peE`, and previous `.HbzLrW` controls. The downloader uses those controls first and keyboard arrows as a fallback, checks carousel state changes, and visits both directions from the clicked photo. The supplied gallery has five product photos and one video poster; the poster is skipped.

Extraction stays within the gallery, reads image/source srcsets and lazy data-src attributes, accepts Flipkart CDN hosts, and deduplicates image identities across size/query variants. It prefers larger dimensions, never reduces an already larger source, retries the source URL if the larger request fails, and checks decoded image dimensions before saving a real JPEG. Manual mode retains the page after login/challenge instead of reloading it.

## Verification

Run the offline suite after installing dependencies and Chromium:

```powershell
python -m unittest discover -s tests -v
```

The regression suite includes CSV URL validation and duplicate detection tests, plus local headless Chromium tests. The fixtures are minimal DOM excerpts extracted from the supplied HTML, without scripts or session data. Test handlers simulate the gallery opening and carousel movement because saved HTML alone does not preserve the live React event handlers. Tests cover the photo overlay click, delayed opening, video exclusion, navigation from a middle slide, srcset extraction, deduplication, CSV parsing, download fallback over a local HTTP server, invalid/small image rejection, JPEG conversion, original PNG preservation, and product manifest output.

On October 21, 2026, all 46 regression tests passed, including an actual minimized Chromium window, paused animation frames, image layout delays, bounded waits, and retry ordering. A live headless check of the reported Flaying jeans URL opened its gallery, extracted four photo URLs, and downloaded/decoded one valid JPEG in a temporary folder without changing the CSV. This verifies that example; other products and future markup can differ. Saved browser HTML rewrites image sources to local companion files, so those saved paths are not treated as downloadable CDN URLs.
