"""Append Flipkart product tabs from Firefox's local saved session to a CSV."""
import argparse
import configparser
import csv
import io
import json
import os
import shutil
import sqlite3
import tempfile
from contextlib import closing
from datetime import datetime
from pathlib import Path

import lz4.block

from downloader import DONE_COLUMN, product_url_identity

ROOT = Path(__file__).resolve().parent
SESSION_FILES = ('sessionstore-backups/recovery.jsonlz4',
                 'sessionstore-backups/recovery.baklz4', 'sessionstore.jsonlz4')


def grouped_profiles(root, store_id):
    """Read selectable profiles without altering Firefox's group database."""
    if not store_id or not store_id.isalnum():
        return []
    database = root / 'Profile Groups' / f'{store_id}.sqlite'
    if not database.is_file():
        return []
    with closing(sqlite3.connect(database.resolve().as_uri() + '?mode=ro', uri=True, timeout=1)) as connection:
        return connection.execute('SELECT path, name FROM Profiles ORDER BY id').fetchall()


def firefox_profiles(root):
    """Discover legacy profiles and selectable profiles within profile groups."""
    parser = configparser.ConfigParser(interpolation=None)
    parser.read(root / 'profiles.ini', encoding='utf-8-sig')
    profiles = {}
    for section in parser.sections():
        if section.startswith('Profile') and parser.has_option(section, 'Path'):
            path = Path(parser.get(section, 'Path'))
            if parser.get(section, 'IsRelative', fallback='1') == '1':
                path = root / path
            path = path.resolve()
            name = parser.get(section, 'Name', fallback=path.name)
            try:
                children = grouped_profiles(root, parser.get(section, 'StoreID', fallback=''))
            except sqlite3.Error as exc:
                print(f'Warning: could not read profiles in {name}: {exc}. Showing its configured profile only.')
                children = []
            for child_path, child_name in children:
                child = Path(child_path)
                if not child.is_absolute():
                    child = root / child
                profiles[child.resolve()] = f'{name} / {child_name}'
            profiles.setdefault(path, name)
    for section in parser.sections():
        if section.startswith('Install') and parser.has_option(section, 'Default'):
            path = (root / parser.get(section, 'Default')).resolve()
            profiles.setdefault(path, path.name)
    return profiles


def prompt_profile(profiles):
    """Require a numbered selection, even when only one profile is configured."""
    if not profiles:
        raise ValueError('No Firefox profiles found. Open Firefox once or use --profile with its profile-folder path.')
    choices = list(profiles.items())
    print('\nChoose a Firefox profile:')
    for number, (path, name) in enumerate(choices, 1):
        available = any((path / relative).is_file() for relative in SESSION_FILES)
        status = '' if available else ' - no saved session'
        print(f'  {number}. {name} ({path.name}){status}')
    while True:
        try:
            answer = input('Profile number (or Q to cancel): ').strip()
        except (EOFError, KeyboardInterrupt):
            print()
            return None
        if answer.casefold() == 'q':
            return None
        if answer.isascii() and answer.isdecimal() and 1 <= int(answer) <= len(choices):
            return choices[int(answer) - 1][0]
        print(f'Enter a number from 1 to {len(choices)}, or Q to cancel.')


def choose_session(profiles, selection=None):
    if selection:
        matches = {path: name for path, name in profiles.items()
                   if selection.casefold() in (name.casefold(), name.rsplit(' / ', 1)[-1].casefold(),
                                              str(path).casefold(), path.name.casefold())}
        if not matches and Path(selection).is_dir():
            path = Path(selection).resolve()
            matches[path] = path.name
        if not matches:
            raise ValueError('Firefox profile not found. Use --list-profiles or an absolute profile-folder path.')
        if len(matches) > 1:
            raise ValueError('Profile name is ambiguous; use its absolute folder path.')
        profiles = matches
    candidates = []
    for profile, name in profiles.items():
        for relative in SESSION_FILES:
            path = profile / relative
            if path.is_file():
                candidates.append((path.stat().st_mtime, path, name))
    if not candidates:
        raise ValueError('No Firefox saved session found. Open product tabs in a normal Firefox window, wait a few seconds, then retry.')
    _, path, name = max(candidates, key=lambda candidate: candidate[0])
    return path, name


def read_session(path):
    raw = path.read_bytes()
    if raw.startswith(b'mozLz40\0'):
        raw = lz4.block.decompress(raw[8:])
    session = json.loads(raw)
    if not isinstance(session, dict) or not isinstance(session.get('windows'), list):
        raise ValueError('Session file has no open-window list.')
    return session


def tab_urls(session):
    """Only each open tab's selected history entry, never closed tabs/history."""
    for window in session.get('windows', []):
        if not isinstance(window, dict):
            continue
        if window.get('isPrivate') or window.get('privateBrowsingId'):
            continue
        for tab in window.get('tabs', []):
            if not isinstance(tab, dict):
                continue
            if tab.get('isPrivate') or tab.get('privateBrowsingId'):
                continue
            entries = tab.get('entries', [])
            if not isinstance(entries, list):
                continue
            index = tab.get('index', len(entries))
            if not isinstance(index, int) or not 1 <= index <= len(entries):
                continue
            entry = entries[index - 1]
            url = entry.get('url') if isinstance(entry, dict) else None
            if isinstance(url, str):
                yield url


def append_product_urls(path, urls, dry_run=False):
    """Preserve existing CSV bytes/columns; append only new product identities."""
    original = path.read_bytes() if path.exists() else b''
    fields, column, seen = ['url', DONE_COLUMN], 'url', set()
    if original:
        reader = csv.DictReader(io.StringIO(original.decode('utf-8-sig'), newline=''))
        fields = reader.fieldnames
        column = next((name for name in fields or [] if name and name.strip().lower()
                       in ('url', 'product_url', 'product link', 'link')), None)
        if not column:
            raise ValueError('Existing CSV has no url, product_url, product link, or link column; file was not changed.')
        for row in reader:
            try:
                seen.add(product_url_identity((row.get(column) or '').strip()))
            except ValueError:
                pass
    additions, duplicates, skipped = [], 0, 0
    for url in urls:
        try:
            identity = product_url_identity(url)
        except ValueError:
            skipped += 1
            continue
        if identity in seen:
            duplicates += 1
        else:
            seen.add(identity)
            additions.append(url)
    stats = {'added': len(additions), 'duplicates': duplicates, 'skipped': skipped}
    if dry_run or not additions:
        return stats
    output = io.StringIO(newline='')
    writer = csv.DictWriter(output, fieldnames=fields)
    if not original:
        writer.writeheader()
    elif not original.endswith((b'\n', b'\r')):
        output.write('\r\n')
    for url in additions:
        writer.writerow({column: url})
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(dir=path.parent, prefix='.firefox-import-', delete=False) as stream:
            temporary = Path(stream.name)
            stream.write(original + output.getvalue().encode('utf-8'))
        if (path.read_bytes() if path.exists() else b'') != original:
            raise ValueError('CSV changed during import. Please run the importer again.')
        if original:
            shutil.copy2(path, path.with_name(path.name + '.bak'))
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
    return stats


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--csv', type=Path, default=ROOT / 'urls.csv')
    parser.add_argument('--profile', help='Firefox profile name or profile-folder path')
    parser.add_argument('--session', type=Path, help='Read a specific saved session JSON/JSONLZ4 file')
    parser.add_argument('--list-profiles', action='store_true')
    parser.add_argument('--dry-run', action='store_true', help='Show counts without changing the CSV')
    args = parser.parse_args()
    try:
        root = Path(os.environ.get('APPDATA', Path.home() / 'AppData/Roaming')) / 'Mozilla/Firefox'
        profiles = firefox_profiles(root)
        if args.list_profiles:
            for path, name in profiles.items():
                print(f'{name}: {path}')
            return
        if args.session:
            source, name = args.session, args.session.parent.name
        else:
            selection = args.profile
            if not selection:
                chosen = prompt_profile(profiles)
                if chosen is None:
                    print('Import cancelled. CSV was not changed.')
                    return
                selection = str(chosen)
            source, name = choose_session(profiles, selection)
        print(f'Firefox profile: {name}')
        print(f'Saved session: {source.name}, updated {datetime.fromtimestamp(source.stat().st_mtime):%Y-%m-%d %H:%M:%S}')
        print('Reading saved open tabs across all windows in this profile. Private tabs are excluded.')
        stats = append_product_urls(args.csv, tab_urls(read_session(source)), args.dry_run)
        verb = 'Would add' if args.dry_run else 'Added'
        print(f"{verb} {stats['added']} product URLs; skipped {stats['duplicates']} duplicates and {stats['skipped']} other/invalid tabs.")
        print(f'CSV: {args.csv.resolve()}')
        print('If recent tabs are missing, wait a few seconds and run again. Saved sessions can lag behind Firefox.')
    except (OSError, ValueError, csv.Error, configparser.Error, lz4.block.LZ4BlockError) as exc:
        parser.exit(1, f'Import failed: {exc}\n')


if __name__ == '__main__':
    main()
