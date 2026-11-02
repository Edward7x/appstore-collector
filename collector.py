#!/usr/bin/env python3
"""Public App Store metadata and screenshot archiver. Python 3.10+; Excel export uses openpyxl and Pillow."""
import argparse
import csv
import hashlib
import html
import json
import logging
import os
import random
import re
import socket
import tempfile
import time
from datetime import datetime, timezone, timedelta
from email.utils import parsedate_to_datetime
from http.client import IncompleteRead
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qs, urlparse, urlencode, urlunparse
from urllib.request import Request, urlopen
from excel_export import write_excel

LOG = logging.getLogger('appstore')
UA = 'AppStoreResearchArchive/1.0 (personal metadata research; urllib)'
HEADERS = {'User-Agent': UA, 'Accept': 'application/json, text/html, image/*'}
APP_ID_RE = re.compile(r'(?<!\d)(?:id)(\d{5,15})(?!\d)', re.I)
COLUMNS = ['app_id','region','trackName','artistName','bundleId','version','primaryGenreName',
           'averageUserRating','userRatingCount','price','formattedPrice','currency',
           'releaseDate','currentVersionReleaseDate','trackViewUrl','description',
           'releaseNotes','icon_file','screenshot_count','status','error','fetched_at']


def now():
    return datetime.now(timezone.utc).isoformat(timespec='seconds')


def atomic_write(path: Path, contents: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile('w', encoding='utf-8', dir=path.parent, delete=False) as tmp:
        tmp.write(contents)
        name = tmp.name
    os.replace(name, path)


def dump(path, obj):
    atomic_write(path, json.dumps(obj, ensure_ascii=False, indent=2) + '\n')


def read_json(path, default=None):
    try:
        return json.loads(path.read_text('utf-8'))
    except (OSError, ValueError):
        return default


def parse_id(value):
    value = value.strip()
    if value.isdigit() and 5 <= len(value) <= 15:
        return value
    parsed = urlparse(value)
    if parsed.scheme not in ('https', 'http') or parsed.hostname not in ('apps.apple.com', 'itunes.apple.com'):
        raise ValueError('Expected an App Store URL or numeric App ID')
    match = APP_ID_RE.search(parsed.path)
    if match:
        return match.group(1)
    ident = parse_qs(parsed.query).get('id', [''])[0]
    if ident.isdigit() and 5 <= len(ident) <= 15:
        return ident
    raise ValueError('No App ID found in URL')


class Client:
    def __init__(self, interval=3.5, retries=4, timeout=25):
        self.interval, self.retries, self.timeout = interval, retries, timeout
        self.last_api = None

    def get(self, url, api=False):
        for attempt in range(self.retries + 1):
            if api:
                if self.last_api is not None:
                    delay = self.interval - (time.monotonic() - self.last_api)
                    if delay > 0: time.sleep(delay)
                self.last_api = time.monotonic()
            try:
                with urlopen(Request(url, headers=HEADERS), timeout=self.timeout) as resp:
                    return resp.read(), resp.headers.get('Content-Type', '')
            except HTTPError as exc:
                if exc.code not in (408, 429, 500, 502, 503, 504) or attempt == self.retries:
                    raise
                retry_after = exc.headers.get('Retry-After')
                pause = 0
                if retry_after:
                    try: pause = max(0, float(retry_after))
                    except ValueError:
                        try: pause = max(0, (parsedate_to_datetime(retry_after) - datetime.now(timezone.utc)).total_seconds())
                        except (ValueError, TypeError, OverflowError): pass
                pause = max(pause, min(60, 2**attempt + random.uniform(0, 1)))
                LOG.warning('HTTP %s; retry %s/%s in %.1fs', exc.code, attempt+1, self.retries, pause)
                time.sleep(pause)
            except (URLError, TimeoutError, socket.timeout, ConnectionError, IncompleteRead) as exc:
                if attempt == self.retries: raise
                pause = min(60, 2**attempt + random.uniform(0, 1))
                LOG.warning('%s; retry %s/%s in %.1fs', exc, attempt+1, self.retries, pause)
                time.sleep(pause)
        raise RuntimeError('Unreachable')

    def lookup(self, app_id, region):
        url = 'https://itunes.apple.com/lookup?' + urlencode({'id':app_id,'country':region,'entity':'software'})
        raw, _ = self.get(url, api=True)
        payload = json.loads(raw.decode('utf-8-sig'))
        matches = [a for a in payload.get('results', []) if str(a.get('trackId')) == app_id and a.get('kind') == 'software']
        return matches[0] if matches else None


def valid_media_url(url):
    if not isinstance(url, str): return False
    p = urlparse(url)
    # Apple-hosted artwork CDN only; reduces malicious/accidental external requests.
    return p.scheme == 'https' and (p.hostname == 'mzstatic.com' or (p.hostname or '').endswith('.mzstatic.com'))


def extension(url, content_type):
    ct = content_type.split(';')[0].lower().strip()
    ext = {'image/png':'.png','image/jpeg':'.jpg','image/webp':'.webp','image/avif':'.avif'}.get(ct)
    if ext: return ext
    suffix = Path(urlparse(url).path).suffix.lower()
    return suffix if suffix in ('.png','.jpg','.jpeg','.webp','.avif') else '.img'


def find_page_images(text):
    """Best-effort discovery from public HTML only; does not bypass protections.

    This collects only Apple CDN image URLs near screenshot-related markup.
    Its output is NOT guaranteed to be exhaustive or device-accurate.
    """
    text = html.unescape(text).replace('\\/', '/')
    candidates = re.findall(r'https://[^\s"\'<>\\]+?\.mzstatic\.com/[^\s"\'<>\\]+', text, re.I)
    images = []
    for raw in candidates:
        url = raw.rstrip('),;]')
        # App Store's screenshots generally contain image/thumb/PurpleSource in the CDN path.
        if ('image/thumb/' in url or 'PurpleSource/' in url) and valid_media_url(url):
            images.append(url)
    return list(dict.fromkeys(images))


def add_media(media, url, kind, source):
    if not valid_media_url(url): return
    # Identify identical source URLs across duplicate fields, not across devices/regions.
    if any(m['url'] == url for m in media): return
    media.append({'url':url,'kind':kind,'source':source})


def download_media(client, entries, base, state, refresh=False):
    files = state.setdefault('files', {})
    succeeded = 0
    for index, entry in enumerate(entries, 1):
        url = entry['url']
        key = hashlib.sha256(url.encode()).hexdigest()
        old = files.get(key, {})
        oldpath = base / old.get('path', '__missing__')
        if not refresh and old.get('ok') and oldpath.is_file() and oldpath.stat().st_size:
            succeeded += 1
            continue
        try:
            data, ct = client.get(url)
            if not ct.lower().startswith('image/') or not data:
                raise ValueError(f'Unexpected media Content-Type: {ct}')
            subdir = 'icon' if entry['kind'] == 'icon' else 'screenshots/' + entry['kind']
            filename = f'{index:03d}_{key[:12]}{extension(url, ct)}'
            path = Path(subdir) / filename
            target = base / path
            target.parent.mkdir(parents=True, exist_ok=True)
            with tempfile.NamedTemporaryFile('wb', dir=target.parent, delete=False) as tmp:
                tmp.write(data)
                temp = tmp.name
            os.replace(temp, target)
            files[key] = {'url':url,'path':path.as_posix(),'sha256':hashlib.sha256(data).hexdigest(),
                          'bytes':len(data),'ok':True,'kind':entry['kind'],'source':entry['source']}
            succeeded += 1
        except (HTTPError, URLError, TimeoutError, ValueError, OSError, IncompleteRead) as exc:
            LOG.error('Media failed: %s: %s', url, exc)
            files[key] = {'url':url,'ok':False,'error':str(exc),'kind':entry['kind'],'source':entry['source']}
        dump(base / 'manifest.json', state)
    return succeeded


def collect(client, app_id, region, root, fallback=False, refresh=False):
    base = root / app_id / region
    base.mkdir(parents=True, exist_ok=True)
    state = read_json(base / 'manifest.json', {}) or {}
    # Freeze the legacy timestamp before this run updates last_success.
    state.setdefault('metadata_fetched_at', state.get('last_success', ''))
    state.update({'app_id':app_id,'region':region,'last_attempt':now()})
    try:
        if (refresh or not (base / 'metadata.json').exists()
                or not state.get('metadata_current', True)
                or state.get('status') == 'not_available'):
            app = client.lookup(app_id, region)
            if app is None:
                state.update({'status':'not_available','error':'Not returned in this storefront',
                              'metadata_current':False})
                dump(base / 'manifest.json', state)
                return make_row(app_id, region, {}, state)
            dump(base / 'metadata.json', app)
            state.update({'metadata_current':True, 'metadata_fetched_at':now()})
            dump(base / 'manifest.json', state)
        else:
            app = read_json(base / 'metadata.json')
            if not isinstance(app, dict):
                raise ValueError('Invalid cached metadata: use --refresh')
        atomic_write(base / 'description.md', f'# {app.get("trackName",app_id)}\n\n## Description\n\n{app.get("description", "")}\n\n## Release notes\n\n{app.get("releaseNotes", "")}\n')
        media = []
        icon = next((app[k] for k in ('artworkUrl512','artworkUrl100','artworkUrl60') if valid_media_url(app.get(k))), None)
        if icon: add_media(media, icon, 'icon', 'lookup')
        for field, kind in [('screenshotUrls','iphone'),('ipadScreenshotUrls','ipad'),('appletvScreenshotUrls','appletv')]:
            for url in app.get(field, []) or []: add_media(media, url, kind, 'lookup')
        if fallback:
            # Public web page may include additional artwork, but cannot reliably distinguish screenshots
            # from icon, recommendations, in-app event creatives, or alternate language variants.
            page = f'https://apps.apple.com/{region}/app/id{app_id}'
            try:
                html_data, _ = client.get(page)
                for url in find_page_images(html_data.decode('utf-8', errors='replace')):
                    add_media(media, url, 'web-unclassified', 'web-heuristic')
            except (HTTPError, URLError, TimeoutError, ValueError, OSError, IncompleteRead) as exc:
                LOG.warning('Web fallback failed for %s %s: %s', app_id, region, exc)
        state['media'] = media
        dump(base / 'manifest.json', state)
        count = download_media(client, media, base, state, refresh=refresh)
        failed = len(media) - count
        state.update({'status':'partial' if failed else 'complete','error':f'{failed} media failed' if failed else '', 'last_success':now()})
        dump(base / 'manifest.json', state)
        return make_row(app_id, region, app, state)
    except (HTTPError, URLError, TimeoutError, ValueError, OSError, IncompleteRead) as exc:
        state.update({'status':'error','error':str(exc)})
        dump(base / 'manifest.json', state)
        LOG.error('App failed %s/%s: %s', app_id, region, exc)
        return make_row(app_id, region, {}, state)


def make_row(app_id, region, app, state):
    files = list(state.get('files', {}).values())
    icons = [f['path'] for f in files if f.get('ok') and f.get('kind') == 'icon']
    row = {name:app.get(name,'') for name in COLUMNS}
    row.update({'app_id':app_id,'region':region,'icon_file':icons[0] if icons else '',
                'screenshot_count':sum(bool(f.get('ok')) and f.get('kind') != 'icon' for f in files),
                'status':state.get('status',''), 'error':state.get('error',''),
                'fetched_at':state.get('metadata_fetched_at', state.get('last_success',''))})
    return row


def workbook_filename(product_name):
    name = product_name.strip()
    if not name or re.search(r'[<>:"/\\|?*\x00-\x1f]', name) or name.endswith('.'):
        raise ValueError('Product name must be nonempty and contain no filename separators or reserved characters')
    date = datetime.now(timezone(timedelta(hours=8))).strftime('%Y-%m-%d')
    return f'{name}_AppStore竞品宣传图与介绍_{date}.xlsx'


def rebuild_index(root, expected, excel_name='index.xlsx', excel_language='zh-CN'):
    rows = []
    for app_id in expected:
        for region in ('us','cn'):
            base = root / app_id / region
            state = read_json(base / 'manifest.json')
            if state is not None:
                app = read_json(base / 'metadata.json')
                if not isinstance(app, dict):
                    app = {}
                    if state.get('status') not in ('error', 'not_available'):
                        state = dict(state, status='error', error='Invalid cached metadata: use --refresh')
                rows.append(make_row(app_id, region, app, state))
    root.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile('w', encoding='utf-8-sig', newline='', dir=root, delete=False) as tmp:
        writer = csv.DictWriter(tmp, fieldnames=COLUMNS, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(rows)
        filename = tmp.name
    os.replace(filename, root / 'index.csv')
    dump(root / 'index.json', rows)
    write_excel(root / excel_name, COLUMNS, rows, language=excel_language)
    return rows


def main(argv=None):
    p = argparse.ArgumentParser(description='Archive public App Store metadata and images in US and CN storefronts.')
    p.add_argument('--input', default='apps.txt', help='One App ID or URL per line')
    p.add_argument('--output', default='output')
    p.add_argument('--product-name', required=True, help='Research product name used in the Excel filename')
    p.add_argument('--excel-language', choices=['zh-CN', 'en'], default='zh-CN', help='Language of Excel labels and sheet names (source content is unchanged)')
    p.add_argument('--regions', nargs='+', choices=['us','cn'], default=['us','cn'])
    p.add_argument('--interval', type=float, default=3.5, help='Minimum seconds between Lookup requests')
    p.add_argument('--retries', type=int, default=4)
    p.add_argument('--timeout', type=float, default=25)
    p.add_argument('--web-fallback', action='store_true', help='Heuristic: extra Apple CDN images from public HTML; unclassified')
    p.add_argument('--refresh', action='store_true', help='Fetch metadata/assets again')
    p.add_argument('--verbose', action='store_true')
    a = p.parse_args(argv)
    try: excel_name = workbook_filename(a.product_name)
    except ValueError as exc: p.error(str(exc))
    logging.basicConfig(level=logging.DEBUG if a.verbose else logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
    if a.interval < 3 or a.retries < 0 or a.timeout <= 0:
        p.error('--interval must be >= 3s, --retries >= 0, --timeout > 0')
    try: lines = Path(a.input).read_text('utf-8-sig').splitlines()
    except OSError as exc: p.error(str(exc))
    ids = []
    for line_no, line in enumerate(lines, 1):
        line = line.strip()
        if not line or line.startswith('#'): continue
        try:
            ident = parse_id(line)
            if ident not in ids: ids.append(ident)
        except ValueError as exc: LOG.warning('Skipped line %s (%s): %s', line_no, line, exc)
    if not ids: p.error('No valid IDs or URLs in input')
    root = Path(a.output)
    client = Client(a.interval, a.retries, a.timeout)
    for i, app_id in enumerate(ids, 1):
        for region in a.regions:
            LOG.info('[%s/%s] %s %s', i, len(ids), app_id, region)
            collect(client, app_id, region, root, a.web_fallback, a.refresh)
            rebuild_index(root, ids, excel_name, a.excel_language)
    rows = rebuild_index(root, ids, excel_name, a.excel_language)
    LOG.info('Finished: %s entries; complete=%s, partial=%s, errors/unavailable=%s', len(rows),
             sum(r['status']=='complete' for r in rows),sum(r['status']=='partial' for r in rows),
             sum(r['status'] in ('error','not_available') for r in rows))
    LOG.info('Index: %s', root / excel_name)

if __name__ == '__main__':
    main()
