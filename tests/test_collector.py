import json
import tempfile
import unittest
from http.client import IncompleteRead
from pathlib import Path
from unittest.mock import MagicMock, patch
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import collector

class FakeClient:
    def __init__(self): self.calls=[]
    def lookup(self, app_id, region):
        self.calls.append(('lookup',app_id,region))
        if region=='cn': return None
        return {'trackId':int(app_id),'kind':'software','trackName':'Demo','description':'Hello',
                'screenshotUrls':['https://is1-ssl.mzstatic.com/image/thumb/one.png'],
                'ipadScreenshotUrls':['https://is1-ssl.mzstatic.com/image/thumb/one.png'],
                'artworkUrl512':'https://is1-ssl.mzstatic.com/image/thumb/icon.png'}
    def get(self,url,api=False):
        self.calls.append(('get',url))
        return b'fake-image-data', 'image/png'

class TestCollector(unittest.TestCase):
    def test_parse(self):
        self.assertEqual(collector.parse_id('https://apps.apple.com/cn/app/example/id123456789?l=zh'), '123456789')
        self.assertEqual(collector.parse_id('123456789'), '123456789')
        with self.assertRaises(ValueError): collector.parse_id('https://example.org/id123456789')
    def test_resume_dedup_and_unavailable(self):
        with tempfile.TemporaryDirectory() as folder:
            base=Path(folder); client=FakeClient()
            r=collector.collect(client,'123456789','us',base)
            self.assertEqual(r['status'],'complete')
            self.assertEqual(r['screenshot_count'],1)
            self.assertEqual(len([x for x in client.calls if x[0]=='get']),2)
            collector.collect(client,'123456789','us',base)
            self.assertEqual(len([x for x in client.calls if x[0]=='get']),2)
            r=collector.collect(client,'123456789','cn',base)
            self.assertEqual(r['status'],'not_available')
            rows=collector.rebuild_index(base,['123456789'])
            self.assertEqual(len(rows),2)
            self.assertTrue((base/'index.csv').exists())
    def test_public_cdn_only(self):
        self.assertFalse(collector.valid_media_url('http://is1-ssl.mzstatic.com/a'))
        self.assertFalse(collector.valid_media_url('https://evil.com/a'))
        self.assertTrue(collector.valid_media_url('https://is1-ssl.mzstatic.com/image/thumb/a.png'))

    def test_unavailable_refresh_does_not_resurrect_old_cache(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder); client = FakeClient()
            collector.collect(client, '123456789', 'us', root)
            metadata = root / '123456789/us/metadata.json'
            original = metadata.read_bytes()
            with patch.object(client, 'lookup', return_value=None) as lookup:
                for refresh in (True, False):
                    row = collector.collect(client, '123456789', 'us', root, refresh=refresh)
                    self.assertEqual(row['status'], 'not_available')
                    self.assertEqual(collector.rebuild_index(root, ['123456789'])[0]['status'], 'not_available')
                self.assertEqual(lookup.call_count, 2)
            self.assertEqual(metadata.read_bytes(), original)
            # Even a failed recheck must not make historical metadata current again.
            with patch.object(client, 'lookup', side_effect=OSError('offline')):
                self.assertEqual(collector.collect(client, '123456789', 'us', root)['status'], 'error')
            with patch.object(client, 'lookup', wraps=client.lookup) as lookup:
                self.assertEqual(collector.collect(client, '123456789', 'us', root)['status'], 'complete')
                lookup.assert_called_once()

    def test_cache_run_preserves_fetch_time_and_refresh_updates_it(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder); client = FakeClient()
            first, second = '2020-01-01T00:00:00+00:00', '2026-11-03T00:00:00+00:00'
            with patch('collector.now', return_value=first):
                collector.collect(client, '123456789', 'us', root)
            with patch('collector.now', return_value=second):
                row = collector.collect(client, '123456789', 'us', root)
                self.assertEqual(row['fetched_at'], first)
                self.assertEqual(collector.rebuild_index(root, ['123456789'])[0]['fetched_at'], first)
                row = collector.collect(client, '123456789', 'us', root, refresh=True)
                self.assertEqual(row['fetched_at'], second)

    def test_legacy_cache_timestamp_is_frozen(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder); client = FakeClient()
            collector.collect(client, '123456789', 'us', root)
            manifest = root / '123456789/us/manifest.json'
            state = collector.read_json(manifest)
            state.pop('metadata_fetched_at')
            state['last_success'] = '2020-01-01T00:00:00+00:00'
            collector.dump(manifest, state)
            for _ in range(2):
                self.assertEqual(collector.collect(client, '123456789', 'us', root)['fetched_at'], state['last_success'])

    def test_invalid_metadata_does_not_abort_report_or_next_app(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            base = root / '123456789/us'; base.mkdir(parents=True)
            for content in ('[1]', 'null', '42', '"text"', '{broken'):
                with self.subTest(content=content):
                    (base / 'metadata.json').write_text(content, encoding='utf-8')
                    collector.dump(base / 'manifest.json', {'status':'complete'})
                    # Export must also tolerate corruption before collect runs.
                    self.assertEqual(collector.rebuild_index(root, ['123456789'])[0]['status'], 'error')
                    row = collector.collect(FakeClient(), '123456789', 'us', root)
                    self.assertEqual(row['status'], 'error')
                    collector.collect(FakeClient(), '987654321', 'us', root)
                    rows = collector.rebuild_index(root, ['123456789', '987654321'])
                    self.assertEqual([r['status'] for r in rows], ['error', 'complete'])

    def test_incomplete_response_read_is_retried(self):
        response = MagicMock()
        response.__enter__.return_value = response
        response.read.side_effect = [IncompleteRead(b'part', 10), b'complete']
        response.headers = {'Content-Type':'image/png'}
        with patch('collector.urlopen', return_value=response) as request, patch('collector.time.sleep'):
            self.assertEqual(collector.Client(retries=1).get('https://is1-ssl.mzstatic.com/a'),
                             (b'complete', 'image/png'))
            self.assertEqual(request.call_count, 2)

    def test_exhausted_incomplete_reads_record_failure_and_continue(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            with patch('collector.urlopen', side_effect=IncompleteRead(b'part', 10)), patch('collector.time.sleep'):
                row = collector.collect(collector.Client(retries=1), '123456789', 'us', root)
            self.assertEqual(row['status'], 'error')
            self.assertIn('IncompleteRead', row['error'])
            client = FakeClient()
            with patch.object(client, 'get', side_effect=IncompleteRead(b'part', 10)):
                row = collector.collect(client, '987654321', 'us', root)
            self.assertEqual(row['status'], 'partial')
            self.assertEqual(row['screenshot_count'], 0)
            collector.collect(FakeClient(), '555555555', 'us', root)
            rows = collector.rebuild_index(root, ['123456789', '987654321', '555555555'])
            self.assertEqual([r['status'] for r in rows], ['error', 'partial', 'complete'])
if __name__=='__main__': unittest.main()
