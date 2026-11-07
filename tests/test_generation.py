"""Exercise the real generator with local CSVs and isolated external services."""
import contextlib
import csv
import datetime
import io
import json
from pathlib import Path
import re
import runpy
import shutil
import tempfile
import unittest
from unittest.mock import patch
from urllib.parse import urlparse

import requests

ROOT = Path(__file__).resolve().parents[1]
CATEGORIES = ['2026年秋アニメ', '2026年夏アニメ', '2026年春アニメ', '2026年冬アニメ', '2025年秋アニメ']


class GenerationTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.work = Path(self.tmp.name)
        shutil.copy2(ROOT / 'update_list.py', self.work)
        shutil.copytree(ROOT / 'viewer_template', self.work / 'viewer_template')
        shutil.copy2(ROOT / 'cool_analysis.csv', self.work)
        self.catalog = (self.work / 'cool_analysis.csv').read_bytes()
        self.expected = {}
        current = None
        for row in csv.reader(io.StringIO(self.catalog.decode('utf-8-sig'))):
            if row and row[0] in CATEGORIES:
                current = row[0]
                self.expected[current] = set()
            elif current and row and any(row) and row[0] != '作品名':
                self.expected[current].add(tuple(row[:4]))
        self.autumn = next(iter(sorted(self.expected[CATEGORIES[0]])))
        anime, kind, artist, song = self.autumn
        today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).strftime('%Y/%m/%d')
        self.dates = ['2026/10/01', '2026/12/31 23:59:59', '2027/01/01', today]
        with (self.work / 'history.csv').open('w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['取得日', '部屋主', '順番', '曲名（ファイル名）', '作品名', '歌手名', '歌った人'])
            for order, date in enumerate(self.dates, 1):
                writer.writerow([date, 'つぼはち部屋', order, song, anime, artist, '歌唱者A'])
        (self.work / 'index.html').write_text('previous index')
        (self.work / 'viewer.html').write_text('previous viewer')
        self.catalog_down = False
        self.history_down = False
        self.writes = []

    def response(self, session, method, url, **kwargs):
        if method.upper() != 'GET':
            self.writes.append(method)
            raise AssertionError('Unexpected external write')
        result = requests.Response()
        result.url = url
        result.status_code = 404
        result._content = b''
        host = urlparse(url).hostname
        if host == 'script.google.com':
            filename = (kwargs.get('params') or {}).get('filename', '')
            if filename == 'cool_analysis.csv' and not self.catalog_down:
                result.status_code = 200
                result._content = self.catalog + '\n2027年冬アニメ,,,\n混入禁止,OP,歌手,混入禁止曲\n'.encode()
            elif self.history_down and filename.startswith('history'):
                result.status_code = 503
        elif host == 'raw.githubusercontent.com':
            # Offline creation counts are not required for these history consistency checks.
            pass
        elif host != 'ykr.moe':
            raise AssertionError(f'Unexpected request host: {host}')
        result.encoding = 'utf-8-sig'
        return result

    def generate(self):
        import os
        previous = Path.cwd()
        os.chdir(self.work)
        self.addCleanup(os.chdir, previous)
        # Patch the normal Requests transport; execute the repository's unchanged entry point.
        with patch.object(requests.Session, 'request', lambda session, method, url, **kw: self.response(session, method, url, **kw)), contextlib.redirect_stdout(io.StringIO()):
            return runpy.run_path(str(self.work / 'update_list.py'), run_name='__main__')

    def check_catalog(self, result):
        self.assertEqual(result['ALLOWED_CATEGORIES'], CATEGORIES)
        self.assertEqual(result['TREND_TARGET_CATEGORY'], CATEGORIES[0])
        self.assertEqual(list(result['categorized_data']), CATEGORIES)
        for category, expected in self.expected.items():
            actual = {(r['anime'], r['type'], r['artist'], r['song']) for r in result['categorized_data'][category]}
            self.assertEqual(actual, expected, category)
        self.assertEqual(len(result['categorized_data'][CATEGORIES[0]]), 150)
        for name in ['index.html', 'viewer.html']:
            html = (self.work / name).read_text()
            cats = json.loads(re.search(r'const CATS = ([^\n]+);', html).group(1))
            self.assertEqual(cats, CATEGORIES)
            self.assertIn('推移の対象: 2026年秋アニメ', html)
            self.assertIn('2026/01/01〜2026/12/31', html)
            self.assertNotRegex(html, r'__(?:COOL|RANK|TREND|HISTORY|COUNT_PERIOD)\w*__')
        self.assertEqual(self.writes, [])

    def test_autumn_catalog_history_ranking_and_trend(self):
        result = self.generate()
        self.check_catalog(result)
        songs = [s for w in result['cool_data_for_js'][CATEGORIES[0]]['works'] for s in w['songs']]
        match = next(s for s in songs if (s['anime'], s['type'], s['artist'], s['song']) == self.autumn)
        expected_count = sum('2026/01/01' <= d < '2027/01/01' for d in self.dates)
        self.assertEqual(match['count'], expected_count)
        self.assertEqual(match['user_count'], 1)
        rank = next(s for s in result['ranking_data_by_cat'][CATEGORIES[0]] if s['song'] == self.autumn[3] and s['artist'] == self.autumn[2])
        self.assertEqual(rank['count'], expected_count)
        now = result['now_dt']
        recent_count = sum(now - result['pd'].Timedelta(days=7) <= result['pd'].Timestamp(d) < now for d in self.dates)
        trend = next(s for s in result['trend_data_for_js']['7']['items'] if s['song'] == self.autumn[3])
        self.assertEqual(trend['cur_count'], recent_count)
        self.assertEqual(trend['all_count'], expected_count)
        self.assertEqual(len(result['history_for_js']), len(self.dates))

    def test_gas_outage_uses_complete_local_catalog(self):
        self.catalog_down = True
        self.check_catalog(self.generate())

    def test_incomplete_history_preserves_existing_outputs(self):
        self.history_down = True
        before = (self.work / 'history.csv').read_bytes()
        with self.assertRaisesRegex(RuntimeError, '履歴の読み込みが不完全'):
            self.generate()
        self.assertEqual((self.work / 'index.html').read_text(), 'previous index')
        self.assertEqual((self.work / 'viewer.html').read_text(), 'previous viewer')
        self.assertEqual((self.work / 'history.csv').read_bytes(), before)
        self.assertEqual(self.writes, [])


if __name__ == '__main__':
    unittest.main()
