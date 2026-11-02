import json
import tempfile
import unittest
from pathlib import Path
from PIL import Image
from openpyxl import load_workbook
from collector import COLUMNS, rebuild_index
from excel_export import SHEETS, write_excel
from excel_locale import SHEETS_EN


class TestExcelLanguage(unittest.TestCase):
    def test_both_languages_preserve_content_images_and_links(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            base = root / '123456' / 'us'
            base.mkdir(parents=True)
            Image.new('RGB', (20, 40), 'green').save(base / 'image.png')
            state = {'status': 'complete', 'files': {
                'one': {'kind':'iphone', 'ok':True, 'path':'image.png', 'source':'lookup'},
                'two': {'kind':'ipad', 'ok':False, 'path':'missing.png', 'source':'lookup'},
            }}
            app = {'trackName':'原始名称', 'description':'原始介绍 unchanged', 'releaseNotes':'更新原文'}
            (base / 'manifest.json').write_text(json.dumps(state), encoding='utf-8')
            (base / 'metadata.json').write_text(json.dumps(app), encoding='utf-8')
            for language, names in [('zh-CN', SHEETS), ('en', SHEETS_EN)]:
                with self.subTest(language=language):
                    filename = language + '.xlsx'
                    rebuild_index(root, ['123456'], filename, language)
                    wb = load_workbook(root / filename)
                    self.assertEqual(wb.sheetnames, names)
                    self.assertEqual(wb.worksheets[0]['C2'].value, app['trackName'])
                    self.assertEqual(wb.worksheets[3]['F2'].value, app['description'])
                    self.assertEqual(wb.worksheets[5].cell(2,COLUMNS.index('description')+1).value, app['description'])
                    self.assertEqual(len(wb.worksheets[1]._images), 1)
                    self.assertEqual(len(wb.worksheets[2]._images), 1)
                    for ws in wb:
                        for row in ws:
                            for cell in row:
                                if cell.hyperlink and cell.hyperlink.target.startswith('#'):
                                    target = cell.hyperlink.target.split("'")[1]
                                    self.assertIn(target, names)
                    if language == 'en':
                        self.assertEqual(wb.worksheets[0]['B1'].value, 'Region')
                        self.assertEqual(wb.worksheets[2]['D5'].value, 'Image missing or unreadable; see Image Index')
                        self.assertEqual(wb.worksheets[6]['A2'].value, 'Navigation')
                    wb.close()

    def test_invalid_language(self):
        with self.assertRaises(ValueError):
            write_excel(Path('unused.xlsx'), COLUMNS, [], language='fr')
