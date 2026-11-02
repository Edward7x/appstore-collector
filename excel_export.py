from excel_locale import translate, SHEETS_EN
"""Seven-sheet App Store research workbook with embedded image galleries."""
import io
import json
import os
import re
import tempfile
from pathlib import Path
from collections import Counter
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
SHEETS = ['研究总览', '宣传图预览', '全部宣传图', '应用介绍', '宣传图索引', '来源数据', '采集说明']
COLORS = ['89603D', '396953', '42627D', '715D86']

def text(cell, value):
    if isinstance(value, str):
        value = re.sub('[\\x00-\\x08\\x0b\\x0c\\x0e-\\x1f]', '', value)
    cell.value = value
    if isinstance(value, str):
        cell.data_type = 's'
    cell.font = Font(name='微软雅黑', size=11, color='263648')
    cell.alignment = Alignment(vertical='top', wrap_text=True)

def band(ws, row, title, end, color='24476A'):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end)
    c = ws.cell(row, 1)
    text(c, title)
    c.fill = PatternFill('solid', fgColor=color)
    c.font = Font(name='微软雅黑', size=14, bold=True, color='FFFFFF')
    ws.row_dimensions[row].height = 30

def table(ws, headers, values, widths):
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    for r, items in enumerate([headers, *values], 1):
        ws.row_dimensions[r].height = 32 if r == 1 else 54
        for c, value in enumerate(items, 1):
            cell = ws.cell(r, c)
            text(cell, value)
            if r == 1:
                cell.fill = PatternFill('solid', fgColor='24476A')
                cell.font = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
            elif r % 2 == 0:
                cell.fill = PatternFill('solid', fgColor='F0F5FA')
    ws.freeze_panes = 'C2'
    ws.auto_filter.ref = ws.dimensions

def link(cell, target):
    cell.hyperlink = target
    cell.font = Font(name='微软雅黑', size=11, color='167D9A', underline='single')

def local_file(base, relative):
    path = (base / relative).resolve()
    return path if path.is_relative_to(base.resolve()) else None

def embed(ws, path, anchor, width, height, language='zh-CN'):
    try:
        if path is None:
            raise OSError('missing path')
        with PILImage.open(path) as original:
            original.thumbnail((width, height))
            image = original.convert('RGBA')
            stream = io.BytesIO()
            image.save(stream, format='PNG')
        stream.seek(0)
        pic = Image(stream)
        ws.add_image(pic, anchor)
        return True
    except (OSError, ValueError):
        text(ws[anchor], translate('图片缺失或无法解码，请查看宣传图索引', language))
        return False

def write_excel(path, columns, rows, language='zh-CN'):
    if language not in ('zh-CN', 'en'):
        raise ValueError('Unsupported Excel language: ' + language)
    path = Path(path)
    wb = Workbook()
    wb.remove(wb.active)
    for name in (SHEETS_EN if language == 'en' else SHEETS):
        ws = wb.create_sheet(name)
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 85
    overview, preview, gallery, intro, index, raw, notes = wb.worksheets
    for ws in (preview, gallery):
        for col, width in {'A': 38, 'B': 32, 'C': 3, 'D': 32, 'E': 3, 'F': 32, 'G': 38}.items():
            ws.column_dimensions[col].width = width
        ws.freeze_panes = 'B4'
        band(ws, 1, ws.title + translate(' · App Store 研究档案', language), 7)
        ws.merge_cells('A2:G2')
        text(ws['A2'], translate('按应用 / 地区分组；图片已嵌入。预览优先前三张 iPhone 图；全部图集包含各设备及未分类图片。', language))
        ws.row_dimensions[2].height = 32
    overview_rows, image_rows, intro_rows = ([], [], [])
    gallery_row = 4
    for number, row in enumerate(rows, 1):
        ident, region = (str(row['app_id']), row['region'])
        base = path.parent / ident / region
        try:
            state = json.loads((base / 'manifest.json').read_text('utf-8'))
        except (OSError, ValueError):
            state = {}
        files = list(state.get('files', {}).values())
        media = [f for f in files if f.get('kind') != 'icon']
        counts = Counter((f.get('kind', 'unknown') for f in media if f.get('ok')))
        title = f"{number}. {row.get('trackName') or ident} | {region.upper()}"
        start = 4 + (number - 1) * 17
        selected = [f for f in media if f.get('kind') == 'iphone'][:3] or media[:3]
        icon = next((f for f in files if f.get('kind') == 'icon' and f.get('ok')), None)

        def block(ws, top, pictures, heading):
            band(ws, top, heading, 7, COLORS[(number - 1) % len(COLORS)])
            for r in range(top + 1, top + 16):
                ws.row_dimensions[r].height = 24
            if icon:
                embed(ws, local_file(base, icon['path']), f'A{top + 1}', 64, 64, language=language)
            info = [f"{translate('类别：', language)}{row.get('primaryGenreName', '')}", f"{translate('评分：', language)}{row.get('averageUserRating', '')}（{row.get('userRatingCount', '')}）", f"{translate('版本：', language)}{row.get('version', '')}", translate('宣传图：', language) + ' / '.join((f'{k} {v}' for k, v in counts.items())), f"{translate('介绍：', language)}{len(row.get('description', ''))}{translate(' 字', language)}", f"{translate('状态：', language)}{row.get('status', '')}"]
            for offset, value in enumerate(info, 5):
                text(ws.cell(top + offset, 1), value)
            for col, item in zip(('B', 'D', 'F'), pictures):
                embed(ws, local_file(base, item.get('path', '__missing__')), f'{col}{top + 1}', 215, 360, language=language)
                text(ws[f'{col}{top + 13}'], item.get('kind', '') + ' · ' + item.get('source', ''))
            if not pictures:
                text(ws.cell(top + 2, 2), translate('暂无宣传图', language))
            text(ws.cell(top + 1, 7), translate('App Store 页面', language))
            if row.get('trackViewUrl'):
                link(ws.cell(top + 1, 7), row['trackViewUrl'])
            text(ws.cell(top + 3, 7), translate('查看宣传图索引', language))
            link(ws.cell(top + 3, 7), translate("#'宣传图索引'!A1", language))
        block(preview, start, selected, title)
        gallery_start = gallery_row
        for offset in range(0, max(1, len(media)), 3):
            block(gallery, gallery_row, media[offset:offset + 3], title + f"{translate(' · 图 ', language)}{offset + 1}–{min(offset + 3, len(media))}" if media else title)
            gallery_row += 17
        text(preview.cell(start + 5, 7), translate('查看全部宣传图', language))
        link(preview.cell(start + 5, 7), f"""{translate("#'全部宣传图'!A", language)}{gallery_start}""")
        overview_rows.append([ident, region.upper(), row.get('trackName', ''), row.get('artistName', ''), row.get('primaryGenreName', ''), row.get('averageUserRating', ''), row.get('userRatingCount', ''), row.get('formattedPrice', ''), row.get('version', ''), sum(counts.values()), row.get('status', ''), translate('查看预览', language), row.get('trackViewUrl', '')])
        for seq, item in enumerate(media, 1):
            p = local_file(base, item.get('path', '__missing__'))
            image_rows.append([ident, region.upper(), row.get('trackName', ''), seq, item.get('kind', ''), item.get('source', ''), translate('可用', language) if item.get('ok') and p and p.is_file() else translate('缺失 / 失败', language), str(p or ''), item.get('url', ''), item.get('sha256', ''), item.get('error', '')])
        for field, label in [('description', translate('应用介绍', language)), ('releaseNotes', translate('更新说明', language))]:
            content = row.get(field, '') or ''
            chunks = [content[i:i + 900] for i in range(0, len(content), 900)] or ['']
            for part, chunk in enumerate(chunks, 1):
                intro_rows.append([ident, region.upper(), row.get('trackName', ''), label, part, chunk])
    table(overview, ['App ID', translate('地区', language), translate('应用名称', language), translate('开发者', language), translate('类别', language), translate('评分', language), translate('评价数', language), translate('价格', language), translate('版本', language), translate('宣传图数', language), translate('状态', language), translate('宣传图预览', language), translate('商店链接', language)], overview_rows, [18, 10, 36, 30, 22, 12, 14, 16, 16, 14, 20, 18, 45])
    for i, row in enumerate(rows, 2):
        link(overview.cell(i, 12), f"""{translate("#'宣传图预览'!A", language)}{4 + (i - 2) * 17}""")
        if row.get('trackViewUrl'):
            link(overview.cell(i, 13), row['trackViewUrl'])
    table(intro, ['App ID', translate('地区', language), translate('应用名称', language), translate('内容类型', language), translate('分段', language), translate('正文', language)], intro_rows, [18, 10, 32, 18, 10, 110])
    for r in range(2, intro.max_row + 1):
        intro.row_dimensions[r].height = 330
    table(index, ['App ID', translate('地区', language), translate('应用名称', language), translate('序号', language), translate('设备', language), translate('来源', language), translate('本地状态', language), translate('本地文件', language), translate('原图地址', language), 'SHA256', translate('错误', language)], image_rows, [18, 10, 32, 10, 22, 20, 18, 60, 65, 40, 40])
    for r in range(2, index.max_row + 1):
        if index.cell(r, 8).value:
            link(index.cell(r, 8), Path(index.cell(r, 8).value).as_uri())
        if index.cell(r, 9).value:
            link(index.cell(r, 9), index.cell(r, 9).value)
    table(raw, columns, [[r.get(c, '') for c in columns] for r in rows], [65 if c in ('description', 'releaseNotes') else 28 for c in columns])
    table(notes, [translate('项目', language), translate('说明', language)], [[translate('工作表导航', language), translate('研究总览：比较与导航；宣传图预览：前三张图；全部宣传图：完整图集；应用介绍：分段全文；宣传图索引：原图与文件；来源数据：原始汇总字段。', language)], [translate('采集范围', language), translate('公开 iTunes Lookup API；US/CN 分别记录，不混合不同地区。未返回不代表应用不存在。', language)], [translate('图片范围', language), translate('宣传图是接口暴露的截图；web-unclassified 是网页启发式发现的未分类图片，不能保证是宣传图。', language)], [translate('图片与链接', language), translate('缩略图嵌入工作簿，保持比例；高清图片见索引。本地链接依赖原目录，网络链接依赖网络与源站。', language)], [translate('数据完整性', language), translate('缺失图片显示占位文字。来源数据的 Excel 单元格最多 32767 字；完整原始内容见 metadata.json；应用介绍分段保存全文。', language)], [translate('统计口径', language), translate('总览按 manifest 中下载成功的非图标文件计数，索引另行检查本地是否存在；可能包括历史下载文件。', language)], [translate('采集时间', language), translate('每条记录的 fetched_at 见来源数据，采用 UTC。重新运行采集会更新工作簿。', language)], [translate('使用说明', language), translate('筛选总览或索引；点击总览进入预览。工作簿用于研究，素材权利属于原权利人。', language)]], [24, 110])
    for r in range(2, notes.max_row + 1):
        notes.row_dimensions[r].height = 65
    with tempfile.NamedTemporaryFile(dir=path.parent, suffix='.xlsx', delete=False) as tmp:
        temporary = tmp.name
    try:
        wb.save(temporary)
        os.replace(temporary, path)
    finally:
        wb.close()
        if os.path.exists(temporary):
            os.unlink(temporary)
