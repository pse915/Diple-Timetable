"""Small, dependency-free OOXML (.xlsx) reader/writer for the project's tabular template.

Limits: does not preserve workbook formatting/formulas on export, does not support .xls/.xlsm,
and intentionally imports only tabular data from sheets. No spreadsheet code is evaluated.
"""
from __future__ import annotations
import io
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr

NS = {'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
MAIN = NS['m']
PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
OFFICE_REL = NS['r']
MAX_XLSX = 15 * 1024 * 1024
MAX_EXPANDED = 75 * 1024 * 1024
MAX_CELLS = 250_000

class WorkbookError(ValueError):
    pass


def _colnum(a1):
    chars = re.match(r'^[A-Z]+', a1.upper())
    if not chars: raise WorkbookError('잘못된 셀 참조: '+str(a1))
    c = 0
    for letter in chars.group(0): c = c * 26 + ord(letter) - 64
    return c - 1


def _read_xml(z, name):
    try:
        return ET.fromstring(z.read(name))
    except (KeyError, ET.ParseError) as e:
        raise WorkbookError(f'엑셀 내부 파일 오류 ({name}): {e}') from e


def read_xlsx(path: str | Path) -> dict[str, list[list[str]]]:
    path = Path(path)
    if path.suffix.lower() != '.xlsx':
        raise WorkbookError('현재는 .xlsx 형식만 지원합니다. .xls 파일은 엑셀에서 다른 이름으로 저장하세요.')
    if path.stat().st_size > MAX_XLSX:
        raise WorkbookError('15MB보다 큰 엑셀 파일은 가져올 수 없습니다.')
    try:
        z = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError) as e:
        raise WorkbookError(f'엑셀 파일을 열 수 없습니다: {e}') from e
    with z:
        infos = z.infolist()
        if sum(i.file_size for i in infos) > MAX_EXPANDED or len(infos) > 300:
            raise WorkbookError('엑셀 내부 데이터 크기가 제한을 초과했습니다.')
        workbook = _read_xml(z, 'xl/workbook.xml')
        relroot = _read_xml(z, 'xl/_rels/workbook.xml.rels')
        relmap = {r.attrib['Id']:r.attrib['Target'] for r in relroot.findall('{'+PKG_REL+'}Relationship')}
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            ss = _read_xml(z, 'xl/sharedStrings.xml')
            for si in ss.findall('m:si',NS):
                shared.append(''.join(n.text or '' for n in si.findall('.//m:t', NS)))
        result = {}
        used = 0
        for s in workbook.findall('.//m:sheets/m:sheet', NS):
            name = s.attrib['name']
            relid = s.attrib.get('{'+OFFICE_REL+'}id')
            target = relmap.get(relid)
            if not target: raise WorkbookError('워크시트 관계 정보가 없습니다: '+name)
            target = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl',target))
            if target not in z.namelist() or '..' in target.split('/'):
                raise WorkbookError('잘못된 시트 위치: '+name)
            root = _read_xml(z, target)
            rows = []
            for row in root.findall('.//m:sheetData/m:row', NS):
                rd = {}
                for c in row.findall('m:c', NS):
                    col = _colnum(c.attrib.get('r','A1'))
                    if col >= 60: raise WorkbookError('시트 열 수가 너무 많습니다.')
                    t = c.attrib.get('t')
                    if t == 'inlineStr':
                        v = ''.join(n.text or '' for n in c.findall('.//m:t', NS))
                    else:
                        ve = c.find('m:v', NS)
                        v = '' if ve is None or ve.text is None else ve.text
                        if t == 's' and v:
                            idx = int(v)
                            if idx >= len(shared): raise WorkbookError('문자열 테이블 손상')
                            v = shared[idx]
                    rd[col] = str(v).strip()
                    used += 1
                    if used > MAX_CELLS: raise WorkbookError('셀 데이터가 너무 많습니다.')
                if rd:
                    rows.append([rd.get(i,'') for i in range(max(rd)+1)])
                else:
                    rows.append([])
                if len(rows) > 50000: raise WorkbookError('행 수가 너무 많습니다.')
            result[name] = rows
        return result


def _col_letter(idx):
    x = idx + 1; result = ''
    while x:
        x, r = divmod(x-1,26); result=chr(65+r)+result
    return result


def _xml(s):
    return escape(str(s), {'"':'&quot;',"'":'&apos;'})


def _sheet_xml(rows):
    bits = [f'<worksheet xmlns="{MAIN}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>']
    for ridx,row in enumerate(rows,1):
        cells = []
        for cidx,v in enumerate(row):
            v = '' if v is None else str(v)
            if not v: continue
            coord = f'{_col_letter(cidx)}{ridx}'
            cells.append(f'<c r="{coord}" s="{1 if ridx==1 else 0}" t="inlineStr"><is><t xml:space="preserve">{_xml(v)}</t></is></c>')
        bits.append(f'<row r="{ridx}">{"".join(cells)}</row>')
    bits.append('</sheetData></worksheet>')
    return ''.join(bits).encode('utf-8')


def write_xlsx(path: str | Path, sheets: dict[str, list[list[object]]]):
    path = Path(path)
    if not sheets or len(sheets)>25: raise WorkbookError('시트 수 제한 초과')
    for name in sheets:
        if not name or len(name)>31 or any(c in name for c in '[]:*?/\\'):
            raise WorkbookError('잘못된 시트명: '+name)
    content_types = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    workbook_sheets = []
    workbook_rels = []
    for i,name in enumerate(sheets,1):
        content_types.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        workbook_sheets.append(f'<sheet name={quoteattr(name)} sheetId="{i}" r:id="rId{i}"/>')
        workbook_rels.append(f'<Relationship Id="rId{i}" Type="{OFFICE_REL}/worksheet" Target="worksheets/sheet{i}.xml"/>')
    workbook_rels.append(f'<Relationship Id="rId{len(sheets)+1}" Type="{OFFICE_REL}/styles" Target="styles.xml"/>')
    content_types.append('</Types>')
    # Two basic styles; style 1 marks header bold.
    styles = f'''<styleSheet xmlns="{MAIN}"><fonts count="2"><font><sz val="11"/><name val="맑은 고딕"/></font><font><b/><sz val="11"/><name val="맑은 고딕"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>'''
    path.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as z:
        z.writestr('[Content_Types].xml',''.join(content_types))
        z.writestr('_rels/.rels',f'<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr('xl/workbook.xml',f'<workbook xmlns="{MAIN}" xmlns:r="{OFFICE_REL}"><sheets>{"".join(workbook_sheets)}</sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels',f'<Relationships xmlns="{PKG_REL}">{"".join(workbook_rels)}</Relationships>')
        z.writestr('xl/styles.xml',styles)
        for i,rows in enumerate(sheets.values(),1):
            z.writestr(f'xl/worksheets/sheet{i}.xml', _sheet_xml(rows))
