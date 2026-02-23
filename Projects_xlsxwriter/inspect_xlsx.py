import zipfile
import xml.etree.ElementTree as ET
import os
from pathlib import Path


def strip_ns(tag):
    return tag.split('}', 1)[-1] if '}' in tag else tag


def find_child_by_localname(parent, localname):
    for ch in parent:
        if strip_ns(ch.tag) == localname:
            return ch
    return None


def parse_xlsx(path):
    path = Path(path)
    if not path.exists():
        print('Arquivo não encontrado:', path)
        return

    with zipfile.ZipFile(path, 'r') as z:
        # Ler sheet1
        sheet_xml = z.read('xl/worksheets/sheet1.xml')
        sheet_root = ET.fromstring(sheet_xml)

        cells = {}
        for c in sheet_root.findall('.//'):
            if strip_ns(c.tag) == 'c':
                r = c.attrib.get('r')
                s = c.attrib.get('s')
                cells[r] = s

        # Ler styles
        styles = {}
        if 'xl/styles.xml' in z.namelist():
            styles_xml = z.read('xl/styles.xml')
            styles_root = ET.fromstring(styles_xml)

            fills = []
            fills_el = find_child_by_localname(styles_root, 'fills')
            if fills_el is not None:
                for f in fills_el:
                    # procurar patternFill/fgColor
                    pattern = None
                    for p in f:
                        if strip_ns(p.tag) == 'patternFill':
                            fg = None
                            for item in p:
                                if strip_ns(item.tag) in ('fgColor', 'bgColor'):
                                    fg = item.attrib
                            pattern = (strip_ns(p.tag), fg)
                    fills.append(pattern)

            # map xfs -> fillId
            xfs = []
            cellxfs = find_child_by_localname(styles_root, 'cellXfs')
            if cellxfs is not None:
                for xf in cellxfs:
                    fillId = xf.attrib.get('fillId')
                    xfs.append(fillId)

            styles['fills'] = fills
            styles['xfs'] = xfs

        # Mostrar resultados para algumas células
        print('Inspeção de:', path)
        for coord in ('A1', 'B1', 'C1', 'A2', 'B2'):
            s = cells.get(coord)
            if s is None:
                print(f'{coord}: sem valor ou sem estilo')
            else:
                try:
                    si = int(s)
                except Exception:
                    si = None
                print(f'{coord}: style_index={s}')
                if styles:
                    if si is not None and si < len(styles.get('xfs', [])):
                        fillId = styles['xfs'][si]
                        print(f'  -> xf index {si} -> fillId {fillId}')
                        if fillId is not None:
                            fi = int(fillId)
                            if fi < len(styles['fills']):
                                print('  -> fill:', styles['fills'][fi])
                            else:
                                print('  -> fillId fora do range')
                    else:
                        print('  -> sem xf aplicável')


if __name__ == '__main__':
    xlsx_path = Path.cwd() / 'Projects_xlsxwriter' / 'dados.xlsx'
    parse_xlsx(xlsx_path)
