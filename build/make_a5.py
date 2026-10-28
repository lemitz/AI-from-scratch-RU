"""Derive the A5 booklet edition from the A4 source .docx.

A5 is A4 scaled by 1/sqrt(2), so all geometry (page, margins, tables, images) is scaled
down proportionally, while fonts shrink less (x0.9) to stay readable in print.
Usage: python make_a5.py <A4.docx> <A5.docx>
"""
import re
import sys
import zipfile

from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

# A4 text widths (twips) in the source
A4_PORTRAIT_W = 11906 - 1247 - 1247          # 9412
A4_LANDSCAPE_W = 16838 - 1020 - 1020         # 14798

# A5 page setup (twips)
PORTRAIT = dict(w=8391, h=11906, top=794, bottom=794, left=851, right=624, header=397, footer=340)
LANDSCAPE = dict(w=11906, h=8391, top=624, bottom=567, left=624, right=624, header=284, footer=284)
A5_PORTRAIT_W = PORTRAIT['w'] - PORTRAIT['left'] - PORTRAIT['right']      # 6916
A5_LANDSCAPE_W = LANDSCAPE['w'] - LANDSCAPE['left'] - LANDSCAPE['right']  # 10658
A5_LANDSCAPE_H = LANDSCAPE['h'] - LANDSCAPE['top'] - LANDSCAPE['bottom']  # 7200

FX = A5_PORTRAIT_W / A4_PORTRAIT_W      # ~0.735, tables & portrait images
FX_LAND = A5_LANDSCAPE_W / A4_LANDSCAPE_W
FONT = 0.9
SPACE = 0.8
INDENT = 0.85
EMU_PER_TWIP = 635


def scale_attr(el, attr, k, minimum=0):
    v = el.get(W + attr)
    if v is not None and re.fullmatch(r'-?\d+', v):
        el.set(W + attr, str(max(minimum, round(int(v) * k))))


def fix_section(sect):
    sz = sect.find(W + 'pgSz')
    landscape = sz.get(W + 'orient') == 'landscape'
    g = LANDSCAPE if landscape else PORTRAIT
    sz.set(W + 'w', str(g['w']))
    sz.set(W + 'h', str(g['h']))
    mar = sect.find(W + 'pgMar')
    for k in ('top', 'bottom', 'left', 'right', 'header', 'footer'):
        mar.set(W + k, str(g[k]))
    mar.set(W + 'gutter', '0')


def transform_document(root):
    for sect in root.iter(W + 'sectPr'):
        fix_section(sect)
    for tag in ('sz', 'szCs'):
        for el in root.iter(W + tag):
            scale_attr(el, 'val', FONT, minimum=13)
    for el in root.iter(W + 'spacing'):
        scale_attr(el, 'before', SPACE)
        scale_attr(el, 'after', SPACE)
    for el in root.iter(W + 'ind'):
        for a in ('left', 'right', 'hanging', 'firstLine', 'start', 'end'):
            scale_attr(el, a, INDENT)
    for el in root.iter(W + 'tcW'):
        scale_attr(el, 'w', FX)
    for el in root.iter(W + 'gridCol'):
        scale_attr(el, 'w', FX)
    for mar in root.iter(W + 'tcMar'):
        for side in mar:
            scale_attr(side, 'w', SPACE)
    for el in root.iter(W + 'tab'):
        scale_attr(el, 'pos', FX)
    for inline in root.iter(WP + 'inline'):
        ext = inline.find(WP + 'extent')
        cx, cy = int(ext.get('cx')), int(ext.get('cy'))
        if cx > A4_PORTRAIT_W * EMU_PER_TWIP * 1.05:     # landscape figure page
            k = FX_LAND
            max_h = (A5_LANDSCAPE_H - 700) * EMU_PER_TWIP  # leave room for the caption
            if cy * k > max_h:
                k = max_h / cy
        else:
            k = FX
        ncx, ncy = str(round(cx * k)), str(round(cy * k))
        ext.set('cx', ncx)
        ext.set('cy', ncy)
        for x in inline.iter(A + 'ext'):
            x.set('cx', ncx)
            x.set('cy', ncy)


def transform_styles(root):
    for tag in ('sz', 'szCs'):
        for el in root.iter(W + tag):
            scale_attr(el, 'val', FONT, minimum=13)
    for el in root.iter(W + 'tab'):
        scale_attr(el, 'pos', FX)
    for el in root.iter(W + 'ind'):
        for a in ('left', 'hanging', 'firstLine'):
            scale_attr(el, a, INDENT)
    # tighter table of contents so it fits on fewer A5 pages
    for st in root.iter(W + 'style'):
        if st.get(W + 'styleId') in ('TOC1', 'TOC2'):
            for sp in st.iter(W + 'spacing'):
                scale_attr(sp, 'before', 0.5)
                scale_attr(sp, 'after', 0.5)
                sp.set(W + 'line', '240')


def transform_hdrftr(root):
    for tag in ('sz', 'szCs'):
        for el in root.iter(W + tag):
            scale_attr(el, 'val', FONT, minimum=13)
    for jc in root.iter(W + 'jc'):     # mirrored pages: keep running header centred
        jc.set(W + 'val', 'center')


def transform_settings(root):
    if root.find(W + 'mirrorMargins') is None:
        el = etree.Element(W + 'mirrorMargins')
        root.find(W + 'zoom').addnext(el)


def main(src, dst):
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            name = item.filename
            fn = None
            if name == 'word/document.xml':
                fn = transform_document
            elif name == 'word/styles.xml':
                fn = transform_styles
            elif re.match(r'word/(header|footer)\d+\.xml$', name):
                fn = transform_hdrftr
            elif name == 'word/settings.xml':
                fn = transform_settings
            if fn:
                root = etree.fromstring(data)
                fn(root)
                data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
            zout.writestr(item, data)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
