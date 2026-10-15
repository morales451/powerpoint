"""Layout polish on the tightened original deck: margins, column balance,
readable type, consistent bullets, clearer titles, filled gaps, color legend."""
import sys, re, copy
from lxml import etree
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
NS = {'a': A[1:-1], 'p': P[1:-1]}
root = sys.argv[1]

MARGIN, WIDTH = 304800, 11582400               # 1/3" side margins on a 13.33" slide
COLS = [2300000, 900000, 3900000, 4482400]      # narrow the color-only column, widen Notes
TITLES = {2: 'Leaky Bucket: Service and Logistics', 3: 'Leaky Bucket: Material, Systems and National Accounts'}
FILL = {  # empty Fix cells -> text, keyed by the row's detractor text
    'No notifications/tracking': 'Covered by the Logistics fixes above',
    'Inflexible on scope/specs': 'None needed; protect this strength',
    'Missing system components': 'Assign an owner to close the gaps',
}
HEAD = {'NOTES': 'Notes', 'FIX': 'Fix'}
LEGEND = [('00B050', 'Strong'), ('FFFF00', 'Fair'), ('FFC000', 'Weak'), ('FF9999', 'Poor')]

def text(el): return ''.join(t.text or '' for t in el.iter(A + 't'))

def legend_xml(next_id, x, y, label):
    parts, cx = [], x
    def box(i, x, w, fill, txt, bold=False):
        fillxml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill><a:ln w="6350"><a:solidFill><a:srgbClr val="7F7F7F"/></a:solidFill></a:ln>' if fill else '<a:noFill/><a:ln><a:noFill/></a:ln>'
        return (f'<p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:nvSpPr><p:cNvPr id="{i}" name="Legend {i}"/><p:cNvSpPr txBox="{0 if fill else 1}"/><p:nvPr/></p:nvSpPr>'
                f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="228600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fillxml}</p:spPr>'
                f'<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"{" b=\"1\"" if bold else ""}><a:solidFill><a:srgbClr val="404040"/></a:solidFill></a:rPr><a:t>{txt}</a:t></a:r></a:p></p:txBody></p:sp>')
    parts.append(box(next_id, cx, 1300000, None, label, True)); cx += 1350000; next_id += 1
    for fill, name in LEGEND:
        parts.append(box(next_id, cx, 228600, fill, '')); next_id += 1
        parts.append(box(next_id, cx + 300000, 650000, None, name)); next_id += 1
        cx += 1000000
    return [etree.fromstring(p) for p in parts]

def add_legend(tree, x, y, label):
    spTree = tree.find('.//p:cSld/p:spTree', NS)
    ids = [int(e.get('id')) for e in tree.iter(P + 'cNvPr')]
    for el in legend_xml(max(ids) + 1, x, y, label): spTree.append(el)

for n in (2, 3):
    path = f'{root}/ppt/slides/slide{n}.xml'
    tree = etree.parse(path)
    # Title: clearer, left-aligned with the table
    title = [sp for sp in tree.iter(P + 'sp') if sp.find('.//p:ph[@type="title"]', NS) is not None][0]
    title.find('.//a:off', NS).set('x', str(MARGIN))
    runs = title.findall('.//a:r', NS); runs[0].find(A + 't').text = TITLES[n]
    for r in runs[1:]: r.getparent().remove(r)
    # Table frame: inside the margins, same top on both slides
    gf = tree.find('.//p:graphicFrame', NS)
    gf.find('p:xfrm/a:off', NS).set('x', str(MARGIN)); gf.find('p:xfrm/a:off', NS).set('y', '815004')
    gf.find('p:xfrm/a:ext', NS).set('cx', str(WIDTH))
    for col, w in zip(gf.iter(A + 'gridCol'), COLS): col.set('w', str(w))
    # Readable type and one bullet indent
    for rpr in gf.iter(A + 'rPr', A + 'endParaRPr'):
        if rpr.get('sz') == '1000': rpr.set('sz', '1100')
    for ppr in gf.iter(A + 'pPr'):
        if ppr.get('marL') == '285750': ppr.set('marL', '171450'); ppr.set('indent', '-171450')
    # Header case and empty Fix cells
    for row in gf.iter(A + 'tr'):
        cells = row.findall(A + 'tc')
        for c in cells:
            if text(c) in HEAD:
                c.find('.//a:t', NS).text = HEAD[text(c)]
        key = text(cells[0])
        if key in FILL and not text(cells[3]).strip():
            p = cells[3].find('.//a:p', NS)
            r = etree.SubElement(p, A + 'r'); rp = etree.SubElement(r, A + 'rPr'); rp.set('lang', 'en-US'); rp.set('sz', '1100')
            etree.SubElement(r, A + 't').text = FILL[key]
            end = p.find(A + 'endParaRPr')
            if end is not None: p.remove(end); p.append(end)
    add_legend(tree, MARGIN, 6250000, "Henry's position:")
    tree.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)

# Slide 1: legend for the tree colors
path = f'{root}/ppt/slides/slide1.xml'
tree = etree.parse(path)
add_legend(tree, MARGIN, 6250000, 'Color key:')
tree.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)
print('polished')
