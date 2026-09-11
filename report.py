"""Native potential-impact report from the computed dependency graph."""
from collections import defaultdict
from PIL import Image, ImageDraw, ImageFont

BG, INK, MUTED = '#f4f1e9', '#202b35', '#5b656f'
RED, BLUE, PANEL = '#a73932', '#226873', '#ffffff'


def font(size):
    try: return ImageFont.truetype('DejaVuSans.ttf', size)
    except OSError: return ImageFont.load_default(size=size)


def text(d, x, y, value, size=30, fill=INK, width=880):
    value = str(value); face = font(size)
    if d.textlength(value, font=face) > width:
        while value and d.textlength(value+'...', font=face) > width: value = value[:-1]
        value += '...'
    d.text((x,y), value, font=face, fill=fill)


def render(plan, output):
    top = plan['impacts'][0] if plan['impacts'] else None
    groups = defaultdict(list)
    if top:
        for cell in top['reachable_formulas']:
            groups[plan['formulas'][cell]['sheet']].append(cell)
    visibility = {s['name']: s['visibility'] for s in plan['sheets']}
    height = 920 + 155*len(groups)
    im=Image.new('RGB',(1000,height),BG); d=ImageDraw.Draw(im)
    text(d,50,35,'WORKBOOK RISK MAP',24,BLUE)
    text(d,50,100,'One cell. How far',59)
    text(d,50,172,'can its impact reach?',59)
    text(d,50,268,top['cell'] if top else 'No issue seeds found',37,RED)
    count = len(top['reachable_formulas']) if top else 0
    text(d,50,332,f'{count} reachable formula cells',43,BLUE)
    text(d,50,398,'Potential impact along parsed references',26,MUTED)
    text(d,50,448,'Grouped by sheet, not dependency order',24,MUTED)
    y=514
    for sheet, cells in sorted(groups.items()):
        d.rounded_rectangle((50,y,950,y+126),radius=18,fill=PANEL)
        text(d,75,y+16,sheet+' / '+visibility[sheet],29,INK,655)
        text(d,780,y+16,str(len(cells)),39,BLUE,130)
        # Complete cell paths remain in audit.json; the image is a summary.
        text(d,75,y+69,', '.join(plan['formulas'][c]['cell'] for c in cells),27,MUTED,820)
        y += 155
    text(d,50,y+24,f"{len(plan['cycles'])} cycle group(s) found",32,RED)
    text(d,50,y+81,f"{plan['coverage']['unresolved_items']} unresolved reference/function item(s)",30,RED)
    text(d,50,y+143,'Incomplete paths are reported, never silently ignored.',25,MUTED)
    text(d,50,y+212,'Potential impact is not a count of wrong results.',26,INK)
    text(d,50,y+266,'All cells, edges and limitations are in audit.json.',25,MUTED)
    im.save(output/'impact.png')
