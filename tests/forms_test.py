"""Form checks on MAGIC SHOW.pdf (a field with 3 copies: 1 on page 1, 2 on page 2), like filling it in the app:
  python tests/forms_test.py
- an emptied field stays empty in every copy (the old value never comes back)
- after changing a field many times, its text is centered in its box in every copy, every time,
  in our own render and in PDFium's (Chrome's engine), after saving and reopening."""
import os, sys, shutil, tempfile
T = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(T)
sys.path.insert(0, ROOT)
import server, pymupdf
SRC = os.path.join(T, 'pdfs', 'MAGIC SHOW.pdf')
if not os.path.exists(SRC): print('MAGIC SHOW.pdf is not in tests/pdfs: skipped'); sys.exit(0)
c = server.app.test_client(); work = tempfile.mkdtemp(prefix='toast-forms-')
results = []
def check(name, ok, info=''):
    results.append(ok); print(('  PASS ' if ok else '  FAIL ') + name + (f'  ({info})' if info else ''))

def widgets(path, name):
    d = pymupdf.open(path)
    out = [(pno, w.xref, w.field_value, pymupdf.Rect(w.rect)) for pno, p in enumerate(d) for w in p.widgets() if w.field_name == name]
    d.close(); return out

def fill(path, name, value):
    pno, xref = widgets(path, name)[0][:2]
    r = c.post('/fill_form', json={'filepath': path, 'values': [{'page': pno, 'xref': xref, 'value': value}]}).get_json()
    out = os.path.join(work, f'step{len(os.listdir(work))}.pdf'); shutil.copyfile(r['new_file'], out); return out

def ink_box(img, rect, scale):
    """Where the dark pixels are inside a field's box (in the box's own coordinates), or None if blank."""
    x0, y0, x1, y1 = [int(v * scale) for v in (rect.x0 + 2, rect.y0 + 2, rect.x1 - 2, rect.y1 - 2)]
    crop = img.crop((x0, y0, x1, y1)).convert('L').point(lambda v: 255 if v < 110 else 0)
    return crop.getbbox(), crop.size

def centered(path, name, engine):
    """How far the text's middle sits from the box's middle, as a share of the box height, in every copy."""
    from PIL import Image
    offs = []
    for pno, xref, val, rect in widgets(path, name):
        if engine == 'pdfium':
            import pypdfium2 as pdfium
            pdf = pdfium.PdfDocument(path); pdf.init_forms()
            img = pdf[pno].render(scale=4, may_draw_forms=True).to_pil(); pdf.close()
        else:
            d = pymupdf.open(path); px = d[pno].get_pixmap(dpi=288); img = Image.frombytes('RGB', (px.width, px.height), px.samples); d.close()
        bb, size = ink_box(img, rect, 4)
        if not bb: offs.append(None); continue
        offs.append(round(abs((bb[1] + bb[3]) / 2 - size[1] / 2) / size[1], 3))
    return offs

print('Emptying a field (Date picked up)')
f = fill(SRC, 'DATE PICKED UPRow1', '11/6/26')
check('a date goes into all 3 copies', [w[2] for w in widgets(f, 'DATE PICKED UPRow1')] == ['11/6/26'] * 3)
f = fill(f, 'DATE PICKED UPRow1', '')
vals = [w[2] for w in widgets(f, 'DATE PICKED UPRow1')]
check('emptied: every copy is empty', vals == [''] * 3, str(vals))
check('emptied: nothing shows in any copy (our render)', centered(f, 'DATE PICKED UPRow1', 'mupdf') == [None] * 3)
check("emptied: nothing shows in any copy (Chrome's engine)", centered(f, 'DATE PICKED UPRow1', 'pdfium') == [None] * 3)
d = pymupdf.open(f); top = [d.xref_get_key(w[1], 'V') for w in widgets(f, 'DATE PICKED UPRow1')]; d.close()
check('emptied: the saved value (V) is empty, not the old date', all(t[1] in ('()', 'null', '<>') or t[1] == '' for t in top), str(top))

print('\nChanging Conditions again and again')
f2 = SRC
for i, v in enumerate(['NEW', 'OLD', 'NEW', '', 'NEW', 'OLD', 'A much longer condition note here', 'NEW', 'OLD', 'NEW']):
    f2 = fill(f2, 'CONDITIONSRow1', v)
    vals = [w[2] for w in widgets(f2, 'CONDITIONSRow1')]
    ok_vals = vals == [v] * 3
    if not v:
        check(f'change {i + 1}: emptied in all copies', ok_vals and centered(f2, 'CONDITIONSRow1', 'mupdf') == [None] * 3, str(vals))
        continue
    ours, chrome = centered(f2, 'CONDITIONSRow1', 'mupdf'), centered(f2, 'CONDITIONSRow1', 'pdfium')
    good = lambda offs: all(o is not None and o < 0.12 for o in offs)
    check(f'change {i + 1} ({v[:12]}): all 3 copies say it and are centered, here and in Chrome\'s engine', ok_vals and good(ours) and good(chrome), f'ours {ours} chrome {chrome}')
# saved and opened again
re_path = os.path.join(work, 'reopened.pdf'); pymupdf.open(f2).save(re_path)
check('after saving and opening again: still centered in all copies', all(o is not None and o < 0.12 for o in centered(re_path, 'CONDITIONSRow1', 'pdfium')))
d = pymupdf.open(f2); da = [d.xref_get_key(w[1], 'DA') for w in widgets(f2, 'CONDITIONSRow1')]; d.close()
check('the auto font size is never written into the copies (no fixed size in DA)', all(t[0] == 'null' or ' 0 Tf' in t[1] for t in da), str(da))

print('\nA signature on every copy')
sig = [60, 324, 160, 348]   # inside the SHIPPER'S SIGNATURE box on page 1
r = c.post('/matching_spots', json={'filepath': SRC, 'page': 0, 'rect': sig}).get_json()
spots = r.get('spots', [])
check('in the shipper box: the other 2 copies are found (page 2, top and bottom)', len(spots) == 2 and all(sp['page'] == 1 for sp in spots), f"{[(sp['page'], [round(v) for v in sp['rect']]) for sp in spots]}")
check('each copy sits in the same place inside its box', all(abs(sp['rect'][0] - 60) < 2 and abs((sp['rect'][2] - sp['rect'][0]) - 100) < 3 for sp in spots))
import base64, io
from PIL import Image as _I
buf = io.BytesIO(); im = _I.new('RGBA', (300, 72), (0, 0, 0, 0)); im.paste((10, 10, 120, 255), (10, 30, 290, 40)); im.save(buf, 'PNG')
b64 = 'data:image/png;base64,' + base64.b64encode(buf.getvalue()).decode()
r2 = c.post('/place_images', json={'filepath': SRC, 'b64': b64, 'placements': spots}).get_json()
dd = pymupdf.open(r2['new_file']); imgs = [len(pg.get_images()) for pg in dd]; before = [len(pg.get_images()) for pg in pymupdf.open(SRC)]
check('placed: page 2 gets the 2 signatures in one go', imgs[1] - before[1] >= 1 and len(dd[1].get_image_info()) - len(pymupdf.open(SRC)[1].get_image_info()) == 2, f'{before} -> {imgs}')
dd.close()
r3 = c.post('/matching_spots', json={'filepath': SRC, 'page': 0, 'rect': [318, 104, 398, 118]}).get_json()
check('under CONSIGNEE: (its own form box): the 2 copies on page 2 are found', len(r3.get('spots', [])) == 2 and all(sp['page'] == 1 for sp in r3['spots']), f"{r3.get('label')} {[(sp['page'], [round(v) for v in sp['rect']]) for sp in r3.get('spots', [])]}")
r4 = c.post('/matching_spots', json={'filepath': SRC, 'page': 0, 'rect': [20, 740, 90, 760]}).get_json()
check('nowhere special: nothing is offered', r4.get('spots') == [], str(r4))

# a form without boxes: the label next to the signature is used to find the other copies
lab = os.path.join(work, 'labels.pdf'); dl = pymupdf.open()
for i in range(2):
    pg = dl.new_page(); pg.insert_text((72, 300), "Receiver's signature", fontsize=10); pg.insert_text((72, 640), "Receiver's signature", fontsize=10)
dl.save(lab)
r5 = c.post('/matching_spots', json={'filepath': lab, 'page': 0, 'rect': [72, 306, 172, 330]}).get_json()
check("no boxes: the same label (Receiver's signature) finds the other 3 spots", len(r5.get('spots', [])) == 3 and r5.get('label') == "Receiver's signature", f"{r5.get('label')} {len(r5.get('spots', []))}")

shutil.rmtree(work, ignore_errors=True)
print(f'\n{sum(results)} of {len(results)} checks passed')
sys.exit(0 if all(results) else 1)
