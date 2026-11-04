"""The PDF test library: opens, shows, fills, crops and dry-run prints every test file, and checks each result
against PDFium (Chrome's PDF engine).  python tests/print_library.py
Files: tests/pdfs (real work files, kept off GitHub), tests/samples/pdf, plus ones made here on purpose:
a scanned page, rotated and landscape pages, mixed sizes, fonts that aren't embedded, a damaged file, a
password-protected file and a 220-page file."""
import os, sys, io, re, time, shutil, tempfile
T = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(T)
sys.path.insert(0, ROOT)
import server, pymupdf, pypdfium2 as pdfium
from PIL import Image, ImageChops, ImageStat, ImageFilter

c = server.app.test_client()
work = tempfile.mkdtemp(prefix='mypdfkit-lib-')
results = []
def check(name, ok, info=''):
    results.append(ok); print(('  PASS ' if ok else '  FAIL ') + name + (f'  ({info})' if info else ''))

# ----- the made-on-purpose files -----
def make_files():
    out = {}
    def text_page(doc, w, h, label, rotate=0):
        p = doc.new_page(width=w, height=h)
        p.insert_text((36, 60), label, fontname='helv', fontsize=22)
        p.insert_text((36, 100), 'Times text, not embedded', fontname='tiro', fontsize=14)
        p.insert_text((36, 130), 'Courier text, not embedded', fontname='cour', fontsize=14)
        p.draw_rect(pymupdf.Rect(36, 150, w - 36, h - 36), color=(0, 0, 0), width=1.5)
        p.draw_line((36, 150), (w - 36, h - 36), color=(0.8, 0, 0), width=2)
        if rotate: p.set_rotation(rotate)
        return p
    d = pymupdf.open(); text_page(d, 612, 792, 'Rotated 90'); d[0].set_rotation(90); text_page(d, 612, 792, 'Rotated 270', 270)
    out['rotated.pdf'] = d
    d = pymupdf.open(); text_page(d, 792, 612, 'Landscape page'); out['landscape.pdf'] = d
    d = pymupdf.open()
    for w, h, n in ((612, 792, 'Letter'), (595.28, 841.89, 'A4'), (288, 432, '4x6 label'), (612, 1008, 'Legal'), (432, 288, '6x4 label')): text_page(d, w, h, n)
    out['mixed_sizes.pdf'] = d
    d = pymupdf.open(); text_page(d, 612, 792, 'Fonts not embedded: Helvetica, Times, Courier'); out['no_embedded_fonts.pdf'] = d
    d = pymupdf.open(); p = text_page(d, 612, 792, 'CropBox smaller than MediaBox'); p.set_cropbox(pymupdf.Rect(36, 36, 576, 500)); out['cropbox.pdf'] = d
    # a scanned page: just a picture of a page
    src = pymupdf.open(); text_page(src, 612, 792, 'Scanned page')
    pix = src[0].get_pixmap(dpi=150, colorspace=pymupdf.csGRAY)
    d = pymupdf.open(); p = d.new_page(width=612, height=792); p.insert_image(p.rect, pixmap=pix); out['scanned.pdf'] = d
    d = pymupdf.open()
    for i in range(220): d.new_page(width=612, height=792).insert_text((72, 72), f'Page {i + 1} of 220', fontname='helv', fontsize=18)
    out['long_220_pages.pdf'] = d
    paths = {}
    for name, doc in out.items():
        paths[name] = os.path.join(work, name); doc.save(paths[name])
    # password-protected (password: test)
    d = pymupdf.open(); text_page(d, 612, 792, 'Password protected')
    paths['password.pdf'] = os.path.join(work, 'password.pdf')
    d.save(paths['password.pdf'], encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw='test', owner_pw='owner')
    # damaged: the table of contents at the end cut off
    raw = open(paths['landscape.pdf'], 'rb').read()
    paths['damaged.pdf'] = os.path.join(work, 'damaged.pdf')
    open(paths['damaged.pdf'], 'wb').write(raw[:raw.rfind(b'xref')] + b'\n%%EOF\n')
    return paths

def gray(img, size=None):
    img = img.convert('L')
    return img.resize(size) if size else img

def diff(a, b):
    """Average difference of two pictures, 0 (same) to 255, after a slight blur: the two engines shade the
    edges of tiny letters a little differently, which isn't a difference anyone sees."""
    b = b.resize(a.size)
    blur = ImageFilter.GaussianBlur(1.2)
    return ImageStat.Stat(ImageChops.difference(a.filter(blur), b.filter(blur))).mean[0]

def pdfium_page(path, i, scale, password=None):
    pdf = pdfium.PdfDocument(path, password=password)
    try: pdf.init_forms()
    except Exception: pass
    return gray(pdf[i].render(scale=scale, may_draw_forms=True).to_pil())

def mupdf_page(path, i, dpi=50):
    p = pymupdf.open(path)[i].get_pixmap(dpi=dpi, annots=True)
    return gray(Image.frombytes('RGB', (p.width, p.height), p.samples))

files = make_files()
for folder in (os.path.join(T, 'pdfs'), os.path.join(T, 'samples', 'pdf')):
    if os.path.isdir(folder):
        for n in sorted(os.listdir(folder)):
            if n.lower().endswith('.pdf'): files[n] = os.path.join(folder, n)

LIMIT = 12   # average difference (out of 255) still counted as looking the same
for name, path in files.items():
    print(f'\n{name}')
    t0 = time.time()
    r = c.post('/process_file', json={'filepath': path}).get_json()
    pw = None
    if r.get('needs_password'):
        r = c.post('/unlock', json={'filepath': path, 'password': 'test'}).get_json(); pw = 'test'
        check('opens with its password', bool(r.get('pdf_path')))
    if not r.get('pdf_path'): check('opens', False, str(r)[:120]); continue
    pdf = r['pdf_path']
    if r.get('message'): print('   note:', r['message'])
    doc = pymupdf.open(pdf); n = doc.page_count
    check('opens', n > 0, f'{n} pages, {time.time() - t0:.1f}s')
    # shows like Chrome (first, middle, last page)
    worst = 0
    for i in sorted({0, n // 2, n - 1}):
        a = mupdf_page(pdf, i); b = pdfium_page(path if pw is None else pdf, i, 50 / 72)
        worst = max(worst, diff(a, b))
    check('looks the same as in Chrome', worst < LIMIT, f'worst page differs by {worst:.1f}')
    # dry-run print, whole pages on Letter: every sheet the paper's size, and matching Chrome's page (turned as needed)
    pages = sorted({1, (n + 1) // 2, n}) if n > 3 else list(range(1, n + 1))
    pr = c.post('/print_prepare', json={'filepath': pdf, 'pages': pages, 'paper': [612, 792], 'printable': [0, 0, 612, 792], 'mode': 'page'}).get_json()
    if 'error' in pr: check('prints whole pages on Letter', False, pr['error']); continue
    sheets = pymupdf.open(pr['path'])
    sizes_ok = all(abs(p.rect.width - 612) < 1 and abs(p.rect.height - 792) < 1 for p in sheets)
    worst = 0
    for k, pg in enumerate(pages):
        sp = sheets[k].get_pixmap(dpi=40); s_img = gray(Image.frombytes('RGB', (sp.width, sp.height), sp.samples))
        ref = pdfium_page(path if pw is None else pdf, pg - 1, 40 / 72)
        # the page as Chrome shows it, fitted onto the sheet the same way (turned if the sheet holds it turned)
        best = 255
        for rot in (0, 90, 270):
            r_img = ref.rotate(rot, expand=True)
            f = min(s_img.width / r_img.width, s_img.height / r_img.height)
            fitted = r_img.resize((max(1, int(r_img.width * f)), max(1, int(r_img.height * f))))
            canvas = Image.new('L', s_img.size, 255); canvas.paste(fitted, ((s_img.width - fitted.width) // 2, (s_img.height - fitted.height) // 2))
            best = min(best, diff(s_img, canvas))
        worst = max(worst, best)
    check('prints whole pages on Letter like Chrome', sizes_ok and worst < LIMIT, f'{sheets.page_count} sheets, worst differs by {worst:.1f}')
    # dry-run print on a 4x6 label printer: labels found, blank pages left out, every sheet exactly 4x6
    pr = c.post('/print_prepare', json={'filepath': pdf, 'pages': pages, 'paper': [288, 432], 'printable': [0, 0, 288, 432], 'mode': 'label', 'skip_blank': True}).get_json()
    if 'error' in pr: check('prints labels on 4x6', 'Nothing to print' in pr['error'], pr['error'])
    else:
        sh = pymupdf.open(pr['path'])
        check('prints labels on 4x6', all(abs(p.rect.width - 288) < 1 and abs(p.rect.height - 432) < 1 for p in sh), f'{sh.page_count} sheets, skipped {pr["skipped"]}')
        rr = c.post('/print_render', json={'path': pr['path'], 'dpi': 203, 'first': 0, 'count': 1}).get_json()
        im = Image.open(rr['pages'][0]['path'])
        check('label sheet pictures at the printer resolution', abs(im.width - 812) <= 2 and abs(im.height - 1218) <= 2, f'{im.width}x{im.height} px at 203 dpi')
    # crop: save the labels as their own 4x6 PDF (vector, nothing re-drawn as a picture)
    sl = c.post('/save_labels', json={'filepath': pdf, 'pages': pages[:2], 'original_path': os.path.join(work, 'out_' + name), 'size': [288, 432]}).get_json()
    if sl.get('success'):
        o = pymupdf.open(sl['path'])
        src_words = sum(len(doc[p - 1].get_text('words')) for p in pages[:2])
        out_words = sum(len(p.get_text('words')) for p in o)
        check('crops to 4x6 and keeps text sharp (vector)', out_words >= min(src_words, 1) or src_words == 0, f'{o.page_count} pages, {out_words} words kept')
    else: check('crops to 4x6', 'No labels' in (sl.get('error') or ''), sl.get('error'))
    # fill: the first text field, on every copy, shown the same by Chrome
    fields = [(p.number, w.xref, w.field_name) for p in doc for w in (p.widgets() or []) if w.field_type == pymupdf.PDF_WIDGET_TYPE_TEXT]
    if fields:
        pno, xref, fname = fields[0]
        ff = c.post('/fill_form', json={'filepath': pdf, 'values': [{'page': pno, 'xref': xref, 'value': 'TEST 123'}]}).get_json()
        filled = ff['new_file']
        copies = sum(1 for p in pymupdf.open(filled) for x in (p.widgets() or []) if x.field_name == fname and x.field_value == 'TEST 123')
        want = sum(1 for f in fields if f[2] == fname)
        a = mupdf_page(filled, pno); b = pdfium_page(filled, pno, 50 / 72)
        check('fills every copy of a field, shown like Chrome', copies == want and diff(a, b) < LIMIT, f'{copies} of {want} copies, differs by {diff(a, b):.1f}')

# ----- layouts: several per sheet, booklet, poster, sticker sheets, receipts, exact size, a whole batch -----
print('\nlayouts')
mixed = files['mixed_sizes.pdf']; L = {'paper': [612, 792], 'printable': [18, 18, 594, 774], 'mode': 'page'}
def sheets_of(d):
    r = c.post('/print_prepare', json={**L, **d}).get_json()
    return (pymupdf.open(r['path']), r) if 'path' in r else (None, r)
doc, r = sheets_of({'filepath': files['long_220_pages.pdf'], 'pages': list(range(1, 9)), 'layout': 'nup', 'nup': 4})
check('4 per sheet: 8 pages on 2 sheets', doc and doc.page_count == 2 and len(doc[0].get_text('words')) >= 16, f'{doc.page_count if doc else r}')
doc, r = sheets_of({'filepath': files['long_220_pages.pdf'], 'pages': list(range(1, 7)), 'layout': 'booklet'})
order = [[int(x) for x in re.findall(r'Page (\d+) of', p.get_text())] for p in doc] if doc else []
check('booklet: fold order (8|1, 2|7, 6|3, 4|5; 7 and 8 are empty)', order == [[1], [2], [6, 3], [4, 5]], f'{order}')
check('booklet: sheets lie sideways', doc and all(p.rect.width > p.rect.height for p in doc))
doc, r = sheets_of({'filepath': files['landscape.pdf'], 'layout': 'poster', 'cols': 3, 'rows': 3})
check('poster 3 x 3: nine sheets with cut lines', doc and doc.page_count == 9 and any(dr.get('dashes') for dr in doc[4].get_drawings()), f'{doc.page_count if doc else r}')
doc, r = sheets_of({'filepath': files['long_220_pages.pdf'], 'pages': list(range(1, 12)), 'layout': 'sheet', 'sheet': '5163', 'start': 3})
check('sticker sheet (10 per sheet) from sticker 3: 11 items on 2 sheets', doc and doc.page_count == 2 and r['count'] == 11, f'{doc.page_count if doc else r}')
first = sorted(fitz_rect.y0 for fitz_rect in [pymupdf.Rect(b[:4]) for b in doc[0].get_text('blocks')])[0] if doc else 0
check('sticker sheet: the first two stickers are left empty', first > 2 * 72, f'first item starts {first / 72:.2f} in down')
doc, r = sheets_of({'filepath': mixed, 'pages': [3], 'layout': 'receipt', 'paper': [226.77, 1000], 'printable': None})
check('80 mm receipt: full width, only as long as needed', doc and abs(doc[0].rect.width - 226.77) < 1 and doc[0].rect.height < 1000, f'{doc[0].rect if doc else r}')
doc, r = sheets_of({'filepath': files['landscape.pdf'], 'scale': 'pct', 'percent': 50, 'orient': 'landscape'})
blk = pymupdf.Rect(); [blk.include_rect(dr['rect']) for dr in doc[0].get_drawings()]
check('exact 50%: the page is drawn at half size', doc and abs(max(blk.width, blk.height) - 0.5 * (792 - 72)) < 6, f'{blk.width:.0f} x {blk.height:.0f} pt')
doc, r = sheets_of({'files': [files['landscape.pdf'], files['rotated.pdf'], files['mixed_sizes.pdf']]})
check('whole tab: every page of every file', doc and doc.page_count == 1 + 2 + 5, f'{doc.page_count if doc else r}')
rr = c.post('/print_render', json={'path': r['path'], 'dpi': 100, 'toner': True, 'indices': [2, 0]}).get_json()
im = Image.open(rr['pages'][0]['path']).convert('L')
check('toner saver prints black as gray; sheets in the order asked', min(im.getdata()) > 80 and rr['pages'][0]['path'].endswith('3.png'), f'darkest {min(im.getdata())}')

shutil.rmtree(work, ignore_errors=True)
print(f'\n{sum(results)} of {len(results)} checks passed')
sys.exit(0 if all(results) else 1)
