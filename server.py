import os
import sys
import subprocess
import hashlib
import tempfile
import pythoncom
import shutil
import re
import base64
import json
from io import BytesIO
from flask import Flask, request, jsonify
from flask_cors import CORS
import fitz  # PyMuPDF
import win32com.client as win32
from PIL import Image

# PyInstaller --noconsole leaves stdout/stderr as None, which crashes Flask's logging
if sys.stdout is None: sys.stdout = open(os.devnull, 'w')
if sys.stderr is None: sys.stderr = open(os.devnull, 'w')

app = Flask(__name__)
CORS(app)

TEMP_DIR = os.path.join(tempfile.gettempdir(), "UniversalEditor_Backend")
os.makedirs(TEMP_DIR, exist_ok=True)

def get_unique_filename(path):
    if not os.path.exists(path): return path
    d, f = os.path.split(path)
    b, e = os.path.splitext(f)
    match = re.search(r' \((\d+)\)$', b)
    c = int(match.group(1)) if match else 1
    b = b[:match.start()] if match else b
    while True:
        new_path = os.path.join(d, f"{b} ({c}){e}")
        if not os.path.exists(new_path): return new_path
        c += 1

def parse_color(c_str):
    try:
        c_str = str(c_str).strip()
        if c_str.startswith('rgb'):
            nums = re.findall(r'\d+', c_str)
            return tuple(int(n)/255.0 for n in nums[:3])
        else:
            h = c_str.lstrip('#')
            if len(h) < 6: h = "000000"
            return tuple(int(h[i:i+2], 16)/255.0 for i in (0, 2, 4))
    except:
        return (0.0, 0.0, 0.0)

# What can be opened, and what turns it into a PDF
WORD_EXT = ('.doc', '.docx', '.docm', '.dot', '.dotx', '.dotm', '.rtf', '.odt', '.wps', '.wpd')
EXCEL_EXT = ('.xls', '.xlsx', '.xlsm', '.xlsb', '.xlt', '.xltx', '.csv', '.tsv', '.ods')
POWERPOINT_EXT = ('.ppt', '.pptx', '.pptm', '.pps', '.ppsx', '.pot', '.potx', '.odp')
IMAGE_EXT = ('.jpg', '.jpeg', '.jfif', '.png', '.bmp', '.gif', '.tif', '.tiff', '.webp', '.ico', '.pnm', '.jxr')
BOOK_EXT = ('.xps', '.oxps', '.epub', '.fb2', '.cbz', '.mobi', '.svg')   # MuPDF reads these itself
WEB_EXT = ('.html', '.htm', '.mhtml', '.mht', '.xhtml')
TEXT_EXT = ('.txt', '.md', '.markdown', '.log', '.json', '.xml', '.yaml', '.yml', '.ini', '.cfg', '.conf', '.py', '.js',
            '.ts', '.css', '.java', '.c', '.cpp', '.h', '.cs', '.go', '.rs', '.rb', '.php', '.sql', '.sh', '.bat', '.ps1')
OPENABLE_EXT = ('.pdf',) + WORD_EXT + EXCEL_EXT + POWERPOINT_EXT + IMAGE_EXT + BOOK_EXT + WEB_EXT + TEXT_EXT

def _images_to_pdf(src, dst):
    """Every picture in the file (a multi-page TIFF or animated GIF has several) becomes a page its own size."""
    out = fitz.open()
    try:
        with Image.open(src) as im: frames = getattr(im, 'n_frames', 1)
    except Exception: frames = 1
    try:
        if frames > 1 and src.lower().endswith(('.gif', '.webp')): raise ValueError('animated: take every frame')
        img = fitz.open(src)            # MuPDF reads most formats directly, keeping the image as is
        out.insert_pdf(fitz.open('pdf', img.convert_to_pdf()))
    except Exception:
        from PIL import ImageSequence
        with Image.open(src) as im:
            for frame in ImageSequence.Iterator(im):
                buf = BytesIO(); frame.convert('RGB').save(buf, 'PNG')
                dpi = (frame.info.get('dpi') or (96, 96))[0] or 96
                w, h = frame.size[0] * 72 / dpi, frame.size[1] * 72 / dpi
                pg = out.new_page(width=w, height=h)
                pg.insert_image(pg.rect, stream=buf.getvalue())
    out.save(dst, garbage=3, deflate=True)

def _text_to_pdf(src, dst):
    """A plain text or code file, typeset in a monospaced font on Letter pages."""
    raw = open(src, 'rb').read()
    for enc in ('utf-8-sig', 'utf-16', 'cp1252', 'latin-1'):
        try: text = raw.decode(enc); break
        except UnicodeDecodeError: continue
    import html as _html
    story = fitz.Story(html=f'<pre style="font-family:monospace;font-size:9pt;white-space:pre-wrap">{_html.escape(text)}</pre>')
    writer = fitz.DocumentWriter(dst)
    page_rect, where = fitz.paper_rect('letter'), fitz.paper_rect('letter') + (54, 54, -54, -54)
    more = True
    while more:
        dev = writer.begin_page(page_rect)
        more, _ = story.place(where)
        story.draw(dev)
        writer.end_page()
    writer.close()

def _find_browser():
    for p in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
              r"C:\Program Files\Google\Chrome\Application\chrome.exe", r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"):
        if os.path.exists(p): return p
    return None

def _web_to_pdf(src, dst):
    """A web page, printed the way the browser shows it (Edge or Chrome, run hidden)."""
    exe = _find_browser()
    if not exe: raise RuntimeError("Opening web pages needs Microsoft Edge or Chrome")
    from pathlib import Path
    prof = tempfile.mkdtemp(dir=TEMP_DIR, prefix='web_')
    try:
        subprocess.run([exe, '--headless', '--disable-gpu', '--no-pdf-header-footer', f'--user-data-dir={prof}',
                        f'--print-to-pdf={dst}', Path(src).as_uri()], timeout=120, capture_output=True, creationflags=0x08000000)
    finally: shutil.rmtree(prof, ignore_errors=True)
    if not os.path.exists(dst): raise RuntimeError("The browser couldn't print this page")

def _com_to_pdf(src, dst):
    """Convert a Word/Excel/PowerPoint file to PDF using a private Office instance
    (DispatchEx) so we never touch documents the user already has open, and always Quit().
    Pictures, e-books, web pages and text files have their own converters."""
    low = src.lower()
    if low.endswith(IMAGE_EXT): return _images_to_pdf(src, dst)
    if low.endswith(BOOK_EXT):
        d = fitz.open(src); fitz.open('pdf', d.convert_to_pdf()).save(dst, garbage=3, deflate=True); return
    if low.endswith(WEB_EXT): return _web_to_pdf(src, dst)
    if low.endswith(TEXT_EXT): return _text_to_pdf(src, dst)
    if low.endswith(WORD_EXT):
        app_ = win32.DispatchEx('Word.Application')
        try:
            app_.DisplayAlerts = 0
            doc = app_.Documents.Open(src, ReadOnly=True)
            try: doc.ExportAsFixedFormat(dst, 17)
            finally: doc.Close(False)
        finally: app_.Quit()
    elif low.endswith(EXCEL_EXT):
        app_ = win32.DispatchEx('Excel.Application')
        try:
            app_.DisplayAlerts = False
            wb = app_.Workbooks.Open(src, ReadOnly=True)
            try:
                try: _widen_cut_off_columns(wb, src)
                except Exception as e: print('column fit skipped:', e, file=sys.stderr)
                wb.ExportAsFixedFormat(0, dst)
            finally: wb.Close(False)   # never saved: the user's file stays as it was
        finally: app_.Quit()
    elif low.endswith(POWERPOINT_EXT):
        app_ = win32.DispatchEx('PowerPoint.Application')
        try:
            # PowerPoint takes Office's -1 / 0 for yes / no (a Python True can't be passed to it)
            pres = app_.Presentations.Open(src, -1, 0, 0)   # read-only, keep its name, no window
            try: pres.SaveAs(dst, 32)                       # 32 = PDF
            finally: pres.Close()
        finally: app_.Quit()
    else:
        raise ValueError("Unsupported format")

def _widen_cut_off_columns(wb, src):
    """What double-clicking a column edge does, but only where it's needed: text that would be cut off
    by the cell next to it (and numbers, which turn into ####) gets its column widened to fit. Text that
    just spills into empty space, like a label or a note under a table, is left alone, and columns never
    get narrower. Works on the open copy in Excel, which is never saved."""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    if not src.lower().endswith(('.xlsx', '.xlsm')): return
    xwb = load_workbook(src)
    for ws in xwb.worksheets:
        merged = {c for r in ws.merged_cells.ranges for row in ws.iter_rows(min_row=r.min_row, max_row=r.max_row, min_col=r.min_col, max_col=r.max_col) for c in row}
        cells = {}   # column letter -> addresses of cells that must fit
        for row in ws.iter_rows():
            for c in row:
                if c.value is None or str(c.value).strip() == '' or c in merged: continue
                a = c.alignment
                if a.wrap_text or a.shrink_to_fit or ws.column_dimensions[c.column_letter].hidden: continue
                # text spills toward the side it reads to (Excel hides thin lines under it, so a label
                # spilling into an empty cell looks fine); it's cut off when that side's cell isn't empty
                num = isinstance(c.value, (int, float))
                h = a.horizontal or ('right' if num else 'left')
                sides = {'left': ['right'], 'right': ['left']}.get(h, ['left', 'right'])
                def blocked(side):
                    k = c.column + (1 if side == 'right' else -1)
                    if k < 1: return True
                    nb = ws.cell(row=c.row, column=k)
                    return nb.value not in (None, '') or nb in merged
                if num or any(blocked(sd) for sd in sides):
                    cells.setdefault(c.column, []).append(c.row)
        if not cells: continue
        xs = wb.Worksheets(ws.title)
        for k, rows in cells.items():
            col = xs.Columns(k)
            orig = best = col.ColumnWidth
            # AutoFit on part of a column fits it to just those cells; do each unbroken run of rows
            runs, start = [], rows[0]
            for a, b in zip(rows, rows[1:] + [None]):
                if b != a + 1: runs.append((start, a)); start = b
            for r0, r1 in runs:
                xs.Range(xs.Cells(r0, k), xs.Cells(r1, k)).Columns.AutoFit()
                best = max(best, col.ColumnWidth)
            col.ColumnWidth = best + 1 if best > orig else orig   # a little breathing room, like the original spacing

def find_soffice():
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe", r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(p): return p
    return shutil.which('soffice')

def libreoffice_to_pdf(src, dst):
    exe = find_soffice()
    outdir = tempfile.mkdtemp(dir=TEMP_DIR)
    try:
        subprocess.run([exe, '--headless', '--convert-to', 'pdf', '--outdir', outdir, src],
                       check=True, timeout=180, capture_output=True, creationflags=0x08000000)
        produced = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + '.pdf')
        shutil.move(produced, dst)
    finally:
        shutil.rmtree(outdir, ignore_errors=True)

def office_to_pdf(src, dst):
    """Use Microsoft Office when it is installed, otherwise LibreOffice."""
    if src.lower().endswith(IMAGE_EXT + BOOK_EXT + WEB_EXT + TEXT_EXT):
        return _com_to_pdf(src, dst)
    try:
        _com_to_pdf(src, dst)
    except Exception as com_err:
        if not find_soffice():
            raise RuntimeError("No Office converter found. Install Microsoft Office or the free LibreOffice, then try again.")
        libreoffice_to_pdf(src, dst)

CONVERTER_VERSION = '3'   # bump when office_to_pdf output changes (2: cut-off Excel columns widened, 3: many more formats)

# ---------- forms ----------
FIELD_KINDS = {fitz.PDF_WIDGET_TYPE_TEXT: 'text', fitz.PDF_WIDGET_TYPE_CHECKBOX: 'checkbox', fitz.PDF_WIDGET_TYPE_RADIOBUTTON: 'radio',
               fitz.PDF_WIDGET_TYPE_COMBOBOX: 'combo', fitz.PDF_WIDGET_TYPE_LISTBOX: 'list', fitz.PDF_WIDGET_TYPE_SIGNATURE: 'signature',
               fitz.PDF_WIDGET_TYPE_BUTTON: 'button'}

def _page_fields(page):
    """The fill-in fields on a page, where they are (as shown, after the page's rotation) and what they hold."""
    out = []
    for w in page.widgets() or []:
        kind = FIELD_KINDS.get(w.field_type, 'other')
        if kind in ('button', 'other'): continue
        r = w.rect * page.rotation_matrix
        out.append({"xref": w.xref, "name": w.field_name, "kind": kind, "rect": [r.x0, r.y0, r.x1, r.y1],
                    "value": w.field_value, "options": w.choice_values or [],
                    "on": w.on_state() if kind in ('checkbox', 'radio') else None,
                    "readonly": bool(w.field_flags & 1), "multiline": bool(w.field_flags & 4096) if kind == 'text' else False,
                    "size": w.text_fontsize or 0, "maxlen": w.text_maxlen or 0, "label": w.field_label or ''})
    return out

@app.route('/page_info', methods=['POST'])
def page_info():
    """What the page has beyond its drawing: fill-in fields, and its paragraphs (for editing them whole)."""
    d = request.json
    try:
        doc = fitz.open(d['filepath']); page = doc[int(d.get('page', 0))]
        notes = []
        for a in page.annots() or []:   # comments and markup, with what they say (shown on hover)
            text = (a.info.get('content') or '').strip()
            r = a.rect * page.rotation_matrix
            notes.append({"xref": a.xref, "kind": a.type[1], "rect": [r.x0, r.y0, r.x1, r.y1], "text": text,
                          "author": a.info.get('title') or ''})
        info = {"fields": _page_fields(page), "paragraphs": _paragraphs(page), "notes": notes}
        doc.close()
        return jsonify(info)
    except Exception as e: return jsonify({"error": str(e), "fields": [], "paragraphs": []})

@app.route('/fill_form', methods=['POST'])
def fill_form():
    """Write typed values into the PDF's own fields: [{page, xref, value}], value true/false for boxes."""
    d = request.json; fp = d['filepath']
    try:
        doc = fitz.open(fp)
        by_page = {}
        for v in d.get('values', []): by_page.setdefault(int(v['page']), []).append(v)
        for pno, vals in by_page.items():
            page = doc[pno]
            widgets = {w.xref: w for w in page.widgets() or []}
            for v in vals:
                w = widgets.get(int(v['xref']))
                if not w: continue
                kind = FIELD_KINDS.get(w.field_type)
                if kind in ('checkbox', 'radio'):
                    w.field_value = w.on_state() if v['value'] else 'Off'
                else:
                    w.field_value = str(v['value'])
                w.update()
                widgets = {x.xref: x for x in page.widgets() or []}   # updating can renew the list
        if d.get('flatten'):   # turn the fields into plain page content
            try: doc.bake(annots=False, widgets=True)
            except Exception as e: print('flatten skipped:', e, file=sys.stderr)
        new_path = os.path.join(TEMP_DIR, f"mod_{hashlib.md5((fp + str(os.path.getmtime(fp)) + json.dumps(d.get('values'))).encode()).hexdigest()}.pdf")
        doc.save(new_path, garbage=3, deflate=True); doc.close()
        return jsonify({"success": True, "new_file": new_path})
    except Exception as e: return jsonify({"error": str(e)}), 500

def _paragraphs(page):
    """Blocks of two or more lines of running text (for editing a paragraph as a whole): the block's box,
    each line's box, and the text."""
    out = []
    for b in page.get_text('dict')['blocks']:
        if b['type'] != 0: continue
        lines = [l for l in b['lines'] if ''.join(s['text'] for s in l['spans']).strip() and l['dir'] == (1, 0)]
        if len(lines) < 2: continue
        out.append({"rect": list(b['bbox']), "lines": [list(l['bbox']) for l in lines],
                    "text": ' '.join(''.join(s['text'] for s in l['spans']).strip() for l in lines)})
    return out

# ---------- compare ----------
def _doc_words(doc):
    """Every word in reading order: (page, Rect, text)."""
    out = []
    for pno, page in enumerate(doc):
        # by where they sit (line by line, left to right), not the order they're stored in, which edits change
        words = sorted(page.get_text('words'), key=lambda w: ((w[1] + w[3]) / 2, w[0]))
        lines, cur = [], []
        for w in words:
            if cur and abs((w[1] + w[3]) / 2 - (cur[-1][1] + cur[-1][3]) / 2) > max(2.5, (w[3] - w[1]) * 0.4):
                lines.append(cur); cur = []
            cur.append(w)
        if cur: lines.append(cur)
        for line in lines:
            for w in sorted(line, key=lambda w: w[0]): out.append((pno, fitz.Rect(w[:4]), w[4]))
    return out

def _changed_areas(pa, pb, skip):
    """Boxes (in page B's points) where the two pages look different, apart from the `skip` boxes:
    pictures, lines and other drawing that a word comparison can't see."""
    if abs(pa.rect.width - pb.rect.width) > 1 or abs(pa.rect.height - pb.rect.height) > 1: return []
    dpi = 40; k = 72 / dpi
    A, B = pa.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY), pb.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
    if (A.width, A.height) != (B.width, B.height): return []
    w, h, sa, sb = A.width, A.height, A.samples, B.samples
    cell = 8
    hot = set()
    for cy in range(0, h, cell):
        for cx in range(0, w, cell):
            diff = 0
            for y in range(cy, min(cy + cell, h)):
                row = y * w
                for x in range(cx, min(cx + cell, w)):
                    if abs(sa[row + x] - sb[row + x]) > 48: diff += 1
            if diff >= 3:
                r = fitz.Rect(cx * k, cy * k, (cx + cell) * k, (cy + cell) * k)
                if not any(r.intersects(z) for z in skip): hot.add((cx // cell, cy // cell))
    # join touching cells into boxes
    boxes, seen = [], set()
    for c in hot:
        if c in seen: continue
        stack, group = [c], []
        seen.add(c)
        while stack:
            x, y = stack.pop(); group.append((x, y))
            for n in ((x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)):
                if n in hot and n not in seen: seen.add(n); stack.append(n)
        xs, ys = [g[0] for g in group], [g[1] for g in group]
        boxes.append(fitz.Rect(min(xs) * cell * k, min(ys) * cell * k, (max(xs) + 1) * cell * k, (max(ys) + 1) * cell * k))
    return boxes

@app.route('/compare', methods=['POST'])
def compare_pdfs():
    """Compare two PDFs. The report is the newer one (b) with changes marked: changed and added words
    highlighted (hover shows the old words), removed words flagged where they were, changed pictures or
    drawings boxed; a first page lists every change."""
    import difflib
    d = request.json
    try:
        A, B = fitz.open(d['a']), fitz.open(d['b'])
        name_a, name_b = d.get('a_name') or os.path.basename(d['a']), d.get('b_name') or os.path.basename(d['b'])
        wa, wb = _doc_words(A), _doc_words(B)
        norm = lambda t: t.strip().lower()
        ops = difflib.SequenceMatcher(None, [norm(w[2]) for w in wa], [norm(w[2]) for w in wb], autojunk=False).get_opcodes()
        report = fitz.open(); report.insert_pdf(B)
        rp = [report[k] for k in range(len(report))]   # one page object each (notes must stay tied to it)
        changes, marked = [], {}
        YELLOW, GREEN, RED = (1, 0.85, 0.2), (0.45, 0.85, 0.45), (0.9, 0.2, 0.2)
        for op, i1, i2, j1, j2 in ops:
            if op == 'equal': continue
            old = ' '.join(w[2] for w in wa[i1:i2]); new = ' '.join(w[2] for w in wb[j1:j2])
            if op in ('replace', 'insert'):
                for pno in sorted({w[0] for w in wb[j1:j2]}):
                    quads = [w[1] for w in wb[j1:j2] if w[0] == pno]
                    a = rp[pno].add_highlight_annot(quads)
                    a.set_colors(stroke=YELLOW if op == 'replace' else GREEN)
                    a.set_info(title='Compare', content=(f'Was: "{old}"' if op == 'replace' else 'Added'))
                    a.update()
                    marked.setdefault(pno, []).extend(quads)
                pg = wb[j1][0]
            else:   # delete: flag it where the removed words were (just before the next word that's still there)
                at = wb[j1] if j1 < len(wb) else (wb[-1] if wb else None)
                pg = at[0] if at else 0
                spot = fitz.Point(at[1].x0 - 2, at[1].y0) if at else fitz.Point(36, 36)
                a = rp[pg].add_text_annot(spot, f'Removed: "{old}"', icon='Comment')
                a.set_colors(stroke=RED); a.update()
            changes.append((pg, op, old, new))
        # pictures and drawings that changed
        visual = 0
        for pno in range(min(len(A), len(B))):
            for r in _changed_areas(A[pno], B[pno], marked.get(pno, [])):
                a = rp[pno].add_rect_annot(r + (-2, -2, 2, 2)); a.set_colors(stroke=RED); a.set_border(width=1.2)
                a.set_info(title='Compare', content='Something here looks different (a picture, line or drawing)'); a.update()
                visual += 1
        if len(A) != len(B): changes.append((min(len(A), len(B)), 'pages', f'{len(A)} pages', f'{len(B)} pages'))
        # summary page first
        lines = [f'Comparing "{name_a}" (before) with "{name_b}" (after)', '',
                 f'{sum(1 for c in changes if c[1] != "pages")} text changes, {visual} other changed areas', '']
        label = {'replace': 'Changed', 'insert': 'Added', 'delete': 'Removed', 'pages': 'Page count'}
        for pg, op, old, new in changes[:400]:
            if op == 'replace': lines.append(f'Page {pg + 1}  Changed: "{old[:80]}"  ->  "{new[:80]}"')
            elif op == 'insert': lines.append(f'Page {pg + 1}  Added: "{new[:120]}"')
            elif op == 'delete': lines.append(f'Page {pg + 1}  Removed: "{old[:120]}"')
            else: lines.append(f'Page count changed: {old} -> {new}')
        if not changes and not visual: lines.append('No differences found.')
        per = 46
        for start in range(0, len(lines), per):
            sp = report.new_page(pno=start // per, width=612, height=792)
            y = 60
            for k, line in enumerate(lines[start:start + per]):
                bold = start == 0 and k == 0
                sp.insert_text((50, y), line, fontsize=13 if bold else 9.5, fontname='hebo' if bold else 'helv',
                               color=(0, 0, 0) if bold or not line.startswith('Page') else (0.15, 0.15, 0.15))
                y += 22 if bold else 15
        out = get_unique_filename(os.path.join(os.path.dirname(d['b']) if not d.get('out_dir') else d['out_dir'],
                                               f"Compare - {os.path.splitext(name_a)[0]} vs {os.path.splitext(name_b)[0]}.pdf"))
        report.save(out, garbage=3, deflate=True)
        return jsonify({"success": True, "path": out, "changes": len(changes), "visual": visual,
                        "message": f"Found {len(changes)} text changes and {visual} other changed areas"})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/formats', methods=['GET'])
def formats():
    return jsonify({"extensions": [e[1:] for e in OPENABLE_EXT]})

@app.route('/ping', methods=['GET'])
def ping():
    return jsonify({"status": "connected"})

@app.route('/process_file', methods=['POST'])
def process_file():
    data = request.json
    path = data.get('filepath')
    if not path or not os.path.exists(path): return jsonify({"error": "File not found"}), 400

    ext = path.lower()
    if ext.endswith('.pdf'):
        # a locked PDF needs its password first; a damaged one is opened from a repaired copy
        try:
            d = fitz.open(path)
            if d.needs_pass: d.close(); return jsonify({"needs_password": True})
            if d.is_repaired:
                fixed = os.path.join(TEMP_DIR, 'fixed_' + hashlib.md5((path + str(os.path.getmtime(path))).encode()).hexdigest() + '.pdf')
                d.save(fixed, garbage=3, deflate=True); d.close()
                return jsonify({"pdf_path": fixed, "message": "This PDF was damaged; it was repaired to open it"})
            d.close()
        except Exception as e:
            return jsonify({"error": f"This file couldn't be read as a PDF ({e})"}), 400
        return jsonify({"pdf_path": path})

    # CONVERTER_VERSION is in the key so a change to how files are converted redoes old previews
    hasher = hashlib.md5((path + str(os.path.getmtime(path)) + CONVERTER_VERSION).encode('utf-8')).hexdigest()
    temp_target = os.path.join(TEMP_DIR, f"prev_{hasher}.pdf")

    if not os.path.exists(temp_target):
        pythoncom.CoInitialize()
        try:
            office_to_pdf(path, temp_target)
        except ValueError as e:
            return jsonify({"error": str(e)}), 400
        except Exception as e:
            return jsonify({"error": str(e)}), 500
        finally:
            pythoncom.CoUninitialize()

    return jsonify({"pdf_path": temp_target})

@app.route('/unlock', methods=['POST'])
def unlock():
    """Open a password-protected PDF: an unlocked working copy (the original file keeps its password)."""
    d = request.json; path, pw = d.get('filepath'), d.get('password', '')
    try:
        doc = fitz.open(path)
        if doc.needs_pass and not doc.authenticate(pw):
            doc.close(); return jsonify({"error": "That password isn't right"}), 400
        out = os.path.join(TEMP_DIR, 'open_' + hashlib.md5((path + str(os.path.getmtime(path))).encode()).hexdigest() + '.pdf')
        doc.save(out, encryption=fitz.PDF_ENCRYPT_NONE, garbage=3, deflate=True); doc.close()
        return jsonify({"pdf_path": out})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/process_all_files', methods=['POST'])
def process_all_files():
    data = request.json
    files_data = data.get('files_data', [])
    duplicate_approved = data.get('duplicate_approved', False)
    processed_count = 0

    try:
        pythoncom.CoInitialize()
        for fd in files_data:
            orig_path = fd['original']
            curr_path = fd['current']
            
            if not os.path.exists(orig_path): continue
            
            ext = orig_path.lower()
            base_pdf = os.path.splitext(orig_path)[0] + ".pdf"
            target_pdf = get_unique_filename(base_pdf) if os.path.exists(base_pdf) and duplicate_approved else base_pdf

            source_to_process = curr_path if (curr_path and os.path.exists(curr_path)) else orig_path

            if source_to_process.lower().endswith('.pdf'):
                if source_to_process != target_pdf: shutil.copy(source_to_process, target_pdf)
                processed_count += 1
            else:
                office_to_pdf(source_to_process, target_pdf)
                processed_count += 1
    except Exception as e: return jsonify({"error": str(e)}), 500
    finally: pythoncom.CoUninitialize()
    
    return jsonify({"success": True, "count": processed_count})

# ---------- rewriting edited text ----------
# The aim is an edit nobody can spot: the new words go in with the original's own font, size, color and
# baseline, only the original glyphs are removed (table lines and pictures around them stay), and
# words that didn't change are left alone.

_system_fonts = None
def _font_key(name):
    """'ABCDEF+Calibri-Bold', 'Calibri,Bold', 'Arial-BoldMT' and 'Calibri Bold' all become 'calibribold'.
    Junk some PDF makers put before the name (Firefox writes things like 'AAAAAD\u00e5\u00bc\u00abMontserrat-Regular')
    is dropped too."""
    name = (name or '').split('+')[-1]
    name = re.sub(r'^.*[^\x00-\x7f]', '', name) or name
    parts = [re.sub(r'(PSMT|PS|MT)$', '', p) for p in re.split(r'[-,]', name)]   # Adobe's 'MT'/'PS' name tails
    key = re.sub(r'[^a-z0-9]', '', ''.join(parts).lower())
    return re.sub(r'regular$', '', key) or key

def _find_system_font(name):
    """Path of the installed Windows font with this name, from the registry's font list."""
    global _system_fonts
    if _system_fonts is None:
        _system_fonts = {}
        import winreg
        for hive, folder in ((winreg.HKEY_LOCAL_MACHINE, os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts')),
                             (winreg.HKEY_CURRENT_USER, '')):
            try: k = winreg.OpenKey(hive, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts')
            except OSError: continue
            i = 0
            while True:
                try: label, file, _ = winreg.EnumValue(k, i)
                except OSError: break
                i += 1
                if not str(file).lower().endswith(('.ttf', '.otf', '.ttc')): continue
                label = re.sub(r'\s*\(.*\)$', '', label)
                parts = label.split(' & ')
                if str(file).lower().endswith('.ttc'): parts = parts[:1]   # only a collection's first font can be loaded
                for part in parts:
                    _system_fonts.setdefault(_font_key(part), os.path.join(folder, file))
    key = _font_key(name)
    if key in _system_fonts: return _system_fonts[key]
    # a name cut short ('Montserrat-Regu'): the longest installed name it starts with
    best = max((k for k in _system_fonts if len(k) >= 4 and key.startswith(k)), key=len, default=None)
    return _system_fonts.get(best) if best else None

def _font_key_raw(name):
    """A font name squashed to letters and digits, keeping 'Regular' (for comparing names cut short)."""
    name = (name or '').split('+')[-1]
    name = re.sub(r'^.*[^\x00-\x7f]', '', name) or name
    return re.sub(r'[^a-z0-9]', '', name.lower())

def _same_font(a, b):
    """Do two font names (one maybe cut short or with junk in front) name the same font?"""
    ka, kb = _font_key(a), _font_key(b)
    return bool(ka and kb) and (ka == kb or (min(len(ka), len(kb)) >= 6 and (ka.startswith(kb) or kb.startswith(ka))))

def _embedded_font(doc, page, span_font, cache):
    """The PDF's own copy of a font as a fitz.Font that can type letters: an exact name match first,
    else the font whose name the (cut short) span name starts. If the copy has no letter map (Word
    leaves some out), one is rebuilt from the glyphs the page's text already uses."""
    key = ('emb', span_font)
    if key in cache: return cache[key]
    sk = _font_key(span_font)
    fonts = [(xref, ext, base) for xref, ext, _, base, *_ in page.get_fonts(full=True) if ext not in ('n/a', '')]
    exact = [f for f in fonts if _font_key(f[2]) == sk]
    rk = lambda n: _font_key_raw(n)   # name as is, for cut-short ones
    sr = rk(span_font)
    cut = sorted((f for f in fonts if len(sr) >= 6 and rk(f[2]).startswith(sr) and f not in exact), key=lambda f: len(rk(f[2])))
    found = None
    for xref, ext, base in exact + cut:
        try:
            buf = doc.extract_font(xref)[3]
            if not buf: continue
            font = fitz.Font(fontbuffer=buf)
            used = _used_glyphs(page, base)
            if ext == 'ttf' and (_has_private_codes(buf) or (used and not all(font.has_glyph(ord(ch)) for ch in used if not ch.isspace()))):
                fixed = _with_letter_map(buf, used)
                if fixed: font = fitz.Font(fontbuffer=fixed)
            found = font; break
        except Exception as e: print('embedded font skipped:', e, file=sys.stderr)
    cache[key] = found
    return found

def _used_glyphs(page, base):
    """{character: glyph id} for the text this page draws in the font named `base`."""
    out = {}
    try:
        for sp in page.get_texttrace():
            if _same_font(sp.get('font', ''), base):
                for ch in sp['chars']:
                    if ch[0] > 0 and ch[1] >= 0: out.setdefault(chr(ch[0]), ch[1])
    except Exception: pass
    return out

def _has_private_codes(buf):
    """Does the font's letter map also point at private-use codes?"""
    try:
        from fontTools.ttLib import TTFont
        tt = TTFont(BytesIO(buf), lazy=True)
        return 'cmap' in tt and any(0xE000 <= cp <= 0xF8FF for t in tt['cmap'].tables if t.isUnicode() for cp in t.cmap)
    except Exception: return False

def _with_letter_map(buf, used):
    """The font file with a Unicode letter map added (TrueType only), so new text can be typed in it."""
    try:
        from fontTools.ttLib import TTFont, newTable
        from fontTools.ttLib.tables._c_m_a_p import cmap_format_4
        tt = TTFont(BytesIO(buf))
        order = tt.getGlyphOrder()
        pua = lambda cp: 0xE000 <= cp <= 0xF8FF
        m = {ord(ch): order[g] for ch, g in used.items() if g < len(order) and ord(ch) < 0x10000 and not pua(ord(ch))}
        if 'cmap' in tt:
            for t in tt['cmap'].tables:
                if t.isUnicode():
                    # private-use codes (some makers map every glyph to one too) would make the edited
                    # words copy and search as symbols, so they're left out
                    for cp, name in t.cmap.items():
                        if not pua(cp) and cp < 0x10000: m.setdefault(cp, name)
        sub = cmap_format_4(4); sub.platformID, sub.platEncID, sub.language, sub.cmap = 3, 1, 0, m
        cm = newTable('cmap'); cm.tableVersion = 0; cm.tables = [sub]
        tt['cmap'] = cm
        out = BytesIO(); tt.save(out); return out.getvalue()
    except Exception as e:
        print('letter map not added:', e, file=sys.stderr); return None

def _font_for(doc, page, span_font, text, cache, prefer_embedded=True):
    """A fitz.Font that draws `text` like `span_font`: the PDF's own copy when it has every letter,
    else the installed font of that name, else the closest built-in font."""
    need = [ch for ch in text if not ch.isspace()]
    if prefer_embedded:
        f = _embedded_font(doc, page, span_font, cache)
        if f is not None and all(f.has_glyph(ord(ch)) for ch in need): return f
    path = _find_system_font(span_font)
    if path and os.path.exists(path):
        try:
            f = fitz.Font(fontfile=path)
            if all(f.has_glyph(ord(ch)) for ch in need): return f
        except Exception: pass
    low = span_font.lower()
    b, it = 'bold' in low or 'black' in low or 'heavy' in low, 'italic' in low or 'oblique' in low
    fam = 'co' if 'cour' in low or 'mono' in low else 'ti' if not low or 'times' in low or 'serif' in low and 'sans' not in low else 'he'
    return fitz.Font({('co', 0, 0): 'cour', ('co', 1, 0): 'cobo', ('co', 0, 1): 'coit', ('co', 1, 1): 'cobi',
                      ('ti', 0, 0): 'tiro', ('ti', 1, 0): 'tibo', ('ti', 0, 1): 'tiit', ('ti', 1, 1): 'tibi',
                      ('he', 0, 0): 'helv', ('he', 1, 0): 'hebo', ('he', 0, 1): 'heit', ('he', 1, 1): 'hebi'}[(fam, int(b), int(it))])

def _layout(chars, text, font, size, keep_positions, hscale=1.0):
    """Where each letter of `text` goes. Letters that were already there stay exactly where they were (PDF
    makers often space letters a little differently from the font's own widths, and copying those spots
    is what keeps an edit invisible); new letters are spaced by the font, and what follows a change moves
    along by the difference."""
    import difflib
    x0 = chars[0]['origin'][0]
    sp = [c for c in chars if c['c'] == ' ']
    def adv(ch):
        if ch == ' ' and not font.has_glyph(32):   # fonts with no space: use the original's spacing
            return (sum(c['bbox'][2] - c['bbox'][0] for c in sp) / len(sp)) if sp else size * 0.27
        return font.text_length(ch, size) * hscale
    if not keep_positions:
        out, x = [], x0
        for ch in text: out.append((x, ch)); x += adv(ch)
        return out
    old = ''.join(c['c'] for c in chars)
    out, cur = [], x0
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, old, text, autojunk=False).get_opcodes():
        if op == 'delete' and i1 == 0 and not old[i1:i2].strip() and i2 < len(chars):
            cur = chars[i2]['origin'][0]   # leading spaces dropped: the words stay where they were
        elif op == 'equal':
            shift = cur - chars[i1]['origin'][0]
            for k in range(i2 - i1): out.append((chars[i1 + k]['origin'][0] + shift, text[j1 + k]))
            last = chars[i2 - 1]
            cur = (chars[i2]['origin'][0] + shift) if i2 < len(chars) else (last['origin'][0] + shift + adv(last['c']))
        elif op in ('replace', 'insert'):
            for ch in text[j1:j2]: out.append((cur, ch)); cur += adv(ch)
            if op == 'replace' and i2 < len(chars):   # old letters gone; continue from where they'd have ended
                pass
    return out

def _text_font(doc, page, span_font, text, cache):
    """A font name usable in insert_text that draws `text` like `span_font`: the copy embedded in this PDF
    when it has every letter needed (it is exactly the original), else the installed font of that name,
    else the closest built-in font."""
    if (span_font, text) in cache: return cache[(span_font, text)]
    alias = None
    for xref, ext, _, base, *_ in page.get_fonts(full=True):
        if ext in ('n/a', '') or not _same_font(base, span_font): continue
        try:
            buf = doc.extract_font(xref)[3]
            if buf and all(fitz.Font(fontbuffer=buf).has_glyph(ord(ch)) for ch in text if not ch.isspace()):
                alias = 'X%d' % xref
                page.insert_font(fontname=alias, fontbuffer=buf)
                cache['font:' + alias] = fitz.Font(fontbuffer=buf)
                break
        except Exception: pass
    if not alias:
        path = _find_system_font(span_font)
        if path and os.path.exists(path):
            alias = 'E' + hashlib.md5(path.encode()).hexdigest()[:8]
            page.insert_font(fontname=alias, fontfile=path)
    if not alias:
        low = span_font.lower()
        b, it = 'bold' in low or 'black' in low or 'heavy' in low, 'italic' in low or 'oblique' in low
        fam = 'co' if 'cour' in low or 'mono' in low else 'ti' if 'times' in low or 'serif' in low and 'sans' not in low else 'he'
        alias = {('co', 0, 0): 'cour', ('co', 1, 0): 'cobo', ('co', 0, 1): 'coit', ('co', 1, 1): 'cobi',
                 ('ti', 0, 0): 'tiro', ('ti', 1, 0): 'tibo', ('ti', 0, 1): 'tiit', ('ti', 1, 1): 'tibi',
                 ('he', 0, 0): 'helv', ('he', 1, 0): 'hebo', ('he', 0, 1): 'heit', ('he', 1, 1): 'hebi'}[(fam, int(b), int(it))]
    cache[(span_font, text)] = alias
    return alias

def _pdf_tokens(data):
    """Operands and operators of a content stream; strings, arrays and dicts come back as placeholders."""
    i, n = 0, len(data)
    ws = b' \t\r\n\x0c\x00'
    while i < n:
        ch = data[i:i + 1]
        if ch in (b' ', b'\t', b'\r', b'\n', b'\x0c', b'\x00'): i += 1; continue
        if ch == b'%':
            while i < n and data[i:i + 1] not in (b'\r', b'\n'): i += 1
            continue
        if ch == b'(':
            depth, i = 1, i + 1
            while i < n and depth:
                c = data[i:i + 1]
                if c == b'\\': i += 2; continue
                depth += (c == b'(') - (c == b')'); i += 1
            yield ('str', None); continue
        if data[i:i + 2] in (b'<<', b'>>'): yield ('dict', None); i += 2; continue
        if ch == b'<':
            j = data.find(b'>', i); i = (j + 1) if j >= 0 else n; yield ('str', None); continue
        if ch in (b'[', b']', b'{', b'}'): yield ('arr', None); i += 1; continue
        j = i + 1
        while j < n and data[j:j + 1] not in ws and data[j:j + 1] not in (b'(', b')', b'<', b'>', b'[', b']', b'/', b'%', b'{', b'}'): j += 1
        tok = data[i:j]; i = j
        if ch == b'/': yield ('name', tok); continue
        try: yield ('num', float(tok))
        except ValueError: yield ('op', tok.decode('latin-1'))

def _text_slants(page):
    """Where each run of text starts on the page and how much it leans (as a shear factor), read from the
    drawing commands. Some PDF makers (browsers printing a web page, for one) draw italics by slanting a
    regular font this way instead of using an italic font."""
    runs = []
    try: data = page.read_contents()
    except Exception: return runs
    ident = fitz.Matrix(1, 0, 0, 1, 0, 0)
    ctm, stack, tm, tlm, ops, th = fitz.Matrix(ident), [], fitz.Matrix(ident), fitz.Matrix(ident), [], 1.0
    to_page = page.transformation_matrix
    for kind, v in _pdf_tokens(data):
        if kind != 'op': ops.append(v); continue
        nums = [x for x in ops if isinstance(x, float)]; ops = []
        if v == 'q': stack.append((fitz.Matrix(ctm), th))
        elif v == 'Q': ctm, th = stack.pop() if stack else (ctm, th)
        elif v == 'Tz' and nums: th = nums[-1] / 100   # horizontal stretch, in percent
        elif v == 'cm' and len(nums) >= 6: ctm = fitz.Matrix(*nums[-6:]) * ctm
        elif v == 'BT': tm, tlm = fitz.Matrix(ident), fitz.Matrix(ident)
        elif v == 'Tm' and len(nums) >= 6: tm = tlm = fitz.Matrix(*nums[-6:])
        elif v in ('Td', 'TD') and len(nums) >= 2: tm = tlm = fitz.Matrix(1, 0, 0, 1, nums[-2], nums[-1]) * tlm
        elif v in ('Tj', 'TJ', "'", '"'):
            m = tm * ctm * to_page
            ux, uy, vx, vy = m.a, m.b, m.c, m.d
            cross = ux * vy - uy * vx
            if abs(cross) < 1e-9: continue
            lu = (ux * ux + uy * uy) ** 0.5
            height = abs(cross) / lu                          # glyph height across the baseline
            k = (ux * vx + uy * vy) / (lu * height)           # lean along the baseline, per unit of height
            hs = th * lu / height                             # width stretch (Tz, or a squashed matrix)
            o = fitz.Point(0, 0) * m
            runs.append((o.x, o.y, k, hs))
    return runs

def _run_at(runs, x, y):
    """(lean, width stretch) of the text run that starts at (x, y), or the one just before it on that line."""
    if not runs: return 0, 1
    best = min(runs, key=lambda r: (r[0] - x) ** 2 + (r[1] - y) ** 2)
    if (best[0] - x) ** 2 + (best[1] - y) ** 2 < 2.5: return best[2], best[3]
    near = [r for r in runs if abs(r[1] - y) < 1.5 and r[0] <= x + 1.5]
    r = max(near, key=lambda r: r[0]) if near else None
    return (r[2], r[3]) if r else (0, 1)

def _slant_at(runs, x, y):
    return _run_at(runs, x, y)[0]

def _trace_looks(page):
    """How each run of text is painted: [(first letter's origin, mode, opacity)], mode as PDF's text render
    mode (0 filled, 1 outlined, 3 invisible like an OCR layer)."""
    out = []
    try:
        for sp in page.get_texttrace():
            if not sp['chars']: continue
            t = sp.get('type', 0)
            mode = {0: 0, 1: 1, 3: 3}.get(t, 0)
            out.append((sp['chars'][0][2], mode, sp.get('opacity', 1) or 1))
    except Exception: pass
    return out

def _page_spans(page):
    """Every run of text on the page, letter by letter, with its line's direction."""
    spans = []
    for b in page.get_text('rawdict')['blocks']:
        if b['type'] != 0: continue
        for l in b['lines']:
            for sp in l['spans']:
                sp['text'] = ''.join(c['c'] for c in sp['chars'])
                sp['dir'] = l.get('dir', (1, 0))
                if sp['text'].strip(): spans.append(sp)
    return spans

def _glyph_key(c):
    return (c['c'], round(c['origin'][0], 1), round(c['origin'][1], 1))

def _tight_box(chars):
    ink = [c['bbox'] for c in chars if c['c'].strip()]
    if not ink: return None
    return fitz.Rect(min(b[0] for b in ink), min(b[1] for b in ink), max(b[2] for b in ink), max(b[3] for b in ink)) + (0.3, 0.5, -0.3, -0.5)

def _repair_neighbours(page, spans, edited):
    """Removing words can make the PDF's own cleanup slide the text after them (when that text was placed
    relative to the removed words). Find letters that are no longer where they were, take away the slid
    copies, and return the runs to write back in their original places."""
    before = {_glyph_key(c) for sp in spans for c in sp['chars'] if c['c'].strip()}
    restore, done = [], set(edited)
    for _ in range(3):
        now = {_glyph_key(c): c for sp in _page_spans(page) for c in sp['chars'] if c['c'].strip()}
        stray = [c for k, c in now.items() if k not in before]
        damaged = [sp for sp in spans if id(sp) not in done and any(_glyph_key(c) not in now for c in sp['chars'] if c['c'].strip())]
        if not stray and not damaged: break
        for c in stray: page.add_redact_annot(fitz.Rect(c['bbox']) + (0.3, 0.5, -0.3, -0.5), fill=False, cross_out=False)
        for sp in damaged:   # take away what's left of it too, then write it back whole
            left = [c for c in sp['chars'] if _glyph_key(c) in now]
            box = _tight_box(left)
            if box: page.add_redact_annot(box, fill=False, cross_out=False)
            done.add(id(sp)); restore.append(sp)
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE)
    return restore

RTL_TEXT = re.compile('[\u0590-\u08ff\ufb1d-\ufdff\ufe70-\ufefe]')   # Hebrew, Arabic and their joined forms

def _html_font_family(font):
    """A font name as CSS knows it ('ABCDEF+NotoNaskhArabic-Regular' -> 'NotoNaskhArabic')."""
    return re.sub(r'^.*[^\x00-\x7f]', '', (font or '').split('+')[-1]).split('-')[0].split(',')[0] or 'sans-serif'

def _rewrite_text(doc, page, edits):
    """Replace edited words with the same look the originals had. Each edit says where the original text
    sat (rect, its baseline y) and what it should say now; moved or restyled text also carries
    dx/dy and the new size, color, bold and italic."""
    spans = _page_spans(page)
    plans = []
    for e in edits:
        r = fitz.Rect(*map(float, e['rect']))
        by = float(e.get('base_y', r.y1))
        def inside(sp):   # most of the run's letters sit in the edited area (not a neighbour that touches it)
            xs = [(c['bbox'][0] + c['bbox'][2]) / 2 for c in sp['chars'] if c['c'].strip()]
            return xs and sum(r.x0 - 1 <= x <= r.x1 + 1 for x in xs) * 2 >= len(xs)
        hit = [s for s in spans if abs(s['origin'][1] - by) < max(2, s['size'] * 0.3) and inside(s)]
        if not hit: continue
        hit.sort(key=lambda s: s['origin'][0])
        plans.append((e, hit))
        for s in hit:   # remove just these glyphs: no white box, lines, pictures and neighbours untouched
            ink = [c['bbox'] for c in s['chars'] if c['c'].strip()]
            box = fitz.Rect(min(b[0] for b in ink), min(b[1] for b in ink), max(b[2] for b in ink), max(b[3] for b in ink)) if ink else fitz.Rect(s['bbox'])
            page.add_redact_annot(box + (0.3, 0.5, -0.3, -0.5), fill=False, cross_out=False)
    if not plans: return
    slants = _text_slants(page)   # read before the redactions rewrite the page
    looks = _trace_looks(page)
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE)
    for sp in _repair_neighbours(page, spans, {id(s) for _, hit in plans for s in hit}):
        plans.append(({'new_text': sp['text']}, [sp]))
    _rewrite_text_plans(doc, page, plans, slants, looks)

def _rewrite_text_plans(doc, page, plans, slants, looks):
    """Write each (edit, original runs) pair: see _rewrite_text."""
    cache = {}
    for e, hit in plans:
        text = str(e.get('new_text', '')).replace('\n', ' ').strip()
        if not text: continue
        s = hit[0]
        size, font, flags = s['size'], s['font'], s['flags']
        c = s['color']; rgb = ((c >> 16 & 255) / 255, (c >> 8 & 255) / 255, (c & 255) / 255)
        restyled = False
        if e.get('styled'):   # the user changed the look on purpose; anything they didn't touch stays
            if e.get('size_ratio') and abs(float(e['size_ratio']) - 1) > 0.01: size *= float(e['size_ratio']); restyled = True
            if e.get('color'): rgb = parse_color(e['color'])
            low = font.lower()
            bold0 = bool(flags & 16) or 'bold' in low
            italic0 = bool(flags & 2) or 'italic' in low or 'oblique' in low
            bold = e['bold'] if 'bold' in e else bold0
            italic = e['italic'] if 'italic' in e else italic0
            if bold != bold0 or italic != italic0 or e.get('font_family'):
                fam = e.get('font_family') or re.sub(r'^.*[^\x00-\x7f]', '', font.split('+')[-1]).split('-')[0].split(',')[0]
                style = ('Bold' if bold else '') + ('Italic' if italic else '')
                font = fam + ('-' + style if style else ''); restyled = True
        k, hs = _run_at(slants, *s['origin'])
        if e.get('styled') and 'italic' in e: k = 0
        if abs(hs - 1) > 0.02: size /= hs ** 0.5   # PDF reports stretched text at the mean of width and height; take the height
        else: hs = 1
        f = _font_for(doc, page, font, text, cache, prefer_embedded=not restyled)
        # the line may run in any direction (turned or tilted text): lay it out flat from its first letter,
        # then turn the result to match
        ddx, ddy = s['dir']
        o0 = fitz.Point(s['origin'])
        flat = [dict(c, origin=(o0.x + (c['origin'][0] - o0.x) * ddx + (c['origin'][1] - o0.y) * ddy, o0.y))
                for t in hit for c in t['chars']]
        dx, dy = float(e.get('dx', 0)), float(e.get('dy', 0))
        x, y = o0.x + dx, o0.y + dy
        placed = _layout(flat, text, f, size, keep_positions=not restyled, hscale=hs)
        look = min(looks, key=lambda L: (L[0][0] - o0.x) ** 2 + (L[0][1] - o0.y) ** 2) if looks else None
        mode, opacity = (look[1], look[2]) if look and (look[0][0] - o0.x) ** 2 + (look[0][1] - o0.y) ** 2 < 4 else (0, 1)
        # (morphs work with the page's y axis flipped, so the turn goes the other way round)
        m = fitz.Matrix(hs, 0, k, 1, 0, 0) * fitz.Matrix(ddx, -ddy, ddy, ddx, 0, 0)
        morph = (fitz.Point(x, y), m) if (abs(k) > 0.01 or hs != 1 or abs(ddx - 1) > 1e-3) else None
        op = opacity if opacity < 1 else -1
        old_text = ''.join(c['c'] for t in hit for c in t['chars']).strip()
        if RTL_TEXT.search(text) and text != old_text and abs(ddx - 1) < 1e-3 and mode != 3:   # (unchanged: keep its exact glyphs)
            # Arabic, Hebrew and the like: letters must join up and run right to left, so let the layout
            # engine (with its text shaping) set the words, ending where the original line ended
            right = max(c['bbox'][2] for t in hit for c in t['chars']) + dx
            css = (f"* {{margin:0;padding:0}} p {{font-family:'{_html_font_family(font)}',sans-serif; font-size:{size:.2f}px;"
                   f" color:rgb({rgb[0]*255:.0f},{rgb[1]*255:.0f},{rgb[2]*255:.0f}); white-space:nowrap}}")
            import html as _h
            body = f'<p dir="rtl">{_h.escape(text)}</p>'
            # set it once on a scratch page to learn its width and where its baseline falls, then place it
            # so it ends where the original line ended, at full size
            scratch = fitz.open().new_page(width=4000, height=size * 4 + 20)
            scratch.insert_htmlbox(fitz.Rect(0, 0, 4000, size * 4 + 20), body, css=css + ' p {text-align:left}')
            got = [sp for b in scratch.get_text('dict')['blocks'] if b['type'] == 0 for l in b['lines'] for sp in l['spans'] if sp['text'].strip()]
            if got:
                left = min(sp['bbox'][0] for sp in got); width = max(sp['bbox'][2] for sp in got) - left
                base = max(sp['origin'][1] for sp in got)
                box = fitz.Rect(right - width - 3, y - base, right + 1, y - base + size * 4 + 20)
                page.insert_htmlbox(box, body, css=css + ' p {text-align:right}', opacity=1 if op < 0 else op)
        elif abs(ddx - 1) < 1e-3:   # level text: one go
            tw = fitz.TextWriter(page.rect)
            for px, ch in placed:
                if ch.isspace(): continue
                # positions are already stretched; the morph stretches again, so take that back out first
                tw.append((x + (px - o0.x) / hs, y), ch, font=f, fontsize=size)
            tw.write_text(page, color=rgb, morph=morph, render_mode=mode, opacity=op)
        else:                     # turned text: each letter at its spot along the line, turned in place
            for px, ch in placed:
                if ch.isspace(): continue
                t = px - o0.x
                at = fitz.Point(x + t * ddx, y + t * ddy)
                tw = fitz.TextWriter(page.rect)
                tw.append(at, ch, font=f, fontsize=size)
                tw.write_text(page, color=rgb, morph=(at, m), render_mode=mode, opacity=op)
        if e.get('underline') and placed and mode != 3:
            end_x = placed[-1][0] + dx + f.text_length(placed[-1][1], size) * hs
            page.draw_line((x, y + size * 0.12), (end_x, y + size * 0.12), color=rgb, width=max(0.5, size / 18))

@app.route('/fonts', methods=['GET'])
def fonts():
    """Installed font families (from the Windows font list), for the font picker."""
    _find_system_font('')   # builds the list
    import winreg
    fams = set()
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try: k = winreg.OpenKey(hive, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts')
        except OSError: continue
        i = 0
        while True:
            try: label, file, _ = winreg.EnumValue(k, i)
            except OSError: break
            i += 1
            if not str(file).lower().endswith(('.ttf', '.otf')): continue
            for part in re.sub(r'\s*\(.*\)$', '', label).split(' & '):
                fam = re.sub(r'\s+(Bold|Italic|Oblique|Regular)(\s+(Bold|Italic|Oblique))*$', '', part).strip()
                if fam and not fam.startswith('@'): fams.add(fam)
    return jsonify({"fonts": sorted(fams, key=str.lower)})

# ---------- paragraphs ----------
def _rewrap_paragraph(doc, page, rect, text):
    """Replace the paragraph inside `rect` with `text`, rewrapped to the paragraph's width in its own font,
    size, color, line spacing and alignment. Returns how many lines it now takes, and how many it took."""
    rect = fitz.Rect(rect)
    spans = _page_spans(page)
    mine = [sp for sp in spans if sp['dir'] == (1, 0) and fitz.Rect(sp['bbox']).intersects(rect)
            and rect.contains(fitz.Point((sp['bbox'][0] + sp['bbox'][2]) / 2, (sp['bbox'][1] + sp['bbox'][3]) / 2))]
    if not mine: raise ValueError("Couldn't find that paragraph")
    # the look most of its letters have
    weight = {}
    for sp in mine: key = (sp['font'], round(sp['size'], 2), sp['color'], sp['flags']); weight[key] = weight.get(key, 0) + len(sp['text'].strip())
    font, size, color, flags = max(weight, key=weight.get)
    rgb = ((color >> 16 & 255) / 255, (color >> 8 & 255) / 255, (color & 255) / 255)
    # its lines: baselines, edges
    base = sorted({round(sp['origin'][1], 1) for sp in mine})
    lines = []
    for y in base:
        row = [sp for sp in mine if abs(sp['origin'][1] - y) < 0.6]
        lines.append((y, min(c['origin'][0] for sp in row for c in sp['chars'] if c['c'].strip()), max(c['bbox'][2] for sp in row for c in sp['chars'] if c['c'].strip())))
    pitch = sorted(b - a for a, b in zip(base, base[1:]))[len(base) // 2 - 1] if len(base) > 1 else size * 1.2
    left = min(l[1] for l in lines); right = max(l[2] for l in lines)
    width = right - left
    full = lines[:-1] or lines
    if len(lines) > 1 and all(abs(l[2] - right) < 1.5 for l in full) and all(abs(l[1] - left) < 1.5 for l in lines):
        align = 'justify'
    elif len(lines) > 1 and all(abs((l[1] + l[2]) / 2 - (left + right) / 2) < 2 for l in lines) and any(abs(l[1] - left) > 3 for l in lines):
        align = 'center'
    elif len(lines) > 1 and all(abs(l[2] - right) < 1.5 for l in lines) and any(abs(l[1] - left) > 3 for l in lines):
        align = 'right'
    else:
        align = 'left'
    slants = _text_slants(page); k, hs = _run_at(slants, *mine[0]['origin'])
    if abs(hs - 1) > 0.02: size /= hs ** 0.5
    else: hs = 1
    looks = _trace_looks(page)
    f = _font_for(doc, page, font, text, {})
    # take the old words away (only them), and put back anything nearby that slid
    for sp in mine:
        box = _tight_box(sp['chars'])
        if box: page.add_redact_annot(box, fill=False, cross_out=False)
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE)
    restore = _repair_neighbours(page, spans, {id(sp) for sp in mine})
    # each original word with its own look, so bold or slanted words stay that way
    styles, words0 = {}, []
    for sp in mine:
        c0 = sp['color']
        col = ((c0 >> 16 & 255) / 255, (c0 >> 8 & 255) / 255, (c0 & 255) / 255)
        word = []
        for ch in sp['chars'] + [{'c': ' ', 'origin': (0, 0)}]:
            if ch['c'].strip(): word.append(ch); continue
            if not word: continue
            # a run can mix upright and slanted words: look the slant up where this word starts
            kk, hh = _run_at(slants, *word[0]['origin'])
            sz = sp['size'] / (hh ** 0.5) if abs(hh - 1) > 0.02 else sp['size']
            words0.append((''.join(c['c'] for c in word), (sp['font'], round(sz, 2), col, round(kk, 3), round(hh if abs(hh - 1) > 0.02 else 1, 3))))
            word = []
    main = (font, size, rgb, round(k, 3), hs)
    import difflib
    new_words = [(para_i, w) for para_i, para in enumerate(text.replace('\r', '').split('\n')) for w in (para.split() or [''])]
    look_of = [main] * len(new_words)
    sm = difflib.SequenceMatcher(None, [w for w, _ in words0], [w for _, w in new_words], autojunk=False)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        for jj in range(j1, j2):
            if op == 'equal': look_of[jj] = words0[i1 + jj - j1][1]
            elif words0: look_of[jj] = words0[min(max(i1 - 1, 0), len(words0) - 1)][1]   # new words look like the one before
    fonts = {}
    def font_of(st):
        if st[0] not in fonts: fonts[st[0]] = _font_for(doc, page, st[0], text, {})
        return fonts[st[0]]
    wlen = lambda w, st: font_of(st).text_length(w, st[1]) * st[4]
    space = (f.text_length(' ', size) * hs) or size * 0.25
    # wrap
    out_lines, cur, last_para = [], [], None
    for (pi, w), st in zip(new_words, look_of):
        if last_para is not None and pi != last_para:
            out_lines.append((cur, True)); cur = []
        last_para = pi
        if not w: continue
        if cur and sum(wlen(a, b) for a, b in cur) + space * len(cur) + wlen(w, st) > width + 0.5:
            out_lines.append((cur, False)); cur = []
        cur.append((w, st))
    out_lines.append((cur, True))
    y = lines[0][0]
    for words, last in out_lines:
        if words:
            ink = sum(wlen(a, b) for a, b in words)
            gap = space
            if align == 'justify' and not last and len(words) > 1: gap = (width - ink) / (len(words) - 1)
            line_w = ink + gap * (len(words) - 1)
            x = left if align in ('left', 'justify') else (right - line_w if align == 'right' else left + (width - line_w) / 2)
            for w, st in words:
                fo, sz, col, kk, hh = font_of(st), st[1], st[2], st[3], st[4]
                tw = fitz.TextWriter(page.rect)
                tw.append((x, y), w, font=fo, fontsize=sz)
                m = fitz.Matrix(hh, 0, kk, 1, 0, 0)
                tw.write_text(page, color=col, morph=(fitz.Point(x, y), m) if (abs(kk) > 0.01 or hh != 1) else None)
                x += wlen(w, st) + gap
        y += pitch
    for sp in restore:   # neighbours that slid, back where they were
        _rewrite_text_plans(doc, page, [({'new_text': sp['text']}, [sp])], slants, looks)
    return len([l for l in out_lines]), len(lines)

@app.route('/rewrite_paragraph', methods=['POST'])
def rewrite_paragraph():
    d = request.json; fp = d['filepath']
    try:
        doc = fitz.open(fp); page = doc[int(d['page'])]
        now, before = _rewrap_paragraph(doc, page, d['rect'], d['text'])
        try: doc.subset_fonts()
        except Exception: pass
        new_path = os.path.join(TEMP_DIR, f"mod_{hashlib.md5((fp + str(os.path.getmtime(fp)) + d['text']).encode()).hexdigest()}.pdf")
        doc.save(new_path, garbage=3, deflate=True); doc.close()
        note = '' if now <= before else f'. It now takes {now - before} more {"line" if now - before == 1 else "lines"}; check it doesn\'t run into what\'s below'
        return jsonify({"success": True, "new_file": new_path, "message": "Paragraph updated" + note})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/burn_edits', methods=['POST'])
def burn_edits():
    data = request.json
    filepath = data.get('filepath')
    page_index = int(data.get('page', 0))

    try:
        doc = fitz.open(filepath)
        page = doc[page_index]
        
        x_off = page.rect.x0
        y_off = page.rect.y0

        # 1. Edited text: rewritten so it looks untouched (see _rewrite_text)
        edits = data.get('text_edits', [])
        if edits:
            try: _rewrite_text(doc, page, edits)
            except Exception as e: print(f"Skipped text edits: {e}", file=sys.stderr)

        # 2. Add New Text ("T"): in the picked font (any installed one), with real bold/italic faces
        font_cache = {}
        for nt in data.get('new_texts', []):
            try:
                rgb = parse_color(nt['color'])
                fam = str(nt.get('font') or 'Arial').replace('"', '').replace("'", '').split(',')[0].strip()
                is_bold = bool(nt.get('bold'))
                is_italic = bool(nt.get('italic'))
                style = ('Bold' if is_bold else '') + ('Italic' if is_italic else '')
                text = str(nt['text'])
                alias = _text_font(doc, page, fam + ('-' + style if style else ''), text, font_cache)
                actual_size = float(nt['font_size']) * 0.82
                page.insert_text(fitz.Point(float(nt['x']) + x_off, float(nt['y']) + y_off - 1.0), text,
                                 fontsize=actual_size, color=rgb, fontname=alias)
            except Exception as e:
                print(f"Skipped new text: {e}")

        # 3. Insert Images/Signatures
        for img in data.get('images', []):
            try:
                rgb_data = img['b64']
                if rgb_data.startswith('data:image'): rgb_data = rgb_data.split(',')[1]
                img_bytes = base64.b64decode(rgb_data)
                
                rect = fitz.Rect(
                    float(img['x']) + x_off, 
                    float(img['y']) + y_off, 
                    float(img['x']) + float(img['w']) + x_off, 
                    float(img['y']) + float(img['h']) + y_off
                )
                page.insert_image(rect, stream=img_bytes)
            except Exception as e:
                print(f"Skipped image: {e}")

        # 4. Burn Drawings ("H" & Pen)
        hl_groups = {} 
        pens = []

        for draw in data.get('drawings', []):
            if len(draw['points']) > 1:
                if draw['type'] == 'highlighter':
                    key = (draw['color'], float(draw['size']))
                    if key not in hl_groups:
                        hl_groups[key] = []
                    
                    pdf_points = [(float(p[0]) + x_off, float(p[1]) + y_off) for p in draw['points']]
                    hl_groups[key].append(pdf_points)
                else:
                    pens.append(draw)

        for (color, size), ink_list in hl_groups.items():
            try:
                rgb = parse_color(color)
                annot = page.add_ink_annot(ink_list)
                annot.set_colors(stroke=rgb)
                annot.set_border({"width": size})
                annot.set_opacity(0.75) 
                try:
                    annot.update(blend_mode=fitz.PDF_BM_Multiply)
                except:
                    annot.update()
            except Exception as e:
                print(f"Skipped highlighter group: {e}")

        for draw in pens:
            try:
                rgb = parse_color(draw['color'])
                pdf_points = [(float(p[0]) + x_off, float(p[1]) + y_off) for p in draw['points']]
                ink_list = [pdf_points]
                annot = page.add_ink_annot(ink_list)
                annot.set_colors(stroke=rgb)
                annot.set_border({"width": float(draw['size'])})
                annot.set_opacity(float(draw.get('opacity', 1.0)))
                annot.update()
            except Exception as e:
                print(f"Skipped pen: {e}")

        hasher = hashlib.md5((filepath + str(os.path.getmtime(filepath))).encode()).hexdigest()
        new_path = os.path.join(TEMP_DIR, f"mod_{hasher}.pdf")
        if edits or data.get('new_texts'):   # an installed font goes in whole; keep only the letters used so the file stays small
            try: doc.subset_fonts()
            except Exception as e: print('font subsetting skipped:', e, file=sys.stderr)
        doc.save(new_path, garbage=3, deflate=True)
        doc.close()
        return jsonify({"success": True, "new_file": new_path})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/compress_pdf', methods=['POST'])
def compress_pdf():
    data = request.json
    filepath = data.get('filepath')
    try:
        doc = fitz.open(filepath)
        hasher = hashlib.md5((filepath + str(os.path.getmtime(filepath))).encode()).hexdigest()
        new_path = os.path.join(TEMP_DIR, f"comp_{hasher}.pdf")
        doc.save(new_path, garbage=4, deflate=True)
        doc.close()
        return jsonify({"success": True, "new_file": new_path})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/crop_page', methods=['POST'])
def crop_page():
    data = request.json
    filepath = data.get('filepath')
    page_index = int(data.get('page', 0))
    crop_rect = data.get('rect') 

    try:
        doc = fitz.open(filepath)
        page = doc[page_index]
        
        x_off = page.rect.x0
        y_off = page.rect.y0
        
        rect = fitz.Rect(
            crop_rect[0] + x_off,
            crop_rect[1] + y_off,
            crop_rect[2] + x_off,
            crop_rect[3] + y_off
        )
        
        page.set_cropbox(rect)

        hasher = hashlib.md5((filepath + str(os.path.getmtime(filepath))).encode()).hexdigest()
        new_path = os.path.join(TEMP_DIR, f"mod_{hasher}.pdf")
        doc.save(new_path)
        doc.close()
        return jsonify({"success": True, "new_file": new_path})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/number_pages', methods=['POST'])
def number_pages():
    data = request.json
    filepath = data.get('filepath')
    start_num = int(data.get('start_num', 1))

    try:
        doc = fitz.open(filepath)
        for i in range(len(doc)):
            page = doc[i]
            rect = page.rect
            text = str(start_num + i)
            p = fitz.Point(rect.x1 - 40, rect.y1 - 30)
            page.insert_text(p, text, fontsize=12, color=(0,0,0), fontname="helv")

        hasher = hashlib.md5((filepath + str(os.path.getmtime(filepath))).encode()).hexdigest()
        new_path = os.path.join(TEMP_DIR, f"mod_{hasher}.pdf")
        doc.save(new_path)
        doc.close()
        return jsonify({"success": True, "new_file": new_path})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/delete_page', methods=['POST'])
def delete_page():
    data = request.json
    filepath = data.get('filepath')
    page_index = data.get('page') 
    try:
        doc = fitz.open(filepath)
        if len(doc) <= 1: return jsonify({"error": "Cannot delete the last page."}), 400
        doc.delete_page(page_index)
        hasher = hashlib.md5((filepath + str(os.path.getmtime(filepath))).encode()).hexdigest()
        new_path = os.path.join(TEMP_DIR, f"mod_{hasher}.pdf")
        doc.save(new_path)
        doc.close()
        return jsonify({"success": True, "new_file": new_path})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/reorder_pages', methods=['POST'])
def reorder_pages():
    data = request.json
    filepath = data.get('filepath')
    new_order = data.get('new_order')
    try:
        if not new_order or len(new_order) == 0:
            return jsonify({"error": "Invalid order array received."}), 400
        
        doc = fitz.open(filepath)
        # Secure the array: only select valid integer indices to prevent the "bad page number" crash
        valid_order = [int(i) for i in new_order if isinstance(i, (int, str)) and str(i).isdigit()]
        
        doc.select(valid_order)
        hasher = hashlib.md5((filepath + str(os.path.getmtime(filepath))).encode()).hexdigest()
        new_path = os.path.join(TEMP_DIR, f"mod_{hasher}.pdf")
        # Add light garbage collection to keep the file perfectly optimized after reordering
        doc.save(new_path, garbage=3)
        doc.close()
        return jsonify({"success": True, "new_file": new_path})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/rotate_page', methods=['POST'])
def rotate_page():
    data = request.json
    filepath = data.get('filepath')
    page_index = data.get('page')
    try:
        doc = fitz.open(filepath)
        page = doc[page_index]
        page.set_rotation((page.rotation + 90) % 360)
        hasher = hashlib.md5((filepath + str(os.path.getmtime(filepath))).encode()).hexdigest()
        new_path = os.path.join(TEMP_DIR, f"mod_{hasher}.pdf")
        doc.save(new_path)
        doc.close()
        return jsonify({"success": True, "new_file": new_path})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/pages_op', methods=['POST'])
def pages_op():
    """Several pages at once (0-based `pages`): delete them, turn them 90 degrees, or save them as a new PDF
    next to the original."""
    d = request.json; fp = d['filepath']; op = d['op']
    pages = sorted({int(i) for i in d.get('pages', [])})
    try:
        doc = fitz.open(fp)
        pages = [i for i in pages if 0 <= i < len(doc)]
        if not pages: return jsonify({"error": "No pages picked"}), 400
        if op == 'extract':
            orig = d.get('original_path') or fp
            base = os.path.splitext(os.path.basename(orig))[0]
            label = _page_label(pages)
            out = get_unique_filename(os.path.join(os.path.dirname(orig), f"{base} (pages {label}).pdf"))
            new = fitz.open()
            for i in pages: new.insert_pdf(doc, from_page=i, to_page=i)
            new.save(out, garbage=3, deflate=True); new.close(); doc.close()
            return jsonify({"success": True, "path": out, "message": f"Saved {os.path.basename(out)}"})
        if op == 'delete':
            if len(pages) >= len(doc): return jsonify({"error": "A PDF needs at least one page"}), 400
            doc.delete_pages(pages)
        elif op == 'rotate':
            for i in pages: doc[i].set_rotation((doc[i].rotation + 90) % 360)
        else: return jsonify({"error": "Unknown action"}), 400
        hasher = hashlib.md5((fp + str(os.path.getmtime(fp)) + op + str(pages)).encode()).hexdigest()
        new_path = os.path.join(TEMP_DIR, f"mod_{hasher}.pdf")
        doc.save(new_path, garbage=3); doc.close()
        return jsonify({"success": True, "new_file": new_path})
    except Exception as e: return jsonify({"error": str(e)}), 500

def _page_label(pages):
    """[0,1,2,4] -> '1-3, 5'"""
    runs, start = [], pages[0]
    for a, b in zip(pages, pages[1:] + [None]):
        if b != a + 1: runs.append(f"{start + 1}" if start == a else f"{start + 1}-{a + 1}"); start = b
    return ', '.join(runs)

@app.route('/render_pages', methods=['POST'])
def render_pages():
    """Page pictures for printing: JPEGs at print quality, with each page's size in points."""
    fp = request.json['filepath']
    try:
        doc = fitz.open(fp)
        out = tempfile.mkdtemp(dir=TEMP_DIR, prefix='print_')
        pages = []
        for i, p in enumerate(doc):
            f = os.path.join(out, f'{i + 1}.jpg')
            p.get_pixmap(dpi=200).save(f, jpg_quality=92)
            pages.append({"path": f, "w": p.rect.width, "h": p.rect.height})
        doc.close()
        return jsonify({"success": True, "pages": pages})
    except Exception as e: return jsonify({"error": str(e)}), 500

async def _ocr_png(png):
    """Words and their boxes (in pixels) found by Windows' built-in text recognition."""
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.graphics.imaging import BitmapDecoder
    from winrt.windows.storage.streams import InMemoryRandomAccessStream, DataWriter
    stream = InMemoryRandomAccessStream()
    w = DataWriter(stream); w.write_bytes(png); await w.store_async(); await w.flush_async(); w.detach_stream()
    stream.seek(0)
    bmp = await (await BitmapDecoder.create_async(stream)).get_software_bitmap_async()
    engine = OcrEngine.try_create_from_user_profile_languages()
    if engine is None: raise RuntimeError("Windows text recognition isn't available. Add a language with OCR in Windows Settings.")
    res = await engine.recognize_async(bmp)
    return [(wd.text, wd.bounding_rect) for line in res.lines for wd in line.words]

@app.route('/ocr', methods=['POST'])
def ocr_pdf():
    """Make scanned pages searchable: recognized words go onto the page as invisible text right over the
    picture of them, so Find, copy and export work while the page looks exactly the same.
    Pages that already have text are left alone."""
    import asyncio
    fp = request.json['filepath']
    try:
        doc = fitz.open(fp)
        done = words = 0
        for page in doc:
            if len(page.get_text().strip()) >= 20: continue
            dpi = 300; k = 72 / dpi
            found = asyncio.run(_ocr_png(page.get_pixmap(dpi=dpi).tobytes('png')))
            for text, r in found:
                x0, y0, w, h = r.x * k, r.y * k, r.width * k, r.height * k
                size = min(w / max(fitz.get_text_length(text, fontname='helv', fontsize=1), 0.01), h * 1.3)
                page.insert_text((x0, y0 + h * 0.85), text, fontsize=max(1, size), fontname='helv', render_mode=3)
            done += 1; words += len(found)
        if not done:
            doc.close()
            return jsonify({"success": True, "new_file": None, "message": "This PDF already has searchable text"})
        hasher = hashlib.md5((fp + str(os.path.getmtime(fp)) + 'ocr').encode()).hexdigest()
        new_path = os.path.join(TEMP_DIR, f"mod_{hasher}.pdf")
        doc.save(new_path, garbage=3, deflate=True); doc.close()
        return jsonify({"success": True, "new_file": new_path,
                        "message": f"Recognized {words} words on {done} {'page' if done == 1 else 'pages'}. Find and export now work on them"})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/split_pdf', methods=['POST'])
def split_pdf():
    data = request.json
    filepath = data.get('filepath')
    original_folder = os.path.dirname(data.get('original_path', filepath))
    base_name = os.path.splitext(os.path.basename(data.get('original_path')))[0]
    try:
        doc = fitz.open(filepath)
        doc_pages = list(range(len(doc)))
        for i in range(len(doc)):
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=i, to_page=i)
            out_path = get_unique_filename(os.path.join(original_folder, f"{base_name}_Page_{i+1}.pdf"))
            new_doc.save(out_path)
            new_doc.close()
        doc.close()
        return jsonify({"success": True, "message": f"Split into {len(doc_pages)} pages"})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/merge_queue', methods=['POST'])
def merge_queue():
    data = request.json
    files = data.get('files', []) 
    original_folder = os.path.dirname(data.get('first_original_path', ''))
    try:
        merged_doc = fitz.open()
        for pdf_path in files:
            if os.path.exists(pdf_path):
                pdf_to_insert = fitz.open(pdf_path)
                merged_doc.insert_pdf(pdf_to_insert)
                pdf_to_insert.close()
        out_path = get_unique_filename(os.path.join(original_folder, "Combined_Document.pdf"))
        merged_doc.save(out_path)
        merged_doc.close()
        return jsonify({"success": True, "message": f"Saved {os.path.basename(out_path)}"})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/save_modified', methods=['POST'])
def save_modified():
    data = request.json
    current_temp_pdf = data.get('filepath')
    original_path = data.get('original_path')
    try:
        original_folder = os.path.dirname(original_path)
        base_name = os.path.splitext(os.path.basename(original_path))[0]
        out_path = get_unique_filename(os.path.join(original_folder, f"{base_name}_Modified.pdf"))
        shutil.copy(current_temp_pdf, out_path)
        return jsonify({"success": True, "message": f"Saved {os.path.basename(out_path)}"})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/save_current_page', methods=['POST'])
def save_current_page():
    data = request.json
    filepath = data.get('filepath')
    page_index = data.get('page')
    original_path = data.get('original_path')
    try:
        doc = fitz.open(filepath)
        new_doc = fitz.open()
        new_doc.insert_pdf(doc, from_page=page_index, to_page=page_index)
        original_folder = os.path.dirname(original_path)
        base_name = os.path.splitext(os.path.basename(original_path))[0]
        out_path = get_unique_filename(os.path.join(original_folder, f"{base_name}_Extracted_Page_{page_index+1}.pdf"))
        new_doc.save(out_path)
        new_doc.close()
        doc.close()
        return jsonify({"success": True, "message": f"Saved {os.path.basename(out_path)}"})
    except Exception as e: return jsonify({"error": str(e)}), 500


# ---------- PDF -> Word / Excel / PowerPoint ----------
def _cell(v):
    """Turn a cell's text into a number when it clearly is one, so Excel can sum it."""
    if v is None: return None
    s = str(v).strip()
    if not s: return None
    t = s.replace(',', '').replace('$', '')
    if re.fullmatch(r'-?\d+(\.\d+)?', t) and not (len(t) > 1 and t.startswith('0') and not t.startswith('0.')):
        try: return float(t) if '.' in t else int(t)
        except ValueError: pass
    return s

def _text_rows(page):
    """For pages without a detected table: rebuild rows from word positions.
    A wide horizontal gap starts a new cell, and each cell goes in the column
    whose horizontal band it sits in, so blank cells don't shift a row left."""
    words = sorted(page.get_text('words'), key=lambda w: (round(w[3]), w[0]))
    lines = []
    for w in words:
        if lines and abs(lines[-1][0] - w[3]) < 3: lines[-1][1].append(w)
        else: lines.append([w[3], [w]])
    segs_by_line = []   # each cell as (x0, x1, text)
    for _, ws in lines:
        ws.sort(key=lambda w: w[0])
        segs, cur = [], None
        for w in ws:
            if cur and w[0] - cur[1] <= 7: cur = [cur[0], w[2], cur[2] + ' ' + w[4]]
            else:
                if cur: segs.append(cur)
                cur = [w[0], w[2], w[4]]
        if cur: segs.append(cur)
        segs_by_line.append(segs)
    # Column bands: merge the x-ranges of cells from multi-cell lines.
    # Single-cell lines (titles, paragraphs) would join every band, so they don't count.
    bands = []
    for x0, x1, _ in sorted(s for segs in segs_by_line if len(segs) > 1 for s in segs):
        if bands and x0 <= bands[-1][1] + 2: bands[-1][1] = max(bands[-1][1], x1)
        else: bands.append([x0, x1])
    rows = []
    for segs in segs_by_line:
        if len(segs) == 1 or not bands:
            rows.append([s[2] for s in segs]); continue
        row = [None] * len(bands)
        for x0, x1, text in segs:
            mid = (x0 + x1) / 2
            j = min(range(len(bands)), key=lambda k: 0 if bands[k][0] <= mid <= bands[k][1] else min(abs(mid - bands[k][0]), abs(mid - bands[k][1])))
            row[j] = text if row[j] is None else row[j] + ' ' + text
        while row and row[-1] is None: row.pop()
        rows.append(row)
    return rows

def _number_format(s):
    """Excel number format that shows a converted number the way the PDF wrote it ($1,200.00 stays $1,200.00)."""
    t = s.strip()
    dec = len(t.split('.')[1]) if '.' in t else 0
    fmt = ('"$"' if '$' in t else '') + ('#,##0' if ',' in t else '0') + ('.' + '0' * dec if dec else '')
    return None if fmt == '0' else fmt

def _ruling(page):
    """The page's drawn lines as (horizontal, vertical, fills).
    horizontal: (y, x0, x1, thickness); vertical: (x, y0, y1, thickness); fills: (rect, rgb) for shaded areas.
    Office draws cell borders as thin filled boxes, so those count as lines too."""
    hs, vs, fills = [], [], []
    def seg(x0, y0, x1, y1, w):
        if abs(y1 - y0) <= abs(x1 - x0): hs.append(((y0 + y1) / 2, min(x0, x1), max(x0, x1), w))
        else: vs.append(((x0 + x1) / 2, min(y0, y1), max(y0, y1), w))
    def straight(it):   # exactly horizontal or vertical
        return it[0] == 're' or (it[0] == 'l' and (abs(it[1].x - it[2].x) < 0.6 or abs(it[1].y - it[2].y) < 0.6))
    for d in page.get_drawings():
        fill, color, w = d.get('fill'), d.get('color'), d.get('width') or 1
        # Freehand ink (highlights, pen, signatures) is many little segments at odd angles, or very
        # thick; treating its pieces as borders would cut the sheet into dozens of slivers.
        if not all(straight(it) for it in d['items']) or (d['type'] == 's' and w > 4):
            continue
        for it in d['items']:
            if it[0] == 'l' and color is not None:
                seg(it[1].x, it[1].y, it[2].x, it[2].y, w)
            elif it[0] == 're':
                r = it[1]
                if fill is not None and min(r.width, r.height) <= 2.5 and sum(fill) < 2.4:
                    # a box this thin is a line
                    if r.width >= r.height: hs.append(((r.y0 + r.y1) / 2, r.x0, r.x1, r.height))
                    else: vs.append(((r.x0 + r.x1) / 2, r.y0, r.y1, r.width))
                    continue
                if fill is not None and sum(fill) < 2.9 and r.width > 2.5 and r.height > 2.5:
                    fills.append((r, fill))
                if color is not None and d['type'] in ('s', 'fs'):
                    seg(r.x0, r.y0, r.x1, r.y0, w); seg(r.x0, r.y1, r.x1, r.y1, w)
                    seg(r.x0, r.y0, r.x0, r.y1, w); seg(r.x1, r.y0, r.x1, r.y1, w)
    return hs, vs, fills

def _cluster(values, tol=2.0):
    out = []
    for v in sorted(values):
        if out and v - out[-1][-1] <= tol: out[-1].append(v)
        else: out.append([v])
    return [sum(c) / len(c) for c in out]

def _layout_sheet(page, ws):
    """Rebuild a ruled page (a form, invoice or packing list) as a sheet that looks like it:
    columns and rows sit where the PDF's lines and text are, boxed text becomes merged cells,
    and fonts, alignment, borders, shading and pictures are carried over. Returns True if any text was found."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    hs, vs, fills = _ruling(page)

    # text pieces: spans on one line that touch are one piece
    pieces = []
    for b in page.get_text('dict')['blocks']:
        if b['type'] != 0: continue
        for l in b['lines']:
            for s in l['spans']:
                if not s['text'].strip(): continue
                x0, y0, x1, y1 = s['bbox']
                p = pieces[-1] if pieces else None
                if p and abs(p['y0'] - y0) < 1.5 and 0 <= x0 - p['x1'] < 1.5:
                    p['text'] += s['text']; p['x1'] = x1; continue
                pieces.append({'x0': x0, 'y0': y0, 'x1': x1, 'y1': y1, 'text': s['text'], 'size': s['size'],
                               'font': s['font'], 'bold': bool(s['flags'] & 16) or 'bold' in s['font'].lower(),
                               'italic': bool(s['flags'] & 2), 'color': s['color']})
    if not pieces: return False

    # the box drawn around each piece, if it has one on all four sides
    def box(p):
        cx, cy = (p['x0'] + p['x1']) / 2, (p['y0'] + p['y1']) / 2
        left = [x for x, a, b, _ in vs if x <= p['x0'] + 2 and a - 1 <= cy <= b + 1]
        right = [x for x, a, b, _ in vs if x >= p['x1'] - 2 and a - 1 <= cy <= b + 1]
        top = [y for y, a, b, _ in hs if y <= p['y0'] + 2 and a - 1 <= cx <= b + 1]
        bottom = [y for y, a, b, _ in hs if y >= p['y1'] - 2 and a - 1 <= cx <= b + 1]
        if left and right and top and bottom: return max(left), max(top), min(right), min(bottom)
        return None
    for p in pieces: p['box'] = box(p)

    # grid lines: every drawn line, line ends, and the left/top edge of text that isn't boxed
    xs = [x for x, *_ in vs] + [v for _, a, b, _ in hs for v in (a, b)]
    ys = [y for y, *_ in hs] + [v for _, a, b, _ in vs for v in (a, b)]
    for p in pieces:
        if p['box']: continue
        xs.append(p['x0']); ys.append(p['y0'])
    xs.append(max(p['x1'] for p in pieces)); ys.append(max(p['y1'] for p in pieces))
    imgs = [i for i in page.get_image_info(xrefs=True) if i.get('xref')]
    for im in imgs:
        xs.append(im['bbox'][0]); ys.append(im['bbox'][1])
    X, Y = _cluster(xs), _cluster(ys)
    if len(X) < 2 or len(Y) < 2: return False
    def col(x): return min(range(len(X)), key=lambda k: abs(X[k] - x))
    def row(y): return min(range(len(Y)), key=lambda k: abs(Y[k] - y))
    def col_at(x): return max([k for k in range(len(X) - 1) if X[k] <= x + 2] or [0])
    def row_at(y): return max([k for k in range(len(Y) - 1) if Y[k] <= y + 2] or [0])

    # A PDF printed from Excel was usually shrunk to fit the page. Build the sheet at its real size
    # (everything / s) and print it at s, so both the sheet and the printout match the original.
    s = _print_scale(page, hs, vs)
    z = 1 / s

    # sizes: Excel column width is in characters, which print about 5.64pt wide (5.25 on screen); row height in points
    for k in range(len(X) - 1):
        ws.column_dimensions[get_column_letter(k + 1)].width = max(0.3, round((X[k + 1] - X[k]) * z / 5.64, 2))
    for k in range(len(Y) - 1):
        ws.row_dimensions[k + 1].height = max(2, round((Y[k + 1] - Y[k]) * z, 1))

    # text: boxed pieces fill their box (merged), others sit in the cell where they start
    cells = {}
    for p in pieces:
        if p['box']:
            l, t, r, b = p['box']
            key = (row(t), col(l), max(row(b) - 1, row(t)), max(col(r) - 1, col(l)))
            bx0, bx1 = l, r
        else:
            # unboxed text covers every row and column it overlaps, plus some room on the side it reads
            # toward (Excel draws text a bit wider than the PDF), so a thin row or a neighbour can't cut it off.
            # Text that ends right against a line was right-aligned to it.
            rr, re_, cc = row_at(p['y0']), row_at(p['y1'] - 2), col_at(p['x0'])
            cy, w = (p['y0'] + p['y1']) / 2, p['x1'] - p['x0']
            stop = [x for x, a, b, _ in vs if p['x1'] - 2 <= x <= p['x1'] + 8 and a - 1 <= cy <= b + 1]
            if stop:
                ce = max(cc, col(min(stop)) - 1)
                cc = min(cc, col_at(p['x0'] - 0.3 * w))
            else:
                ce = max(cc, min(col_at(p['x1'] + 0.3 * w), len(X) - 2))
            key = (rr, cc, max(rr, re_), ce)
            bx0, bx1 = X[cc], X[min(ce + 1, len(X) - 1)]
        c = cells.get(key)
        if c:   # more lines of text in the same box
            c['lines'].append(p); continue
        mid, half = (p['x0'] + p['x1']) / 2, (bx0 + bx1) / 2
        if p['box'] and abs(mid - half) < max(3, (bx1 - bx0) * 0.08): align = 'center'
        elif p['box'] and bx1 - p['x1'] < p['x0'] - bx0: align = 'right'
        elif not p['box'] and stop: align = 'right'
        else: align = 'left'
        cells[key] = {'lines': [p], 'align': align}

    used = set()
    # boxed text first, so loose text's extra room gives way to it
    for (r0, c0, r1, c1), c in sorted(cells.items(), key=lambda kv: (not kv[1]['lines'][0]['box'], kv[0])):
        span = {(r, k) for r in range(r0, r1 + 1) for k in range(c0, c1 + 1)}
        if span & used:   # overlaps another: shrink to the cells it actually starts in
            p0 = c['lines'][0]
            c0 = c1 = col_at(p0['x0']) if not p0['box'] else c0
            r1 = r0
            while r1 + 1 < len(Y) - 1 and Y[r1 + 1] < p0['y1'] - 2 and (r1 + 1, c0) not in used: r1 += 1
            span = {(r, c0) for r in range(r0, r1 + 1)}
            if (r0, c0) in used: continue
        used |= span
        lines = sorted(c['lines'], key=lambda p: (round(p['y0']), p['x0']))
        text = '\n'.join(p['text'].strip() for p in lines)
        p = lines[0]
        cell = ws.cell(row=r0 + 1, column=c0 + 1)
        val = _cell(text)
        cell.value = val
        if isinstance(val, (int, float)):
            fmt = _number_format(text)
            if fmt: cell.number_format = fmt
        name = re.sub(r'^[A-Z]{6}\+', '', p['font']).split('-')[0].split(',')[0] or 'Calibri'
        cell.font = Font(name=name, size=max(1, round(p['size'] * z * 2) / 2), bold=p['bold'], italic=p['italic'],
                         color='%06X' % (p['color'] & 0xFFFFFF))
        cell.alignment = Alignment(horizontal=c['align'], vertical='center', wrap_text=len(lines) > 1)
        if (r1, c1) != (r0, c0):
            ws.merge_cells(start_row=r0 + 1, start_column=c0 + 1, end_row=r1 + 1, end_column=c1 + 1)

    # shading
    for rect, rgb in fills:
        f = PatternFill('solid', fgColor='%02X%02X%02X' % tuple(round(v * 255) for v in rgb[:3]))
        for r in range(len(Y) - 1):
            for k in range(len(X) - 1):
                if rect.x0 - 1 <= (X[k] + X[k + 1]) / 2 <= rect.x1 + 1 and rect.y0 - 1 <= (Y[r] + Y[r + 1]) / 2 <= rect.y1 + 1:
                    ws.cell(row=r + 1, column=k + 1).fill = f

    # borders: each drawn line becomes the edge of every cell it runs along
    # Excel prints thin / medium / thick as about 0.96 / 1.92 / 2.88pt at 100%; pick the nearest at this scale
    def side(w):
        w *= z
        return Side(style='thick' if w > 2.4 else 'medium' if w > 1.44 else 'thin', color='000000')
    edges = {}   # (row, col) -> {'top': Side, ...}
    for y, a, b, w in hs:
        r = row(y)
        for k in range(len(X) - 1):
            if X[k] >= a - 2 and X[k + 1] <= b + 2:
                if r < len(Y) - 1: edges.setdefault((r, k), {})['top'] = side(w)
                if r > 0: edges.setdefault((r - 1, k), {})['bottom'] = side(w)
    for x, a, b, w in vs:
        k = col(x)
        for r in range(len(Y) - 1):
            if Y[r] >= a - 2 and Y[r + 1] <= b + 2:
                if k < len(X) - 1: edges.setdefault((r, k), {})['left'] = side(w)
                if k > 0: edges.setdefault((r, k - 1), {})['right'] = side(w)
    for (r, k), e in edges.items():
        ws.cell(row=r + 1, column=k + 1).border = Border(**e)

    # pictures (logos, signatures), placed at the same spot
    try:
        from openpyxl.drawing.image import Image as XLImage
        from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
        from openpyxl.drawing.xdr import XDRPositiveSize2D
        EMU = 12700   # per point
        smasks = {g[0]: g[1] for g in page.get_images()}   # transparency masks, or logos come out as black boxes
        for im in imgs:
            x0, y0, x1, y1 = im['bbox']
            pix = fitz.Pixmap(page.parent, im['xref'])
            if smasks.get(im['xref']):   # the shape lives in the mask; the image itself may be one flat color
                if pix.alpha: pix = fitz.Pixmap(pix, 0)
                if pix.n > 3: pix = fitz.Pixmap(fitz.csRGB, pix)
                pix = fitz.Pixmap(pix, fitz.Pixmap(page.parent, smasks[im['xref']]))
            elif pix.n - pix.alpha > 3: pix = fitz.Pixmap(fitz.csRGB, pix, pix.alpha)   # CMYK -> RGB
            pic = XLImage(BytesIO(pix.tobytes('png')))
            c, r = col_at(x0), row_at(y0)
            E = EMU * z
            pic.anchor = OneCellAnchor(_from=AnchorMarker(col=c, colOff=int((x0 - X[c]) * E), row=r, rowOff=int((y0 - Y[r]) * E)),
                                       ext=XDRPositiveSize2D(int((x1 - x0) * E), int((y1 - y0) * E)))
            ws.add_image(pic)
    except Exception as e:
        print('xlsx image skipped:', e, file=sys.stderr)

    ws.sheet_view.showGridLines = False   # the PDF's own lines are the grid
    # print like the PDF: same paper, orientation and margins, scaled to fill the same area on one page
    W, H = page.rect.width, page.rect.height
    ws.page_setup.orientation = 'landscape' if W > H else 'portrait'
    short, long_ = sorted((W, H))
    for size, (a, b) in {1: (612, 792), 5: (612, 1008), 9: (595, 842), 8: (842, 1191), 11: (420, 595), 3: (792, 1224)}.items():
        if abs(short - a) < 6 and abs(long_ - b) < 6: ws.page_setup.paperSize = size
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1
    pm = ws.page_margins
    pm.left, pm.top = round(X[0] / 72, 3), round(Y[0] / 72, 3)
    pm.right, pm.bottom = round(max(0.1, W - X[-1] - 1) / 72, 3), round(max(0.1, min(H - Y[-1], Y[0])) / 72, 3)
    pm.header = pm.footer = 0
    return True

def _print_scale(page, hs, vs):
    """How much Excel shrank the sheet when it printed this PDF (1 if it wasn't made in Excel).
    Excel prints its thin / medium / thick borders at 0.96 / 1.92 / 2.88pt, so the scale that turns
    those into the line widths we see is the one it used."""
    meta = page.parent.metadata or {}
    if 'excel' not in ((meta.get('creator') or '') + (meta.get('producer') or '')).lower(): return 1
    widths = [w for *_, w in hs + vs if 0.05 < w < 6]
    if len(widths) < 4: return 1
    def miss(s): return sum(min(abs(w - b * s) for b in (0.96, 1.92, 2.88)) / w for w in widths)
    best = min(range(100, 29, -1), key=lambda pct: miss(pct / 100))   # ties go to the larger scale
    return best / 100

def export_xlsx(doc, out):
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Border, Side
    thin = Side(style='thin', color='000000')
    grid = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb = Workbook(); wb.remove(wb.active)
    found_text = False
    for i, page in enumerate(doc):
        ws = wb.create_sheet(f'Page {i + 1}')
        hs, vs, _ = _ruling(page)
        if len(hs) >= 2 and len(vs) >= 2:   # a ruled page: copy its layout
            if _layout_sheet(page, ws): found_text = True; continue
        rows, boxes = [], []   # boxes: (first row, row count, column count) of each table, to draw its grid
        try: tables = page.find_tables().tables
        except Exception: tables = []
        for t in tables:
            cells = t.extract()
            boxes.append((len(rows) + 1, len(cells), max((len(r) for r in cells), default=0)))
            rows += cells; rows.append([])
        if not tables: rows = _text_rows(page)
        widths = {}
        for r in rows:
            vals = [_cell(c) for c in r]
            if any(v is not None for v in vals): found_text = True
            ws.append(vals)
            for j, v in enumerate(vals):
                if v is not None: widths[j] = max(widths.get(j, 0), min(len(str(v)), 60))
        for top, nrows, ncols in boxes:
            for r in range(top, top + nrows):
                for c in range(1, ncols + 1):
                    ws.cell(row=r, column=c).border = grid
        for j, w in widths.items():
            ws.column_dimensions[get_column_letter(j + 1)].width = max(8, w + 2)
    wb.save(out)
    return found_text

def export_pptx(doc, out):
    from pptx import Presentation
    from pptx.util import Emu, Inches
    prs = Presentation()
    r0 = doc[0].rect
    sw = Inches(10)
    sh = int(sw * r0.height / r0.width)
    sh = max(Inches(1), min(sh, Inches(56)))   # PowerPoint's slide size limits
    prs.slide_width, prs.slide_height = sw, sh
    for page in doc:
        slide = prs.slides.add_slide(prs.slide_layouts[6])   # blank layout
        img = BytesIO(page.get_pixmap(dpi=150).tobytes('png'))
        r = page.rect
        scale = min(sw / r.width, sh / r.height)
        w, h = int(r.width * scale), int(r.height * scale)
        slide.shapes.add_picture(img, (sw - w) // 2, (sh - h) // 2, Emu(w), Emu(h))
        text = page.get_text().strip()
        if text: slide.notes_slide.notes_text_frame.text = text
    prs.save(out)

def export_docx(src_pdf, out):
    try:
        from pdf2docx import Converter
    except ImportError:
        raise RuntimeError("Word export needs the pdf2docx package. Run setup-and-build.bat again.")
    cv = Converter(src_pdf)
    try: cv.convert(out)
    finally: cv.close()

@app.route('/export', methods=['POST'])
def export_pdf():
    d = request.json; fp = d['filepath']; orig = d.get('original_path') or fp; fmt = d['format']
    folder = os.path.dirname(orig); base = os.path.splitext(os.path.basename(orig))[0]
    try:
        src_ext = os.path.splitext(orig)[1].lower()
        if fmt == 'xlsx' and src_ext in ('.xlsx', '.xlsm', '.xls') and os.path.exists(orig):
            # The document was opened from a spreadsheet: hand back a copy of that workbook, which keeps
            # every real column width, row height, formula and format (rebuilding it from the PDF can't).
            out = get_unique_filename(os.path.join(folder, base + src_ext))
            shutil.copyfile(orig, out)
            return jsonify({"success": True, "path": out,
                            "message": f"Saved {os.path.basename(out)}, copied from the original Excel file"})
        if fmt == 'auto':   # each kind of page to its own format: tables to Excel, text to Word, slides to PowerPoint
            src = fitz.open(fp)
            parts = _page_parts(src)
            if len(parts) == 1: fmt = parts[0][0]; src.close()
            else:
                outs = []
                try:
                    for kind, pages in parts:
                        sub = fitz.open()
                        for i in pages: sub.insert_pdf(src, from_page=i, to_page=i)
                        out = get_unique_filename(os.path.join(folder, f'{base} ({KIND_NAMES[kind][0]}).{kind}'))
                        if kind == 'docx':
                            tmp = os.path.join(TEMP_DIR, f'part_{hashlib.md5(out.encode()).hexdigest()}.pdf')
                            sub.save(tmp); export_docx(tmp, out)
                        elif kind == 'xlsx': export_xlsx(sub, out)
                        else: export_pptx(sub, out)
                        sub.close(); outs.append(out)
                finally: src.close()
                return jsonify({"success": True, "path": outs[0], "paths": outs,
                                "message": "Saved " + ", ".join(os.path.basename(o) for o in outs)})
        if fmt in ('docx', 'xlsx', 'pptx'):
            out = get_unique_filename(os.path.join(folder, f'{base}.{fmt}'))
            note = ''
            if fmt == 'docx':
                export_docx(fp, out)
            else:
                doc = fitz.open(fp)
                try:
                    if fmt == 'xlsx':
                        if not export_xlsx(doc, out): note = ' (no text found; this PDF may be a scan)'
                    else:
                        export_pptx(doc, out)
                finally: doc.close()
            return jsonify({"success": True, "path": out, "message": f"Exported to {os.path.basename(out)}{note}"})

        doc = fitz.open(fp)
        if fmt in ('png', 'jpeg'):
            out = get_unique_filename(os.path.join(folder, base + '_images')); os.makedirs(out)
            for i, p in enumerate(doc):
                p.get_pixmap(dpi=150).save(os.path.join(out, f"{base}_{i+1}." + ('jpg' if fmt == 'jpeg' else 'png')))
        else:
            out = get_unique_filename(os.path.join(folder, base + ('.html' if fmt == 'html' else '.txt')))
            with open(out, 'w', encoding='utf-8') as f:
                f.write('\n'.join(p.get_text('html' if fmt == 'html' else 'text') for p in doc))
        doc.close()
        return jsonify({"success": True, "path": out, "message": f"Exported to {os.path.basename(out)}"})
    except PermissionError:
        return jsonify({"error": "Couldn't write the file. Close it if it's open in Office, then try again."}), 500
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/protect', methods=['POST'])
def protect_pdf():
    d = request.json; fp = d['filepath']; orig = d.get('original_path') or fp
    if not d.get('password'): return jsonify({"error": "Enter a password"}), 400
    try:
        doc = fitz.open(fp)
        out = get_unique_filename(os.path.join(os.path.dirname(orig), os.path.splitext(os.path.basename(orig))[0] + '_Protected.pdf'))
        doc.save(out, encryption=fitz.PDF_ENCRYPT_AES_256, user_pw=d['password'], owner_pw=d['password'])
        doc.close()
        return jsonify({"success": True, "message": f"Saved {os.path.basename(out)}"})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/header_footer', methods=['POST'])
def header_footer():
    d = request.json
    try:
        doc = fitz.open(d['filepath']); n = len(doc)
        for i, p in enumerate(doc):
            r = p.rect
            for text, box in ((d.get('header', ''), fitz.Rect(36, 16, r.width - 36, 46)), (d.get('footer', ''), fitz.Rect(36, r.height - 44, r.width - 36, r.height - 14))):
                if text: p.insert_textbox(box, text.replace('{page}', str(i + 1)).replace('{pages}', str(n)), fontsize=10, align=1)
        out = os.path.join(TEMP_DIR, f"hf_{hashlib.md5(os.urandom(8)).hexdigest()}.pdf")
        doc.save(out); doc.close()
        return jsonify({"success": True, "new_file": out})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/create_blank', methods=['POST'])
def create_blank():
    try:
        doc = fitz.open(); doc.new_page(width=595, height=842)
        out = get_unique_filename(os.path.join(TEMP_DIR, 'Blank page.pdf'))
        doc.save(out); doc.close()
        return jsonify({"success": True, "path": out})
    except Exception as e: return jsonify({"error": str(e)}), 500

KIND_NAMES = {'xlsx': ('tables', 'Excel'), 'docx': ('text', 'Word'), 'pptx': ('slides', 'PowerPoint')}

def _page_kind(page):
    """What a page mostly is: 'xlsx' (tables or a form), 'pptx' (a slide) or 'docx' (text, or a scan).
    A page counts as a table when most of its words sit in tables: ruled ones, or borderless ones whose
    columns line up and hold mostly numbers. A page of text with a small table stays text."""
    words = page.get_text('words')
    n = len(words)
    if n < 5: return 'docx'   # a scan, a picture or an empty page
    def share(rects):
        return sum(1 for w in words if any(r.contains(fitz.Point((w[0] + w[2]) / 2, (w[1] + w[3]) / 2)) for r in rects)) / n
    tables = 0
    try:
        ruled = [fitz.Rect(t.bbox) for t in page.find_tables().tables if t.row_count >= 2 and t.col_count >= 2]
        tables = share(ruled)
        if tables < 0.5:
            def numeric(t):
                cells = [c for row in t.extract() for c in row if c and c.strip()]
                return cells and sum(bool(re.fullmatch(r'[-+$€£(]?[\d.,]+%?\)?', c.strip())) for c in cells) / len(cells) >= 0.3
            loose = [fitz.Rect(t.bbox) for t in page.find_tables(strategy='text').tables
                     if t.row_count >= 4 and t.col_count >= 3 and numeric(t)]
            tables = max(tables, share(ruled + loose))
    except Exception:
        hs, vs, _ = _ruling(page)
        tables = 1 if len(hs) >= 4 and len(vs) >= 4 else 0
    if tables >= 0.5: return 'xlsx'
    r = page.rect
    sizes = sorted(s['size'] for b in page.get_text('dict')['blocks'] if b['type'] == 0 for l in b['lines'] for s in l['spans'] if s['text'].strip())
    big_type = sizes and sizes[len(sizes) // 2] >= 16
    if r.width > r.height * 1.25 and n < 150 and big_type: return 'pptx'
    return 'docx'

def _page_parts(doc):
    """Pages grouped by kind, in order of first appearance: [('docx', [0, 2]), ('xlsx', [1])]."""
    parts = {}
    for i, page in enumerate(doc): parts.setdefault(_page_kind(page), []).append(i)
    return list(parts.items())

def _pages_text(pages):
    """[0, 1, 2, 4] -> 'pages 1-3, 5'; [3] -> 'page 4'"""
    return ('page ' if len(pages) == 1 else 'pages ') + _page_label(pages)

def suggest_export(fp, orig):
    """Best export format for this document and why: the file it was opened from, then whether it mixes
    kinds of pages (tables and text, say: then each part gets its own format), then the program that made
    the PDF (it records that), then what the pages look like."""
    ext = os.path.splitext(orig or '')[1].lower()
    by_ext = {'.xlsx': 'xlsx', '.xlsm': 'xlsx', '.xls': 'xlsx', '.docx': 'docx', '.doc': 'docx', '.pptx': 'pptx', '.ppt': 'pptx'}
    names = {'xlsx': 'Excel', 'docx': 'Word', 'pptx': 'PowerPoint'}
    if ext in by_ext:
        return by_ext[ext], f"You opened this from an {names[by_ext[ext]]} file" if by_ext[ext] == 'xlsx' else f"You opened this from a {names[by_ext[ext]]} file"
    doc = fitz.open(fp)
    try:
        if len(doc) > 1:
            parts = _page_parts(doc)
            if len(parts) > 1:
                return 'auto', '; '.join(f"{KIND_NAMES[k][0].capitalize()} ({_pages_text(pg)}) to {KIND_NAMES[k][1]}" for k, pg in parts)
        made_by = f"{doc.metadata.get('creator') or ''} {doc.metadata.get('producer') or ''}".lower()
        for key, fmt, app_name in (('excel', 'xlsx', 'Excel'), ('calc', 'xlsx', 'a spreadsheet app'), ('google sheets', 'xlsx', 'Google Sheets'),
                                   ('powerpoint', 'pptx', 'PowerPoint'), ('impress', 'pptx', 'a slides app'), ('keynote', 'pptx', 'Keynote'), ('google slides', 'pptx', 'Google Slides'),
                                   ('word', 'docx', 'Word'), ('writer', 'docx', 'a word processor'), ('pages', 'docx', 'Pages'), ('google docs', 'docx', 'Google Docs')):
            if key in made_by:
                return fmt, f"This PDF was made in {app_name}"
        pages = list(doc)[:5]
        words = sum(len(p.get_text('words')) for p in pages)
        if words < 5 * len(pages) and any(p.get_images() for p in pages):
            return 'docx', "Looks like a scan; Word keeps the page images"
        kinds = [_page_kind(p) for p in pages]
        best = max(('xlsx', 'pptx', 'docx'), key=kinds.count)
        return best, {'xlsx': "Mostly tables", 'pptx': "Looks like slides", 'docx': "Mostly text"}[best]
    finally:
        doc.close()

@app.route('/suggest_export', methods=['POST'])
def suggest_export_route():
    d = request.json
    try:
        fmt, why = suggest_export(d['filepath'], d.get('original_path'))
        return jsonify({"format": fmt, "reason": why})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/search', methods=['POST'])
def search_pdf():
    d = request.json
    try:
        doc = fitz.open(d['filepath']); pages = []; count = 0; boxes = {}
        for i, p in enumerate(doc):
            hits = p.search_for(d['query'])
            if hits:
                pages.append(i + 1); count += len(hits)
                # where each match sits, in points from the top-left of the page as displayed (after rotation)
                boxes[i + 1] = [[round(v, 2) for v in (r * p.rotation_matrix)] for r in hits]
        doc.close()
        return jsonify({"pages": pages, "count": count, "boxes": boxes})
    except Exception as e: return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, threaded=True)