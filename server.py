import os
import sys
import subprocess
import hashlib
import tempfile
import pythoncom
import shutil
import re
import base64
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

def _com_to_pdf(src, dst):
    """Convert a Word/Excel/PowerPoint file to PDF using a private Office instance
    (DispatchEx) so we never touch documents the user already has open, and always Quit()."""
    low = src.lower()
    if low.endswith(('.docx', '.doc')):
        app_ = win32.DispatchEx('Word.Application')
        try:
            app_.DisplayAlerts = 0
            doc = app_.Documents.Open(src, ReadOnly=True)
            try: doc.ExportAsFixedFormat(dst, 17)
            finally: doc.Close(False)
        finally: app_.Quit()
    elif low.endswith(('.xlsx', '.xls')):
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
    elif low.endswith(('.pptx', '.ppt')):
        app_ = win32.DispatchEx('PowerPoint.Application')
        try:
            pres = app_.Presentations.Open(src, WithWindow=False, ReadOnly=True)
            try: pres.ExportAsFixedFormat(dst, 2)
            finally: pres.Close()
        finally: app_.Quit()
    elif low.endswith(('.jpg', '.jpeg', '.png')):
        Image.open(src).convert('RGB').save(dst)
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
    if src.lower().endswith(('.jpg', '.jpeg', '.png')):
        return _com_to_pdf(src, dst)
    try:
        _com_to_pdf(src, dst)
    except Exception as com_err:
        if not find_soffice():
            raise RuntimeError("No Office converter found. Install Microsoft Office or the free LibreOffice, then try again.")
        libreoffice_to_pdf(src, dst)

CONVERTER_VERSION = '2'   # bump when office_to_pdf output changes (2: cut-off Excel columns are widened)

@app.route('/ping', methods=['GET'])
def ping():
    return jsonify({"status": "connected"})

@app.route('/process_file', methods=['POST'])
def process_file():
    data = request.json
    path = data.get('filepath')
    if not path or not os.path.exists(path): return jsonify({"error": "File not found"}), 400

    ext = path.lower()
    if ext.endswith('.pdf'): return jsonify({"pdf_path": path})

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
    """'ABCDEF+Calibri-Bold', 'Calibri,Bold', 'Arial-BoldMT' and 'Calibri Bold' all become 'calibribold'."""
    name = re.sub(r'^[A-Z]{6}\+', '', name or '')
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
                if not str(file).lower().endswith(('.ttf', '.otf')): continue   # .ttc collections hold several faces
                label = re.sub(r'\s*\(.*\)$', '', label)
                for part in label.split(' & '):
                    _system_fonts.setdefault(_font_key(part), os.path.join(folder, file))
    return _system_fonts.get(_font_key(name))

def _text_font(doc, page, span_font, text, cache):
    """A font name usable in insert_text that draws `text` like `span_font`:
    the installed font of that name, else the copy embedded in the PDF if it has every letter,
    else the closest built-in font."""
    if (span_font, text) in cache: return cache[(span_font, text)]
    alias = None
    path = _find_system_font(span_font)
    if path and os.path.exists(path):
        alias = 'E' + hashlib.md5(path.encode()).hexdigest()[:8]
        page.insert_font(fontname=alias, fontfile=path)
    else:
        for xref, ext, _, base, *_ in page.get_fonts(full=True):
            if _font_key(base) != _font_key(span_font) or ext in ('n/a', ''): continue
            try:
                buf = doc.extract_font(xref)[3]
                if buf and all(fitz.Font(fontbuffer=buf).has_glyph(ord(ch)) for ch in text if not ch.isspace()):
                    alias = 'X%d' % xref
                    page.insert_font(fontname=alias, fontbuffer=buf)
                    break
            except Exception: pass
    if not alias:
        low = span_font.lower()
        b, it = 'bold' in low or 'black' in low or 'heavy' in low, 'italic' in low or 'oblique' in low
        fam = 'co' if 'cour' in low or 'mono' in low else 'ti' if 'times' in low or 'serif' in low and 'sans' not in low else 'he'
        alias = {('co', 0, 0): 'cour', ('co', 1, 0): 'cobo', ('co', 0, 1): 'coit', ('co', 1, 1): 'cobi',
                 ('ti', 0, 0): 'tiro', ('ti', 1, 0): 'tibo', ('ti', 0, 1): 'tiit', ('ti', 1, 1): 'tibi',
                 ('he', 0, 0): 'helv', ('he', 1, 0): 'hebo', ('he', 0, 1): 'heit', ('he', 1, 1): 'hebi'}[(fam, int(b), int(it))]
    cache[(span_font, text)] = alias
    return alias

def _rewrite_text(doc, page, edits):
    """Replace edited words with the same look the originals had. Each edit says where the original text
    sat (rect, its baseline y) and what it should say now; moved or restyled text also carries
    dx/dy and the new size, color, bold and italic."""
    spans = [s for b in page.get_text('dict')['blocks'] if b['type'] == 0 for l in b['lines'] for s in l['spans'] if s['text'].strip()]
    plans = []
    for e in edits:
        r = fitz.Rect(*map(float, e['rect']))
        by = float(e.get('base_y', r.y1))
        hit = [s for s in spans if abs(s['origin'][1] - by) < max(2, s['size'] * 0.3)
               and fitz.Rect(s['bbox']).x1 > r.x0 + 1 and fitz.Rect(s['bbox']).x0 < r.x1 - 1]
        if not hit: continue
        hit.sort(key=lambda s: s['origin'][0])
        plans.append((e, hit))
        for s in hit:   # remove just these glyphs: no white box, lines and pictures untouched
            page.add_redact_annot(fitz.Rect(s['bbox']), fill=False, cross_out=False)
    if not plans: return
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE)

    cache = {}
    for e, hit in plans:
        text = str(e.get('new_text', '')).replace('\n', ' ').strip()
        if not text: continue
        s = hit[0]
        size, font, flags = s['size'], s['font'], s['flags']
        c = s['color']; rgb = ((c >> 16 & 255) / 255, (c >> 8 & 255) / 255, (c & 255) / 255)
        if e.get('styled'):   # the user changed the look on purpose; anything they didn't touch stays
            if e.get('size_ratio'): size *= float(e['size_ratio'])
            if e.get('color'): rgb = parse_color(e['color'])
            low = font.lower()
            bold = e['bold'] if 'bold' in e else bool(flags & 16) or 'bold' in low
            italic = e['italic'] if 'italic' in e else bool(flags & 2) or 'italic' in low or 'oblique' in low
            fam = e.get('font_family') or re.sub(r'^[A-Z]{6}\+', '', font).split('-')[0].split(',')[0]
            style = ('Bold' if bold else '') + ('Italic' if italic else '')
            font = fam + ('-' + style if style else '')
        x, y = s['origin']
        x += float(e.get('dx', 0)); y += float(e.get('dy', 0))
        alias = _text_font(doc, page, font, text, cache)
        page.insert_text(fitz.Point(x, y), text, fontsize=size, fontname=alias, color=rgb)
        if e.get('underline'):
            w = fitz.get_text_length(text, fontname='helv', fontsize=size) if alias in ('helv',) else None
            try: w = fitz.Font(fontfile=_find_system_font(font)).text_length(text, size) if w is None else w
            except Exception: w = fitz.get_text_length(text, fontsize=size)
            page.draw_line((x, y + size * 0.12), (x + w, y + size * 0.12), color=rgb, width=max(0.5, size / 18))

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

def suggest_export(fp, orig):
    """Best export format for this document and why: the file it was opened from, then the program
    that made the PDF (it records that), then what the pages look like."""
    ext = os.path.splitext(orig or '')[1].lower()
    by_ext = {'.xlsx': 'xlsx', '.xlsm': 'xlsx', '.xls': 'xlsx', '.docx': 'docx', '.doc': 'docx', '.pptx': 'pptx', '.ppt': 'pptx'}
    names = {'xlsx': 'Excel', 'docx': 'Word', 'pptx': 'PowerPoint'}
    if ext in by_ext:
        return by_ext[ext], f"You opened this from an {names[by_ext[ext]]} file" if by_ext[ext] == 'xlsx' else f"You opened this from a {names[by_ext[ext]]} file"
    doc = fitz.open(fp)
    try:
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
        ruled = sum(1 for p in pages if (lambda hv: len(hv[0]) >= 4 and len(hv[1]) >= 4)(_ruling(p)))
        if ruled * 2 >= len(pages):
            return 'xlsx', "Mostly tables"
        r = pages[0].rect
        if r.width > r.height * 1.25 and words / len(pages) < 150:
            return 'pptx', "Looks like slides"
        return 'docx', "Mostly text"
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