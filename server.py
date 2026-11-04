import os
import sys

def _recognize_ink(strokes):
    """Windows' handwriting recognizer: the words it reads, each with its other guesses."""
    import asyncio
    async def run():
        from winrt.windows.ui.input.inking import InkStrokeBuilder, InkStrokeContainer, InkRecognizerContainer, InkRecognitionTarget, InkPoint
        from winrt.windows.foundation import Point
        from winrt.windows.foundation.numerics import Matrix3x2
        b, cont = InkStrokeBuilder(), InkStrokeContainer()
        for pts in strokes:
            if len(pts) < 2: pts = pts + [[pts[0][0] + 0.3, pts[0][1] + 0.3]]
            cont.add_stroke(b.create_stroke_from_ink_points([InkPoint(Point(float(x) * 4, float(y) * 4), 0.5) for x, y in pts], Matrix3x2(1, 0, 0, 1, 0, 0)))
        rc = InkRecognizerContainer()   # held on to: if it's let go mid-way, the answer never comes
        res = await rc.recognize_async(cont, InkRecognitionTarget.ALL)
        return [list(r.get_text_candidates()) for r in res]
    # on a fresh thread: a thread where COM was already set up (pywin32 does that) never gets the answer
    import threading
    out = {}
    def work():
        try: out['words'] = asyncio.run(run())
        except Exception as e: out['error'] = e
    t = threading.Thread(target=work, daemon=True); t.start(); t.join(30)
    if 'error' in out: raise out['error']
    if 'words' not in out: raise TimeoutError('no answer from the recognizer')
    return out['words']
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
# design, CAD and photo files (architects, fashion and product designers, photographers)
CAD_EXT = ('.dxf', '.dwg', '.dwf', '.dwfx', '.plt', '.hpgl', '.hpg', '.hgl', '.prn')
PICTURE_EXT = ('.psd', '.psb', '.heic', '.heif', '.avif', '.tga', '.dds', '.pcx', '.icns', '.jp2', '.j2k', '.emf', '.wmf', '.ppm', '.pgm', '.pbm', '.xbm', '.svgz')
VECTOR_EXT = ('.ai', '.eps', '.ps', '.epsf')
MODEL_EXT = ('.stl', '.obj', '.ply', '.off')
# files whose own program isn't on this PC: the preview picture (or preview PDF) saved inside them
PREVIEW_EXT = ('.skp', '.rvt', '.rfa', '.rte', '.3dm', '.max', '.blend', '.indd', '.idml', '.cdr', '.afdesign', '.afphoto', '.afpub', '.sketch',
               '.fig', '.xd', '.pages', '.numbers', '.key', '.vsd', '.vsdx', '.pub', '.ifc', '.dgn', '.step', '.stp', '.iges', '.igs', '.sldprt',
               '.sldasm', '.slddrw', '.ipt', '.iam', '.f3d', '.c4d', '.fbx', '.dae', '.3ds', '.cr2', '.cr3', '.nef', '.arw', '.dng', '.raf', '.orf',
               '.rw2', '.pef', '.srw', '.clo', '.zprj', '.avt', '.pds', '.mdl', '.aama', '.rul',
               '.zpac', '.zfab', '.zsel', '.bw', '.vsp', '.u3m', '.iba', '.vet', '.mrk', '.tuk',
               '.pln', '.pla', '.bimx', '.vwx', '.nwd', '.nwc', '.nwf', '.lcf', '.tcp', '.dwfx3d')
MAIL_EXT = ('.eml',)
DESIGN_EXT = CAD_EXT + PICTURE_EXT + VECTOR_EXT + MODEL_EXT + PREVIEW_EXT + MAIL_EXT
OPENABLE_EXT = ('.pdf',) + WORD_EXT + EXCEL_EXT + POWERPOINT_EXT + IMAGE_EXT + BOOK_EXT + WEB_EXT + TEXT_EXT + DESIGN_EXT

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
    local = os.environ.get('LOCALAPPDATA', '')
    for p in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
              r"C:\Program Files\Google\Chrome\Application\chrome.exe", r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
              os.path.join(local, r"Microsoft\Edge\Application\msedge.exe"), os.path.join(local, r"Google\Chrome\Application\chrome.exe")):
        if os.path.exists(p): return p
    try:   # wherever Windows says Edge or Chrome is
        import winreg
        for name in ('msedge.exe', 'chrome.exe'):
            for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                try:
                    with winreg.OpenKey(root, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\" + name) as k:
                        v = winreg.QueryValue(k, None)
                        if v and os.path.exists(v.strip('"')): return v.strip('"')
                except OSError: continue
    except Exception: pass
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

# ---------- design, CAD and photo files ----------
KIND_NAME = {'.skp': 'SketchUp', '.rvt': 'Revit', '.rfa': 'Revit', '.rte': 'Revit', '.3dm': 'Rhino', '.max': '3ds Max', '.blend': 'Blender',
             '.indd': 'InDesign', '.idml': 'InDesign', '.cdr': 'CorelDRAW', '.afdesign': 'Affinity Designer', '.afphoto': 'Affinity Photo',
             '.afpub': 'Affinity Publisher', '.sketch': 'Sketch', '.fig': 'Figma', '.xd': 'Adobe XD', '.pages': 'Pages', '.numbers': 'Numbers',
             '.key': 'Keynote', '.vsd': 'Visio', '.vsdx': 'Visio', '.pub': 'Publisher', '.ifc': 'a BIM program', '.dgn': 'MicroStation',
             '.step': 'a CAD program', '.stp': 'a CAD program', '.iges': 'a CAD program', '.igs': 'a CAD program', '.sldprt': 'SolidWorks',
             '.sldasm': 'SolidWorks', '.slddrw': 'SolidWorks', '.ipt': 'Inventor', '.iam': 'Inventor', '.f3d': 'Fusion 360', '.c4d': 'Cinema 4D',
             '.fbx': 'a 3D program', '.dae': 'a 3D program', '.3ds': 'a 3D program', '.dwg': 'AutoCAD', '.dwf': 'Design Review',
             '.clo': 'CLO 3D', '.zprj': 'CLO 3D', '.avt': 'CLO 3D', '.pds': 'Optitex', '.mdl': 'a pattern program', '.aama': 'a pattern program', '.rul': 'a pattern program',
             '.zpac': 'CLO 3D', '.zfab': 'CLO 3D', '.zsel': 'CLO 3D', '.bw': 'Browzwear VStitcher', '.vsp': 'Browzwear VStitcher', '.u3m': 'a 3D fabric program',
             '.iba': 'Lectra Modaris', '.vet': 'Lectra Modaris', '.mrk': 'Optitex', '.tuk': 'Tukatech',
             '.pln': 'Archicad', '.pla': 'Archicad', '.bimx': 'Archicad BIMx', '.vwx': 'Vectorworks', '.nwd': 'Navisworks', '.nwc': 'Navisworks',
             '.nwf': 'Navisworks', '.lcf': 'Vectorworks', '.tcp': 'TurboCAD', '.dwfx3d': 'Autodesk Design Review'}

def _program_path(*globs):
    import glob
    for g in globs:
        hits = sorted(glob.glob(g), reverse=True)
        if hits: return hits[0]
    return None

def _dxf_to_pdf(src, dst):
    """A DXF drawing (AutoCAD, pattern makers' AAMA/ASTM files too): the model and every sheet layout, each
    on a page fitted to the drawing, on white, in the drawing's own colors and line weights."""
    from ezdxf import recover
    from ezdxf.addons.drawing import RenderContext, Frontend, config, layout
    from ezdxf.addons.drawing.pymupdf import PyMuPdfBackend
    doc, auditor = recover.readfile(src)
    cfg = config.Configuration(background_policy=config.BackgroundPolicy.WHITE, lineweight_scaling=1.0)
    out = fitz.open()
    def add(lay, fit):
        backend = PyMuPdfBackend()
        Frontend(RenderContext(doc), backend, config=cfg).draw_layout(lay, finalize=True)
        if fit:   # the model is drawn life-size (meters of wall): fit it on an A3 sheet, turned to match its shape
            from ezdxf import bbox
            ext = bbox.extents(lay, fast=True)
            wide = not ext.has_data or ext.size.x >= ext.size.y
            page = layout.Page(420 if wide else 297, 297 if wide else 420, layout.Units.mm, margins=layout.Margins.all(10))
        else:     # a sheet layout already has its paper size
            page = layout.Page(0, 0, layout.Units.mm, margins=layout.Margins.all(5))
        out.insert_pdf(fitz.open('pdf', backend.get_pdf_bytes(page)))
    add(doc.modelspace(), True)
    for name in doc.layout_names_in_taborder():
        if name.lower() == 'model': continue
        lay = doc.paperspace(name)
        if sum(1 for e in lay if e.dxftype() != 'VIEWPORT') == 0: continue
        try: add(lay, False)
        except Exception as e: print('layout skipped:', name, e, file=sys.stderr)
    if not len(out): raise RuntimeError("This drawing is empty")
    out.save(dst, garbage=3, deflate=True)

def _dwg_to_pdf(src, dst):
    """DWG: the free ODA File Converter turns it into DXF; without it, the preview picture inside the file."""
    oda = _program_path(r'C:\Program Files\ODA\ODAFileConverter*\ODAFileConverter.exe', r'C:\Program Files (x86)\ODA\ODAFileConverter*\ODAFileConverter.exe')
    if oda:
        tmp_in, tmp_out = tempfile.mkdtemp(dir=TEMP_DIR, prefix='dwg_'), tempfile.mkdtemp(dir=TEMP_DIR, prefix='dxf_')
        try:
            shutil.copy(src, tmp_in)
            subprocess.run([oda, tmp_in, tmp_out, 'ACAD2018', 'DXF', '0', '1'], timeout=300, capture_output=True, creationflags=0x08000000)
            dxf = [os.path.join(tmp_out, f) for f in os.listdir(tmp_out) if f.lower().endswith('.dxf')]
            if dxf: return _dxf_to_pdf(dxf[0], dst)
        finally:
            shutil.rmtree(tmp_in, ignore_errors=True); shutil.rmtree(tmp_out, ignore_errors=True)
    try:
        return _preview_to_pdf(src, dst, note="To open the full drawing, install the free ODA File Converter (opendesign.com/guestfiles/oda_file_converter), then open this file again.")
    except Exception:
        raise RuntimeError("DWG drawings open with the free ODA File Converter. Install it from opendesign.com, then open this file again.")

def _hpgl_to_pdf(src, dst):
    """HP-GL plotter files (.plt), as cutters and pattern systems write them: pen moves become lines."""
    raw = open(src, 'rb').read().decode('latin-1')
    raw = re.sub(r'\x1b[^A-Za-z]*[A-Za-z]', '', raw)   # printer escape codes
    paths, cur, pen_down, pos, absolute, color = [], [], False, (0.0, 0.0), True, 1
    colors = {1: (0, 0, 0), 2: (0.85, 0, 0), 3: (0, 0.55, 0), 4: (0.9, 0.7, 0), 5: (0, 0, 0.8), 6: (0.7, 0, 0.7), 7: (0, 0.6, 0.7), 8: (0.4, 0.4, 0.4)}
    def flush():
        nonlocal cur
        if len(cur) > 1: paths.append((color, cur))
        cur = []
    for cmd in re.split(r'[;\n\r]+', raw):
        cmd = cmd.strip()
        if len(cmd) < 2: continue
        op, args = cmd[:2].upper(), cmd[2:]
        nums = [float(n) for n in re.findall(r'-?\d+(?:\.\d+)?', args)]
        pts = list(zip(nums[0::2], nums[1::2]))
        if op == 'SP': flush(); color = int(nums[0]) if nums else 1
        elif op == 'PA': absolute = True
        elif op == 'PR': absolute = False
        elif op in ('PU', 'PD'):
            if op == 'PU': flush()
            else:
                if not cur: cur = [pos]
            pen_down = op == 'PD'
            for x, y in pts:
                pos = (x, y) if absolute else (pos[0] + x, pos[1] + y)
                if pen_down: cur.append(pos)
            if op == 'PU' and pts: cur = []
        elif op in ('PA', 'PR') or op == 'IN': pass
    flush()
    if not paths: raise RuntimeError("No drawing found in this plotter file")
    xs = [x for _, p in paths for x, _ in p]; ys = [y for _, p in paths for _, y in p]
    x0, y0, x1, y1 = min(xs), min(ys), max(xs), max(ys)
    k = 72 / 1016   # plotter units are 1/40 mm
    w, h = (x1 - x0) * k + 36, (y1 - y0) * k + 36
    doc = fitz.open(); page = doc.new_page(width=max(w, 72), height=max(h, 72))
    for c, p in paths:
        sh = page.new_shape()
        sh.draw_polyline([(18 + (x - x0) * k, h - 18 - (y - y0) * k) for x, y in p])
        sh.finish(color=colors.get(c, (0, 0, 0)), width=0.6, closePath=False); sh.commit()
    doc.save(dst, garbage=3, deflate=True)

def _picture_to_pdf(src, dst):
    """Photoshop, iPhone (HEIC), TGA, EMF/WMF and the like: their picture (Photoshop's merged layers) as a page."""
    low = src.lower()
    if low.endswith(('.heic', '.heif', '.avif')):
        import pillow_heif; pillow_heif.register_heif_opener()
    if low.endswith('.svgz'):
        import gzip
        svg = os.path.join(TEMP_DIR, 'svgz_' + hashlib.md5(src.encode()).hexdigest() + '.svg')
        open(svg, 'wb').write(gzip.open(src).read())
        d = fitz.open(svg); fitz.open('pdf', d.convert_to_pdf()).save(dst, garbage=3, deflate=True); return
    with Image.open(src) as im:
        if low.endswith(('.wmf', '.emf')):
            try: im.load(dpi=200)
            except TypeError: im.load()
        dpi = (im.info.get('dpi') or (72, 72))[0] or 72
        rgb = im.convert('RGBA') if im.mode in ('RGBA', 'LA', 'P') else im.convert('RGB')
        if rgb.mode == 'RGBA':   # see-through parts on white, like the designer sees them
            bg = Image.new('RGB', rgb.size, 'white'); bg.paste(rgb, mask=rgb.split()[3]); rgb = bg
        buf = BytesIO(); rgb.save(buf, 'PNG')
        w, h = rgb.size[0] * 72 / dpi, rgb.size[1] * 72 / dpi
    doc = fitz.open(); pg = doc.new_page(width=w, height=h); pg.insert_image(pg.rect, stream=buf.getvalue())
    doc.save(dst, garbage=3, deflate=True)

def _vector_to_pdf(src, dst):
    """Illustrator (.ai is a PDF inside when saved with PDF compatibility, the default), EPS and PostScript
    (through Ghostscript when it's installed, otherwise the preview inside the EPS)."""
    if src.lower().endswith('.ai'):
        try:
            d = fitz.open(src, filetype='pdf')
            if d.page_count: d.save(dst, garbage=3, deflate=True); return
        except Exception: pass
    gs = _program_path(r'C:\Program Files\gs\gs*\bin\gswin64c.exe', r'C:\Program Files (x86)\gs\gs*\bin\gswin32c.exe')
    if gs:
        r = subprocess.run([gs, '-dSAFER', '-dBATCH', '-dNOPAUSE', '-dEPSCrop', '-sDEVICE=pdfwrite', f'-sOutputFile={dst}', src],
                           timeout=300, capture_output=True, creationflags=0x08000000)
        if os.path.exists(dst) and os.path.getsize(dst) > 0: return
    try:
        return _preview_to_pdf(src, dst, note="To open the full artwork, install the free Ghostscript (ghostscript.com), then open this file again.")
    except Exception:
        raise RuntimeError("This file needs the free Ghostscript to open (ghostscript.com). Install it, then open this file again.")

def _load_mesh(src):
    """Triangles of a 3D model: STL (text or binary), OBJ, PLY (text), OFF."""
    import numpy as np, struct
    low = src.lower(); raw = open(src, 'rb').read()
    if low.endswith('.stl'):
        if raw[:5].lower() == b'solid' and b'facet' in raw[:2000]:
            v = [list(map(float, m)) for m in re.findall(rb'vertex\s+(\S+)\s+(\S+)\s+(\S+)', raw)]
            return np.array(v, dtype=float).reshape(-1, 3, 3)
        n = struct.unpack('<I', raw[80:84])[0]
        a = np.frombuffer(raw[84:84 + n * 50], dtype=np.dtype([('n', '<f4', 3), ('v', '<f4', (3, 3)), ('a', '<u2')]))
        return a['v'].astype(float)
    text = raw.decode('latin-1').splitlines()
    verts, faces = [], []
    if low.endswith('.obj'):
        for line in text:
            if line.startswith('v '): verts.append([float(x) for x in line.split()[1:4]])
            elif line.startswith('f '):
                idx = [int(t.split('/')[0]) for t in line.split()[1:]]
                idx = [i - 1 if i > 0 else len(verts) + i for i in idx]
                faces += [[idx[0], idx[i], idx[i + 1]] for i in range(1, len(idx) - 1)]
    elif low.endswith('.off'):
        rows = [l for l in text if l.strip() and not l.startswith('#')]
        nv, nf = map(int, rows[1].split()[:2])
        verts = [list(map(float, r.split()[:3])) for r in rows[2:2 + nv]]
        for r in rows[2 + nv:2 + nv + nf]:
            idx = list(map(int, r.split()))[1:]; faces += [[idx[0], idx[i], idx[i + 1]] for i in range(1, len(idx) - 1)]
    else:   # ply, text form
        head = text.index('end_header'); nv = nf = 0
        for l in text[:head]:
            if l.startswith('element vertex'): nv = int(l.split()[2])
            if l.startswith('element face'): nf = int(l.split()[2])
        verts = [list(map(float, l.split()[:3])) for l in text[head + 1:head + 1 + nv]]
        for l in text[head + 1 + nv:head + 1 + nv + nf]:
            idx = list(map(int, l.split()))[1:]; faces += [[idx[0], idx[i], idx[i + 1]] for i in range(1, len(idx) - 1)]
    v = np.array(verts, dtype=float)
    return v[np.array(faces, dtype=int)]

def _model_to_pdf(src, dst):
    """A 3D model drawn from four sides (3D view, front, top, side), shaded, with its size."""
    import numpy as np
    tri = _load_mesh(src)
    if not len(tri): raise RuntimeError("No 3D shapes found in this file")
    if len(tri) > 60000: tri = tri[np.random.default_rng(0).choice(len(tri), 60000, replace=False)]
    lo, hi = tri.reshape(-1, 3).min(0), tri.reshape(-1, 3).max(0); size = hi - lo
    tri = tri - (lo + hi) / 2
    def view(R):
        t = tri @ R.T
        n = np.cross(t[:, 1] - t[:, 0], t[:, 2] - t[:, 0]); ln = np.linalg.norm(n, axis=1) + 1e-12
        shade = np.abs(n[:, 2]) / ln
        order = np.argsort(t[:, :, 2].mean(1))   # far ones first
        return t, shade, order
    def rot(ax, ay):
        cx, sx, cy, sy = np.cos(ax), np.sin(ax), np.cos(ay), np.sin(ay)
        return np.array([[cy, 0, sy], [0, 1, 0], [-sy, 0, cy]]) @ np.array([[1, 0, 0], [0, cx, -sx], [0, sx, cx]])
    views = [('3D view', rot(np.radians(-60), np.radians(35))), ('Front', rot(np.radians(-90), 0)), ('Top', np.eye(3)), ('Side', rot(np.radians(-90), np.radians(90)))]
    doc = fitz.open(); page = doc.new_page(width=792, height=612)
    cells = [fitz.Rect(36, 50, 396, 320), fitz.Rect(396, 50, 756, 320), fitz.Rect(36, 330, 396, 590), fitz.Rect(396, 330, 756, 590)]
    for (name, R), cell in zip(views, cells):
        t, shade, order = view(R)
        xs, ys = t[:, :, 0], t[:, :, 1]
        span = max(xs.max() - xs.min(), ys.max() - ys.min()) or 1
        k = min(cell.width - 30, cell.height - 40) / span
        cx, cy = cell.x0 + cell.width / 2, cell.y0 + 20 + (cell.height - 20) / 2
        sh = page.new_shape()
        for i in order:
            pts = [(cx + t[i, j, 0] * k, cy - t[i, j, 1] * k) for j in range(3)]
            g = 0.35 + 0.6 * shade[i]
            sh.draw_polyline(pts + [pts[0]]); sh.finish(fill=(g * 0.75, g * 0.82, g), color=None, width=0)
        sh.commit()
        page.insert_text((cell.x0 + 6, cell.y0 + 14), name, fontsize=10, color=(0.3, 0.3, 0.3))
    page.insert_text((36, 32), f"{os.path.basename(src)}  ·  {size[0]:.4g} × {size[1]:.4g} × {size[2]:.4g} (file units)  ·  {len(tri):,} triangles", fontsize=11)
    doc.save(dst, garbage=3, deflate=True)

def _preview_to_pdf(src, dst, note=None):
    """A file whose program isn't on this PC: the preview the program saved inside it. Apple's Pages,
    Numbers and Keynote keep a whole preview PDF; others (SketchUp, Revit, InDesign, CorelDRAW, Figma, camera
    RAW photos...) keep a picture."""
    import zipfile
    ext = os.path.splitext(src)[1].lower()
    best, pdf = None, None
    if zipfile.is_zipfile(src):
        with zipfile.ZipFile(src) as z:
            names = z.namelist()
            for n in names:
                if n.lower().endswith('preview.pdf') or (n.lower().endswith('.pdf') and 'quicklook' in n.lower()): pdf = z.read(n); break
            if not pdf:
                pics = [n for n in names if n.lower().endswith(('.png', '.jpg', '.jpeg')) and re.search(r'thumb|preview|quicklook|docprops', n, re.I)]
                pics = pics or [n for n in names if n.lower().endswith(('.png', '.jpg', '.jpeg'))]
                if pics: best = max((z.read(n) for n in pics), key=len)
    if pdf:
        fitz.open('pdf', pdf).save(dst, garbage=3, deflate=True); return
    if best is None:   # look through the file for the largest picture stored in it
        raw = open(src, 'rb').read()
        found = []
        for m in re.finditer(rb'\x89PNG\r\n\x1a\n', raw):
            end = raw.find(b'IEND', m.start())
            if end > 0: found.append(raw[m.start():end + 8])
        for m in re.finditer(rb'\xff\xd8\xff[\xe0-\xef\xdb]', raw):
            end = raw.find(b'\xff\xd9', m.start() + 4)
            while end > 0 and end - m.start() < 200:   # a tiny segment: keep looking for the real end
                end = raw.find(b'\xff\xd9', end + 2)
            if end > 0: found.append(raw[m.start():end + 2])
        good = []
        for f in found:
            try:
                with Image.open(BytesIO(f)) as im: im.load(); good.append((im.size[0] * im.size[1], f))
            except Exception: continue
        if not good: raise RuntimeError(f"{os.path.basename(src)} has no preview inside, and {KIND_NAME.get(ext, 'its program')} isn't on this PC")
        best = max(good)[1]
    with Image.open(BytesIO(best)) as im:
        w, h = im.size; buf = BytesIO(); im.convert('RGB').save(buf, 'PNG')
    k = min(1.0, 1400 / max(w, h)) * 0.75
    pw, ph = max(w * k, 300), h * k + 30
    doc = fitz.open(); page = doc.new_page(width=pw, height=ph)
    page.insert_image(fitz.Rect((pw - w * k) / 2, 0, (pw + w * k) / 2, h * k), stream=buf.getvalue())
    prog = KIND_NAME.get(ext)
    caption = note or (f"Preview saved inside this file. Open it in {prog} for the full version." if prog else "Preview picture saved inside this file.")
    page.insert_textbox(fitz.Rect(10, h * k + 6, pw - 10, ph - 2), caption, fontsize=7.5, color=(0.45, 0.45, 0.45), align=1)
    doc.save(dst, garbage=3, deflate=True)

def _mail_to_pdf(src, dst):
    """An email (.eml): who it's from and to, the subject and date, its message, and its attachments listed."""
    import email, html as _h
    from email import policy
    msg = email.message_from_bytes(open(src, 'rb').read(), policy=policy.default)
    body = msg.get_body(preferencelist=('html', 'plain'))
    content = body.get_content() if body else ''
    if body is not None and body.get_content_type() == 'text/plain': content = '<pre style="white-space:pre-wrap;font-family:inherit">' + _h.escape(content) + '</pre>'
    files = [p.get_filename() for p in msg.iter_attachments() if p.get_filename()]
    head = ''.join(f'<tr><td style="color:#666;padding-right:12px">{k}</td><td>{_h.escape(str(msg.get(k) or ""))}</td></tr>' for k in ('From', 'To', 'Cc', 'Date') if msg.get(k))
    page = (f'<html><head><meta charset="utf-8"></head><body style="font-family:Segoe UI,Arial;font-size:11pt">'
            f'<h2 style="margin:0 0 8px">{_h.escape(str(msg.get("Subject") or "(no subject)"))}</h2><table style="font-size:10pt;margin-bottom:10px">{head}</table>'
            + (f'<p style="font-size:10pt;color:#666">Attachments: {_h.escape(", ".join(files))}</p>' if files else '') + '<hr>' + content + '</body></html>')
    tmp = os.path.join(TEMP_DIR, 'mail_' + hashlib.md5(src.encode()).hexdigest() + '.html')
    open(tmp, 'w', encoding='utf-8').write(page)
    _web_to_pdf(tmp, dst)

def _design_to_pdf(src, dst):
    low = src.lower()
    try:
        if low.endswith('.dxf'): return _dxf_to_pdf(src, dst)
        if low.endswith('.dwg'): return _dwg_to_pdf(src, dst)
        if low.endswith('.dwfx'):
            d = fitz.open(src, filetype='xps'); fitz.open('pdf', d.convert_to_pdf()).save(dst, garbage=3, deflate=True); return
        if low.endswith(('.plt', '.hpgl', '.hpg', '.hgl', '.prn')): return _hpgl_to_pdf(src, dst)
        if low.endswith(PICTURE_EXT): return _picture_to_pdf(src, dst)
        if low.endswith(VECTOR_EXT): return _vector_to_pdf(src, dst)
        if low.endswith(MODEL_EXT): return _model_to_pdf(src, dst)
        if low.endswith(MAIL_EXT): return _mail_to_pdf(src, dst)
    except Exception as e:
        if low.endswith(('.dwg', '.ai', '.eps', '.ps', '.epsf')): raise
        try: return _preview_to_pdf(src, dst)   # last try: a preview inside the file
        except Exception: raise RuntimeError(f"Couldn't open {os.path.basename(src)}: {e}")
    return _preview_to_pdf(src, dst)

# ---------- Office, kept quick ----------
# Starting an Office app takes about a second and shutting it down a few more, and the old way did both for
# every file while the user waited. Now hidden Word and Excel are started with the app, PowerPoint on its first
# file, and each is kept for the next file (closed after 15 idle minutes, or when the app closes). Everything Office runs on one worker thread (COM likes that).
import queue as _queue, threading as _threading, time as _time

class _OfficeWorker:
    IDLE_QUIT = 900   # seconds of no use before the hidden Office apps are closed

    # name, how to count what's open in it (never close one that has someone's own files in it)
    APPS = {'word': ('Word.Application', 'Documents'), 'excel': ('Excel.Application', 'Workbooks'),
            'ppt': ('PowerPoint.Application', 'Presentations')}

    def __init__(self):
        self.jobs = _queue.Queue(); self.apps = {}; self.last = 0
        _threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        pythoncom.CoInitialize()
        while True:
            try: fn, box = self.jobs.get(timeout=20)
            except _queue.Empty:
                if self.apps and _time.time() - self.last > self.IDLE_QUIT: self._quit_all()
                continue
            try: box['result'] = fn(self)
            except Exception as e: box['error'] = e
            finally:
                self.last = _time.time()
                if 'done' in box: box['done'].set()

    def get(self, kind):
        a = self.apps.get(kind)
        if a is not None:
            try: getattr(a, self.APPS[kind][1]).Count   # still running?
            except Exception: a = None
        if a is None:
            a = win32.DispatchEx(self.APPS[kind][0])
            if kind == 'word': a.Visible = False; a.DisplayAlerts = 0
            elif kind == 'excel': a.Visible = False; a.DisplayAlerts = False; a.ScreenUpdating = False
            self.apps[kind] = a
        return a

    def get_word(self): return self.get('word')

    def forget(self, kind): self.apps.pop(kind, None)

    def _quit(self, kind):
        a = self.apps.pop(kind, None)
        if a is None: return
        try:
            if getattr(a, self.APPS[kind][1]).Count == 0: a.Quit()
        except Exception: pass

    def _quit_all(self):
        for k in list(self.apps): self._quit(k)

    def run(self, fn, timeout=150):
        box = {'done': _threading.Event()}
        self.jobs.put((fn, box))
        if not box['done'].wait(timeout):
            # Office is stuck (often a hidden first-run or sign-in message): leave it and start over
            # with a fresh worker, so the next file isn't stuck behind it
            global OFFICE
            if OFFICE is self: OFFICE = _OfficeWorker()
            raise TimeoutError("Office didn't answer. Open Word or Excel once, close any message it shows, then try again")
        if 'error' in box: raise box['error']
        return box['result']

    def later(self, fn):
        self.jobs.put((fn, {}))

    def shutdown(self):
        self.run(lambda o: o._quit_all(), timeout=20)

OFFICE = _OfficeWorker()

def _prewarm_office():
    """Have Word, Excel and PowerPoint ready before the first file comes in (the ones installed)."""
    import winreg
    _time.sleep(3)
    if NO_OFFICE_TEST: return
    for kind, (prog, _) in _OfficeWorker.APPS.items():
        try: winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, prog))
        except OSError: continue
        if kind == 'ppt': continue   # PowerPoint is one shared copy per user: started only when a slide file comes in
        OFFICE.later(lambda o, k=kind: o.get(k))

def _com_to_pdf(src, dst):
    """Convert a Word/Excel/PowerPoint file to PDF using a private Office instance
    (DispatchEx) so we never touch documents the user already has open.
    Pictures, e-books, web pages and text files have their own converters."""
    low = src.lower()
    if low.endswith(IMAGE_EXT): return _images_to_pdf(src, dst)
    if low.endswith(DESIGN_EXT): return _design_to_pdf(src, dst)
    if low.endswith(BOOK_EXT):
        d = fitz.open(src); fitz.open('pdf', d.convert_to_pdf()).save(dst, garbage=3, deflate=True); return
    if low.endswith(WEB_EXT): return _web_to_pdf(src, dst)
    if low.endswith(TEXT_EXT): return _text_to_pdf(src, dst)
    def kept(o, kind, open_):
        """Open the file in the kept app; if that app went away, start a fresh one once."""
        try: return open_(o.get(kind))
        except Exception:
            o.forget(kind); return open_(o.get(kind))
    if low.endswith(WORD_EXT):
        def word_job(o):
            doc = kept(o, 'word', lambda w: w.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False, Visible=False))
            try: doc.ExportAsFixedFormat(dst, 17)
            finally: doc.Close(False)
        OFFICE.run(word_job)
    elif low.endswith(EXCEL_EXT):
        def excel_job(o):
            wb = kept(o, 'excel', lambda x: x.Workbooks.Open(src, ReadOnly=True, AddToMru=False))
            try:
                try: _widen_cut_off_columns(wb, src)
                except Exception as e: print('column fit skipped:', e, file=sys.stderr)
                wb.ExportAsFixedFormat(0, dst)
            finally: wb.Close(False)   # never saved: the user's file stays as it was
        OFFICE.run(excel_job)
    elif low.endswith(POWERPOINT_EXT):
        def ppt_job(o):
            # PowerPoint takes Office's -1 / 0 for yes / no (a Python True can't be passed to it)
            pres = kept(o, 'ppt', lambda a: a.Presentations.Open(src, -1, 0, 0))   # read-only, keep its name, no window
            try: pres.SaveAs(dst, 32)                                              # 32 = PDF
            finally: pres.Close()
        OFFICE.run(ppt_job)
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

# ---------- Word, Excel and PowerPoint without Office ----------
# A new PC may have neither Microsoft Office nor LibreOffice. Rather than refuse, these files are laid out as a
# web page (fonts, colors, tables, pictures, page size) and Edge, which every Windows 10/11 has, prints it to
# PDF. Close to the original, not exact: Office (or LibreOffice) is still used whenever it's installed.
import html as _html

BUILTIN_EXT = ('.docx', '.docm', '.dotx', '.dotm', '.xlsx', '.xlsm', '.xltx', '.xltm', '.xls', '.csv', '.tsv',
               '.pptx', '.pptm', '.ppsx', '.ppsm', '.potx', '.potm')

def _esc(t): return _html.escape(str(t), quote=True)

def _img_uri(blob, ctype=''):
    """A picture as a data: link the browser can show (EMF/WMF and other kinds are turned into PNG)."""
    if ctype in ('image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/webp', 'image/svg+xml'):
        return f'data:{ctype};base64,' + base64.b64encode(blob).decode()
    try:
        with Image.open(BytesIO(blob)) as im:
            buf = BytesIO(); im.convert('RGBA').save(buf, 'PNG')
        return 'data:image/png;base64,' + base64.b64encode(buf.getvalue()).decode()
    except Exception: return None

def _html_to_pdf(page_html, dst):
    tmp = os.path.join(TEMP_DIR, 'builtin_' + hashlib.md5(dst.encode()).hexdigest() + '.html')
    with open(tmp, 'w', encoding='utf-8') as f: f.write(page_html)
    try: _web_to_pdf(tmp, dst)
    finally:
        try: os.remove(tmp)
        except OSError: pass

def _shade(hex6, amount):
    """Lighten (amount > 0) or darken (< 0) a color the way Office tints and shades theme colors."""
    r, g, b = (int(hex6[i:i + 2], 16) for i in (0, 2, 4))
    f = (lambda c: c + (255 - c) * amount) if amount > 0 else (lambda c: c * (1 + amount))
    return '%02X%02X%02X' % tuple(max(0, min(255, round(f(c)))) for c in (r, g, b))

_PAGE_CSS = "body{margin:0;-webkit-print-color-adjust:exact;print-color-adjust:exact} p{margin:0} table{border-collapse:collapse}"

# ----- Word -----
def _docx_to_html(src):
    import docx
    from docx.oxml.ns import qn
    from docx.text.run import Run
    from docx.table import _Cell
    d = docx.Document(src)
    W = lambda tag: qn('w:' + tag)
    sec = d.sections[0]
    inch = lambda v, dflt: (v / 914400) if v is not None else dflt
    pw, ph = inch(sec.page_width, 8.5), inch(sec.page_height, 11)
    mt, mb, ml, mr = inch(sec.top_margin, 1), inch(sec.bottom_margin, 1), inch(sec.left_margin, 1), inch(sec.right_margin, 1)
    hd = inch(sec.header_distance, 0.5)

    # document defaults (font, size, spacing) that styles build on
    styles_el = d.styles.element
    def dflt(path, attr):
        el = styles_el.find(W('docDefaults') + '/' + path)
        return el.get(W(attr)) if el is not None else None
    # the theme's fonts: most documents say "the theme's body font" rather than naming one
    tfonts = {}
    try:
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from lxml import etree
        th = etree.fromstring(d.part.part_related_by(RT.THEME).blob)
        for kind in ('minor', 'major'):
            lat = th.find(f".//{qn('a:' + kind + 'Font')}/{qn('a:latin')}")
            if lat is not None and lat.get('typeface'): tfonts[kind] = lat.get('typeface')
    except Exception: pass
    def font_of(rfonts):
        if rfonts is None: return None
        if rfonts.get(W('ascii')): return rfonts.get(W('ascii'))
        t = rfonts.get(W('asciiTheme')) or rfonts.get(W('hAnsiTheme'))
        return tfonts.get('major' if t and t.startswith('major') else 'minor') if t else None
    base_font = font_of(styles_el.find(f"{W('docDefaults')}/{W('rPrDefault')}/{W('rPr')}/{W('rFonts')}")) or tfonts.get('minor') or 'Calibri'
    style_font = lambda st: st.font.name or font_of(st.element.find(f"{W('rPr')}/{W('rFonts')}"))
    base_size = int(dflt(f"{W('rPrDefault')}/{W('rPr')}/{W('sz')}", 'val') or 22) / 2
    base_after = int(dflt(f"{W('pPrDefault')}/{W('pPr')}/{W('spacing')}", 'after') or 0) / 20
    base_line = dflt(f"{W('pPrDefault')}/{W('pPr')}/{W('spacing')}", 'line')
    base_line = int(base_line) / 240 if base_line else 1.0

    # numbering: list number formats per list and level
    nums, counters = {}, {}
    try:
        nx = d.part.numbering_part.element
        absmap = {}
        for an in nx.findall(W('abstractNum')):
            lv = {}
            for l in an.findall(W('lvl')):
                fmt, txt = l.find(W('numFmt')), l.find(W('lvlText'))
                lv[int(l.get(W('ilvl')))] = (fmt.get(W('val')) if fmt is not None else 'bullet', txt.get(W('val')) if txt is not None else '•')
            absmap[an.get(W('abstractNumId'))] = lv
        for n in nx.findall(W('num')):
            a = n.find(W('abstractNumId'))
            if a is not None: nums[n.get(W('numId'))] = absmap.get(a.get(W('val')), {})
    except Exception: pass

    def roman(n):
        out = ''
        for v, s in ((1000, 'm'), (900, 'cm'), (500, 'd'), (400, 'cd'), (100, 'c'), (90, 'xc'), (50, 'l'), (40, 'xl'), (10, 'x'), (9, 'ix'), (5, 'v'), (4, 'iv'), (1, 'i')):
            while n >= v: out += s; n -= v
        return out
    def number(n, fmt):
        if fmt == 'lowerLetter': return chr(96 + (n - 1) % 26 + 1)
        if fmt == 'upperLetter': return chr(64 + (n - 1) % 26 + 1)
        if fmt == 'lowerRoman': return roman(n)
        if fmt == 'upperRoman': return roman(n).upper()
        return str(n)
    def marker(p_el, style):
        numPr = p_el.find(f"{W('pPr')}/{W('numPr')}")
        s = style
        while numPr is None and s is not None:
            numPr = s.element.find(f"{W('pPr')}/{W('numPr')}"); s = s.base_style
        if numPr is None: return None, 0
        nid = numPr.find(W('numId')); il = numPr.find(W('ilvl'))
        nid = nid.get(W('val')) if nid is not None else None
        il = int(il.get(W('val'))) if il is not None else 0
        if not nid or nid == '0' or nid not in nums: return None, 0
        fmt, txt = nums[nid].get(il, ('bullet', '•'))
        c = counters.setdefault(nid, {})
        c[il] = c.get(il, 0) + 1
        for k in [k for k in c if k > il]: del c[k]
        if fmt == 'bullet':
            ch = txt or '•'
            return ('•' if (not ch or ord(ch[0]) >= 0xF000 or ch in ('o',)) and ch != '-' else ch), il
        if fmt == 'none': return '', il
        return re.sub(r'%(\d)', lambda m: number(c.get(int(m.group(1)) - 1, 1), nums[nid].get(int(m.group(1)) - 1, (fmt,))[0]), txt or '%1.'), il

    def chain(style, get):
        while style is not None:
            v = get(style)
            if v is not None: return v
            style = style.base_style
        return None

    HL = {'YELLOW': 'yellow', 'BRIGHT_GREEN': '#00ff00', 'TURQUOISE': '#00ffff', 'PINK': '#ff00ff', 'BLUE': '#0000ff', 'RED': '#ff0000',
          'DARK_BLUE': '#000080', 'TEAL': '#008080', 'GREEN': '#008000', 'VIOLET': '#800080', 'DARK_RED': '#800000', 'DARK_YELLOW': '#808000',
          'GRAY_25': '#c0c0c0', 'GRAY_50': '#808080', 'BLACK': '#000000'}

    def run_html(r, pstyle):
        f = r.font
        cstyle = r.style if r.style is not None and r.style.name != 'Default Paragraph Font' else None
        def prop(name):
            v = getattr(f, name)
            if v is None and cstyle is not None: v = chain(cstyle, lambda s: getattr(s.font, name))
            if v is None: v = chain(pstyle, lambda s: getattr(s.font, name))
            return v
        css = []
        if prop('bold'): css.append('font-weight:bold')
        if prop('italic'): css.append('font-style:italic')
        deco = []
        if prop('underline'): deco.append('underline')
        if prop('strike'): deco.append('line-through')
        if deco: css.append('text-decoration:' + ' '.join(deco))
        size = prop('size')
        if size is not None: css.append(f'font-size:{size.pt}pt')
        name = f.name or font_of(r._r.find(f"{W('rPr')}/{W('rFonts')}"))
        if name is None and cstyle is not None: name = chain(cstyle, style_font)
        if name is None: name = chain(pstyle, style_font)
        if name: css.append(f"font-family:'{name}','{base_font}',Calibri,Arial,sans-serif")
        try:
            c = f.color.rgb if f.color is not None and f.color.type is not None else None
            if c is None and cstyle is not None: c = chain(cstyle, lambda s: s.font.color.rgb if s.font.color.type is not None else None)
            if c is None: c = chain(pstyle, lambda s: s.font.color.rgb if s.font.color.type is not None else None)
            if c is not None: css.append(f'color:#{c}')
        except Exception: pass
        try:
            hl = f.highlight_color
            if hl is not None and getattr(hl, 'name', '') in HL: css.append('background:' + HL[hl.name])
        except Exception: pass
        if prop('superscript'): css.append('vertical-align:super;font-size:smaller')
        elif prop('subscript'): css.append('vertical-align:sub;font-size:smaller')
        if prop('all_caps'): css.append('text-transform:uppercase')
        elif prop('small_caps'): css.append('font-variant:small-caps')
        out = []
        for child in r._r.iterchildren():
            tag = child.tag
            if tag == W('t'): out.append(_esc(child.text or ''))
            elif tag == W('tab'): out.append('<span style="display:inline-block;width:0.5in"></span>')
            elif tag in (W('br'), W('cr')):
                out.append('<span style="display:block;break-after:page"></span>' if child.get(W('type')) == 'page' else '<br>')
            elif tag == W('drawing') or tag == W('pict'):
                for blip in child.iter(qn('a:blip')):
                    rid = blip.get(qn('r:embed')) or blip.get(qn('r:link'))
                    part = r.part.related_parts.get(rid) if rid else None
                    if part is None: continue
                    uri = _img_uri(part.blob, part.content_type)
                    if not uri: continue
                    ext = next(child.iter(qn('wp:extent')), None)
                    wpx = int(ext.get('cx')) / 9525 if ext is not None else None
                    hpx = int(ext.get('cy')) / 9525 if ext is not None else None
                    anchor = next(child.iter(qn('wp:anchor')), None)
                    float_css = ''
                    if anchor is not None:   # a floating picture: to the side it's placed on
                        al = next(anchor.iter(qn('wp:align')), None)
                        pos = next(anchor.iter(qn('wp:posOffset')), None)
                        side = al.text if al is not None else ('right' if pos is not None and int(pos.text or 0) / 914400 > (pw - ml - mr) / 2 else 'left')
                        float_css = f'float:{"right" if side == "right" else "left"};margin:0 8px 4px 8px;'
                    size_css = f'width:{wpx:.1f}px;height:{hpx:.1f}px;' if wpx else 'max-width:100%;'
                    out.append(f'<img src="{uri}" style="{float_css}{size_css}vertical-align:bottom">')
                    break
        if not out: return ''
        body = ''.join(out)
        return f'<span style="{";".join(css)}">{body}</span>' if css else body

    def para_html(p_el, parent):
        from docx.text.paragraph import Paragraph
        p = Paragraph(p_el, parent)
        style = p.style
        pf = p.paragraph_format
        def pprop(name):
            v = getattr(pf, name)
            return v if v is not None else chain(style, lambda s: getattr(s.paragraph_format, name))
        css = []
        al = pprop('alignment')
        if al is not None:
            css.append('text-align:' + {0: 'left', 1: 'center', 2: 'right', 3: 'justify'}.get(int(getattr(al, 'value', al)), 'left'))
        sb, sa = pprop('space_before'), pprop('space_after')
        css.append(f'margin-top:{sb.pt if sb is not None else 0}pt;margin-bottom:{sa.pt if sa is not None else base_after}pt')
        ls = pprop('line_spacing')
        if ls is None: css.append(f'line-height:{base_line * 1.15:.2f}')
        elif isinstance(ls, float): css.append(f'line-height:{ls * 1.15:.2f}')
        else: css.append(f'line-height:{ls.pt}pt')
        li, fi, ri = pprop('left_indent'), pprop('first_line_indent'), pprop('right_indent')
        mark, lvl = marker(p_el, style)
        left = li.pt if li is not None else (0.25 * 72 * (lvl + 1) if mark is not None else 0)
        if left: css.append(f'margin-left:{left}pt')
        if ri is not None and ri.pt: css.append(f'margin-right:{ri.pt}pt')
        if fi is not None and fi.pt and mark is None: css.append(f'text-indent:{fi.pt}pt')
        if pprop('page_break_before'): css.append('break-before:page')
        sz = chain(style, lambda s: s.font.size)
        css.append(f'font-size:{sz.pt if sz is not None else base_size}pt')
        fn = chain(style, style_font)
        css.append(f"font-family:'{fn or base_font}',Calibri,Arial,sans-serif")
        shd = p_el.find(f"{W('pPr')}/{W('shd')}")
        if shd is not None and shd.get(W('fill')) not in (None, 'auto'): css.append(f"background:#{shd.get(W('fill'))}")
        # the runs, in order, including those inside links, fields and content controls (not deleted ones)
        parts = []
        for r_el in p_el.iter(W('r')):
            anc = r_el.getparent(); skip = False; link = None
            while anc is not None and anc is not p_el:
                if anc.tag in (W('del'), W('moveFrom'), qn('mc:Fallback'), W('txbxContent')): skip = True; break
                if anc.tag == W('hyperlink'): link = anc
                anc = anc.getparent()
            if skip: continue
            h = run_html(Run(r_el, p), style)
            if link is not None and link.get(qn('r:id')):
                try: h = f'<a href="{_esc(p.part.rels[link.get(qn("r:id"))].target_ref)}" style="color:inherit">{h}</a>'
                except Exception: pass
            parts.append(h)
        inner = ''.join(parts) or '&nbsp;'
        if mark:
            inner = f'<span style="display:inline-block;min-width:0.25in;margin-left:-0.25in">{_esc(mark)}</span>' + inner
        return f'<p style="{";".join(css)}">{inner}</p>'

    def table_html(t_el, parent):
        grid = [int(g.get(W('w')) or 0) / 20 for g in t_el.iterfind(f"{W('tblGrid')}/{W('gridCol')}")]
        bord = t_el.find(f"{W('tblPr')}/{W('tblBorders')}")
        sname = t_el.find(f"{W('tblPr')}/{W('tblStyle')}")
        has_b = bord is not None and any(b.get(W('val')) not in ('nil', 'none', None) for b in bord)
        if not has_b and sname is not None:
            try:
                st = d.styles.element.xpath(f'w:style[@w:styleId="{sname.get(W("val"))}"]')
                has_b = bool(st and st[0].xpath('.//w:tblBorders/*[@w:val!="nil" and @w:val!="none"]'))
            except Exception: pass
        cell_b = 'border:0.75pt solid #000;' if has_b else ''
        rows = [list(tr.iterfind(W('tc'))) for tr in t_el.iterfind(W('tr'))]
        pos = []   # each row: [(grid column, span, tc)]
        for cells in rows:
            col, row = 0, []
            for tc in cells:
                gs = tc.find(f"{W('tcPr')}/{W('gridSpan')}")
                span = int(gs.get(W('val'))) if gs is not None else 1
                row.append((col, span, tc)); col += span
            pos.append(row)
        def vmerge(tc):
            v = tc.find(f"{W('tcPr')}/{W('vMerge')}")
            return None if v is None else (v.get(W('val')) or 'continue')
        out = [f'<table style="border-collapse:collapse;table-layout:fixed;width:{sum(grid) or 468}pt;margin:2pt 0">']
        if grid: out.append('<colgroup>' + ''.join(f'<col style="width:{g}pt">' for g in grid) + '</colgroup>')
        for ri, row in enumerate(pos):
            out.append('<tr>')
            for col, span, tc in row:
                vm = vmerge(tc)
                if vm == 'continue': continue
                rs = 1
                if vm == 'restart':
                    for nxt in pos[ri + 1:]:
                        hit = next((t for c, s, t in nxt if c == col), None)
                        if hit is not None and vmerge(hit) == 'continue': rs += 1
                        else: break
                css = cell_b + 'padding:0 5.4pt;vertical-align:top;'
                tcb = tc.find(f"{W('tcPr')}/{W('tcBorders')}")
                if tcb is not None:
                    for side in ('top', 'left', 'bottom', 'right'):
                        b = tcb.find(W(side))
                        if b is not None:
                            css += f"border-{side}:" + ('none;' if b.get(W('val')) in ('nil', 'none') else f"{max(0.5, int(b.get(W('sz')) or 4) / 8)}pt solid #{(b.get(W('color')) or '000000').replace('auto', '000000')};")
                shd = tc.find(f"{W('tcPr')}/{W('shd')}")
                if shd is not None and shd.get(W('fill')) not in (None, 'auto'): css += f"background:#{shd.get(W('fill'))};"
                va = tc.find(f"{W('tcPr')}/{W('vAlign')}")
                if va is not None: css += 'vertical-align:' + {'center': 'middle', 'bottom': 'bottom'}.get(va.get(W('val')), 'top') + ';'
                cell = _Cell(tc, parent)
                out.append(f'<td colspan="{span}" rowspan="{rs}" style="{css}">{blocks_html(tc, cell)}</td>')
            out.append('</tr>')
        out.append('</table>')
        return ''.join(out)

    def blocks_html(container, parent):
        out = []
        for el in container.iterchildren():
            if el.tag == W('p'): out.append(para_html(el, parent))
            elif el.tag == W('tbl'): out.append(table_html(el, parent))
            elif el.tag == W('sdt'):
                c = el.find(W('sdtContent'))
                if c is not None: out.append(blocks_html(c, parent))
        return ''.join(out)

    body = blocks_html(d.element.body, d._body)
    # the first page's header (letterheads, logos) sits in its top margin
    header = ''
    try:
        hdr = sec.first_page_header if sec.different_first_page_header_footer else sec.header
        if not hdr.is_linked_to_previous: header = blocks_html(hdr._element, hdr)
    except Exception: header = ''
    head_h = max(0.0, mt - hd)
    page_css = (f"@page{{size:{pw}in {ph}in;margin:{mt}in {mr}in {mb}in {ml}in}}"
                + (f"@page :first{{margin-top:{hd}in}}" if header else ''))
    head_div = f'<div style="height:{head_h}in;overflow:hidden">{header}</div>' if header else ''
    return (f'<!doctype html><html><head><meta charset="utf-8"><style>{page_css}{_PAGE_CSS}'
            f"body{{font-family:'{base_font}',Calibri,Arial,sans-serif;font-size:{base_size}pt}} img{{max-width:100%}}</style></head>"
            f'<body>{head_div}{body}</body></html>')

# ----- Excel (and CSV): every sheet as a table laid out like the sheet -----
_THEME_XL = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']

def _xl_color(c):
    if c is None: return None
    try:
        if c.type == 'rgb' and isinstance(c.rgb, str): return c.rgb[-6:]
        if c.type == 'theme' and c.theme is not None and c.theme < len(_THEME_XL):
            return _shade(_THEME_XL[c.theme], c.tint or 0)
        if c.type == 'indexed' and c.indexed is not None:
            from openpyxl.styles.colors import COLOR_INDEX
            if c.indexed < len(COLOR_INDEX): return COLOR_INDEX[c.indexed][-6:]
            return '000000' if c.indexed == 64 else None
    except Exception: pass
    return None

def _xl_fmt(v, fmt):
    import datetime
    if v is None: return ''
    if isinstance(v, bool): return 'TRUE' if v else 'FALSE'
    if isinstance(v, datetime.datetime) and (v.hour or v.minute) and 'h' in (fmt or '').lower():
        return f"{v.month}/{v.day}/{v.year} {v.strftime('%I:%M %p').lstrip('0')}"
    if isinstance(v, (datetime.datetime, datetime.date)):
        f = (fmt or '').lower()
        if 'mmmm' in f: return v.strftime('%B %d, %Y').replace(' 0', ' ')
        if 'mmm' in f: return f"{v.day}-{v.strftime('%b')}" + (f"-{v.strftime('%y')}" if 'yy' in f else '')
        return f"{v.month}/{v.day}/{v.year}"
    if isinstance(v, datetime.time): return v.strftime('%I:%M %p').lstrip('0')
    if isinstance(v, (int, float)):
        f = fmt or 'General'
        if f == 'General' or ('0' not in f and '#' not in f):
            if isinstance(v, float):
                return str(int(v)) if v.is_integer() and abs(v) < 1e15 else f'{v:.10g}'
            return str(v)
        sec = f.split(';')
        body = re.sub(r'"[^"]*"|\[[^\]]*\]|_.|\\.|\*.', '', sec[0])
        pct = '%' in body
        x = v * 100 if pct else v
        m = re.search(r'\.(0+)', body)
        dec = len(m.group(1)) if m else 0
        s = f'{abs(x):,.{dec}f}' if ',' in body else f'{abs(x):.{dec}f}'
        cur = '$' if '$' in sec[0] else '€' if '€' in sec[0] else '£' if '£' in sec[0] else ''
        s = cur + s + ('%' if pct else '')
        if x < 0: s = f'({s})' if len(sec) > 1 and '(' in sec[1] else '-' + s
        return s
    return str(v)

def _xlsx_sheets(src):
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    wb = load_workbook(src, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        if ws.sheet_state != 'visible': continue
        # how far the sheet goes: the print area if it has one, else the last cell with a value or a border/fill
        r0, c0, r1, c1 = 1, 1, 0, 0
        area = None
        try: area = ws.print_area
        except Exception: pass
        if area:
            from openpyxl.utils.cell import range_boundaries
            first = (area if isinstance(area, str) else area[0]).split(',')[0]
            c0, r0, c1, r1 = range_boundaries(first.split('!')[-1].replace('$', ''))
        else:
            for row in ws.iter_rows():
                for cell in row:
                    b = cell.border
                    if cell.value not in (None, '') or (cell.fill is not None and cell.fill.fill_type == 'solid') or \
                            (b is not None and any(getattr(getattr(b, s), 'style', None) for s in ('left', 'right', 'top', 'bottom'))):
                        r1, c1 = max(r1, cell.row), max(c1, cell.column)
            for mr in ws.merged_cells.ranges:
                if mr.min_row <= r1 + 1: r1, c1 = max(r1, mr.max_row), max(c1, mr.max_col)
        if r1 == 0: continue
        widths, hidden_c = {}, set()
        dw = ws.sheet_format.defaultColWidth or 8.43
        for cd in ws.column_dimensions.values():
            for i in range((cd.min or 0), (cd.max or 0) + 1):
                if cd.hidden: hidden_c.add(i)
                if cd.width: widths[i] = cd.width
        dh = ws.sheet_format.defaultRowHeight or 15
        cols = [c for c in range(c0, c1 + 1) if c not in hidden_c]
        rows = [r for r in range(r0, r1 + 1) if not ws.row_dimensions[r].hidden]
        colw = {c: round(widths.get(c, dw) * 7 + 5) for c in cols}
        rowh = {r: (ws.row_dimensions[r].height or dh) * 96 / 72 for r in rows}
        span, covered = {}, set()
        for mr in ws.merged_cells.ranges:
            rr = [r for r in range(mr.min_row, mr.max_row + 1) if r in rowh]
            cc = [c for c in range(mr.min_col, mr.max_col + 1) if c in colw]
            if not rr or not cc: continue
            span[(rr[0], cc[0])] = (len(rr), len(cc))
            for r in rr:
                for c in cc:
                    if (r, c) != (rr[0], cc[0]): covered.add((r, c))
        cells = {}
        for r in rows:
            for c in cols:
                if (r, c) in covered: continue
                cell = ws.cell(r, c)
                css, f = [], cell.font
                if f is not None:
                    if f.b: css.append('font-weight:bold')
                    if f.i: css.append('font-style:italic')
                    if f.u: css.append('text-decoration:underline')
                    if f.sz: css.append(f'font-size:{f.sz}pt')
                    if f.name: css.append(f"font-family:'{f.name}',Calibri,Arial,sans-serif")
                    fc = _xl_color(f.color)
                    if fc: css.append(f'color:#{fc}')
                if cell.fill is not None and cell.fill.fill_type == 'solid':
                    bg = _xl_color(cell.fill.fgColor)
                    if bg: css.append(f'background:#{bg}')
                b = cell.border
                if b is not None:
                    for side in ('left', 'right', 'top', 'bottom'):
                        s = getattr(b, side)
                        if s is not None and s.style:
                            w = 3 if s.style in ('thick', 'double') else 2 if s.style.startswith('medium') else 1
                            css.append(f'border-{side}:{w}px {"dashed" if "dash" in s.style or "dot" in s.style else "double" if s.style == "double" else "solid"} #{_xl_color(s.color) or "000000"}')
                # a merged block's outer border comes from its edge cells
                if (r, c) in span:
                    rs, cs = span[(r, c)]
                    last_c = cols[min(len(cols) - 1, cols.index(c) + cs - 1)]
                    rb = ws.cell(r, last_c).border
                    if rb is not None and rb.right is not None and rb.right.style:
                        css.append(f'border-right:{3 if rb.right.style in ("thick", "double") else 2 if rb.right.style.startswith("medium") else 1}px solid #{_xl_color(rb.right.color) or "000000"}')
                    last_r = rows[min(len(rows) - 1, rows.index(r) + rs - 1)]
                    bb = ws.cell(last_r, c).border
                    if bb is not None and bb.bottom is not None and bb.bottom.style:
                        css.append(f'border-bottom:{3 if bb.bottom.style in ("thick", "double") else 2 if bb.bottom.style.startswith("medium") else 1}px solid #{_xl_color(bb.bottom.color) or "000000"}')
                al = cell.alignment
                num = isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool)
                h = (al.horizontal if al is not None else None) or ('right' if num else None)
                if h in ('center', 'centerContinuous'): css.append('text-align:center')
                elif h == 'right': css.append('text-align:right')
                elif h == 'justify': css.append('text-align:justify')
                v = al.vertical if al is not None else None
                css.append('vertical-align:' + {'top': 'top', 'center': 'middle'}.get(v, 'bottom'))
                wrap = bool(al is not None and al.wrap_text)
                text = _xl_fmt(cell.value, cell.number_format)
                cells[(r, c)] = (text, css, wrap)
        images = []
        xs, ys = {}, {}
        acc = 0
        for c in range(1, c1 + 2):
            xs[c] = acc
            if c in colw: acc += colw[c]
        acc = 0
        for r in range(1, r1 + 2):
            ys[r] = acc
            if r in rowh: acc += rowh[r]
        for img in getattr(ws, '_images', []):
            try:
                a = img.anchor; fr = a._from
                x = xs.get(fr.col + 1, 0) - xs.get(c0, 0) + fr.colOff / 9525
                y = ys.get(fr.row + 1, 0) - ys.get(r0, 0) + fr.rowOff / 9525
                if hasattr(a, 'to') and a.to is not None:
                    w = xs.get(a.to.col + 1, 0) - xs.get(c0, 0) + a.to.colOff / 9525 - x
                    hgt = ys.get(a.to.row + 1, 0) - ys.get(r0, 0) + a.to.rowOff / 9525 - y
                else:
                    w, hgt = a.ext.width / 9525, a.ext.height / 9525
                data = img._data()
                uri = _img_uri(data, 'image/' + (img.format or 'png').lower().replace('jpg', 'jpeg'))
                if uri and w > 0 and hgt > 0: images.append((x, y, w, hgt, uri))
            except Exception: continue
        ps = ws.page_setup
        pm = ws.page_margins
        sheets.append({'name': ws.title, 'rows': rows, 'cols': cols, 'colw': colw, 'rowh': rowh, 'span': span,
                       'cells': cells, 'images': images,
                       'landscape': {'landscape': True, 'portrait': False}.get(ps.orientation),
                       'a4': str(ps.paperSize) == '9',
                       'margins': (pm.top or 0.75, pm.right or 0.7, pm.bottom or 0.75, pm.left or 0.7)})
    return sheets

def _xls_sheets(src):
    """Old .xls workbooks (Excel 97-2003): values, number formats, merged cells, widths, bold and borders."""
    import xlrd
    bk = xlrd.open_workbook(src, formatting_info=True)
    sheets = []
    for sh in bk.sheets():
        if sh.visibility != 0 or sh.nrows == 0: continue
        cols = [c for c in range(sh.ncols) if not (sh.colinfo_map.get(c) and sh.colinfo_map[c].hidden)]
        rows = [r for r in range(sh.nrows) if not (sh.rowinfo_map.get(r) and sh.rowinfo_map[r].hidden)]
        dw = (sh.defcolwidth or 8.43) if sh.defcolwidth else 8.43
        colw = {c: round(((sh.colinfo_map[c].width / 256) if c in sh.colinfo_map else dw) * 7 + 5) for c in cols}
        rowh = {r: ((sh.rowinfo_map[r].height / 20) if r in sh.rowinfo_map else 15) * 96 / 72 for r in rows}
        span, covered = {}, set()
        for rlo, rhi, clo, chi in sh.merged_cells:
            span[(rlo, clo)] = (len([r for r in range(rlo, rhi) if r in rowh]), len([c for c in range(clo, chi) if c in colw]))
            for r in range(rlo, rhi):
                for c in range(clo, chi):
                    if (r, c) != (rlo, clo): covered.add((r, c))
        cells = {}
        for r in rows:
            for c in cols:
                if (r, c) in covered: continue
                cell = sh.cell(r, c)
                css = []
                fmt = None
                try:
                    xf = bk.xf_list[cell.xf_index]
                    font = bk.font_list[xf.font_index]
                    if font.bold: css.append('font-weight:bold')
                    if font.italic: css.append('font-style:italic')
                    if font.underline_type: css.append('text-decoration:underline')
                    css.append(f'font-size:{font.height / 20}pt')
                    css.append(f"font-family:'{font.name}',Arial,sans-serif")
                    fmt = bk.format_map[xf.format_key].format_str
                    bd = xf.border
                    for side, ls, ci in (('left', bd.left_line_style, bd.left_colour_index), ('right', bd.right_line_style, bd.right_colour_index),
                                         ('top', bd.top_line_style, bd.top_colour_index), ('bottom', bd.bottom_line_style, bd.bottom_colour_index)):
                        if ls: css.append(f'border-{side}:{2 if ls in (2, 5, 8) else 1}px solid #000')
                    bg = bk.colour_map.get(xf.background.pattern_colour_index) if xf.background.fill_pattern else None
                    if bg: css.append('background:#%02X%02X%02X' % bg)
                    h = xf.alignment.hor_align
                    if h == 2: css.append('text-align:center')
                    elif h == 3: css.append('text-align:right')
                    wrap = bool(xf.alignment.text_wrapped)
                except Exception: wrap = False
                v = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try: v = xlrd.xldate.xldate_as_datetime(v, bk.datemode)
                    except Exception: pass
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN: v = bool(v)
                elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK): v = None
                elif cell.ctype == xlrd.XL_CELL_ERROR: v = '#ERR'
                if cell.ctype == xlrd.XL_CELL_NUMBER and not any(x.startswith('text-align') for x in css): css.append('text-align:right')
                css.append('vertical-align:bottom')
                cells[(r, c)] = (_xl_fmt(v, fmt), css, wrap)
        sheets.append({'name': sh.name, 'rows': rows, 'cols': cols, 'colw': colw, 'rowh': rowh, 'span': span, 'cells': cells,
                       'images': [], 'landscape': None, 'a4': False, 'margins': (0.75, 0.7, 0.75, 0.7)})
    return sheets

def _csv_sheets(src):
    import csv
    with open(src, newline='', encoding='utf-8-sig', errors='replace') as f:
        data = list(csv.reader(f, delimiter='\t' if src.lower().endswith('.tsv') else ','))
    ncols = max((len(r) for r in data), default=0)
    if not data or not ncols: return []
    rows, cols = list(range(len(data))), list(range(ncols))
    widths = [min(60, max([len(r[c]) if c < len(r) else 0 for r in data] + [6])) for c in cols]
    cells = {}
    for r in rows:
        for c in cols:
            t = data[r][c] if c < len(data[r]) else ''
            num = bool(re.fullmatch(r'[-+$]?[\d,]*\.?\d+%?', t.strip()))
            cells[(r, c)] = (t, ['text-align:right' if num else 'text-align:left', 'border:1px solid #d0d0d0'] + (['font-weight:bold'] if r == 0 else []), False)
    return [{'name': os.path.basename(src), 'rows': rows, 'cols': cols, 'colw': {c: widths[c] * 7 + 10 for c in cols},
             'rowh': {r: 20 for r in rows}, 'span': {}, 'cells': cells, 'images': [], 'landscape': None, 'a4': False,
             'margins': (0.5, 0.5, 0.5, 0.5)}]

def _sheets_to_html(sheets):
    if not sheets: raise ValueError('This spreadsheet is empty')
    css = [_PAGE_CSS, "td{padding:1px 3px;overflow:hidden;white-space:nowrap;font-family:Calibri,Arial,sans-serif;font-size:11pt;box-sizing:border-box}",
           "td.wrap{white-space:normal;word-wrap:break-word} td.spill{overflow:visible} .sheet{position:relative;break-after:page}"]
    pages = []
    for i, s in enumerate(sheets):
        width = sum(s['colw'][c] for c in s['cols'])
        height = sum(s['rowh'][r] for r in s['rows'])
        land = s['landscape'] if s['landscape'] is not None else width > 720
        pw, ph = (8.27, 11.69) if s['a4'] else (8.5, 11)
        if land: pw, ph = ph, pw
        mt, mr, mb, ml = s['margins']
        zoom = min(1.0, (pw - ml - mr) * 96 / max(width, 1))
        name = f'p{i}'
        css.append(f"@page {name}{{size:{pw}in {ph}in;margin:{mt}in {mr}in {mb}in {ml}in}} .{name}{{page:{name}}}")
        out = [f'<div class="sheet {name}"><div style="position:relative;zoom:{zoom:.4f};width:{width}px">',
               f'<table style="table-layout:fixed;width:{width}px"><colgroup>' + ''.join(f'<col style="width:{s["colw"][c]}px">' for c in s['cols']) + '</colgroup>']
        for r in s['rows']:
            out.append(f'<tr style="height:{s["rowh"][r]:.1f}px">')
            for k, c in enumerate(s['cols']):
                if (r, c) not in s['cells']: continue
                text, ccss, wrap = s['cells'][(r, c)]
                rs, cs = s['span'].get((r, c), (1, 1))
                cls = 'wrap' if wrap else ''
                # text runs on into the next cells when they're empty, as in Excel
                if not wrap and text and cs == 1 and k + 1 < len(s['cols']):
                    nxt = s['cells'].get((r, s['cols'][k + 1]))
                    if nxt is not None and not nxt[0]: cls = 'spill'
                attrs = (f' rowspan="{rs}"' if rs > 1 else '') + (f' colspan="{cs}"' if cs > 1 else '')
                out.append(f'<td{attrs} class="{cls}" style="{";".join(ccss)}">{_esc(text)}</td>')
            out.append('</tr>')
        out.append('</table>')
        for x, y, w, h, uri in s['images']:
            out.append(f'<img src="{uri}" style="position:absolute;left:{x:.1f}px;top:{y:.1f}px;width:{w:.1f}px;height:{h:.1f}px">')
        out.append('</div></div>')
        pages.append(''.join(out))
    return f'<!doctype html><html><head><meta charset="utf-8"><style>{"".join(css)}</style></head><body>{"".join(pages)}</body></html>'

# ----- PowerPoint: each slide as a page, shapes where they sit on it -----
def _pptx_to_html(src):
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
    from pptx.oxml.ns import qn
    prs = Presentation(src)
    SW, SH = (prs.slide_width or 12192000) / 9525, (prs.slide_height or 6858000) / 9525
    master = prs.slide_master
    # the theme's colors, for shapes and text that use them
    theme = {}
    try:
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        tx = master.part.part_related_by(RT.THEME).blob
        from lxml import etree
        cs = etree.fromstring(tx).find('.//' + qn('a:clrScheme'))
        for el in cs:
            name = el.tag.split('}')[1]
            srgb = el.find(qn('a:srgbClr')); sysc = el.find(qn('a:sysClr'))
            theme[name] = srgb.get('val') if srgb is not None else (sysc.get('lastClr') if sysc is not None else None)
    except Exception: pass
    THEME_KEYS = {'ACCENT_1': 'accent1', 'ACCENT_2': 'accent2', 'ACCENT_3': 'accent3', 'ACCENT_4': 'accent4', 'ACCENT_5': 'accent5',
                  'ACCENT_6': 'accent6', 'DARK_1': 'dk1', 'TEXT_1': 'dk1', 'LIGHT_1': 'lt1', 'BACKGROUND_1': 'lt1', 'DARK_2': 'dk2',
                  'TEXT_2': 'dk2', 'LIGHT_2': 'lt2', 'BACKGROUND_2': 'lt2', 'HYPERLINK': 'hlink'}
    def color(cf):
        try:
            if cf.type is None: return None
            try: rgb = str(cf.rgb)
            except AttributeError:
                rgb = theme.get(THEME_KEYS.get(getattr(cf.theme_color, 'name', ''), ''), None)
            if not rgb: return None
            br = cf.brightness or 0
            return _shade(rgb, br) if br else rgb
        except Exception: return None

    def inherited(sh, getter):
        """Look up a text setting on the layout's and master's matching placeholder, then the master's text styles."""
        if not sh.is_placeholder: return None
        try:
            idx, typ = sh.placeholder_format.idx, sh.placeholder_format.type
            cands = []
            lay = sh.part.slide.slide_layout
            for ph in lay.placeholders:
                if ph.placeholder_format.idx == idx: cands.append(ph._element); break
            for ph in lay.slide_master.placeholders:
                if ph.placeholder_format.type == typ: cands.append(ph._element); break
            for el in cands:
                v = getter(el)
                if v is not None: return v
            ts = lay.slide_master._element.find(qn('p:txStyles'))
            key = 'p:titleStyle' if typ in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE) else 'p:bodyStyle' if typ in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.SUBTITLE) else 'p:otherStyle'
            if ts is not None and ts.find(qn(key)) is not None: return getter(ts.find(qn(key)))
        except Exception: pass
        return None

    def text_html(sh):
        tf = sh.text_frame
        body = tf._txBody
        bp = body.find(qn('a:bodyPr'))
        scale = 1.0
        na = bp.find(qn('a:normAutofit')) if bp is not None else None
        if na is not None and na.get('fontScale'): scale = int(na.get('fontScale')) / 100000
        anchor = bp.get('anchor') if bp is not None else None
        if anchor is None: anchor = inherited(sh, lambda el: next((b.get('anchor') for b in el.iter(qn('a:bodyPr')) if b.get('anchor')), None))
        is_title = sh.is_placeholder and sh.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
        is_body = sh.is_placeholder and sh.placeholder_format.type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)
        ins = lambda k, d: int(bp.get(k)) / 9525 if bp is not None and bp.get(k) else d
        pad = f'padding:{ins("tIns", 4.8)}px {ins("rIns", 9.6)}px {ins("bIns", 4.8)}px {ins("lIns", 9.6)}px;'
        paras = []
        for p in tf.paragraphs:
            lvl = p.level
            def lvl_sz(el, lvl=lvl):
                for l in el.iter(qn(f'a:lvl{lvl + 1}pPr')):
                    d = l.find(qn('a:defRPr'))
                    if d is not None and d.get('sz'): return int(d.get('sz')) / 100
                return None
            dsz = inherited(sh, lvl_sz) or (44 if is_title else 18)
            al = p.alignment
            if al is None:
                al_inh = inherited(sh, lambda el, lvl=lvl: next((l.get('algn') for l in el.iter(qn(f'a:lvl{lvl + 1}pPr')) if l.get('algn')), None))
                align = {'ctr': 'center', 'r': 'right', 'just': 'justify'}.get(al_inh, 'left')
            else:
                align = {1: 'left', 2: 'center', 3: 'right', 4: 'justify'}.get(int(getattr(al, 'value', al)), 'left')
            runs = []
            for r in p.runs:
                f = r.font
                sz = (f.size.pt if f.size else dsz) * scale
                css = [f'font-size:{sz:.1f}pt']
                if f.bold: css.append('font-weight:bold')
                if f.italic: css.append('font-style:italic')
                if f.underline: css.append('text-decoration:underline')
                if f.name: css.append(f"font-family:'{f.name}',Calibri,Arial,sans-serif")
                c = color(f.color)
                if c: css.append(f'color:#{c}')
                runs.append(f'<span style="{";".join(css)}">{_esc(r.text)}</span>')
            ppr = p._p.find(qn('a:pPr'))
            bullet = ''
            if runs and is_body and not (ppr is not None and ppr.find(qn('a:buNone')) is not None):
                bullet = '<span style="display:inline-block;width:0.9em;margin-left:-0.9em">•</span>'
            if ppr is not None:
                bc = ppr.find(qn('a:buChar'))
                if bc is not None and runs: bullet = f'<span style="display:inline-block;width:0.9em;margin-left:-0.9em">{_esc(bc.get("char") or "•")}</span>'
            indent = (lvl * 0.4 + (0.35 if bullet else 0))
            line = f'font-size:{dsz * scale:.1f}pt'
            paras.append(f'<div style="text-align:{align};margin-left:{indent}in;{line};line-height:1.15">{bullet}{"".join(runs) or "&nbsp;"}</div>')
        jc = {'ctr': 'center', 'b': 'flex-end'}.get(anchor or ('ctr' if is_title else 't'), 'flex-start')
        wrap = 'white-space:nowrap;' if bp is not None and bp.get('wrap') == 'none' else ''
        return f'<div style="position:absolute;inset:0;display:flex;flex-direction:column;justify-content:{jc};{pad}{wrap}box-sizing:border-box">{"".join(paras)}</div>'

    def shape_html(sh, tr):
        """tr maps the shape's own position to the slide (groups have their own coordinates inside)."""
        try:
            x, y = tr(sh.left or 0, sh.top or 0)
            x2, y2 = tr((sh.left or 0) + (sh.width or 0), (sh.top or 0) + (sh.height or 0))
        except Exception: return ''
        w, h = x2 - x, y2 - y
        rot = f'transform:rotate({sh.rotation}deg);' if getattr(sh, 'rotation', 0) else ''
        box = f'position:absolute;left:{x:.1f}px;top:{y:.1f}px;width:{w:.1f}px;height:{h:.1f}px;{rot}'
        st = sh.shape_type
        if st == MSO_SHAPE_TYPE.GROUP:
            g = sh._element.find(qn('p:grpSpPr')).find(qn('a:xfrm'))
            off, ext = g.find(qn('a:off')), g.find(qn('a:ext'))
            cho, che = g.find(qn('a:chOff')), g.find(qn('a:chExt'))
            ox, oy, ew, eh = int(off.get('x')), int(off.get('y')), int(ext.get('cx')) or 1, int(ext.get('cy')) or 1
            cx, cy, cw, chh = int(cho.get('x')), int(cho.get('y')), int(che.get('cx')) or 1, int(che.get('cy')) or 1
            inner = lambda a, b: tr(ox + (a - cx) * ew / cw, oy + (b - cy) * eh / chh)
            return ''.join(shape_html(s, inner) for s in sh.shapes)
        if st == MSO_SHAPE_TYPE.PICTURE or (sh.is_placeholder and hasattr(sh, 'image')):
            try:
                uri = _img_uri(sh.image.blob, sh.image.content_type)
                if uri: return f'<img src="{uri}" style="{box}object-fit:fill">'
            except Exception: pass
            return ''
        if st in (MSO_SHAPE_TYPE.LINE,) or sh._element.tag == qn('p:cxnSp'):
            try:
                c = color(sh.line.color) or '000000'
                lw = max(1, (sh.line.width or 9525) / 9525)
            except Exception: c, lw = '000000', 1
            flipH = sh._element.find('.//' + qn('a:xfrm')).get('flipH') == '1'
            flipV = sh._element.find('.//' + qn('a:xfrm')).get('flipV') == '1'
            ax, bx = (x2, x) if flipH else (x, x2)
            ay, by = (y2, y) if flipV else (y, y2)
            return (f'<svg style="position:absolute;left:0;top:0;overflow:visible" width="1" height="1">'
                    f'<line x1="{ax:.1f}" y1="{ay:.1f}" x2="{bx:.1f}" y2="{by:.1f}" stroke="#{c}" stroke-width="{lw:.1f}"/></svg>')
        if getattr(sh, 'has_table', False) and sh.has_table:
            rows = []
            for row in sh.table.rows:
                cells = []
                for cell in row.cells:
                    if cell.is_spanned: continue
                    bg = ''
                    try:
                        if cell.fill.type is not None:
                            cc = color(cell.fill.fore_color)
                            if cc: bg = f'background:#{cc};'
                    except Exception: pass
                    txt = '<br>'.join(_esc(p.text) for p in cell.text_frame.paragraphs)
                    span = (f' colspan="{cell.span_width}"' if cell.span_width > 1 else '') + (f' rowspan="{cell.span_height}"' if cell.span_height > 1 else '')
                    cells.append(f'<td{span} style="{bg}border:1px solid #fff;padding:4px 8px;font-size:14pt">{txt}</td>')
                rows.append(f'<tr style="height:{row.height / 9525:.1f}px">{"".join(cells)}</tr>')
            return f'<table style="{box}border-collapse:collapse;background:#e8ecf4">{"".join(rows)}</table>'
        if getattr(sh, 'has_chart', False) and sh.has_chart:
            title = ''
            try: title = sh.chart.chart_title.text_frame.text if sh.chart.has_title else ''
            except Exception: pass
            return f'<div style="{box}border:1px dashed #bbb;display:flex;align-items:center;justify-content:center;color:#777;font:14pt Calibri,Arial">{_esc(title or "Chart")}</div>'
        style = box
        try:
            if sh.fill.type == 1:   # solid
                c = color(sh.fill.fore_color)
                if c: style += f'background:#{c};'
        except Exception: pass
        try:
            if sh.line.fill.type == 1:
                c = color(sh.line.color)
                if c: style += f'border:{max(1, (sh.line.width or 9525) / 9525):.1f}px solid #{c};box-sizing:border-box;'
        except Exception: pass
        try:
            name = str(sh.auto_shape_type).upper()
            if 'OVAL' in name: style += 'border-radius:50%;'
            elif 'ROUNDED' in name: style += f'border-radius:{min(w, h) * 0.16:.1f}px;'
        except Exception: pass
        txt = text_html(sh) if getattr(sh, 'has_text_frame', False) and sh.has_text_frame and sh.text_frame.text.strip() else ''
        if not txt and style == box: return ''
        return f'<div style="{style}">{txt}</div>'

    def background(*holders):
        for hd in holders:
            bg = hd._element.find(qn('p:cSld')).find(qn('p:bg'))
            if bg is None: continue
            blip = next(bg.iter(qn('a:blip')), None)
            if blip is not None:
                try:
                    part = hd.part.related_part(blip.get(qn('r:embed')))
                    uri = _img_uri(part.blob, part.content_type)
                    if uri: return f'background:url({uri}) center/cover no-repeat;'
                except Exception: pass
            srgb = next(bg.iter(qn('a:srgbClr')), None)
            if srgb is not None: return f'background:#{srgb.get("val")};'
            sch = next(bg.iter(qn('a:schemeClr')), None)
            if sch is not None:
                key = {'bg1': 'lt1', 'bg2': 'lt2', 'tx1': 'dk1', 'tx2': 'dk2'}.get(sch.get('val'), sch.get('val'))
                if theme.get(key): return f'background:#{theme[key]};'
        return 'background:#fff;'

    ident = lambda a, b: (a / 9525, b / 9525)
    pages = []
    for slide in prs.slides:
        if slide._element.get('show') == '0': continue
        lay = slide.slide_layout
        out = []
        # what the master and layout draw behind the slide (logos, bands), not their empty placeholders
        if slide._element.get('showMasterSp') != '0':
            if lay._element.get('showMasterSp') != '0':
                out += [shape_html(s, ident) for s in lay.slide_master.shapes if not s.is_placeholder]
            out += [shape_html(s, ident) for s in lay.shapes if not s.is_placeholder]
        out += [shape_html(s, ident) for s in slide.shapes]
        pages.append(f'<div style="position:relative;width:{SW:.1f}px;height:{SH:.1f}px;overflow:hidden;break-after:page;'
                     f'{background(slide, lay, lay.slide_master)}font-family:Calibri,Arial,sans-serif">{"".join(out)}</div>')
    if not pages: raise ValueError('This presentation has no slides to show')
    return (f'<!doctype html><html><head><meta charset="utf-8"><style>@page{{size:{SW / 96:.3f}in {SH / 96:.3f}in;margin:0}}{_PAGE_CSS}</style></head>'
            f'<body>{"".join(pages)}</body></html>')

def _builtin_to_pdf(src, dst):
    low = src.lower()
    if low.endswith(('.docx', '.docm', '.dotx', '.dotm')): page = _docx_to_html(src)
    elif low.endswith('.xls'): page = _sheets_to_html(_xls_sheets(src))
    elif low.endswith(('.csv', '.tsv')): page = _sheets_to_html(_csv_sheets(src))
    elif low.endswith(('.xlsx', '.xlsm', '.xltx', '.xltm')): page = _sheets_to_html(_xlsx_sheets(src))
    elif low.endswith(('.pptx', '.pptm', '.ppsx', '.ppsm', '.potx', '.potm')): page = _pptx_to_html(src)
    else: raise ValueError('No built-in converter for this kind of file')
    _html_to_pdf(page, dst)

# a test can make this PC look like a new one with no Office or LibreOffice (MYPDFKIT_NO_OFFICE=1)
NO_OFFICE_TEST = os.environ.get('MYPDFKIT_NO_OFFICE') == '1'

def find_soffice():
    if NO_OFFICE_TEST: return None
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

class NoOfficeError(RuntimeError):
    """Neither Microsoft Office nor LibreOffice is here to open this file (the app offers a LibreOffice download)."""

def _office_has(src):
    """Is the Office app for this kind of file installed?"""
    if NO_OFFICE_TEST: return False
    import winreg
    low = src.lower()
    kind = 'word' if low.endswith(WORD_EXT) else 'excel' if low.endswith(EXCEL_EXT) else 'ppt' if low.endswith(POWERPOINT_EXT) else None
    if not kind: return False
    try: winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, _OfficeWorker.APPS[kind][0])); return True
    except OSError: return False

def office_to_pdf(src, dst):
    """Microsoft Office when it's installed, else LibreOffice, else the built-in converter (Word, Excel,
    PowerPoint and CSV files open on any new PC). Returns 'builtin' when the built-in one made it."""
    if src.lower().endswith(IMAGE_EXT + BOOK_EXT + WEB_EXT + TEXT_EXT + DESIGN_EXT):
        return _com_to_pdf(src, dst)
    builtin_ok = src.lower().endswith(BUILTIN_EXT)
    try:
        if not _office_has(src): raise NoOfficeError('Office is not installed')
        _com_to_pdf(src, dst)
    except Exception as com_err:
        if isinstance(com_err, ValueError): raise
        if find_soffice():
            try: return libreoffice_to_pdf(src, dst)
            except Exception:
                if not builtin_ok: raise
        if builtin_ok:
            try:
                _builtin_to_pdf(src, dst); return 'builtin'
            except Exception as e:
                print('built-in converter failed:', e, file=sys.stderr)
                if isinstance(e, ValueError) and not _office_has(src): raise
        # Office is here but couldn't open this file: say what went wrong, not that Office is missing
        if _office_has(src): raise com_err
        raise NoOfficeError("Opening this file needs Microsoft Office or the free LibreOffice, and neither is installed.")

CONVERTER_VERSION = '5'   # bump when office_to_pdf output changes (2: cut-off Excel columns widened, 3: many more formats)

# ---------- forms ----------
FIELD_KINDS = {fitz.PDF_WIDGET_TYPE_TEXT: 'text', fitz.PDF_WIDGET_TYPE_CHECKBOX: 'checkbox', fitz.PDF_WIDGET_TYPE_RADIOBUTTON: 'radio',
               fitz.PDF_WIDGET_TYPE_COMBOBOX: 'combo', fitz.PDF_WIDGET_TYPE_LISTBOX: 'list', fitz.PDF_WIDGET_TYPE_SIGNATURE: 'signature',
               fitz.PDF_WIDGET_TYPE_BUTTON: 'button'}

def _inherited_q(doc, xref):
    """A field's text alignment (0 left, 1 center, 2 right): on the widget, else its parents, else the form's."""
    seen = 0
    while xref and seen < 12:
        t_, v = doc.xref_get_key(xref, 'Q')
        if t_ == 'int': return int(v)
        t_, v = doc.xref_get_key(xref, 'Parent')
        xref = int(v.split()[0]) if t_ == 'xref' else 0; seen += 1
    t_, v = doc.xref_get_key(doc.pdf_catalog(), 'AcroForm/Q')
    return int(v) if t_ == 'int' else 0

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
                    "size": w.text_fontsize or 0, "maxlen": w.text_maxlen or 0, "label": w.field_label or '',
                    "align": _inherited_q(page.parent, w.xref), "hidden": bool(_flag(page.parent, w.xref) & 2)})
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
            notes.append({"xref": a.xref, "kind": a.type[1], "rect": [r.x0, r.y0, r.x1, r.y1], "text": text, "subject": a.info.get('subject') or '',
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
        # a field can show in several places (copies on other pages): the value goes into every one of them
        names, radios = {}, {}
        for pno, vals in by_page.items():
            page = doc[pno]
            widgets = {w.xref: w for w in page.widgets() or []}
            for v in vals:
                w = widgets.get(int(v['xref']))
                if not w: continue
                kind = FIELD_KINDS.get(w.field_type)
                if kind == 'radio': radios[w.xref] = (w.field_name, bool(v['value']))
                else: names[w.field_name] = (kind, v['value'])
        for page in doc:
            for w in list(page.widgets() or []):
                try:
                    if w.xref in radios:
                        w.field_value = w.on_state() if radios[w.xref][1] else 'Off'; w.update(); continue
                    if w.field_name not in names or w.field_type == fitz.PDF_WIDGET_TYPE_RADIOBUTTON: continue
                    kind, value = names[w.field_name]
                    w.field_value = (w.on_state() if value else 'Off') if kind == 'checkbox' else str(value)
                    _draw_widget(w, doc)
                except Exception as e: print('field skipped:', w.field_name, e, file=sys.stderr)
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

@app.route('/redact', methods=['POST'])
def redact():
    """Black out areas for good: the text, pictures and drawings under each box are removed from the file
    (not just covered), then the box is filled black. boxes: [{page, rect}] in points as shown."""
    d = request.json; fp = d['filepath']
    try:
        doc = fitz.open(fp)
        n = 0
        pages = {}
        for b in d.get('boxes', []): pages.setdefault(int(b['page']), []).append((b['rect'], (b.get('reason') or '').strip()))
        for pno, items in pages.items():
            page = doc[pno]
            for r, reason in items:
                box = fitz.Rect(r) * page.derotation_matrix
                if reason:   # the reason code (like (b)(6)) printed in white on the box, as agencies and courts do
                    size = max(4, min(10, box.height * 0.62, (box.width - 2) / max(1, fitz.get_text_length(reason, fontname='helv', fontsize=1))))
                    page.add_redact_annot(box, text=reason, fontname='helv', fontsize=size, align=fitz.TEXT_ALIGN_CENTER, fill=(0, 0, 0), text_color=(1, 1, 1))
                else: page.add_redact_annot(box, fill=(0, 0, 0))
                n += 1
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_PIXELS, graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED)   # lines just touching a box stay
        left = 0   # proof: no text is left under any box (except the reason codes printed on them)
        reasons = {reason for items in pages.values() for _, reason in items if reason}
        for pno, items in pages.items():
            page = doc[pno]
            for r, reason in items:
                area = fitz.Rect(r) * page.derotation_matrix
                left += sum(1 for w in page.get_text('words') if (fitz.Rect(w[:4]) & area).get_area() > 0.3 * max(fitz.Rect(w[:4]).get_area(), 0.01)
                            and not any(w[4] in rs for rs in reasons))
        if left: doc.close(); return jsonify({"error": f"{left} words were still found under the boxes, so nothing was saved. Please tell us about this file."}), 500
        if d.get('scrub'):   # hidden information: document details, comments, attachments, scripts, old versions
            doc.scrub(metadata=True, xml_metadata=True, attached_files=True, embedded_files=True, javascript=True,
                      remove_links=False, reset_fields=False, reset_responses=True, thumbnails=True, hidden_text=True)
        new_path = os.path.join(TEMP_DIR, f"mod_{hashlib.md5((fp + str(os.path.getmtime(fp)) + json.dumps(d.get('boxes'))).encode()).hexdigest()}.pdf")
        doc.save(new_path, garbage=4, deflate=True, clean=True); doc.close()
        return jsonify({"success": True, "new_file": new_path, "message": f"Blacked out {n} {'area' if n == 1 else 'areas'} for good. Checked: nothing readable is left under them"})
    except Exception as e: return jsonify({"error": str(e)}), 500

STAMPS = {'Approved': fitz.STAMP_Approved, 'Draft': fitz.STAMP_Draft, 'Confidential': fitz.STAMP_Confidential, 'Final': fitz.STAMP_Final,
          'Not approved': fitz.STAMP_NotApproved, 'For comment': fitz.STAMP_ForComment, 'Experimental': fitz.STAMP_Experimental,
          'Expired': fitz.STAMP_Expired, 'As is': fitz.STAMP_AsIs, 'Departmental': fitz.STAMP_Departmental, 'For public release': fitz.STAMP_ForPublicRelease,
          'Not for public release': fitz.STAMP_NotForPublicRelease, 'Sold': fitz.STAMP_Sold, 'Top secret': fitz.STAMP_TopSecret}

@app.route('/annotate', methods=['POST'])
def annotate():
    """Add a comment or markup as a standard PDF annotation (other readers show it too): a note, a box,
    an arrow, struck-out or underlined words, or a stamp. Points are as shown on screen."""
    d = request.json; fp = d['filepath']
    try:
        doc = fitz.open(fp); page = doc[int(d['page'])]
        un = page.derotation_matrix
        color = parse_color(d.get('color') or '#e03c3c')
        kind, text = d['kind'], (d.get('text') or '').strip()
        author = d.get('author') or os.environ.get('USERNAME', '')
        if kind == 'note':
            a = page.add_text_annot(fitz.Point(d['p1']) * un, text or 'Note', icon='Comment')
        elif kind == 'box':
            a = page.add_rect_annot(fitz.Rect(d['rect']) * un); a.set_border(width=1.5)
        elif kind == 'arrow':
            a = page.add_line_annot(fitz.Point(d['p1']) * un, fitz.Point(d['p2']) * un)
            a.set_line_ends(fitz.PDF_ANNOT_LE_NONE, fitz.PDF_ANNOT_LE_CLOSED_ARROW); a.set_border(width=1.8)
        elif kind in ('strike', 'underline'):
            area = fitz.Rect(d['rect']) * un
            words = [fitz.Rect(w[:4]) for w in page.get_text('words') if fitz.Rect(w[:4]).intersects(area)]
            if not words: doc.close(); return jsonify({"error": "There's no text there. Drag across the words"}), 400
            a = (page.add_strikeout_annot if kind == 'strike' else page.add_underline_annot)(words)
        elif kind == 'stamp':
            c = fitz.Point(d['p1']) * un
            a = page.add_stamp_annot(fitz.Rect(c.x - 80, c.y - 24, c.x + 80, c.y + 24), stamp=STAMPS.get(d.get('stamp'), fitz.STAMP_Approved))
        else:
            doc.close(); return jsonify({"error": "Unknown kind"}), 400
        if kind not in ('stamp',): a.set_colors(stroke=color)
        a.set_info(title=author, content=text or a.info.get('content', ''))
        a.update()
        new_path = os.path.join(TEMP_DIR, f"mod_{hashlib.md5((fp + str(os.path.getmtime(fp)) + json.dumps(d)).encode()).hexdigest()}.pdf")
        doc.save(new_path, garbage=3, deflate=True); doc.close()
        return jsonify({"success": True, "new_file": new_path})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/delete_annot', methods=['POST'])
def delete_annot():
    d = request.json; fp = d['filepath']
    try:
        doc = fitz.open(fp); page = doc[int(d['page'])]
        for a in page.annots() or []:
            if a.xref == int(d['xref']): page.delete_annot(a); break
        new_path = os.path.join(TEMP_DIR, f"mod_{hashlib.md5((fp + str(os.path.getmtime(fp)) + 'del' + str(d['xref'])).encode()).hexdigest()}.pdf")
        doc.save(new_path, garbage=3, deflate=True); doc.close()
        return jsonify({"success": True, "new_file": new_path})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- fill a template from a spreadsheet ----------
PLACEHOLDER = re.compile(r'\{\{\s*([^{}]+?)\s*\}\}|<<\s*([^<>]+?)\s*>>')
_norm_key = lambda t: re.sub(r'[^a-z0-9]', '', str(t).lower())

def _read_table(path):
    """(headers, rows as dicts) from the first sheet of an .xlsx/.xlsm or a .csv/.tsv; the first row with
    any text is the header row."""
    low = path.lower()
    if low.endswith(('.csv', '.tsv', '.txt')):
        import csv
        raw = open(path, 'rb').read()
        for enc in ('utf-8-sig', 'cp1252', 'latin-1'):
            try: text = raw.decode(enc); break
            except UnicodeDecodeError: continue
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=',;\t|') if text.strip() else csv.excel
        rows = [r for r in csv.reader(text.splitlines(), dialect) if any(c.strip() for c in r)]
    elif low.endswith(('.xlsx', '.xlsm')):
        from openpyxl import load_workbook
        ws = load_workbook(path, data_only=True, read_only=True).worksheets[0]
        def cell(v):
            if v is None: return ''
            if hasattr(v, 'strftime'): return v.strftime('%m/%d/%Y')
            if isinstance(v, float) and v.is_integer(): return str(int(v))
            return str(v)
        rows = [[cell(v) for v in r] for r in ws.iter_rows(values_only=True) if any(v not in (None, '') for v in r)]
    else:
        raise ValueError("Use an .xlsx or .csv file (save older .xls files as .xlsx first)")
    if not rows: raise ValueError("The spreadsheet is empty")
    headers = [h.strip() for h in rows[0]]
    return headers, [{h: (r[i].strip() if i < len(r) else '') for i, h in enumerate(headers) if h} for r in rows[1:]]

def _template_spots(doc):
    """What the template has to fill: {{Name}} / <<Name>> spots in its text, and its form fields."""
    names, fields = set(), set()
    for page in doc:
        for m in PLACEHOLDER.finditer(page.get_text()): names.add((m.group(1) or m.group(2)).strip())
        for w in page.widgets() or []: fields.add(w.field_name)
    return names, fields

def _fill_one(template, row, lookup):
    """A copy of the template with this row's values in its {{...}} spots and matching form fields."""
    doc = fitz.open(template)
    value = lambda name: row.get(lookup.get(_norm_key(name), ''), '')
    for page in doc:
        edits = []
        for sp in _page_spans(page):
            if PLACEHOLDER.search(sp['text']):
                new = PLACEHOLDER.sub(lambda m: value(m.group(1) or m.group(2)), sp['text'])
                edits.append({'rect': list(sp['bbox']), 'base_y': sp['origin'][1], 'new_text': new if new.strip() else ' '})
        if edits: _rewrite_text(doc, page, edits)
        for w in page.widgets() or []:
            key = lookup.get(_norm_key(w.field_name))
            if key is None: continue
            v = row.get(key, '')
            if w.field_type in (fitz.PDF_WIDGET_TYPE_CHECKBOX, fitz.PDF_WIDGET_TYPE_RADIOBUTTON):
                w.field_value = w.on_state() if v.strip().lower() in ('1', 'yes', 'y', 'true', 'x', 'checked', 'on') else 'Off'
            else: w.field_value = v
            w.update()
    return doc

@app.route('/template_info', methods=['POST'])
def template_info():
    d = request.json
    try:
        headers, rows = _read_table(d['data'])
        doc = fitz.open(d['template']); names, fields = _template_spots(doc); doc.close()
        keys = {_norm_key(h) for h in headers}
        used = sorted({n for n in names | fields if _norm_key(n) in keys})
        missing = sorted({n for n in names if _norm_key(n) not in keys})
        return jsonify({"headers": headers, "rows": len(rows), "used": used, "missing": missing,
                        "spots": len(names) + len(fields)})
    except Exception as e: return jsonify({"error": str(e)}), 400

# ---------- box labels from a packing list ----------
# A customer's label PDF has one page per size with an empty quantity box. The packing list says how many
# cartons of each size there are and how many pieces go in each. This makes one label per carton, in carton
# order, with only the quantity filled in (nothing else on the label is touched), ready to print in one go.
def _sheet_rows(path):
    """Every row of the first sheet as a list of values (.xlsx/.xlsm, .xls through Excel, .csv)."""
    low = path.lower()
    if low.endswith(('.csv', '.tsv')):
        import csv
        with open(path, newline='', encoding='utf-8-sig', errors='replace') as f:
            return [r for r in csv.reader(f, delimiter='\t' if low.endswith('.tsv') else ',')]
    if low.endswith('.xls'):
        try:
            import xlrd
            sh = xlrd.open_workbook(path).sheet_by_index(0)
            return [[None if c.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) else c.value for c in sh.row(r)] for r in range(sh.nrows)]
        except Exception: pass   # not readable directly: let Excel open it
        tmp = os.path.join(TEMP_DIR, 'pl_' + hashlib.md5(path.encode()).hexdigest() + '.xlsx')
        def job(o):
            wb = o.get('excel').Workbooks.Open(path, ReadOnly=True, AddToMru=False)
            try: wb.SaveAs(tmp, 51)
            finally: wb.Close(False)
        OFFICE.run(job); path = tmp
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    return [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]

def _num(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool): return v
    try: return float(str(v).replace(',', '').strip())
    except Exception: return None

def _norm(v): return re.sub(r'\s+', ' ', str(v or '')).strip().upper()

def _read_packing_list(path):
    rows = _sheet_rows(path)
    def after(label):   # the first filled cell to the right of a label like "STYLE:"
        for r in rows:
            for i, v in enumerate(r):
                if _norm(v).rstrip(':').strip() == label:
                    for w in r[i + 1:]:
                        if w not in (None, ''): return w
        return None
    head = next((i for i, r in enumerate(rows) if any(_norm(v) == 'COLOR' for v in r)), None)
    if head is None: raise ValueError("couldn't find the COLOR heading")
    hr = rows[head]
    c_color = next(i for i, v in enumerate(hr) if _norm(v) == 'COLOR')
    end = next((i for i, v in enumerate(hr) if i > c_color and re.search(r'UNITS|TOTAL', _norm(v))), len(hr))
    sizes = {i: _norm(v) for i, v in enumerate(hr) if c_color < i < end and v not in (None, '')}
    if not sizes: raise ValueError("couldn't find the size columns")
    sub = rows[head + 1] if head + 1 < len(rows) else []
    def col(word, default):
        for r in (hr, sub):
            for i, v in enumerate(r):
                if _norm(v) == word: return i
        return default
    c_from, c_to = col('FROM', 0), col('TO', 1)
    cartons, started = [], False
    for r in rows[head + 1:]:
        a = _num(r[c_from]) if c_from < len(r) else None
        if a is None:
            if started: break
            continue
        started = True
        b = _num(r[c_to]) if c_to < len(r) else None
        b = a if b is None else b
        color = _norm(r[c_color]) if c_color < len(r) else ''
        packs = [(sz, int(_num(r[i]))) for i, sz in sizes.items() if i < len(r) and (_num(r[i]) or 0) > 0]
        for n in range(int(a), int(b) + 1):
            cartons.append({'carton': n, 'color': color, 'packs': packs})
    total = _num(after('TOTAL UNITS'))
    po = after('PO NUMBER')
    if isinstance(po, float) and po.is_integer(): po = int(po)   # Excel keeps numbers as 1086173.0
    return {'style': _norm(after('STYLE')), 'po': re.sub(r'\D', '', str(po or '')).lstrip('0'),
            'cartons': cartons, 'total_units': int(total) if total else None}

def _label_pages(doc):
    """What each label page is for: its size, color, style and PO (read from the words printed on it)."""
    out = []
    for i, page in enumerate(doc):
        words = page.get_text('words')
        def right_of(label):
            w0 = next((w for w in words if _norm(w[4]).rstrip(':') == label), None)
            if not w0: return ''
            cy = (w0[1] + w0[3]) / 2
            line = sorted((w for w in words if w[0] > w0[2] - 1 and w[1] < cy < w[3]), key=lambda w: w[0])
            return _norm(' '.join(w[4] for w in line))
        text = page.get_text()
        po = next((re.sub(r'\D', '', m).lstrip('0') for m in re.findall(r'PO\s*#?:?\s*0*\d{5,}', text, re.I)), '')
        boxes = [w for w in page.widgets() if w.field_type == fitz.PDF_WIDGET_TYPE_TEXT]
        qty = [w for w in boxes if re.search(r'QTY|QUANT|PCS|PIECES|UNITS', w.field_name or '', re.I)] or (boxes if len(boxes) == 1 else [])
        out.append({'page': i, 'size': right_of('SIZE'), 'color': right_of('COLOR'), 'style': right_of('STYLE'),
                    'po': po, 'qty_field': qty[0].field_name if qty else None})
    return out

def _qty_box(page, w):
    """The box the quantity goes in, as it looks on the label: the drawn frame around the field (a frame
    that runs off the page edge counts to the edge), inside its border. Without one, the field itself."""
    c = (w.rect.tl + w.rect.br) / 2
    best = None
    for d in page.get_drawings():
        r = d['rect'] & page.rect
        if d.get('color') is None or not r.contains(c) or r.width > page.rect.width * 0.8: continue
        if best is None or r.get_area() < best[0].get_area(): best = (r, d.get('width') or 1)
    if not best: return w.rect
    r, bw = best
    x0 = r.x0 + bw / 2 if r.x0 > page.rect.x0 + 1 else r.x0
    y0 = r.y0 + bw / 2 if r.y0 > page.rect.y0 + 1 else r.y0
    x1 = r.x1 - bw / 2 if r.x1 < page.rect.x1 - 1 else r.x1
    y1 = r.y1 - bw / 2 if r.y1 < page.rect.y1 - 1 else r.y1
    return fitz.Rect(x0, y0, x1, y1)

def _write_centered(page, box, text, size):
    """Text in Helvetica, centered both ways in the box (by the height of the digits, not the font's line)."""
    size = min(size, box.width * 0.8 / max(fitz.get_text_length(text, 'helv', 1), 0.01), box.height * 0.7 / 0.72)
    tw = fitz.get_text_length(text, 'helv', size)
    cx, cy = (box.x0 + box.x1) / 2, (box.y0 + box.y1) / 2
    page.insert_text((cx - tw / 2, cy + 0.716 * size / 2), text, fontname='helv', fontsize=size, color=(0, 0, 0))

@app.route('/box_labels', methods=['POST'])
def box_labels():
    data = request.json or {}
    src, orig, plist = data.get('filepath'), data.get('original_path') or data.get('filepath'), data.get('packing_list')
    if not src or not os.path.exists(src): return jsonify({'error': 'Open the label PDF first'}), 400
    if not plist or not os.path.exists(plist): return jsonify({'error': 'Pick the packing list'}), 400
    try: pl = _read_packing_list(plist)
    except Exception as e: return jsonify({'error': f"Couldn't read the packing list: {e}"}), 400
    if not pl['cartons']: return jsonify({'error': 'No cartons found in the packing list'}), 400
    labels = _label_pages(fitz.open(src))
    warnings = []
    if not any(l['qty_field'] for l in labels): return jsonify({'error': 'This label has no quantity box to fill in'}), 400
    lp = {l['po'] for l in labels if l['po']}
    if pl['po'] and lp and pl['po'] not in lp: warnings.append(f"The labels are for PO {', '.join(sorted(lp))} but the packing list is PO {pl['po']}")
    ls = {l['style'] for l in labels if l['style']}
    if pl['style'] and ls and not any(pl['style'] in x or x in pl['style'] for x in ls): warnings.append(f"The labels are for style {', '.join(sorted(ls))} but the packing list is {pl['style']}")
    def page_for(size, color):
        same = [l for l in labels if l['size'] == size]
        return next((l for l in same if not color or not l['color'] or l['color'] == color), same[0] if same else None)
    plan, missing = [], set()
    for c in pl['cartons']:
        for size, q in c['packs']:
            l = page_for(size, c['color'])
            if l is None: missing.add(size); continue
            if not l['qty_field']: warnings.append(f"The {size} label has no quantity box"); continue
            plan.append((c['carton'], size, q, l))
    if missing: warnings.append(f"No label page for size {', '.join(sorted(missing))}: those cartons were left out")
    if not plan: return jsonify({'error': 'None of the sizes in the packing list have a label page'}), 400
    made = sum(q for _, _, q, _ in plan)
    if pl['total_units'] and made != pl['total_units'] and not missing:
        warnings.append(f"The labels add up to {made} pieces but the packing list says {pl['total_units']}")
    out = fitz.open()
    for carton, size, q, l in plan:
        one = fitz.open(src); one.select([l['page']])
        page = one[0]
        for w in list(page.widgets()):
            if w.field_name == l['qty_field']:
                box, fs = _qty_box(page, w), w.text_fontsize or 48
                page.delete_widget(w)
                _write_centered(page, box, str(q), fs)   # the number becomes part of the label, same in every viewer and printer
        out.insert_pdf(one); one.close()
    out.set_toc([[1, f'Carton {c}: {s} x {q}', i + 1] for i, (c, s, q, _) in enumerate(plan)])
    base = os.path.splitext(os.path.basename(orig))[0]
    name = f'{base} - {len(plan)} box labels.pdf'
    # kept aside until the user picks Save (they look through it first; Print works on it as is)
    folder = os.path.join(TEMP_DIR, 'labels_' + hashlib.md5((orig + plist + str(_time.time())).encode()).hexdigest()[:10])
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, name)
    out.save(path, garbage=3, deflate=True); out.close()
    summary = {}
    for _, s, q, _ in plan: summary.setdefault(s, []).append(q)
    return jsonify({'success': True, 'path': path, 'save_to': os.path.join(os.path.dirname(orig), name), 'labels': len(plan), 'cartons': len({c for c, *_ in plan}), 'units': made,
                    'summary': [[s, qs] for s, qs in summary.items()], 'warnings': warnings})

@app.route('/fill_template', methods=['POST'])
def fill_template():
    """One filled PDF per spreadsheet row (named by a column), or all of them in one PDF."""
    d = request.json
    try:
        headers, rows = _read_table(d['data'])
        lookup = {_norm_key(h): h for h in headers}
        base = os.path.splitext(d.get('template_name') or os.path.basename(d['template']))[0]
        folder = os.path.dirname(d.get('original_path') or d['template'])
        name_col = d.get('name_col') or headers[0]
        if d.get('combine'):
            out_doc = fitz.open()
            for row in rows:
                one = _fill_one(d['template'], row, lookup); out_doc.insert_pdf(one); one.close()
            out = get_unique_filename(os.path.join(folder, f'{base} - all {len(rows)} filled.pdf'))
            try: out_doc.subset_fonts()
            except Exception: pass
            out_doc.save(out, garbage=3, deflate=True)
            return jsonify({"success": True, "path": out, "message": f"Filled {len(rows)} copies into {os.path.basename(out)}"})
        out_dir = get_unique_filename(os.path.join(folder, f'{base} - filled'))
        os.makedirs(out_dir)
        for i, row in enumerate(rows, 1):
            label = re.sub(r'[\\/:*?"<>|]+', '-', row.get(name_col, '').strip())[:80] or f'Row {i}'
            one = _fill_one(d['template'], row, lookup)
            try: one.subset_fonts()
            except Exception: pass
            one.save(get_unique_filename(os.path.join(out_dir, f'{base} - {label}.pdf')), garbage=3, deflate=True); one.close()
        return jsonify({"success": True, "path": out_dir, "message": f"Made {len(rows)} filled PDFs in the folder {os.path.basename(out_dir)}"})
    except Exception as e: return jsonify({"error": str(e)}), 500

WATERMARK_TURN = -1   # a page turned 90 degrees needs its watermark turned back the other way (checked on turned pages)

def _gs_table(doc, page):
    """Where the page's ExtGState (see-through settings) table lives: (xref holding it, key path to it)."""
    t, v = doc.xref_get_key(page.xref, 'Resources')
    holder, path = (int(v.split()[0]), 'ExtGState') if t == 'xref' else (page.xref, 'Resources/ExtGState')
    t2, v2 = doc.xref_get_key(holder, path)
    if t2 == 'xref': return int(v2.split()[0]), ''
    return holder, path

def _gs_names(doc, page):
    holder, path = _gs_table(doc, page)
    if not path: return set(doc.xref_get_keys(holder))
    t, v = doc.xref_get_key(holder, path)
    return set(re.findall(r'/([^\s/<>\[\]()]+)\s*(?:<<|\d+ 0 R)', v)) if t == 'dict' else set()

def _watermark_outline(W, H, d, rgb, opacity):
    """A one-page PDF the size of the page as shown, holding the watermark text turned into outlines."""
    import math
    text = (d.get('text') or 'CONFIDENTIAL').strip()
    diag = d.get('angle', 'diagonal') == 'diagonal'
    scale = float(d['scale']) if d.get('scale') else {'small': .35, 'medium': .6, 'large': .85}.get(d.get('size'), .6)
    room = (math.hypot(W, H) if diag else W) * scale
    size = min(room / max(1, fitz.get_text_length(text, fontname='hebo', fontsize=1)), H * 0.6)
    wlen = fitz.get_text_length(text, fontname='hebo', fontsize=size)
    t = fitz.open(); p = t.new_page(width=W, height=H); c = fitz.Point(W / 2, H / 2)
    turn = math.degrees(math.atan2(H, W)) if diag else 0
    p.insert_text(fitz.Point(c.x - wlen / 2, c.y + size * 0.35), text, fontsize=size, fontname='hebo', color=rgb,
                  fill_opacity=opacity, morph=(c, fitz.Matrix(-turn)) if turn else None)
    return fitz.open('pdf', fitz.open('svg', p.get_svg_image(text_as_path=True).encode()).convert_to_pdf())

def _multiply_in(doc, xref, seen):
    """Every transparency setting inside a drawing (and the drawings in it) set to multiply."""
    if xref in seen: return
    seen.add(xref)
    t, v = doc.xref_get_key(xref, 'Resources/ExtGState')
    if t == 'xref':
        gs_x = int(v.split()[0]); t, v = 'dict', doc.xref_object(gs_x); holder, prefix = gs_x, ''
    else: holder, prefix = xref, 'Resources/ExtGState/'
    if t == 'dict':
        for name, ref in re.findall(r'/([^\s/<>\[\]()]+)\s*(\d+ 0 R|<<)', v):
            if ref == '<<': doc.xref_set_key(holder, prefix + name + '/BM', '/Multiply')
            else: doc.xref_set_key(int(ref.split()[0]), 'BM', '/Multiply')
    t, v = doc.xref_get_key(xref, 'Resources/XObject')
    if t == 'xref': t, v = 'dict', doc.xref_object(int(v.split()[0]))
    if t == 'dict':
        for ref in re.findall(r'(\d+) 0 R', v): _multiply_in(doc, int(ref), seen)

def _multiply_new_xobjects(doc, page, before):
    seen = set()
    for x in page.get_xobjects():
        if x[0] not in before: _multiply_in(doc, x[0], seen)

def _multiply_new_gs(doc, page, before):
    """Make the see-through setting the watermark just added (and only it) blend as 'multiply': the
    watermark then only darkens the white of the page, so it looks behind the content even on pages that
    paint a white background."""
    try:
        holder, path = _gs_table(doc, page)
        for name in _gs_names(doc, page) - before:
            key = (path + '/' if path else '') + name
            t, v = doc.xref_get_key(holder, key)
            if t == 'xref': doc.xref_set_key(int(v.split()[0]), 'BM', '/Multiply')
            else: doc.xref_set_key(holder, key + '/BM', '/Multiply')
    except Exception as e: print('blend not set:', e, file=sys.stderr)

@app.route('/watermark', methods=['POST'])
def watermark():
    """A watermark (text or a picture) or a background color on every page of each file.
    files: [paths]; kind: text | image | color; on_top: over the content or behind it."""
    import math
    d = request.json
    try:
        kind = d.get('kind', 'text'); on_top = bool(d.get('on_top', True))
        rgb = parse_color(d.get('color') or '#c0392b'); opacity = max(0.03, min(1, float(d.get('opacity', 30)) / 100))
        img_bytes = None
        if kind == 'image':
            with Image.open(d['image']) as im:   # the picture, made see-through
                im = im.convert('RGBA'); a = im.getchannel('A').point(lambda v: int(v * opacity)); im.putalpha(a)
                buf = BytesIO(); im.save(buf, 'PNG'); img_bytes = buf.getvalue(); iw, ih = im.size
        outs = []; made = {}
        for fp in d['files']:
            doc = fitz.open(fp)
            for page in doc:
                r = page.rect   # as shown
                c = fitz.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)
                if kind == 'color':
                    page.draw_rect(page.mediabox, color=None, fill=rgb, overlay=False, fill_opacity=1)
                elif kind == 'image':
                    w = r.width * (float(d['scale']) if d.get('scale') else {'small': .25, 'medium': .45, 'large': .7}.get(d.get('size'), .45)); h = w * ih / iw
                    box = fitz.Rect(c.x - w / 2, c.y - h / 2, c.x + w / 2, c.y + h / 2) * page.derotation_matrix
                    page.insert_image(box, stream=img_bytes, overlay=on_top, keep_proportion=True)
                else:
                    # drawn as the letters' outlines, not as text: it looks the same but can't be selected, copied
                    # or found by search in any PDF reader. Made once per page size, then placed on each page.
                    key = (round(r.width, 1), round(r.height, 1))
                    if key not in made: made[key] = _watermark_outline(r.width, r.height, d, rgb, opacity)
                    before = {x[0] for x in page.get_xobjects()}
                    cb = page.cropbox
                    page.show_pdf_page(fitz.Rect(0, 0, cb.width, cb.height), made[key], 0, rotate=page.rotation, overlay=True)
                    if not on_top: _multiply_new_xobjects(doc, page, before)   # tint only: the page's own text stays crisp on top
            out = os.path.join(TEMP_DIR, f"mod_{hashlib.md5((fp + str(os.path.getmtime(fp)) + json.dumps(d)).encode()).hexdigest()}.pdf")
            doc.save(out, garbage=3, deflate=True); doc.close(); outs.append(out)
        return jsonify({"success": True, "files": outs})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- small helpers for the routes below ----------
def _temp_out(fp, tag, extra=''):
    """A new working file for a change to fp (unique per input file, change and options)."""
    h = hashlib.md5((fp + str(os.path.getmtime(fp)) + tag + extra).encode()).hexdigest()
    return os.path.join(TEMP_DIR, f"{tag}_{h}.pdf")

def _pick(doc, pages):
    """1-based page list from the app -> sorted 0-based indexes (all pages when empty)."""
    if not pages: return list(range(len(doc)))
    return sorted({int(n) - 1 for n in pages if 1 <= int(n) <= len(doc)})

# ---------- insert pages from another file ----------
@app.route('/insert_pages', methods=['POST'])
def insert_pages():
    """Copy pages of `source` (1-based `pages`, all when empty) into the open PDF before 0-based `at`
    (the end when -1). Links, comments and form fields come along."""
    d = request.json; fp = d['filepath']
    try:
        doc = fitz.open(fp); src = fitz.open(d['source'])
        at = int(d.get('at', -1)); at = len(doc) if at < 0 or at > len(doc) else at
        picked = _pick(src, d.get('pages'))
        for k, i in enumerate(picked):
            doc.insert_pdf(src, from_page=i, to_page=i, start_at=at + k)
        src.close()
        out = _temp_out(fp, 'ins', json.dumps(d))
        doc.save(out, garbage=3, deflate=True); doc.close()
        return jsonify({"success": True, "new_file": out, "count": len(picked)})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- paper size, white margins ----------
PAPER_PT = {'Letter': (612, 792), 'Legal': (612, 1008), 'Tabloid': (792, 1224), 'A3': (841.89, 1190.55),
            'A4': (595.28, 841.89), 'A5': (419.53, 595.28)}

def _content_box(page):
    """Where anything is drawn on the page (text, pictures, lines), in page coordinates; None if blank."""
    pix = page.get_pixmap(dpi=72, alpha=False, annots=True)
    im = Image.frombytes('RGB', (pix.width, pix.height), pix.samples).convert('L')
    box = im.point(lambda v: 255 if v < 245 else 0).getbbox()
    if not box: return None
    k = page.rect.width / pix.width
    return fitz.Rect(box[0] * k, box[1] * k, box[2] * k, box[3] * k)

@app.route('/page_size', methods=['POST'])
def page_size():
    """mode 'paper': put the chosen pages on another paper size (fit or actual size, centered);
    'trim': cut the white margins down to `margin` points; 'margin': add `margin` points of white."""
    d = request.json; fp = d['filepath']; mode = d.get('mode', 'paper')
    margin = max(0.0, float(d.get('margin', 18)))
    try:
        doc = fitz.open(fp)
        picked = set(_pick(doc, d.get('pages')))
        if mode == 'trim':
            n = 0
            for i in picked:
                page = doc[i]
                box = _content_box(page)
                if not box: continue
                box = fitz.Rect(box.x0 - margin, box.y0 - margin, box.x1 + margin, box.y1 + margin) & page.rect
                # the crop box is given unrotated, relative to the media box's top-left
                r = box * page.derotation_matrix
                mb = page.mediabox
                page.set_cropbox(fitz.Rect(r.x0 + page.cropbox.x0 - mb.x0, r.y0 + page.cropbox.y0 - mb.y0,
                                           r.x1 + page.cropbox.x0 - mb.x0, r.y1 + page.cropbox.y0 - mb.y0) & fitz.Rect(0, 0, mb.width, mb.height))
                n += 1
            out = _temp_out(fp, 'trim', json.dumps(d))
            doc.save(out, garbage=3, deflate=True); doc.close()
            return jsonify({"success": True, "new_file": out, "message": f"Trimmed {n} {'page' if n == 1 else 'pages'}"})
        # 'paper' and 'margin' redraw the page onto a new sheet; comments and fields are flattened first
        try: doc.bake()
        except Exception: pass
        new = fitz.open()
        for i, page in enumerate(doc):
            if i not in picked:
                new.insert_pdf(doc, from_page=i, to_page=i); continue
            pw, ph = page.rect.width, page.rect.height
            if mode == 'margin':
                W, H = pw + 2 * margin, ph + 2 * margin
                target = fitz.Rect(margin, margin, margin + pw, margin + ph)
            else:
                W, H = PAPER_PT.get(d.get('paper'), PAPER_PT['Letter'])
                o = d.get('orient', 'auto')
                if o == 'landscape' or (o == 'auto' and pw > ph): W, H = H, W
                f = min(W / pw, H / ph)
                if d.get('fit') == 'actual': f = min(1, f) if d.get('shrink_only') else 1
                tw, th = pw * f, ph * f
                target = fitz.Rect((W - tw) / 2, (H - th) / 2, (W + tw) / 2, (H + th) / 2)
            np_ = new.new_page(width=W, height=H)
            np_.show_pdf_page(target, doc, i)
        try: new.set_toc(doc.get_toc(simple=False))
        except Exception: pass
        out = _temp_out(fp, 'size', json.dumps(d))
        new.save(out, garbage=3, deflate=True); new.close(); doc.close()
        return jsonify({"success": True, "new_file": out})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- bookmarks (the outline) ----------
@app.route('/toc', methods=['POST'])
def get_toc():
    try:
        doc = fitz.open(request.json['filepath'])
        toc = [{"level": t[0], "title": t[1], "page": t[2]} for t in doc.get_toc(simple=True)]
        doc.close()
        return jsonify({"toc": toc})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/set_toc', methods=['POST'])
def set_toc():
    """Replace the bookmarks: [{level, title, page}] (1-based pages, levels start at 1 and go up by one at most)."""
    d = request.json; fp = d['filepath']
    try:
        doc = fitz.open(fp)
        toc, last = [], 0
        for t in d.get('toc', []):
            lvl = max(1, min(int(t['level']), last + 1)); last = lvl
            toc.append([lvl, str(t['title'])[:300] or 'Untitled', max(1, min(int(t['page']), len(doc)))])
        doc.set_toc(toc)
        out = _temp_out(fp, 'toc', json.dumps(toc))
        doc.save(out, garbage=3, deflate=True); doc.close()
        return jsonify({"success": True, "new_file": out})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/auto_toc', methods=['POST'])
def auto_toc():
    """Suggested bookmarks from the headings: short lines set clearly bigger (or bold at a bigger size)
    than the body text; the biggest sizes become the top level."""
    from collections import Counter
    try:
        doc = fitz.open(request.json['filepath'])
        sizes, lines = Counter(), []
        for pno, page in enumerate(doc):
            for b in page.get_text('dict')['blocks']:
                for l in b.get('lines', []):
                    text = ''.join(s['text'] for s in l['spans']).strip()
                    if not text: continue
                    big = max(l['spans'], key=lambda s: len(s['text'].strip()))
                    size = round(big['size'], 1); bold = bool(big['flags'] & 16) or 'bold' in big['font'].lower()
                    sizes[size] += len(text)
                    lines.append((pno + 1, text, size, bold, l['bbox'][1]))
        if not sizes: return jsonify({"toc": []})
        body = sizes.most_common(1)[0][0]
        heads = [(p, t, sz, b) for p, t, sz, b, y in lines
                 if len(t) <= 90 and not t.endswith(('.', ',', ';')) and re.search(r'[A-Za-z]', t)
                 and (sz >= body * 1.18 or (b and sz >= body * 1.05))]
        levels = sorted({h[2] for h in heads}, reverse=True)[:3]
        toc, seen = [], set()
        for p, t, sz, b in heads:
            if sz not in levels: continue
            key = (p, t)
            if key in seen: continue
            seen.add(key)
            if toc and toc[-1]['page'] == p and toc[-1]['level'] == levels.index(sz) + 1 and len(toc[-1]['title']) + len(t) < 120:
                toc[-1]['title'] += ' ' + t; continue   # one heading set on two lines
            toc.append({"level": levels.index(sz) + 1, "title": t, "page": p})
        # a heading level only below a level that exists
        last = 0
        for t in toc: t['level'] = min(t['level'], last + 1); last = t['level']
        doc.close()
        return jsonify({"toc": toc[:500]})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- make it small: aim for a size ----------
SHRINK_STEPS = [None, (200, 85), (150, 75), (120, 65), (96, 55), (72, 45), (60, 35)]   # (picture dpi, JPEG quality)

def _shrink_once(fp, step, out):
    doc = fitz.open(fp)
    try: doc.subset_fonts()
    except Exception: pass
    if step:
        dpi, q = step
        doc.rewrite_images(dpi_threshold=dpi + 15, dpi_target=dpi, quality=q, lossy=True, lossless=True, bitonal=True, color=True, gray=True)
    try: doc.scrub(metadata=False, xml_metadata=False, attached_files=False, embedded_files=False, javascript=False,
                   thumbnails=True, redactions=False, remove_links=False, reset_fields=False, reset_responses=False, hidden_text=False, clean_pages=False)
    except Exception: pass
    doc.save(out, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, clean=True, use_objstms=1)
    doc.close()
    return os.path.getsize(out)

def shrink_pdf(fp, target_bytes=None, level='balanced'):
    """The smallest file that still looks good (level), or the best-looking one under target_bytes.
    Returns (path, size, met_target)."""
    steps = SHRINK_STEPS
    if not target_bytes:
        steps = {'light': SHRINK_STEPS[:1], 'balanced': [SHRINK_STEPS[2]], 'strong': [SHRINK_STEPS[4]]}.get(level, [SHRINK_STEPS[2]])
    best = None
    for k, step in enumerate(steps):
        out = _temp_out(fp, 'shrink', f'{k}{step}{target_bytes}{level}')
        size = _shrink_once(fp, step, out)
        if best is None or size < best[1]: best = (out, size)
        if target_bytes and size <= target_bytes: return out, size, True
    return best[0], best[1], not target_bytes

@app.route('/shrink', methods=['POST'])
def shrink():
    d = request.json; fp = d['filepath']
    try:
        before = os.path.getsize(fp)
        target = float(d['target_mb']) * 1024 * 1024 if d.get('target_mb') else None
        out, size, met = shrink_pdf(fp, target, d.get('level', 'balanced'))
        mb = lambda b: f"{b / 1024 / 1024:.1f} MB" if b >= 1024 * 1024 else f"{max(1, round(b / 1024))} KB"
        if size >= before:
            return jsonify({"success": True, "new_file": fp, "size": before, "before": before,
                            "message": f"This PDF is already as small as it gets ({mb(before)})"})
        msg = f"{mb(before)} → {mb(size)}" + ("" if met else f". That's the smallest it goes without losing the words; still over {d.get('target_mb')} MB")
        return jsonify({"success": True, "new_file": out, "size": size, "before": before, "met": met, "message": msg})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- digital signatures with a certificate ----------
def _trust_roots():
    """The certificates Windows trusts, for checking who signed."""
    import ssl
    from asn1crypto import x509 as ax509
    roots = []
    try:
        for store in ('ROOT', 'CA'):
            for der, enc, trust in ssl.enum_certificates(store):
                if enc == 'x509_asn':
                    try: roots.append(ax509.Certificate.load(der))
                    except Exception: pass
    except Exception: pass
    return roots

@app.route('/cert_make', methods=['POST'])
def cert_make():
    """Make the user's own digital ID: a self-signed certificate saved as a password-protected .pfx."""
    from cryptography import x509
    from cryptography.x509.oid import NameOID, ExtendedKeyUsageOID
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.hazmat.primitives.serialization import pkcs12
    import datetime
    d = request.json
    try:
        if not d.get('name') or not d.get('password'): return jsonify({"error": "Enter your name and a password"}), 400
        key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        attrs = [x509.NameAttribute(NameOID.COMMON_NAME, d['name'])]
        if d.get('org'): attrs.append(x509.NameAttribute(NameOID.ORGANIZATION_NAME, d['org']))
        if d.get('email'): attrs.append(x509.NameAttribute(NameOID.EMAIL_ADDRESS, d['email']))
        name = x509.Name(attrs)
        now = datetime.datetime.now(datetime.timezone.utc)
        cert = (x509.CertificateBuilder().subject_name(name).issuer_name(name).public_key(key.public_key())
                .serial_number(x509.random_serial_number()).not_valid_before(now - datetime.timedelta(minutes=5))
                .not_valid_after(now + datetime.timedelta(days=365 * 5))
                .add_extension(x509.KeyUsage(digital_signature=True, content_commitment=True, key_encipherment=False, data_encipherment=False,
                                             key_agreement=False, key_cert_sign=True, crl_sign=False, encipher_only=False, decipher_only=False), critical=True)
                .add_extension(x509.ExtendedKeyUsage([ExtendedKeyUsageOID.EMAIL_PROTECTION, x509.ObjectIdentifier('1.2.840.113583.1.1.5')]), critical=False)
                .sign(key, hashes.SHA256()))
        data = pkcs12.serialize_key_and_certificates(d['name'].encode(), key, cert, None,
                                                     serialization.BestAvailableEncryption(d['password'].encode()))
        os.makedirs(os.path.dirname(d['out']), exist_ok=True)
        with open(d['out'], 'wb') as f: f.write(data)
        with open(d['out'] + '.cer', 'wb') as f: f.write(cert.public_bytes(serialization.Encoding.DER))   # for checking own signatures
        return jsonify({"success": True, "path": d['out']})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/cert_info', methods=['POST'])
def cert_info():
    """Who a .pfx/.p12 digital ID belongs to (checks the password)."""
    from cryptography.hazmat.primitives.serialization import pkcs12
    from cryptography.x509.oid import NameOID
    d = request.json
    try:
        key, cert, _ = pkcs12.load_key_and_certificates(open(d['pfx'], 'rb').read(), (d.get('password') or '').encode() or None)
        cn = cert.subject.get_attributes_for_oid(NameOID.COMMON_NAME)
        return jsonify({"name": cn[0].value if cn else cert.subject.rfc4514_string(),
                        "issuer": cert.issuer.rfc4514_string(), "expires": cert.not_valid_after_utc.strftime('%b %d, %Y'),
                        "self_signed": cert.issuer == cert.subject})
    except ValueError: return jsonify({"error": "That password isn't right for this digital ID"}), 400
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/sign_cert', methods=['POST'])
def sign_cert():
    """Sign with a digital ID (.pfx/.p12). The signed copy is saved next to the original; any later change
    to it shows up when the signature is checked. rect (on screen, points) gives a visible signature box."""
    from pyhanko.sign import signers, fields
    from pyhanko.stamp import TextStampStyle
    from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
    from pyhanko.pdf_utils.text import TextBoxStyle
    d = request.json; fp = d['filepath']; orig = d.get('original_path') or fp
    try:
        signer = signers.SimpleSigner.load_pkcs12(d['pfx'], passphrase=(d.get('password') or '').encode() or None)
        if signer is None: return jsonify({"error": "That password isn't right for this digital ID"}), 400
        if not os.path.exists(d['pfx'] + '.cer'):
            try:
                with open(d['pfx'] + '.cer', 'wb') as f: f.write(signer.signing_cert.dump())
            except Exception: pass
        # pyHanko signs a plain PDF best: write a clean copy first
        doc = fitz.open(fp)
        existing = sum(1 for p in doc for w in (p.widgets() or []) if w.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE)
        box = None
        if d.get('rect'):
            page = doc[int(d['page'])]
            r = fitz.Rect(d['rect']) * page.derotation_matrix   # unrotated, top-left origin
            mb = page.mediabox
            box = (int(mb.x0 + r.x0), int(mb.y1 - r.y1), int(mb.x0 + r.x1), int(mb.y1 - r.y0))   # PDF space, bottom-left origin
        clean = _temp_out(fp, 'presign')
        doc.save(clean, garbage=3, deflate=True, no_new_id=False); doc.close()
        field = f"Signature{existing + 1}"
        out = get_unique_filename(os.path.join(os.path.dirname(orig), os.path.splitext(os.path.basename(orig))[0] + ' - signed.pdf'))
        meta = signers.PdfSignatureMetadata(field_name=field, reason=d.get('reason') or None, location=d.get('location') or None,
                                            name=d.get('name') or None)
        style = TextStampStyle(stamp_text='Digitally signed by %(signer)s\nDate: %(ts)s' + (f"\nReason: {d['reason']}" if d.get('reason') else ''),
                               text_box_style=TextBoxStyle(), border_width=1)
        with open(clean, 'rb') as inf:
            w = IncrementalPdfFileWriter(inf, strict=False)
            fields.append_signature_field(w, fields.SigFieldSpec(field, on_page=int(d.get('page', 0)), box=box) if box
                                          else fields.SigFieldSpec(field))
            pdf_signer = signers.PdfSigner(meta, signer=signer, stamp_style=style)
            with open(out, 'wb') as outf: pdf_signer.sign_pdf(w, output=outf)
        return jsonify({"success": True, "path": out, "message": f"Signed and saved {os.path.basename(out)}"})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/verify_signatures', methods=['POST'])
def verify_signatures():
    """Every digital signature in the file: who signed, when, whether the file changed since, and whether
    the signer's ID is one Windows trusts (or one of the user's own IDs)."""
    from pyhanko.pdf_utils.reader import PdfFileReader
    from pyhanko.sign.validation import validate_pdf_signature
    from pyhanko_certvalidator import ValidationContext
    from pyhanko.sign.validation.status import SignatureCoverageLevel
    from asn1crypto import x509 as ax509
    d = request.json
    try:
        own = []
        for path in d.get('own_ids', []):   # the user's own IDs (their certificate saved beside each) count as trusted on this PC
            try: own.append(ax509.Certificate.load(open(path + '.cer', 'rb').read()))
            except Exception: pass
        with open(d['filepath'], 'rb') as f:
            r = PdfFileReader(f, strict=False)
            sigs = list(r.embedded_signatures)
            if not sigs: return jsonify({"signatures": []})
            vc = ValidationContext(trust_roots=_trust_roots() + own, allow_fetching=False)
            out = []
            for sig in sigs:
                st = validate_pdf_signature(sig, vc, skip_diff=False)
                cert = sig.signer_cert
                who = cert.subject.native.get('common_name') or cert.subject.human_friendly
                when = sig.self_reported_timestamp
                whole = st.coverage == SignatureCoverageLevel.ENTIRE_FILE
                mods = getattr(st, 'modification_level', None)
                out.append({"field": sig.field_name, "signer": who, "time": when.strftime('%b %d, %Y %I:%M %p') if when else '',
                            "intact": bool(st.intact and st.valid), "trusted": bool(st.trusted),
                            "changed_after": not whole and (mods is None or getattr(mods, 'name', '') not in ('NONE', 'LTA_UPDATES')),
                            "reason": sig.sig_object.get('/Reason', '') and str(sig.sig_object.get('/Reason')),
                            "issuer": cert.issuer.human_friendly})
        return jsonify({"signatures": out})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- the same action on many files ----------
import threading, uuid
BATCH = {}

BATCH_PDF_ACTIONS = ('compress', 'watermark', 'number', 'headfoot', 'ocr', 'rotate', 'scrub')

def _batch_one(client, src, action, opt, out_dir, pdf=None, inplace=False):
    """Run one action on one file through the app's own routes; returns the saved path. inplace (the files
    open in a tab): changes to the PDF come back as a new working copy for the tab instead of a saved file,
    and new files (exports, protected copies) go next to each original."""
    def call(route, body):
        r = client.post(route, json=body); j = r.get_json() or {}
        if r.status_code >= 400 or j.get('error'): raise RuntimeError(j.get('error') or f'{route} failed')
        return j
    base = os.path.splitext(os.path.basename(src))[0]
    if not pdf or not os.path.exists(pdf):
        r = call('/process_file', {'filepath': src})
        if r.get('needs_password'): raise RuntimeError('It has a password')
        pdf = r['pdf_path']
    if inplace: out_dir = os.path.dirname(src)
    def save(path, suffix=''):
        if inplace and action in BATCH_PDF_ACTIONS: return path
        dst = get_unique_filename(os.path.join(out_dir, base + suffix + '.pdf'))
        shutil.copyfile(path, dst); return dst
    if action == 'convert': return save(pdf)
    if action == 'compress':
        return save(call('/shrink', {'filepath': pdf, 'target_mb': opt.get('target_mb'), 'level': opt.get('level', 'balanced')})['new_file'])
    if action == 'watermark':
        j = call('/watermark', {'files': [pdf], 'kind': 'text', 'text': opt.get('text') or 'CONFIDENTIAL', 'color': opt.get('color') or '#c0392b',
                                'opacity': opt.get('opacity', 30), 'size': 'medium', 'angle': 'diagonal', 'on_top': True})
        return save(j['files'][0])
    if action == 'number': return save(call('/number_pages', {'filepath': pdf, 'start_num': 1})['new_file'])
    if action == 'headfoot': return save(call('/header_footer', {'filepath': pdf, 'header': opt.get('header', ''), 'footer': opt.get('footer', '')})['new_file'])
    if action == 'ocr': return save(call('/ocr', {'filepath': pdf}).get('new_file') or pdf)
    if action == 'rotate':
        doc = fitz.open(pdf)
        for pg in doc: pg.set_rotation((pg.rotation + int(opt.get('angle', 90))) % 360)
        tmp = _temp_out(pdf, 'brot', str(opt)); doc.save(tmp, garbage=3); doc.close(); return save(tmp)
    if action == 'scrub':
        doc = fitz.open(pdf); doc.scrub(); doc.set_metadata({})
        tmp = _temp_out(pdf, 'bscrub'); doc.save(tmp, garbage=4, deflate=True, clean=True); doc.close(); return save(tmp)
    if action == 'protect':
        if not opt.get('password'): raise RuntimeError('No password given')
        doc = fitz.open(pdf); dst = get_unique_filename(os.path.join(out_dir, base + ('_Protected' if inplace else '') + '.pdf'))
        doc.save(dst, encryption=fitz.PDF_ENCRYPT_AES_256, user_pw=opt['password'], owner_pw=opt['password']); doc.close(); return dst
    if action in ('docx', 'xlsx', 'pptx', 'png', 'txt', 'html'):
        j = call('/export', {'filepath': pdf, 'original_path': os.path.join(out_dir, os.path.basename(src)), 'format': action})
        return j.get('path') or out_dir
    raise RuntimeError('Unknown action')

def _batch_run(job, files, action, opt, out_dir, pdfs=None, inplace=False):
    client = app.test_client()
    for i, f in enumerate(files):
        if job['cancel']: break
        job['current'] = os.path.basename(f)
        try:
            out = _batch_one(client, f, action, opt, out_dir, (pdfs or [None] * len(files))[i], inplace)
            job['outputs'].append(out); job['results'][i] = out
        except Exception as e: job['errors'].append(f"{os.path.basename(f)}: {e}")
        job['done'] = i + 1
    job['finished'] = True

@app.route('/batch_start', methods=['POST'])
def batch_start():
    d = request.json
    if d.get('items'):   # the files open in a tab, each with its working copy
        items = [it for it in d['items'] if it.get('src')]
        jid = uuid.uuid4().hex
        BATCH[jid] = job = {"total": len(items), "done": 0, "current": '', "errors": [], "outputs": [], "results": [None] * len(items),
                            "finished": False, "cancel": False, "out_dir": None, "inplace": True}
        threading.Thread(target=_batch_run, args=(job, [it['src'] for it in items], d['action'], d.get('options') or {}, None,
                                                  [it.get('pdf') for it in items], True), daemon=True).start()
        return jsonify({"id": jid, "total": len(items), "out_dir": None})
    first = (d.get('files') or [''])[0]
    out_dir = d.get('out_dir') or os.path.join(first if os.path.isdir(first) else os.path.dirname(first), 'MyPDFKit output')
    skip = os.path.normcase(os.path.abspath(out_dir))
    files = []
    for f in d.get('files', []):
        if os.path.isdir(f):
            for root, _, names in os.walk(f):
                if d.get('subfolders') is False and root != f: continue
                if os.path.normcase(os.path.abspath(root)).startswith(skip) or 'MyPDFKit output' in root: continue
                files += [os.path.join(root, n) for n in sorted(names) if os.path.splitext(n)[1].lower() in OPENABLE_EXT]
        elif os.path.isfile(f): files.append(f)
    if not files: return jsonify({"error": "There are no files there that MyPDFKit can open"}), 400
    os.makedirs(out_dir, exist_ok=True)
    jid = uuid.uuid4().hex
    BATCH[jid] = job = {"total": len(files), "done": 0, "current": '', "errors": [], "outputs": [], "results": [None] * len(files), "finished": False, "cancel": False, "out_dir": out_dir}
    threading.Thread(target=_batch_run, args=(job, files, d['action'], d.get('options') or {}, out_dir), daemon=True).start()
    return jsonify({"id": jid, "total": len(files), "out_dir": out_dir})

@app.route('/batch_status', methods=['POST'])
def batch_status():
    job = BATCH.get(request.json['id'])
    if not job: return jsonify({"error": "No such job"}), 404
    if request.json.get('cancel'): job['cancel'] = True
    return jsonify({k: v for k, v in job.items() if k != 'outputs'} | {"saved": len(job['outputs'])})

# ---------- measurements kept on the page ----------
@app.route('/add_measure', methods=['POST'])
def add_measure():
    """A measurement as standard annotations: the line or outline plus its label (other readers show it)."""
    d = request.json; fp = d['filepath']
    try:
        doc = fitz.open(fp); page = doc[int(d['page'])]; un = page.derotation_matrix
        pts = [fitz.Point(p) * un for p in d['points']]
        color = parse_color(d.get('color') or '#e8527a')
        if d['kind'] == 'distance':
            a = page.add_line_annot(pts[0], pts[-1]); a.set_line_ends(fitz.PDF_ANNOT_LE_BUTT, fitz.PDF_ANNOT_LE_BUTT)
        elif d['kind'] == 'area':
            a = page.add_polygon_annot(pts)
        else:
            a = page.add_polyline_annot(pts)
        a.set_colors(stroke=color); a.set_border(width=1.5)
        a.set_info(title=os.environ.get('USERNAME', ''), content=d.get('label', ''), subject='Measurement'); a.update()
        lp = fitz.Point(d['label_at']) * un
        t = page.add_freetext_annot(fitz.Rect(lp.x, lp.y - 9, lp.x + 7 * len(d.get('label', '')) + 12, lp.y + 7), d.get('label', ''),
                                    fontsize=9, text_color=color, fill_color=(1, 1, 1), rotate=page.rotation)
        t.set_info(subject='Measurement'); t.update()
        out = _temp_out(fp, 'meas', json.dumps(d))
        doc.save(out, garbage=3, deflate=True); doc.close()
        return jsonify({"success": True, "new_file": out})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- pens whose line gets thicker and thinner ----------
def _ink_outline(P, W):
    """The outline of a stroke through points P with widths W, round at both ends (as the app draws it)."""
    import math
    n = len(P)
    if n == 1:
        x, y = P[0]; r = W[0] / 2
        return [(x + r * math.cos(i * math.pi / 6), y + r * math.sin(i * math.pi / 6)) for i in range(12)]
    L, R = [], []
    for i in range(n):
        a, b = P[max(0, i - 1)], P[min(n - 1, i + 1)]
        dx, dy = b[0] - a[0], b[1] - a[1]; ln = math.hypot(dx, dy) or 1; dx /= ln; dy /= ln; h = W[i] / 2
        L.append((P[i][0] - dy * h, P[i][1] + dx * h)); R.append((P[i][0] + dy * h, P[i][1] - dx * h))
    cap = lambda c, h, frm: [(c[0] + h * math.cos(frm - (k + 1) * math.pi / 6), c[1] + h * math.sin(frm - (k + 1) * math.pi / 6)) for k in range(5)]
    ang = lambda a, b: math.atan2(b[1] - a[1], b[0] - a[0])
    return L + cap(P[-1], W[-1] / 2, ang(P[-2], P[-1]) + math.pi / 2) + R[::-1] + cap(P[0], W[0] / 2, ang(P[0], P[1]) - math.pi / 2)

def _add_pen_shape(doc, page, outline, rgb, opacity, center=None, widths=None, info=None, tip=None):
    """A filled shape of ink; the stroke's centre line and widths are kept with it so the eraser can cut it."""
    a = page.add_polygon_annot(outline)
    a.set_colors(stroke=rgb, fill=rgb); a.set_border(width=0); a.set_opacity(opacity)
    a.set_info(title=(info or {}).get('title') or os.environ.get('USERNAME', ''), subject='Pen')
    a.update()
    if center and widths:
        data = json.dumps({"p": [[round(x, 2), round(y, 2)] for x, y in center], "w": [round(w, 3) for w in widths], "t": tip or 'fountain'})
        doc.xref_set_key(a.xref, 'MyPDFKitInk', fitz.get_pdf_str(data))
    return a

# ---------- shapes ----------
LINE_SHAPES = {'line': (fitz.PDF_ANNOT_LE_NONE, fitz.PDF_ANNOT_LE_NONE), 'arrow': (fitz.PDF_ANNOT_LE_NONE, fitz.PDF_ANNOT_LE_CLOSED_ARROW),
               'darrow': (fitz.PDF_ANNOT_LE_CLOSED_ARROW, fitz.PDF_ANNOT_LE_CLOSED_ARROW)}

def _add_shape(doc, page, d, x_off=0, y_off=0):
    pts = [(float(p[0]) + x_off, float(p[1]) + y_off) for p in d['points']]
    rgb = parse_color(d.get('color') or '#111111')
    fill = parse_color(d['fill']) if d.get('fill') else None
    if d['kind'] in LINE_SHAPES:
        a = page.add_line_annot(pts[0], pts[-1])
        a.set_line_ends(*LINE_SHAPES[d['kind']])
        a.set_colors(stroke=rgb, fill=rgb)
    else:
        if len(pts) > 2 and pts[0] == pts[-1]: pts = pts[:-1]
        a = page.add_polygon_annot(pts)
        a.set_colors(stroke=rgb, fill=fill)
    a.set_border(width=max(0.25, float(d.get('size') or 1)))
    a.set_opacity(float(d.get('opacity', 1)))
    a.set_info(subject='Shape', title=os.environ.get('USERNAME', ''))
    a.update()
    keep = {k: d.get(k) for k in ('kind', 'box', 'color', 'fill', 'size', 'opacity', 'rot', 'text', 'textColor', 'textSize', 'group', 'id')}
    doc.xref_set_key(a.xref, 'MyPDFKitShape', fitz.get_pdf_str(json.dumps(keep)))
    if (d.get('text') or '').strip() and d['kind'] not in LINE_SHAPES:
        _shape_text(doc, page, d, x_off, y_off)
    return a

def _shape_text(doc, page, d, x_off=0, y_off=0):
    """The words inside a shape: a borderless text box centered in it (upright, as the app shows it)."""
    x0, y0, x1, y1 = d['box']; x0, x1 = sorted((x0, x1)); y0, y1 = sorted((y0, y1))
    size = float(d.get('textSize') or 12); pad = max(4, (x1 - x0) * 0.06)
    width = max(10, x1 - x0 - 2 * pad)
    lines = 0
    for para in d['text'].split('\n'):   # how many lines it wraps to, to center it up and down
        n, cur = 1, 0.0
        for w in para.split(' '):
            lw = fitz.get_text_length(w + ' ', fontname='helv', fontsize=size)
            if cur and cur + lw > width: n += 1; cur = lw
            else: cur += lw
        lines += n
    h = lines * size * 1.25 + 4
    cy = (y0 + y1) / 2
    r = fitz.Rect(x0 + pad + x_off, cy - h / 2 + y_off, x1 - pad + x_off, cy + h / 2 + y_off)
    t = page.add_freetext_annot(r, d['text'], fontsize=size, fontname='helv', text_color=parse_color(d.get('textColor') or '#111111'),
                                align=fitz.TEXT_ALIGN_CENTER)
    t.set_border(width=0); t.set_info(subject='Shape text'); t.update()
    doc.xref_set_key(t.xref, 'MyPDFKitShapeText', fitz.get_pdf_str(str(d.get('id') or '')))

@app.route('/take_shape', methods=['POST'])
def take_shape():
    """A shape saved earlier comes out of the file so it can be moved, resized or recolored again."""
    d = request.json; fp = d['filepath']
    try:
        doc = fitz.open(fp); page = doc[int(d['page'])]
        for a in page.annots() or []:
            if a.xref != int(d['xref']): continue
            t, v = doc.xref_get_key(a.xref, 'MyPDFKitShape')
            if t != 'string': break
            shape = json.loads(v)
            page.delete_annot(a)
            if shape.get('id'):   # its words go with it
                for t in list(page.annots(types=[fitz.PDF_ANNOT_FREE_TEXT]) or []):
                    tt, tv = doc.xref_get_key(t.xref, 'MyPDFKitShapeText')
                    if tt == 'string' and tv == str(shape['id']): page.delete_annot(t)
            # the rest of its group comes too
            group = []
            if shape.get('group'):
                for b in list(page.annots(types=[fitz.PDF_ANNOT_POLYGON, fitz.PDF_ANNOT_LINE]) or []):
                    bt, bv = doc.xref_get_key(b.xref, 'MyPDFKitShape')
                    if bt != 'string': continue
                    other = json.loads(bv)
                    if other.get('group') != shape['group']: continue
                    page.delete_annot(b); group.append(other)
                    if other.get('id'):
                        for t in list(page.annots(types=[fitz.PDF_ANNOT_FREE_TEXT]) or []):
                            tt, tv = doc.xref_get_key(t.xref, 'MyPDFKitShapeText')
                            if tt == 'string' and tv == str(other['id']): page.delete_annot(t)
            shape['others'] = group
            out = _temp_out(fp, 'tshape', json.dumps(d))
            doc.save(out, garbage=3, deflate=True); doc.close()
            return jsonify({"success": True, "shape": shape, "new_file": out})
        doc.close()
        return jsonify({"success": True, "shape": None})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- find and replace ----------
@app.route('/replace_text', methods=['POST'])
def replace_text():
    """Replace a word or phrase everywhere it appears, in the same font, size and color as the text around it."""
    d = request.json; fp = d['filepath']; find = d.get('find') or ''
    if not find.strip(): return jsonify({"error": "Type what to find"}), 400
    try:
        flags = 0 if d.get('case') else re.I
        pat = re.compile((r'(?<!\w)' + re.escape(find) + r'(?!\w)') if d.get('whole') else re.escape(find), flags)
        repl = d.get('replace') or ''
        doc = fitz.open(fp); count = pages = 0
        for page in doc:
            edits = []
            for sp in _page_spans(page):
                n = len(pat.findall(sp['text']))
                if not n: continue
                new = pat.sub(lambda m: repl, sp['text'])
                edits.append({'rect': list(sp['bbox']), 'base_y': sp['origin'][1], 'new_text': new if new.strip() else ' '}); count += n
            if edits: _rewrite_text(doc, page, edits); pages += 1
        if not count: doc.close(); return jsonify({"success": True, "count": 0})
        try: doc.subset_fonts()
        except Exception: pass
        out = _temp_out(fp, 'repl', json.dumps(d))
        doc.save(out, garbage=3, deflate=True); doc.close()
        return jsonify({"success": True, "count": count, "pages": pages, "new_file": out})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- Bates numbers ----------
@app.route('/bates', methods=['POST'])
def bates():
    """A running number (ABC000001, ABC000002...) on every page of each file, continuing from file to file."""
    d = request.json
    try:
        n = int(d.get('start') or 1); digits = max(1, min(12, int(d.get('digits') or 6)))
        prefix, suffix = d.get('prefix') or '', d.get('suffix') or ''
        size = float(d.get('size') or 10); rgb = parse_color(d.get('color') or '#111111')
        where = d.get('position') or 'br'
        outs = []
        for fp in d['files']:
            doc = fitz.open(fp)
            for page in doc:
                label = f"{prefix}{n:0{digits}d}{suffix}"; n += 1
                r = page.rect; w = fitz.get_text_length(label, fontname='helv', fontsize=size); m = 22
                x = {'l': m, 'c': (r.width - w) / 2, 'r': r.width - m - w}[where[1]]
                y = m + size if where[0] == 't' else r.height - m
                # unrotated page coordinates, so the number sits upright where it reads
                pt = fitz.Point(x, y) * page.derotation_matrix
                page.insert_text(pt, label, fontsize=size, fontname='helv', color=rgb, rotate=page.rotation)
            out = _temp_out(fp, 'bates', json.dumps(d) + str(n))
            doc.save(out, garbage=3, deflate=True); doc.close(); outs.append(out)
        return jsonify({"success": True, "files": outs, "next": n})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- privacy check: what someone could still dig out of a PDF ----------
COVER_ANNOTS = None

def _covered_words(page):
    """Letters that are in the file but can't be seen: under a black (or white) box, a dark highlight, a picture,
    or written in the same color as what's behind them. Found by looking, letter by letter: where each letter
    sits is drawn as one flat color. These are the "redactions" that can simply be copied out (a box over half
    a word hides only those letters, and only those are reported and removed)."""
    import numpy as np
    spans = [sp for sp in page.get_texttrace() if sp.get('type') != 3]   # render mode 3 is OCR's see-through text: meant to be invisible
    chars = [(ch[0], fitz.Rect(ch[3])) for sp in spans for ch in sp['chars'] if chr(ch[0]).strip()]
    chars = [(c, r) for c, r in chars if r.width > 0.3 and r.height > 0.3]
    if not chars: return []
    hmin = max(0.5, min(r.height for _, r in chars))
    dpi = int(min(220, max(72, 10 * 72 / hmin)))
    pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY, alpha=False, annots=True)
    a = np.frombuffer(pix.samples, np.uint8)[:pix.stride * pix.height].reshape(pix.height, pix.stride)[:, :pix.width]
    k = dpi / 72; rm = page.rotation_matrix
    hidden = []
    for c, r in chars:
        shown = r * rm
        x0, y0, x1, y1 = int(shown.x0 * k + 0.5), int(shown.y0 * k + 0.5), int(shown.x1 * k - 0.5), int(shown.y1 * k - 0.5)
        if x1 - x0 < 2 or y1 - y0 < 3: continue
        region = a[max(0, y0):y1, max(0, x0):x1]
        if region.size < 6: continue
        if region.std() < 4:   # nothing drawn there but one flat color: the letter can't be seen
            hidden.append({"rect": [shown.x0, shown.y0, shown.x1, shown.y1], "word": chr(c), "dark": float(region.mean()) < 128, "raw": [r.x0, r.y0, r.x1, r.y1]})
    return hidden

def _group_rects(hits):
    """Hidden letters next to each other on a line become one area (with the hidden text, for the report)."""
    out = []
    for h in sorted(hits, key=lambda h: (round(h['rect'][1]), h['rect'][0])):
        r = fitz.Rect(h['rect'])
        if out and abs(out[-1]['r'].y0 - r.y0) < 3 and r.x0 - out[-1]['r'].x1 < max(4, r.height * 0.8) and out[-1]['dark'] == h['dark']:
            gap = ' ' if r.x0 - out[-1]['r'].x1 > r.height * 0.15 else ''
            out[-1]['r'] |= r; out[-1]['text'] += gap + h['word']
        else: out.append({'r': r, 'text': h['word'], 'dark': h['dark']})
    out = [g for g in out if len(g['text'].replace(' ', '')) >= 1]
    return [{"rect": [g['r'].x0, g['r'].y0, g['r'].x1, g['r'].y1], "words": len(g['text'].split()) or 1, "dark": g['dark'],
             "sample": g['text'][:24] + ('…' if len(g['text']) > 24 else '')} for g in out]

def _privacy_findings(fp, quick=False):
    doc = fitz.open(fp); raw = open(fp, 'rb').read()
    find = []
    covered, pending = {}, {}
    for pno, page in enumerate(doc):
        reds = [a for a in (page.annots() or []) if a.type[0] == fitz.PDF_ANNOT_REDACT]
        if reds: pending[pno + 1] = len(reds)
        if quick:   # only pages that have something that could cover words
            has_cover = any(a.type[0] in (fitz.PDF_ANNOT_SQUARE, fitz.PDF_ANNOT_POLYGON, fitz.PDF_ANNOT_CIRCLE, fitz.PDF_ANNOT_HIGHLIGHT, fitz.PDF_ANNOT_INK,
                                          fitz.PDF_ANNOT_FREE_TEXT, fitz.PDF_ANNOT_STAMP) for a in (page.annots() or []))
            has_cover = has_cover or any(d.get('fill') is not None and max(d['fill']) < 0.3 for d in page.get_drawings()) or bool(page.get_images())
            if not has_cover: continue
        g = _group_rects(_covered_words(page))
        if g: covered[pno + 1] = g
    if covered:
        n = sum(len(v) for v in covered.values())
        find.append({"id": "covered", "severity": "danger", "title": f"Text hidden under boxes or covered up in {n} {'place' if n == 1 else 'places'}",
                     "detail": "They look blacked out (or whited out), but the words are still in the file: anyone can copy them out or remove the box. Fixing removes the words for good and keeps the box.",
                     "pages": covered, "fixable": True})
    if pending:
        find.append({"id": "pending", "severity": "danger", "title": f"{sum(pending.values())} {'redaction' if sum(pending.values()) == 1 else 'redactions'} marked but never applied",
                     "detail": "These are only marks. The text under them is still there until they're applied.", "pages": {k: [] for k in pending}, "fixable": True})
    if quick: doc.close(); return find
    versions = raw.count(b'%%EOF')
    if versions > 1:
        find.append({"id": "versions", "severity": "danger", "title": f"{versions - 1} earlier {'version' if versions == 2 else 'versions'} kept inside the file",
                     "detail": "The file still holds what it looked like before later changes, including text that was deleted or covered. Fixing keeps only the current version.", "fixable": True})
    meta = {k: v for k, v in (doc.metadata or {}).items() if v and k not in ('format', 'encryption')}
    if meta or doc.get_xml_metadata():
        find.append({"id": "metadata", "severity": "warn", "title": "Document details: " + ', '.join(f"{k} \"{str(v)[:30]}\"" for k, v in list(meta.items())[:4]) if meta else "Hidden document details (XML)",
                     "detail": "Who made it, with which program, when, its title and subject. Fixing clears them.", "fixable": True})
    if doc.embfile_count():
        find.append({"id": "attachments", "severity": "warn", "title": f"{doc.embfile_count()} attached {'file' if doc.embfile_count() == 1 else 'files'}: " + ', '.join(doc.embfile_names()[:3]),
                     "detail": "Files tucked inside the PDF. Fixing removes them.", "fixable": True})
    if b'/JavaScript' in raw or b'/JS ' in raw or b'/JS(' in raw:
        find.append({"id": "scripts", "severity": "warn", "title": "Scripts (JavaScript) inside", "detail": "Code that runs when the file opens. Fixing removes it.", "fixable": True})
    notes = sum(1 for p in doc for a in (p.annots() or []) if a.type[0] in (fitz.PDF_ANNOT_TEXT, fitz.PDF_ANNOT_POPUP) or (a.info.get('content') and a.type[0] not in (fitz.PDF_ANNOT_FREE_TEXT,)))
    if notes:
        find.append({"id": "comments", "severity": "warn", "title": f"{notes} {'comment' if notes == 1 else 'comments'} with notes",
                     "detail": "Sticky notes and the text inside comments travel with the file. Fixing removes the notes (the marks you drew stay).", "fixable": True})
    try:
        layers = [c for c in doc.layer_ui_configs() if not c.get('on')]
    except Exception: layers = []
    if layers:
        find.append({"id": "layers", "severity": "warn", "title": f"{len(layers)} hidden {'layer' if len(layers) == 1 else 'layers'}: " + ', '.join(c.get('text', '') for c in layers[:3]),
                     "detail": "Parts of the page switched off from view but still in the file. Fixing shows them, so you can see what's there.", "fixable": True})
    fields = sum(1 for p in doc for w in (p.widgets() or []) if w.field_value not in (None, '', 'Off'))
    if fields:
        find.append({"id": "fields", "severity": "info", "title": f"{fields} filled-in form {'field' if fields == 1 else 'fields'}",
                     "detail": "Anyone can still change what's typed in them. Lock and finalize makes them part of the page.", "fixable": False})
    doc.close()
    return find

@app.route('/privacy_check', methods=['POST'])
def privacy_check():
    try: return jsonify({"findings": _privacy_findings(request.json['filepath'], bool(request.json.get('quick')))})
    except Exception as e: return jsonify({"error": str(e)}), 500

def _privacy_fix(doc, fixes, fp):
    """Fix what the check found (fixes: ids). Returns lines saying what was done."""
    done = []
    if 'covered' in fixes:   # covered words go for good: real redactions that remove only the text, so the box and the page look the same
        n = 0
        for page in doc:
            hits = _covered_words(page)
            if not hits: continue
            for h in hits:   # each hidden letter, shrunk a little so the visible letters beside it stay
                r = fitz.Rect(h['raw']); dx, dy = r.width * 0.2, r.height * 0.2
                page.add_redact_annot(fitz.Rect(r.x0 + dx, r.y0 + dy, r.x1 - dx, r.y1 - dy), fill=None)
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE, text=fitz.PDF_REDACT_TEXT_REMOVE)
            n += len(hits)
        if n: done.append(f"Removed {n} hidden {'letter' if n == 1 else 'letters'} for good")
    if 'pending' in fixes:
        n = 0
        for page in doc:
            k = sum(1 for a in (page.annots() or []) if a.type[0] == fitz.PDF_ANNOT_REDACT)
            if k: page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_PIXELS, graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED); n += k
        if n: done.append(f"Applied {n} {'redaction' if n == 1 else 'redactions'}")
    if 'comments' in fixes:
        n = 0
        for page in doc:
            for a in list(page.annots() or []):
                if a.type[0] in (fitz.PDF_ANNOT_TEXT, fitz.PDF_ANNOT_POPUP): page.delete_annot(a); n += 1
                elif a.info.get('content') and a.type[0] != fitz.PDF_ANNOT_FREE_TEXT: a.set_info(content=''); a.update(); n += 1
        if n: done.append(f"Removed {n} comment {'note' if n == 1 else 'notes'}")
    if 'layers' in fixes:
        try:
            for i, c in enumerate(doc.layer_ui_configs()):
                if not c.get('on'): doc.set_layer_ui_config(i, 0)
            done.append('Turned hidden layers on')
        except Exception: pass
    raw = open(fp, 'rb').read()   # report only what was really there
    present = {'metadata': bool({k: v for k, v in (doc.metadata or {}).items() if v and k not in ('format', 'encryption')}) or bool(doc.get_xml_metadata()),
               'attachments': doc.embfile_count() > 0, 'scripts': b'/JavaScript' in raw or b'/JS ' in raw or b'/JS(' in raw}
    scrub = {k: k in fixes and present[k] for k in ('metadata', 'attachments', 'scripts')}
    if any(scrub.values()):
        doc.scrub(metadata=scrub['metadata'], xml_metadata=scrub['metadata'], attached_files=scrub['attachments'], embedded_files=scrub['attachments'],
                  javascript=scrub['scripts'], thumbnails=True, reset_fields=False, reset_responses=False, remove_links=False, hidden_text=False, redactions=False, clean_pages=False)
        if scrub['metadata']: doc.set_metadata({}); done.append('Cleared the document details')
        if scrub['attachments']: done.append('Removed attached files')
        if scrub['scripts']: done.append('Removed scripts')
    if 'versions' in fixes and raw.count(b'%%EOF') > 1: done.append('Removed the earlier versions kept inside')   # the full save drops them
    return done

@app.route('/privacy_fix', methods=['POST'])
def privacy_fix():
    d = request.json; fp = d['filepath']
    try:
        doc = fitz.open(fp)
        done = _privacy_fix(doc, set(d.get('fixes') or []), fp)
        out = _temp_out(fp, 'priv', json.dumps(d))
        doc.save(out, garbage=4, deflate=True, clean=True); doc.close()   # a fresh, full save: no earlier versions inside
        left = [f['id'] for f in _privacy_findings(out) if f['id'] in ('covered', 'pending', 'versions')]
        return jsonify({"success": True, "new_file": out, "done": done, "still": left})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- lock and finalize ----------
def _outline_text(src):
    """The document with its text turned into drawn outlines: it looks the same and stays sharp at any zoom,
    but there's no text left to select, copy or change. Pictures and lines are kept exactly as they were."""
    out = fitz.open()
    base = fitz.open(); base.insert_pdf(src)     # pictures and lines, text taken out
    txt = fitz.open(); txt.insert_pdf(src)       # the text alone
    for i in range(src.page_count):
        b = base[i]; b.add_redact_annot(b.rect)
        b.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE, text=fitz.PDF_REDACT_TEXT_REMOVE)
        t = txt[i]; t.add_redact_annot(t.rect)
        t.apply_redactions(images=fitz.PDF_REDACT_IMAGE_REMOVE, graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_TOUCHED, text=fitz.PDF_REDACT_TEXT_NONE)
    for i in range(src.page_count):
        page = out.new_page(width=src[i].rect.width, height=src[i].rect.height)
        page.show_pdf_page(page.rect, base, i)
        if txt[i].get_text('words'):
            tp = fitz.open('pdf', fitz.open('svg', txt[i].get_svg_image(text_as_path=True).encode('utf-8')).convert_to_pdf())
            page.show_pdf_page(page.rect, tp, 0)
    return out

@app.route('/unlock_mine', methods=['POST'])
def unlock_mine():
    """A copy without the lock, for a file locked with this app's own password."""
    d = request.json; fp = d['filepath']; orig = d.get('original_path') or fp
    try:
        doc = fitz.open(fp)
        if not (doc.metadata or {}).get('encryption'): return jsonify({"error": "This file isn't locked"}), 400
        if not doc.authenticate(d.get('password') or '') & 4:   # 4: the owner's password, the one that allows changes
            return jsonify({"error": "This file wasn't locked with MyPDFKit on this computer"}), 400
        base = os.path.splitext(os.path.basename(orig))[0].replace(' - final', '')
        out = get_unique_filename(os.path.join(os.path.dirname(orig), base + ' (unlocked).pdf'))
        doc.save(out, encryption=fitz.PDF_ENCRYPT_NONE, garbage=3, deflate=True)
        return jsonify({"success": True, "path": out})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/finalize', methods=['POST'])
def finalize():
    """The finished copy: every edit made part of the page (comments, shapes, signatures, form fields), hidden
    things cleaned out, optionally every page turned into a picture, then protected: editing blocked with a
    password, or certified with a digital ID so any later change shows as tampering."""
    d = request.json; fp = d['filepath']; orig = d.get('original_path') or fp
    try:
        doc = fitz.open(fp); report = []
        if d.get('clean', True):
            report += _privacy_fix(doc, {'covered', 'pending', 'metadata', 'attachments', 'scripts', 'versions'}, fp)
        if d.get('flatten', True):
            n = sum(1 for p in doc for a in (p.annots() or []) if a.type[0] not in (fitz.PDF_ANNOT_LINK,)) + sum(1 for p in doc for _ in (p.widgets() or []))
            doc.bake(annots=True, widgets=True)
            if n: report.append(f"Made {n} {'edit' if n == 1 else 'edits'}, comments and fields part of the page")
        content = d.get('content') or ('picture' if d.get('rasterize') else 'keep')
        if content == 'outline':   # sharp, but no text left to select or copy
            had = sum(len(p.get_text('words')) for p in doc)
            out_doc = _outline_text(doc); doc.close(); doc = out_doc
            report.append(f"Turned the text ({had} words) into outlines: it looks the same and stays sharp, but can't be selected or copied")
        if content == 'picture':   # the strongest: pictures of the pages, nothing left to edit or copy
            dpi = int(d.get('dpi') or 300); pics = fitz.open()
            for page in doc:
                pix = page.get_pixmap(dpi=dpi, annots=True)
                np_ = pics.new_page(width=page.rect.width, height=page.rect.height); np_.insert_image(np_.rect, pixmap=pix)
            doc.close(); doc = pics
            report.append(f"Turned every page into a picture ({dpi} dpi)")
        base = os.path.splitext(os.path.basename(orig))[0]
        out = get_unique_filename(os.path.join(os.path.dirname(orig), base + ' - final.pdf'))
        mode = d.get('protect') or 'none'
        if mode in ('password', 'auto'):
            if not d.get('password'): return jsonify({"error": "Choose a password for unlocking changes"}), 400
            perm = fitz.PDF_PERM_ACCESSIBILITY   # no editing, form filling, comments or page changes
            if d.get('allow_print', True): perm |= fitz.PDF_PERM_PRINT | fitz.PDF_PERM_PRINT_HQ
            if d.get('allow_copy', mode == 'password'): perm |= fitz.PDF_PERM_COPY
            doc.save(out, garbage=4, deflate=True, clean=True, encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw=d['password'], user_pw='', permissions=perm)
            report.append(('Locked: anyone can open it' + (' and print it' if d.get('allow_print', True) else '') + '; editing, forms and page changes are blocked')
                          if mode == 'auto' else 'Changes need your password' + ('' if d.get('allow_copy', True) else '; copying text is blocked') + ('' if d.get('allow_print', True) else '; printing is blocked'))
        else:
            tmp = out if mode == 'none' else _temp_out(fp, 'final', json.dumps(d))
            doc.save(tmp, garbage=4, deflate=True, clean=True)
            if mode == 'certify':
                from pyhanko.sign import signers, fields
                from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
                signer = signers.SimpleSigner.load_pkcs12(d['pfx'], passphrase=(d.get('pfx_password') or '').encode() or None)
                if signer is None: doc.close(); return jsonify({"error": "That password isn't right for this digital ID"}), 400
                meta = signers.PdfSignatureMetadata(field_name='Certification', certify=True, docmdp_permissions=fields.MDPPerm.NO_CHANGES,
                                                    reason=d.get('reason') or 'Final version: no changes allowed')
                with open(tmp, 'rb') as inf:
                    w = IncrementalPdfFileWriter(inf, strict=False)
                    with open(out, 'wb') as outf: signers.PdfSigner(meta, signer=signer).sign_pdf(w, output=outf)
                report.append('Certified with your digital ID: any change from now on shows as tampering')
        doc.close()
        return jsonify({"success": True, "path": out, "report": report, "message": f"Saved {os.path.basename(out)}"})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- find things to black out: phone numbers, emails, card numbers... ----------
REDACT_PATTERNS = {
    'phone': r'(?<![\w/])(?:\+\d{1,3}[\s.-]?)?(?:\(\d{3}\)\s?|\d{3}[\s.-])\d{3}[\s.-]\d{4}(?!\d)',
    'email': r'[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}',
    'card': r'(?<!\d)(?:\d[ -]?){12,18}\d(?!\d)',
    'ssn': r'(?<!\d)\d{3}-\d{2}-\d{4}(?!\d)',
    'date': r'(?<!\d)(?:\d{1,2}[/.-]\d{1,2}[/.-](?:\d{4}|\d{2})|\d{4}-\d{2}-\d{2})(?!\d)|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4}\b|\b\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\.?\s+\d{4}\b',
    'money': r'[$€£¥]\s?\d[\d,]*(?:\.\d{1,2})?|\b\d[\d,]*(?:\.\d{2})?\s?(?:USD|EUR|GBP|dollars)\b',
    'url': r'\bhttps?://[^\s<>"]+|\bwww\.[^\s<>"]+\.[a-z]{2,}[^\s<>"]*',
    'ip': r'(?<![\d.])(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)(?![\d.])',
    'zip': r'(?<!\d)\d{5}(?:-\d{4})(?!\d)',
}

def _luhn(digits):
    s, alt = 0, False
    for ch in reversed(digits):
        d = int(ch)
        if alt: d = d * 2 - 9 if d > 4 else d * 2
        s += d; alt = not alt
    return s % 10 == 0

@app.route('/find_patterns', methods=['POST'])
def find_patterns():
    """Every phone number, email, card number, SSN, date, amount, web address or IP in the file (and any words
    typed), as boxes to black out. Looks line by line, so a number split into pieces is still found."""
    d = request.json; kinds = d.get('kinds') or []
    words_extra = [w.strip() for w in (d.get('words') or '').replace(';', ',').split(',') if w.strip()]
    pats = [(k, re.compile(REDACT_PATTERNS[k], re.I)) for k in kinds if k in REDACT_PATTERNS]
    pats += [('word', re.compile(r'(?<!\w)' + re.escape(w) + r'(?!\w)', re.I)) for w in words_extra]
    try:
        doc = fitz.open(d['filepath']); out = {}; counts = {}
        for pno, page in enumerate(doc):
            rm = page.rotation_matrix
            lines = {}
            for w in page.get_text('words'): lines.setdefault((w[5], w[6]), []).append(w)
            for ws in lines.values():
                ws.sort(key=lambda w: w[0])
                text, spans = '', []
                for w in ws:
                    if text: text += ' '
                    spans.append((len(text), len(text) + len(w[4]), w)); text += w[4]
                for kind, pat in pats:
                    for m in pat.finditer(text):
                        if kind == 'card' and not _luhn(re.sub(r'\D', '', m.group())): continue
                        if kind == 'phone' and len(re.sub(r'\D', '', m.group())) < 10: continue
                        hit = [sp for sp in spans if sp[0] < m.end() and sp[1] > m.start()]
                        if not hit: continue
                        r = fitz.Rect()
                        for a, b, w in hit:   # only the matched letters of the first and last word
                            wr = fitz.Rect(w[:4]); n = max(1, len(w[4])); cw = wr.width / n
                            x0 = wr.x0 + cw * max(0, m.start() - a); x1 = wr.x0 + cw * min(n, m.end() - a)
                            r |= fitz.Rect(x0, wr.y0, x1, wr.y1)
                        r = r * rm   # as shown
                        out.setdefault(pno + 1, []).append({"rect": [r.x0 - 1, r.y0 - 1, r.x1 + 1, r.y1 + 1], "kind": kind, "text": m.group()[:40]})
                        counts[kind] = counts.get(kind, 0) + 1
        doc.close()
        return jsonify({"pages": out, "counts": counts, "count": sum(counts.values())})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- PDF/A: the archive format courts and archives ask for ----------
def _embed_base14(doc):
    """Fonts that were only named, not stored in the file (Helvetica, Times, Courier...), get stored, using
    MuPDF's built-in look-alikes with the same letter widths. PDF/A needs every font in the file."""
    done = 0
    for xref in range(1, doc.xref_length()):
        try:
            if doc.xref_get_key(xref, 'Type')[1] != '/Font': continue
            sub = doc.xref_get_key(xref, 'Subtype')[1]
            if sub not in ('/Type1', '/TrueType', '/MMType1'): continue
            fd = doc.xref_get_key(xref, 'FontDescriptor')
            if fd[0] == 'xref':
                fdx = int(fd[1].split()[0])
                if any(doc.xref_get_key(fdx, k)[0] != 'null' for k in ('FontFile', 'FontFile2', 'FontFile3')): continue
            base = doc.xref_get_key(xref, 'BaseFont')[1].lstrip('/')
            base = re.sub(r'^[A-Z]{6}\+', '', base)
            low = base.lower().replace('-', '').replace(',', '')
            fam = 'cour' if 'cour' in low else 'tiro' if ('times' in low or 'roman' in low or 'serif' in low) else 'symb' if 'symbol' in low else 'zadb' if 'dingbat' in low else 'helv'
            if fam in ('cour', 'tiro', 'helv'):
                bold, ital = 'bold' in low or 'black' in low, 'italic' in low or 'oblique' in low
                fam = {('cour', 0, 0): 'cour', ('cour', 1, 0): 'cobo', ('cour', 0, 1): 'coit', ('cour', 1, 1): 'cobi',
                       ('tiro', 0, 0): 'tiro', ('tiro', 1, 0): 'tibo', ('tiro', 0, 1): 'tiit', ('tiro', 1, 1): 'tibi',
                       ('helv', 0, 0): 'helv', ('helv', 1, 0): 'hebo', ('helv', 0, 1): 'heit', ('helv', 1, 1): 'hebi'}[(fam, int(bold), int(ital))]
            f = fitz.Font(fam)
            data = f.buffer
            ff = doc.get_new_xref(); doc.update_object(ff, '<<>>'); doc.update_stream(ff, data)
            doc.xref_set_key(ff, 'Subtype', '/Type1C')
            bb = f.bbox
            desc = (f"<</Type/FontDescriptor/FontName/{base}/Flags {32 if fam not in ('symb', 'zadb') else 4}"
                    f"/FontBBox[{int(bb.x0 * 1000)} {int(bb.y0 * 1000)} {int(bb.x1 * 1000)} {int(bb.y1 * 1000)}]/ItalicAngle {-12 if fam in ('heit', 'hebi', 'tiit', 'tibi', 'coit', 'cobi') else 0}"
                    f"/Ascent {int(f.ascender * 1000)}/Descent {int(f.descender * 1000)}/CapHeight 700/StemV 80/FontFile3 {ff} 0 R>>")
            fdx = doc.get_new_xref(); doc.update_object(fdx, desc)
            doc.xref_set_key(xref, 'FontDescriptor', f'{fdx} 0 R')
            if doc.xref_get_key(xref, 'Widths')[0] == 'null':   # letter widths, stated in the file as PDF/A wants
                widths = [int(round(f.glyph_advance(f.has_glyph(c) and c or 32) * 1000)) if c >= 32 else 0 for c in range(256)]
                doc.xref_set_key(xref, 'FirstChar', '0'); doc.xref_set_key(xref, 'LastChar', '255')
                doc.xref_set_key(xref, 'Widths', '[' + ' '.join(map(str, widths)) + ']')
            done += 1
        except Exception as e:
            print('font embed skipped:', xref, e, file=sys.stderr)
    return done

@app.route('/pdfa', methods=['POST'])
def save_pdfa():
    """A PDF/A-2b copy next to the original: every font stored in the file, a color profile, the archive
    marker in the document details, comments and fields made part of the page, and nothing PDF/A forbids
    (no password, scripts or attachments)."""
    import datetime
    d = request.json; fp = d['filepath']; orig = d.get('original_path') or fp
    try:
        doc = fitz.open(fp)
        if doc.needs_pass or doc.is_encrypted: return jsonify({"error": "Remove the password first: PDF/A files can't be locked"}), 400
        report = []
        n = sum(1 for p in doc for _ in (p.annots() or [])) + sum(1 for p in doc for _ in (p.widgets() or []))
        doc.bake(annots=True, widgets=True)
        if n: report.append(f"Made {n} comments and fields part of the page")
        doc.scrub(metadata=False, xml_metadata=True, attached_files=True, embedded_files=True, javascript=True, thumbnails=True,
                  reset_fields=False, reset_responses=False, remove_links=False, hidden_text=False, redactions=False, clean_pages=False)
        k = _embed_base14(doc)
        if k: report.append(f"Stored {k} {'font' if k == 1 else 'fonts'} in the file")
        # the color profile everything is shown in (sRGB, from Windows)
        icc_path = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'System32', 'spool', 'drivers', 'color', 'sRGB Color Space Profile.icm')
        icc = open(icc_path, 'rb').read()
        ix = doc.get_new_xref(); doc.update_object(ix, '<</N 3>>'); doc.update_stream(ix, icc)
        oi = doc.get_new_xref(); doc.update_object(oi, f'<</Type/OutputIntent/S/GTS_PDFA1/OutputConditionIdentifier(sRGB IEC61966-2.1)/Info(sRGB IEC61966-2.1)/DestOutputProfile {ix} 0 R>>')
        doc.xref_set_key(doc.pdf_catalog(), 'OutputIntents', f'[{oi} 0 R]')
        # document details, the same in both places, with the PDF/A-2b marker
        meta = {k2: v for k2, v in (doc.metadata or {}).items() if v}
        title = meta.get('title') or os.path.splitext(os.path.basename(orig))[0]
        now = datetime.datetime.now(datetime.timezone.utc).astimezone()
        pdf_date = now.strftime("D:%Y%m%d%H%M%S") + (now.strftime('%z')[:3] + "'" + now.strftime('%z')[3:] + "'")
        iso = now.isoformat(timespec='seconds')
        doc.set_metadata({'title': title, 'author': meta.get('author', ''), 'subject': meta.get('subject', ''), 'keywords': meta.get('keywords', ''),
                          'creator': 'MyPDFKit', 'producer': 'MyPDFKit', 'creationDate': pdf_date, 'modDate': pdf_date})
        from xml.sax.saxutils import escape as esc
        xmp = f"""<?xpacket begin="\ufeff" id="W5M0MpCehiHzreSzNTczkc9d"?>
<x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
<rdf:Description rdf:about="" xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:xmp="http://ns.adobe.com/xap/1.0/" xmlns:pdf="http://ns.adobe.com/pdf/1.3/">
<pdfaid:part>2</pdfaid:part><pdfaid:conformance>B</pdfaid:conformance>
<dc:format>application/pdf</dc:format>
<dc:title><rdf:Alt><rdf:li xml:lang="x-default">{esc(title)}</rdf:li></rdf:Alt></dc:title>
{f'<dc:creator><rdf:Seq><rdf:li>{esc(meta["author"])}</rdf:li></rdf:Seq></dc:creator>' if meta.get('author') else ''}
{f'<dc:description><rdf:Alt><rdf:li xml:lang="x-default">{esc(meta["subject"])}</rdf:li></rdf:Alt></dc:description>' if meta.get('subject') else ''}
{f'<pdf:Keywords>{esc(meta["keywords"])}</pdf:Keywords>' if meta.get('keywords') else ''}
<xmp:CreatorTool>MyPDFKit</xmp:CreatorTool><pdf:Producer>MyPDFKit</pdf:Producer>
<xmp:CreateDate>{iso}</xmp:CreateDate><xmp:ModifyDate>{iso}</xmp:ModifyDate><xmp:MetadataDate>{iso}</xmp:MetadataDate>
</rdf:Description></rdf:RDF></x:xmpmeta>
<?xpacket end="w"?>"""
        doc.set_xml_metadata(xmp)
        raw_fonts_left = []   # a last look: anything that still isn't stored in the file
        for xref in range(1, doc.xref_length()):
            try:
                if doc.xref_get_key(xref, 'Type')[1] == '/Font' and doc.xref_get_key(xref, 'Subtype')[1] in ('/Type1', '/TrueType'):
                    fd = doc.xref_get_key(xref, 'FontDescriptor')
                    if fd[0] != 'xref' or all(doc.xref_get_key(int(fd[1].split()[0]), kk)[0] == 'null' for kk in ('FontFile', 'FontFile2', 'FontFile3')):
                        raw_fonts_left.append(doc.xref_get_key(xref, 'BaseFont')[1])
            except Exception: pass
        cmyk = any(b'/DeviceCMYK' in doc.xref_object(x).encode() for x in range(1, min(doc.xref_length(), 20000)) if doc.xref_object(x))
        base = os.path.splitext(os.path.basename(orig))[0]
        out = get_unique_filename(os.path.join(os.path.dirname(orig), base + ' (PDF-A).pdf'))
        doc.save(out, garbage=3, deflate=True, clean=True, no_new_id=False, pretty=False)
        doc.close()
        warn = []
        if raw_fonts_left: warn.append("Some fonts couldn't be stored: " + ', '.join(sorted(set(raw_fonts_left))[:4]))
        if cmyk: warn.append('It uses print colors (CMYK); strict checkers may flag that')
        report.append('Marked as PDF/A-2b with an sRGB color profile')
        return jsonify({"success": True, "path": out, "report": report, "warnings": warn, "message": f"Saved {os.path.basename(out)}"})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- the eraser, on pen strokes already in the file ----------
def _dense(pts, step):
    """The points of a line with no gap bigger than step between them."""
    out = []
    for a, b in zip(pts, pts[1:]):
        n = max(1, int(((b[0] - a[0]) ** 2 + (b[1] - a[1]) ** 2) ** 0.5 / step))
        out += [(a[0] + (b[0] - a[0]) * k / n, a[1] + (b[1] - a[1]) * k / n) for k in range(n)]
    return out + pts[-1:]

@app.route('/erase_ink', methods=['POST'])
def erase_ink():
    """mode 'stroke': strokes the eraser touched go; 'area': only the ink under the eraser goes (a stroke
    can break in two); 'all': every pen stroke on the page goes. Points are as the pen saved them."""
    d = request.json; fp = d['filepath']
    try:
        doc = fitz.open(fp); page = doc[int(d['page'])]
        mode, r = d.get('mode', 'stroke'), float(d.get('radius', 6))
        path = _dense([tuple(p) for p in d.get('points') or []], max(0.5, r / 3)) if mode != 'all' else []
        if path:
            xs, ys = [p[0] for p in path], [p[1] for p in path]
            reach = fitz.Rect(min(xs) - r, min(ys) - r, max(xs) + r, max(ys) + r)
        def hit(pt, extra=0.0):
            rr = (r + extra) ** 2
            return any((pt[0] - q[0]) ** 2 + (pt[1] - q[1]) ** 2 <= rr for q in path)
        changed = 0
        for a in list(page.annots() or []):
            kind = a.type[0]
            pen_shape = kind == fitz.PDF_ANNOT_POLYGON and a.info.get('subject') == 'Pen'
            if kind != fitz.PDF_ANNOT_INK and not pen_shape: continue
            if mode == 'all': page.delete_annot(a); changed += 1; continue
            if not a.rect.intersects(reach): continue
            width = (a.border or {}).get('width') or 1
            if pen_shape:
                kept = None
                try:
                    t, v = doc.xref_get_key(a.xref, 'MyPDFKitInk')
                    if t == 'string': kept = json.loads(v)
                except Exception: pass
                if not any(hit(v) for v in _dense([tuple(v) for v in a.vertices] + [tuple(a.vertices[0])], 1)): continue
                if mode == 'stroke' or not kept:
                    page.delete_annot(a); changed += 1; continue
                # rub out just the part under the eraser: cut the centre line, redraw each piece
                P, W = kept['p'], kept['w']; dp, dw = [tuple(P[0])], [W[0]]
                for i in range(1, len(P)):
                    seg = _dense([tuple(P[i - 1]), tuple(P[i])], 0.8)[1:]
                    dp += seg; dw += [W[i - 1] + (W[i] - W[i - 1]) * (k + 1) / len(seg) for k in range(len(seg))]
                pieces, run = [], ([], [])
                for q, w in zip(dp, dw):
                    if hit(q, w / 2):
                        if len(run[0]) > 1: pieces.append(run)
                        run = ([], [])
                    else: run[0].append(q); run[1].append(w)
                if len(run[0]) > 1: pieces.append(run)
                rgb, opacity, info = a.colors.get('fill') or a.colors.get('stroke'), a.opacity, a.info
                page.delete_annot(a); changed += 1
                for pp, ww in pieces:
                    _add_pen_shape(doc, page, _ink_outline(pp, ww), rgb, opacity if opacity is not None and opacity >= 0 else 1, pp, ww, info, kept.get('t'))
                continue
            lists = [_dense([tuple(v) for v in l], 1) for l in a.vertices]
            if mode == 'stroke':
                if any(hit(v, width / 2) for l in lists for v in l): page.delete_annot(a); changed += 1
                continue
            pieces, cut = [], False
            for l in lists:
                run = []
                for v in l:
                    if hit(v, width / 2): cut = True; (pieces.append(run) if len(run) > 1 else None); run = []
                    else: run.append(v)
                if len(run) > 1: pieces.append(run)
            if not cut: continue
            color, opacity, info, blend = a.colors.get('stroke'), a.opacity, a.info, a.blendmode
            page.delete_annot(a); changed += 1
            if pieces:
                n = page.add_ink_annot(pieces)
                if color: n.set_colors(stroke=color)
                n.set_border(width=width)
                if opacity is not None and opacity >= 0: n.set_opacity(opacity)
                n.set_info(title=info.get('title', ''), subject=info.get('subject', ''), content=info.get('content', ''))
                try: n.update(blend_mode=blend) if blend else n.update()
                except Exception: n.update()
        if not changed: doc.close(); return jsonify({"success": True, "changed": 0})
        out = _temp_out(fp, 'erase', json.dumps(d))
        doc.save(out, garbage=3, deflate=True); doc.close()
        return jsonify({"success": True, "changed": changed, "new_file": out})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- lasso: pick saved pen strokes up again ----------
def _inside(pt, poly):
    x, y, c, n = pt[0], pt[1], False, len(poly)
    for i in range(n):
        (x1, y1), (x2, y2) = poly[i], poly[(i + 1) % n]
        if (y1 > y) != (y2 > y) and x < (x2 - x1) * (y - y1) / ((y2 - y1) or 1e-9) + x1: c = not c
    return c

def _hex(rgb):
    return '#' + ''.join(f'{max(0, min(255, round(v * 255))):02x}' for v in (rgb or [0, 0, 0])[:3])

@app.route('/take_ink', methods=['POST'])
def take_ink():
    """Pen strokes mostly inside the lasso come out of the file and back to the app as strokes it can move,
    resize, recolor or turn into text (they go back in when the page is saved)."""
    d = request.json; fp = d['filepath']
    try:
        doc = fitz.open(fp); page = doc[int(d['page'])]; poly = [tuple(q) for q in d['polygon']]
        taken = []
        for a in list(page.annots() or []):
            kind = a.type[0]
            if kind not in (fitz.PDF_ANNOT_INK, fitz.PDF_ANNOT_POLYGON): continue
            tip = None
            try:
                t, v = doc.xref_get_key(a.xref, 'MyPDFKitTip'); tip = v.lstrip('/') if t == 'name' else None
            except Exception: pass
            opacity = a.opacity if a.opacity is not None and a.opacity >= 0 else 1
            if kind == fitz.PDF_ANNOT_POLYGON:
                if a.info.get('subject') != 'Pen': continue
                t, v = doc.xref_get_key(a.xref, 'MyPDFKitInk')
                if t != 'string': continue
                k = json.loads(v); pts = k['p']
                if sum(_inside(q, poly) for q in pts) < 0.6 * len(pts): continue
                taken.append([{"type": k.get('t') or 'fountain', "color": _hex(a.colors.get('fill') or a.colors.get('stroke')), "size": max(k['w']),
                               "opacity": opacity, "points": pts, "widths": k['w']}])
            else:
                lists = [[list(v) for v in l] for l in a.vertices]
                allp = [q for l in lists for q in l]
                if not allp or sum(_inside(q, poly) for q in allp) < 0.6 * len(allp): continue
                hl = tip == 'highlighter' or a.blendmode == 'Multiply'
                width = (a.border or {}).get('width') or 1
                taken.append([{"type": 'highlighter' if hl else (tip or 'pen'), "color": _hex(a.colors.get('stroke')), "size": width,
                               "opacity": 1 if hl else opacity, "points": l} for l in lists if len(l) > 1])
            page.delete_annot(a)
        strokes = [x for g in taken for x in g]
        if not strokes: doc.close(); return jsonify({"success": True, "strokes": []})
        out = _temp_out(fp, 'take', json.dumps(d))
        doc.save(out, garbage=3, deflate=True); doc.close()
        return jsonify({"success": True, "strokes": strokes, "new_file": out})
    except Exception as e: return jsonify({"error": str(e)}), 500

# ---------- handwriting to text (Windows' own handwriting recognizer) ----------
@app.route('/ink_to_text', methods=['POST'])
def ink_to_text():
    strokes = request.json.get('strokes') or []
    try:
        words = _recognize_ink(strokes)
        if not words: return jsonify({"error": "Couldn't read that writing"}), 400
        return jsonify({"text": ' '.join(w[0] for w in words if w), "words": words})
    except Exception as e:
        return jsonify({"error": "Handwriting recognition isn't available on this PC (" + str(e)[:120] + ")"}), 500

# ---------- places for others to sign ----------
@app.route('/add_fields', methods=['POST'])
def add_fields():
    """Sign-here boxes and other fields for the people who'll sign: [{page, kind, rect (as shown), signer}].
    kind: signature | initials | date | name | text | checkbox. They're real PDF form fields, so other
    readers show them too."""
    d = request.json; fp = d['filepath']
    try:
        doc = fitz.open(fp)
        names = {w.field_name for p in doc for w in (p.widgets() or [])}
        label = {'signature': 'Signature', 'initials': 'Initials', 'date': 'Date', 'name': 'Name', 'text': 'Text', 'checkbox': 'Check'}
        for f in d.get('fields', []):
            page = doc[int(f['page'])]
            kind, signer = f['kind'], (f.get('signer') or '').strip() or 'Signer'
            base = f"{label.get(kind, 'Field')} - {signer}"; name, n = base, 1
            while name in names: n += 1; name = f'{base} {n}'
            names.add(name)
            w = fitz.Widget()
            w.field_type = {'signature': fitz.PDF_WIDGET_TYPE_SIGNATURE, 'initials': fitz.PDF_WIDGET_TYPE_SIGNATURE,
                            'checkbox': fitz.PDF_WIDGET_TYPE_CHECKBOX}.get(kind, fitz.PDF_WIDGET_TYPE_TEXT)
            w.field_name = name
            w.field_label = {'signature': f'{signer}: sign here', 'initials': f'{signer}: initial here', 'date': f'{signer}: date',
                             'name': f'{signer}: your name'}.get(kind, f'{signer}: fill in')
            w.rect = fitz.Rect(f['rect']) * page.derotation_matrix
            if w.field_type == fitz.PDF_WIDGET_TYPE_TEXT: w.text_fontsize = 0
            w.border_color = (0.18, 0.44, 0.82); w.border_width = 1
            w.fill_color = (0.93, 0.96, 1.0)
            page.add_widget(w)
        out = _temp_out(fp, 'fields', json.dumps(d))
        doc.save(out, garbage=3, deflate=True); doc.close()
        return jsonify({"success": True, "new_file": out})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/sign_fields', methods=['POST'])
def sign_fields():
    """The places still waiting for a signature or initials, page by page."""
    try:
        doc = fitz.open(request.json['filepath']); out = []
        for i, page in enumerate(doc):
            for w in page.widgets(types=[fitz.PDF_WIDGET_TYPE_SIGNATURE]) or []:
                if w.field_value: continue
                r = w.rect * page.rotation_matrix
                out.append({"page": i + 1, "rect": [r.x0, r.y0, r.x1, r.y1], "name": w.field_name, "label": w.field_label or '',
                             "initials": w.field_name.lower().startswith('initials')})
        doc.close()
        return jsonify({"fields": out})
    except Exception as e: return jsonify({"error": str(e), "fields": []})

# ---------- translate, keeping the layout ----------
LANGS = {'es': 'Spanish', 'fr': 'French', 'de': 'German', 'it': 'Italian', 'pt': 'Portuguese', 'nl': 'Dutch', 'pl': 'Polish', 'ru': 'Russian',
         'uk': 'Ukrainian', 'tr': 'Turkish', 'ar': 'Arabic', 'he': 'Hebrew', 'hi': 'Hindi', 'zh-CN': 'Chinese (Simplified)', 'zh-TW': 'Chinese (Traditional)',
         'ja': 'Japanese', 'ko': 'Korean', 'vi': 'Vietnamese', 'th': 'Thai', 'id': 'Indonesian', 'tl': 'Filipino', 'sv': 'Swedish', 'el': 'Greek', 'en': 'English'}
TRANSLATE = {}

def _translate_texts(texts, target):
    """Texts translated with Google Translate, sent in batches of paragraphs."""
    import urllib.request, urllib.parse
    out, batch, size = [None] * len(texts), [], 0
    def flush():
        nonlocal batch, size
        if not batch: return
        q = '\n\n'.join(texts[i] for i in batch)
        req = urllib.request.Request('https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&dt=t&tl=' + urllib.parse.quote(target),
                                     data=urllib.parse.urlencode({'q': q}).encode(), headers={'User-Agent': 'Mozilla/5.0'})
        for attempt in range(3):
            try: j = json.loads(urllib.request.urlopen(req, timeout=30).read()); break
            except Exception:
                if attempt == 2: raise
                import time; time.sleep(1.5 * (attempt + 1))
        parts = ''.join(x[0] for x in j[0] if x and x[0]).split('\n\n')
        if len(parts) != len(batch):   # the paragraphs got merged: translate these one by one
            for i in batch:
                r = urllib.request.Request(req.full_url, data=urllib.parse.urlencode({'q': texts[i]}).encode(), headers={'User-Agent': 'Mozilla/5.0'})
                jj = json.loads(urllib.request.urlopen(r, timeout=30).read()); out[i] = ''.join(x[0] for x in jj[0] if x and x[0])
        else:
            for i, t in zip(batch, parts): out[i] = t.strip()
        batch, size = [], 0
    for i, t in enumerate(texts):
        if size + len(t) > 3500: flush()
        batch.append(i); size += len(t) + 2
    flush()
    return out

def _translate_run(job, fp, target, out_path):
    try:
        doc = fitz.open(fp)
        cjk = target in ('zh-CN', 'zh-TW', 'ja', 'ko')
        for pno, page in enumerate(doc):
            if job['cancel']: break
            job['current'] = pno + 1
            blocks = []
            for b in page.get_text('dict', flags=fitz.TEXTFLAGS_TEXT)['blocks']:
                lines = [l for l in b.get('lines', []) if any(sp['text'].strip() for sp in l['spans'])]
                if not lines: continue
                text = ' '.join(''.join(sp['text'] for sp in l['spans']).strip() for l in lines)
                text = re.sub(r'\s+', ' ', text).strip()
                if not re.search(r'[^\W\d_]', text): continue   # numbers and symbols stay as they are
                spans = [sp for l in lines for sp in l['spans'] if sp['text'].strip()]
                big = max(spans, key=lambda sp: len(sp['text']))
                dirs = lines[0]['dir']
                if abs(dirs[1]) > 0.01: continue   # turned text: leave it
                xs = [l['bbox'][0] for l in lines]
                align = 'center' if len(lines) > 1 and max(xs) - min(xs) > 8 and abs((lines[0]['bbox'][0] + lines[0]['bbox'][2]) / 2 - (lines[-1]['bbox'][0] + lines[-1]['bbox'][2]) / 2) < 3 else 'left'
                blocks.append({"rect": fitz.Rect(b['bbox']), "text": text, "size": big['size'], "color": big['color'],
                               "bold": bool(big['flags'] & 16) or 'bold' in big['font'].lower(), "italic": bool(big['flags'] & 2), "align": align,
                               "serif": bool(big['flags'] & 4)})
            if not blocks: continue
            done = _translate_texts([b['text'] for b in blocks], target)
            for b in blocks: page.add_redact_annot(b['rect'])
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE, text=fitz.PDF_REDACT_TEXT_REMOVE)
            for b, t in zip(blocks, done):
                if not t: continue
                c = b['color']; color = f"#{c:06x}"
                fam = 'serif' if b['serif'] else 'sans-serif'
                css = (f"* {{font-family:{fam}; font-size:{b['size']:.1f}px; color:{color}; line-height:1.15; margin:0; text-align:{b['align']};"
                       f"{' font-weight:bold;' if b['bold'] else ''}{' font-style:italic;' if b['italic'] else ''}}}")
                r = fitz.Rect(b['rect'].x0, b['rect'].y0, b['rect'].x1 + 2, b['rect'].y1 + 2)
                import html as _h
                page.insert_htmlbox(r, _h.escape(t), css=css, scale_low=0)
            job['done'] = pno + 1
        if not job['cancel']:
            try: doc.subset_fonts()
            except Exception: pass
            doc.save(out_path, garbage=3, deflate=True)
            job['path'] = out_path
        doc.close()
    except Exception as e:
        job['error'] = "Couldn't translate (" + str(e)[:160] + "). Check the internet connection."
    job['finished'] = True

@app.route('/translate_start', methods=['POST'])
def translate_start():
    d = request.json; fp = d['filepath']; target = d['target']
    orig = d.get('original_path') or fp
    out = get_unique_filename(os.path.join(os.path.dirname(orig), os.path.splitext(os.path.basename(orig))[0] + f" ({LANGS.get(target, target)}).pdf"))
    n = fitz.open(fp).page_count
    jid = uuid.uuid4().hex
    TRANSLATE[jid] = job = {"total": n, "done": 0, "current": 0, "finished": False, "cancel": False, "path": None, "error": None}
    threading.Thread(target=_translate_run, args=(job, fp, target, out), daemon=True).start()
    return jsonify({"id": jid, "total": n})

@app.route('/translate_status', methods=['POST'])
def translate_status():
    job = TRANSLATE.get(request.json['id'])
    if not job: return jsonify({"error": "No such job"}), 404
    if request.json.get('cancel'): job['cancel'] = True
    return jsonify(job)

@app.route('/formats', methods=['GET'])
def formats():
    return jsonify({"extensions": [e[1:] for e in OPENABLE_EXT]})

@app.route('/cleanup', methods=['POST'])
def cleanup():
    """Free disk space: the working folder keeps only what open tabs use and the previews of the recent
    files; everything else (old previews, edit copies, print pictures) is deleted."""
    d = request.json or {}
    keep = {os.path.normcase(os.path.abspath(p)) for p in d.get('keep_paths') or [] if p}
    for orig in d.get('keep_originals') or []:
        try:
            m = str(os.path.getmtime(orig))
            keep.add(os.path.normcase(os.path.join(TEMP_DIR, f"prev_{hashlib.md5((orig + m + CONVERTER_VERSION).encode('utf-8')).hexdigest()}.pdf")))
            keep.add(os.path.normcase(os.path.join(TEMP_DIR, f"prev_{hashlib.md5((orig + m + CONVERTER_VERSION).encode('utf-8')).hexdigest()}.pdf.builtin")))
            for tag in ('fixed_', 'open_'):
                keep.add(os.path.normcase(os.path.join(TEMP_DIR, tag + hashlib.md5((orig + m).encode()).hexdigest() + '.pdf')))
        except OSError: pass
    keep_dirs = {os.path.dirname(k) for k in keep}
    freed, removed, now = 0, 0, _time.time()
    for name in os.listdir(TEMP_DIR):
        full = os.path.join(TEMP_DIR, name)
        key = os.path.normcase(full)
        try:
            if key in keep or key in keep_dirs or now - os.path.getmtime(full) < 120: continue   # in use, or being made right now
            if os.path.isdir(full):
                size = sum(os.path.getsize(os.path.join(r, f)) for r, _, fs in os.walk(full) for f in fs)
                shutil.rmtree(full)
            else:
                size = os.path.getsize(full); os.remove(full)
            freed += size; removed += 1
        except OSError: continue   # open in another program: next time
    return jsonify({"success": True, "removed": removed, "freed_mb": round(freed / 1048576, 1)})

@app.route('/shutdown', methods=['POST'])
def shutdown():
    """The app is closing: close the hidden Word (nothing else is left running)."""
    try: OFFICE.shutdown()
    except Exception: pass
    return jsonify({"success": True})

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

    marker = temp_target + '.builtin'
    if os.path.exists(temp_target) and os.path.exists(marker) and (_office_has(path) or find_soffice()):
        for f in (temp_target, marker):
            try: os.remove(f)
            except OSError: pass
    note = None
    if not os.path.exists(temp_target):
        pythoncom.CoInitialize()
        try:
            if office_to_pdf(path, temp_target) == 'builtin':
                open(marker, 'w').close()
                note = "Opened without Microsoft Office, so it may look a little different from the original"
        except NoOfficeError as e:
            return jsonify({"error": str(e), "need_office": True}), 500
        except ValueError as e:
            return jsonify({"error": str(e)}), 400
        except Exception as e:
            return jsonify({"error": str(e)}), 500
        finally:
            pythoncom.CoUninitialize()

    return jsonify({"pdf_path": temp_target, **({"message": note} if note else {})})

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
    except NoOfficeError as e: return jsonify({"error": str(e), "need_office": True}), 500
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

class _MixedFont:
    """The PDF's own copy of a font for every letter it has, and a stand-in only for the letters it lacks.
    A PDF usually keeps just the letters it used, so typing a new one (a "J" in a heading that had none)
    shouldn't switch the whole line to another font."""
    def __init__(self, main, fallback): self.main, self.fallback = main, fallback
    def pick(self, ch): return self.main if self.main.has_glyph(ord(ch)) else self.fallback
    def has_glyph(self, c): return self.main.has_glyph(c) or self.fallback.has_glyph(c)
    def text_length(self, text, fontsize=11): return sum(self.pick(ch).text_length(ch, fontsize) for ch in text)
    def __getattr__(self, name): return getattr(self.main, name)

def _tw_add(tw, pos, text, font, size):
    """TextWriter.append that handles a _MixedFont (each run of letters in the font that has them)."""
    if not isinstance(font, _MixedFont): return tw.append(pos, text, font=font, fontsize=size)
    x, y = pos; run, cur = '', None
    for ch in text + '\x00':
        f = font.pick(ch) if ch != '\x00' else None
        if f is not cur and run:
            tw.append((x, y), run, font=cur, fontsize=size); x += cur.text_length(run, size); run = ''
        if ch == '\x00': break
        cur = f; run += ch

def _font_for(doc, page, span_font, text, cache, prefer_embedded=True):
    """A fitz.Font that draws `text` like `span_font`: the PDF's own copy when it has every letter,
    else the installed font of that name, else the PDF's copy for the letters it has with the closest
    built-in font for the rest, else the closest built-in font."""
    need = [ch for ch in text if not ch.isspace()]
    emb = None
    if prefer_embedded:
        emb = _embedded_font(doc, page, span_font, cache)
        if emb is not None and all(emb.has_glyph(ord(ch)) for ch in need): return emb
    path = _find_system_font(span_font)
    if path and os.path.exists(path):
        try:
            f = fitz.Font(fontfile=path)
            if all(f.has_glyph(ord(ch)) for ch in need): return f
        except Exception: pass
    if emb is not None and need and sum(1 for ch in need if emb.has_glyph(ord(ch))) >= len(need) / 2:
        return _MixedFont(emb, _font_for(doc, page, span_font, text, cache, prefer_embedded=False))
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
                tw.append((x + (px - o0.x) / hs, y), ch, font=f.pick(ch) if isinstance(f, _MixedFont) else f, fontsize=size)
            tw.write_text(page, color=rgb, morph=morph, render_mode=mode, opacity=op)
        else:                     # turned text: each letter at its spot along the line, turned in place
            for px, ch in placed:
                if ch.isspace(): continue
                t = px - o0.x
                at = fitz.Point(x + t * ddx, y + t * ddy)
                tw = fitz.TextWriter(page.rect)
                tw.append(at, ch, font=f.pick(ch) if isinstance(f, _MixedFont) else f, fontsize=size)
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
                _tw_add(tw, (x, y), w, fo, sz)
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
                for w in list(page.widgets(types=[fitz.PDF_WIDGET_TYPE_SIGNATURE]) or []):
                    if not w.field_value and (w.rect & rect).get_area() > 0.4 * w.rect.get_area():
                        page.delete_widget(w)
            except Exception as e:
                print(f"Skipped image: {e}")

        # 4. Burn Drawings ("H" & Pen)
        hl_groups = {} 
        pens = []

        polys, shapes = [], []
        for draw in data.get('drawings', []):
            if draw.get('type') == 'shape': shapes.append(draw); continue
            if draw.get('outline'):   # pens whose line gets thicker and thinner: the shape of the ink
                polys.append(draw); continue
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
                doc.xref_set_key(annot.xref, 'MyPDFKitTip', '/highlighter')
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
                annot.set_info(subject='Pen', title=os.environ.get('USERNAME', ''))
                annot.update()
                doc.xref_set_key(annot.xref, 'MyPDFKitTip', '/' + re.sub(r'[^a-z]', '', str(draw.get('type') or 'pen')))
            except Exception as e:
                print(f"Skipped pen: {e}")

        for draw in shapes:   # drawn with the Shapes tool: standard PDF shapes, which MyPDFKit can open for editing again
            try:
                _add_shape(doc, page, draw, x_off, y_off)
            except Exception as e:
                print(f"Skipped shape: {e}")

        for draw in polys:
            try:
                center = [(float(p[0]) + x_off, float(p[1]) + y_off) for p in draw['points']]
                _add_pen_shape(doc, page, _ink_outline(center, draw['widths']) if draw.get('widths') else
                               [(float(p[0]) + x_off, float(p[1]) + y_off) for p in draw['outline']],
                               parse_color(draw['color']), float(draw.get('opacity', 1.0)), center, draw.get('widths'), tip=draw.get('type'))
            except Exception as e:
                print(f"Skipped pen shape: {e}")

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

# ---------- printing: sheets the size of the printer's paper ----------
# The printer decides the paper (as in Chrome): Windows says what paper it holds, its resolution and the part
# it can print on. Each sheet is made here at exactly that size, with the page (or just its label) placed on
# it as vectors, so the preview shows what will come out and barcodes stay sharp.
def _printer_info(name):
    import win32print, win32gui
    h = win32print.OpenPrinter(name)
    try:
        dm = win32print.GetPrinter(h, 2)['pDevMode']
        hdc = win32gui.CreateDC('WINSPOOL', name, dm)
        try:
            g = lambda i: win32print.GetDeviceCaps(hdc, i)
            dx, dy = g(88), g(90)                                   # LOGPIXELSX/Y
            pw, ph = g(110) / dx * 72, g(111) / dy * 72              # PHYSICALWIDTH/HEIGHT
            ox, oy = g(112) / dx * 72, g(113) / dy * 72              # PHYSICALOFFSETX/Y
            hr, vr = g(8) / dx * 72, g(10) / dy * 72                 # HORZRES/VERTRES
        finally: win32gui.DeleteDC(hdc)
        return {'paper': [round(pw, 2), round(ph, 2)], 'printable': [round(ox, 2), round(oy, 2), round(ox + hr, 2), round(oy + vr, 2)],
                'dpi': int(min(dx, dy)), 'form': dm.FormName or '', 'landscape': dm.Orientation == 2}
    finally: win32print.ClosePrinter(h)

@app.route('/printer_info', methods=['POST'])
def printer_info():
    try: return jsonify(_printer_info(request.json['name']))
    except Exception as e: return jsonify({'error': str(e)})

def _flag(doc, xref, key='F'):
    t, v = doc.xref_get_key(xref, key)
    try: return int(v) if t == 'int' else 0
    except ValueError: return 0

def _pdf_str(t):
    return '(' + t.replace('\\', '\\\\').replace('(', '\\(').replace(')', '\\)') + ')'

def _draw_widget(w, doc=None):
    """Make a field's picture of its value, the way Chrome shows it. PyMuPDF draws the box, but it always puts
    the text 2 points from the left and ignores the field's alignment (often set on its parent field), so a
    one-line value is placed here: aligned as the field says, sized to fit when the size is "auto", and in the
    middle up and down (also in boxes that allow several lines, where it would otherwise touch the top edge)."""
    doc = doc or w.parent.parent
    w.update()
    v = w.field_value
    if w.field_type != fitz.PDF_WIDGET_TYPE_TEXT or not isinstance(v, str) or not v.strip() or '\n' in v or '\r' in v: return
    try:
        v.encode('cp1252')
    except UnicodeEncodeError: return          # letters Helvetica can't show: keep PyMuPDF's drawing
    try:
        t, ref = doc.xref_get_key(w.xref, 'AP/N')
        if t != 'xref': return
        ap = int(ref.split()[0]); st = doc.xref_stream(ap).decode('latin-1')
        bt, et = st.find('BT'), st.rfind('ET')
        if bt < 0 or et < 0 or st.count(' Tj') != 1: return
        font = re.search(r'/(\w+) ([\d.]+) Tf', st[bt:et])
        color = re.search(r'([\d.]+ [\d.]+ [\d.]+ rg|[\d.]+ g)\s', st[bt:et])
        fname = font.group(1) if font else 'Helv'
        bw, bh = w.rect.width, w.rect.height
        size = w.text_fontsize or 0
        multi = bool(w.field_flags & 4096)
        one = fitz.get_text_length(v, 'helv', 1)
        if not size:   # auto: as big as the box allows (12 in boxes for several lines), smaller if it's long
            size = 12 if multi else min(bh * 0.68, 24)
        if one * size > bw - 4: size = max(4, (bw - 4) / one)
        if multi and one * size > bw - 4: return   # it wraps: leave PyMuPDF's lines
        tw = one * size
        q = _inherited_q(doc, w.xref)
        x = 2 if q == 0 else (bw - tw) / 2 if q == 1 else bw - 2 - tw
        y = (bh - 0.718 * size) / 2            # the capitals' middle on the box's middle
        text = f"BT\n{color.group(1) if color else '0 g'}\n/{fname} {size:.2f} Tf\n{x:.2f} {y:.2f} Td\n{_pdf_str(v)} Tj\n"
        doc.update_stream(ap, (st[:bt] + text + st[et:]).encode('latin-1'))
    except Exception as e: print('field drawing kept as is:', e, file=sys.stderr)

def _fix_widget_looks(doc, everything=False):
    """Fields filled in another app often have no picture of their value (or ask viewers to make one):
    make it, so the value shows and prints everywhere."""
    need = everything
    try:
        t, v = doc.xref_get_key(doc.pdf_catalog(), 'AcroForm/NeedAppearances')
        need = need or v == 'true'
    except Exception: pass
    for page in doc:
        for w in list(page.widgets() or []):
            try:
                if w.field_type in (fitz.PDF_WIDGET_TYPE_BUTTON, fitz.PDF_WIDGET_TYPE_SIGNATURE): continue
                no_ap = doc.xref_get_key(w.xref, 'AP')[0] == 'null'
                if (need or no_ap) and w.field_value not in (None, '', 'Off'): _draw_widget(w, doc)
            except Exception: continue

def _printable_copy(path, comments=True):
    """The document as it prints: field values drawn in, hidden and not-for-print comments and fields left out
    (and every comment, when comments=False)."""
    doc = fitz.open(path)
    if doc.needs_pass: raise ValueError('This PDF is locked with a password')
    _fix_widget_looks(doc)
    for page in doc:
        for a in list(page.annots() or []):
            f = a.flags
            if f & 2 or not f & 4 or not comments: page.delete_annot(a)   # hidden, not marked to print, or comments left out
        for w in list(page.widgets() or []):
            f = _flag(doc, w.xref)
            if f & 2 or not f & 4: page.delete_widget(w)
    try: doc.bake(annots=True, widgets=True)
    except Exception as e: print('bake skipped:', e, file=sys.stderr)
    return doc

def _is_blank(page):
    pix = page.get_pixmap(dpi=18, colorspace=fitz.csGRAY, annots=True)
    return min(pix.samples) >= 246 if pix.samples else True

def _shape_is(w, h, sizes, tol):
    a, b = sorted((w, h))
    return any(abs(a / b - x / y) <= tol * (x / y) for x, y in sizes)

LABEL_SHAPES = [(4, 6), (4, 4)]

def _label_regions(page):
    """Where the label(s) are on a page, in page coordinates, and how they were found:
    a picture shaped like a label, a drawn box of label size, the block around a barcode, or the content."""
    unrot = lambda r: fitz.Rect(r) * page.rotation_matrix   # to the page as shown
    pr = page.rect
    # a) a picture with a label's shape (4x6 or 4x4), at least 2 inches on its short side
    pics = [fitz.Rect(i['bbox']) for i in page.get_image_info() if i.get('bbox')]
    hits = [r for r in pics if min(r.width, r.height) >= 144 and _shape_is(r.width, r.height, LABEL_SHAPES, 0.04)]
    if hits: return [unrot(r) for r in hits], 'picture'
    drawings = page.get_drawings()
    # b) a drawn border or dashed cut line of label size
    boxes = []
    for dr in drawings:
        r = dr['rect']
        if dr.get('color') is None or r.width < 100 or r.height < 100: continue
        a, b = sorted((r.width, r.height))
        if any(abs(a - x * 72) <= 0.12 * x * 72 and abs(b - y * 72) <= 0.12 * y * 72 for x, y in ((4, 6), (4, 4))):
            if not any(abs(r.x0 - q.x0) < 4 and abs(r.y0 - q.y0) < 4 for q in boxes): boxes.append(r)
    if boxes: return [unrot(r) for r in boxes], 'border'
    # c) the content block around a barcode: many thin bars side by side
    bars = [dr['rect'] for dr in drawings if dr.get('fill') is not None and dr['rect'].width <= 14 and dr['rect'].height >= 12
            and dr['rect'].height >= 3 * dr['rect'].width]
    codes = []
    for r in sorted(bars, key=lambda r: (round(r.y0), r.x0)):
        for c in codes:
            if abs(c['y0'] - r.y0) < 3 and abs(c['y1'] - r.y1) < 3 and r.x0 - c['rect'].x1 < 24:
                c['rect'] |= r; c['n'] += 1; break
        else: codes.append({'rect': fitz.Rect(r), 'y0': r.y0, 'y1': r.y1, 'n': 1})
    codes = [c['rect'] for c in codes if c['n'] >= 8 and c['rect'].width >= 40]
    parts = [fitz.Rect(b[:4]) for b in page.get_text('blocks') if b[4].strip()] + pics + \
            [dr['rect'] for dr in drawings if dr['rect'].get_area() < 0.8 * pr.get_area() and dr['rect'].width > 0.5]
    def clusters(rects, gap=26):
        groups = [fitz.Rect(r) for r in rects if not r.is_empty]
        merged = True
        while merged:
            merged = False
            for i in range(len(groups)):
                for j in range(i + 1, len(groups)):
                    a, b = groups[i], groups[j]
                    if a.x0 - gap <= b.x1 and b.x0 - gap <= a.x1 and a.y0 - gap <= b.y1 and b.y0 - gap <= a.y1:
                        groups[i] = a | b; groups.pop(j); merged = True; break
                if merged: break
        return groups
    if codes:
        found = [g for g in clusters(parts + codes) if any(g.intersects(c) for c in codes)]
        if found: return [unrot(g + (-4, -4, 4, 4)) & pr for g in found], 'barcode'
    # d) otherwise the page's content without its white margins
    pix = page.get_pixmap(dpi=50, colorspace=fitz.csGRAY, annots=True)
    w, h, s = pix.width, pix.height, pix.samples
    rows = [y for y in range(h) if min(s[y * w:(y + 1) * w]) < 235]
    if not rows: return [], 'blank'
    cols = [x for x in range(w) if min(s[x::w]) < 235]
    k = 72 / 50
    r = fitz.Rect(cols[0] * k - 6, rows[0] * k - 6, (cols[-1] + 1) * k + 6, (rows[-1] + 1) * k + 6) & page.rect
    return [r], 'trim'

@app.route('/label_regions', methods=['POST'])
def label_regions():
    d = request.json
    try:
        doc = _printable_copy(d['filepath'])
        pages = d.get('pages') or list(range(1, doc.page_count + 1))
        regions, how, blank = {}, {}, []
        for n in pages:
            page = doc[n - 1]
            if _is_blank(page): blank.append(n); continue
            rs, kind = _label_regions(page)
            regions[n] = [[round(v, 2) for v in r] for r in rs]; how[n] = kind
        return jsonify({'regions': regions, 'how': how, 'blank': blank})
    except Exception as e: return jsonify({'error': str(e)}), 500

def _content_edges(page):
    """Edges of what's on the page (text, pictures, lines), for snapping a dragged area to them."""
    xs, ys = set(), set()
    rects = [fitz.Rect(b[:4]) for b in page.get_text('blocks')] + [fitz.Rect(i['bbox']) for i in page.get_image_info()] + \
            [dr['rect'] for dr in page.get_drawings()]
    for r in rects:
        r = r * page.rotation_matrix
        if r.is_empty: continue
        xs.update((round(r.x0, 1), round(r.x1, 1))); ys.update((round(r.y0, 1), round(r.y1, 1)))
    return sorted(xs), sorted(ys)

@app.route('/content_edges', methods=['POST'])
def content_edges():
    d = request.json
    try:
        doc = fitz.open(d['filepath']); xs, ys = _content_edges(doc[int(d['page']) - 1])
        return jsonify({'x': xs, 'y': ys})
    except Exception as e: return jsonify({'error': str(e), 'x': [], 'y': []})

# Label sheets for Letter paper (Avery and the same sizes from other brands), in inches:
# label width, height, columns, rows, left margin, top margin, step across, step down
LABEL_SHEETS = {
    '5160': ('Address 2⅝ × 1 in, 30 per sheet (5160, 8160)', 2.625, 1, 3, 10, 0.1875, 0.5, 2.75, 1),
    '5163': ('Shipping 4 × 2 in, 10 per sheet (5163, 8163)', 4, 2, 2, 5, 0.15625, 0.5, 4.1875, 2),
    '5164': ('Shipping 4 × 3⅓ in, 6 per sheet (5164, 8164)', 4, 3.3333, 2, 3, 0.15625, 0.5, 4.1875, 3.3333),
    '5168': ('Shipping 3½ × 5 in, 4 per sheet (5168, 8168)', 3.5, 5, 2, 2, 0.5, 0.5, 4.0, 5.0),
    '5126': ('Shipping half sheet 8½ × 5½ in, 2 per sheet (5126, 8126)', 8.5, 5.5, 1, 2, 0, 0, 8.5, 5.5),
    '5167': ('Return address 1¾ × ½ in, 80 per sheet (5167, 8167)', 1.75, 0.5, 4, 20, 0.28125, 0.5, 2.0625, 0.5),
    '22806': ('Square 2 × 2 in, 12 per sheet (22806)', 2, 2, 3, 4, 0.625, 0.625, 2.625, 2.625),
}

@app.route('/print_layouts', methods=['GET'])
def print_layouts():
    return jsonify({'label_sheets': [[k, v[0]] for k, v in LABEL_SHEETS.items()]})

def _place_in(page, src, n, clip, area, orient='auto', scale='fit', pct=100, rot_min_gain=1.15):
    """Draw page n of src (or its clip, in page coordinates as shown) into area of page: turned when asked or
    when it fits clearly bigger, sized by scale ('fit', 'shrink', 'actual' or 'pct'), centered."""
    sp = src[n - 1]
    box = fitz.Rect(clip) if clip else sp.rect
    cw, ch = box.width, box.height
    aw, ah = area.width, area.height
    fit = lambda rot: min(aw / (ch if rot else cw), ah / (cw if rot else ch))
    if orient == 'auto': rot = 90 if fit(True) > fit(False) * rot_min_gain else 0
    else: rot = 0 if (orient == 'landscape') == (page.rect.width > page.rect.height) else 90
    f = fit(rot) if scale == 'fit' else min(1.0, fit(rot)) if scale == 'shrink' else (pct / 100.0 if scale == 'pct' else 1.0)
    tw, th = ((ch, cw) if rot else (cw, ch))
    tw, th = tw * f, th * f
    cx, cy = (area.x0 + area.x1) / 2, (area.y0 + area.y1) / 2
    target = fitz.Rect(cx - tw / 2, cy - th / 2, cx + tw / 2, cy + th / 2)
    page.show_pdf_page(target, src, n - 1, clip=(box * sp.derotation_matrix) if clip else None, rotate=-rot if rot else 0)
    return target

def _items(d):
    """What goes on the sheets: (document, page, area) for every page (or label) of every file asked for."""
    files = d.get('files') or [d['filepath']]
    mode, regions = d.get('mode', 'page'), {int(k): v for k, v in (d.get('regions') or {}).items()}
    items, skipped = [], []
    for fi, fp in enumerate(files):
        src = _printable_copy(fp, comments=d.get('comments', True))
        pages = (d.get('pages') if fi == 0 and not d.get('files') else None) or list(range(1, src.page_count + 1))
        for n in pages:
            if n < 1 or n > src.page_count: continue
            tag = n if len(files) == 1 else f'{fi + 1}:{n}'
            if d.get('skip_blank') and _is_blank(src[n - 1]): skipped.append(tag); continue
            if mode == 'label':
                rs = regions.get(n) if len(files) == 1 else None
                if rs is None: rs, _ = _label_regions(src[n - 1])
                if not rs: skipped.append(tag); continue
                items += [(src, n, r, tag) for r in rs]
            else: items.append((src, n, None, tag))
    return items, skipped

def _sheets(d):
    paper = d.get('paper') or [612, 792]
    pw, ph = paper
    area = fitz.Rect(d.get('printable') or (0, 0, pw, ph)) & fitz.Rect(0, 0, pw, ph)
    if d.get('margin'): area = area + (d['margin'], d['margin'], -d['margin'], -d['margin'])
    items, skipped = _items(d)
    layout = d.get('layout') or 'single'
    orient, scale, pct = d.get('orient', 'auto'), d.get('scale', 'fit'), float(d.get('percent') or 100)
    out, plan = fitz.open(), []
    gap = 6
    if layout == 'nup':   # several pages on each sheet, in a grid that gives them the most room
        nup = int(d.get('nup') or 2)
        grids = {2: (1, 2), 4: (2, 2), 6: (2, 3), 9: (3, 3), 16: (4, 4)}
        c, r = grids.get(nup, (1, 2))
        if items:
            s0, n0, clip0, _ = items[0]
            b = fitz.Rect(clip0) if clip0 else s0[n0 - 1].rect
            def cell_fit(cols, rows):
                cw_, ch_ = (area.width - gap * (cols - 1)) / cols, (area.height - gap * (rows - 1)) / rows
                return max(min(cw_ / b.width, ch_ / b.height), min(cw_ / b.height, ch_ / b.width))
            if cell_fit(r, c) > cell_fit(c, r): c, r = r, c
        cw_, ch_ = (area.width - gap * (c - 1)) / c, (area.height - gap * (r - 1)) / r
        for k, (src, n, clip, tag) in enumerate(items):
            if k % (c * r) == 0: page = out.new_page(width=pw, height=ph)
            i = k % (c * r); col, row = i % c, i // c
            cell = fitz.Rect(area.x0 + col * (cw_ + gap), area.y0 + row * (ch_ + gap), area.x0 + col * (cw_ + gap) + cw_, area.y0 + row * (ch_ + gap) + ch_)
            _place_in(page, src, n, clip, cell, 'auto', 'fit', rot_min_gain=1.0)
            if i == 0: plan.append((tag, clip))
    elif layout == 'booklet':   # fold the stack in half: two pages side by side on each side of a sheet
        lw, lh = max(pw, ph), min(pw, ph)
        la = fitz.Rect(area.y0, area.x0, area.y1, area.x1) if pw < ph else fitz.Rect(area)
        seq = items + [None] * ((-len(items)) % 4)
        total = len(seq)
        for sheet in range(total // 4):
            for left, right in ((total - 1 - 2 * sheet, 2 * sheet), (2 * sheet + 1, total - 2 - 2 * sheet)):
                page = out.new_page(width=lw, height=lh)
                half = la.width / 2
                for idx, cell in ((left, fitz.Rect(la.x0, la.y0, la.x0 + half - gap / 2, la.y1)), (right, fitz.Rect(la.x0 + half + gap / 2, la.y0, la.x1, la.y1))):
                    if seq[idx] is None: continue
                    src, n, clip, tag = seq[idx]
                    _place_in(page, src, n, clip, cell, 'auto', 'fit', rot_min_gain=10)
                plan.append((seq[right][3] if seq[right] else (seq[left][3] if seq[left] else None), None))
    elif layout == 'poster':   # one page across several sheets, overlapping a little, with lines to cut along
        cols, rows = int(d.get('cols') or 2), int(d.get('rows') or 2)
        ov = float(d.get('overlap') or 18)
        big = fitz.Rect(0, 0, cols * area.width - (cols - 1) * ov, rows * area.height - (rows - 1) * ov)
        for src, n, clip, tag in items:
            b = fitz.Rect(clip) if clip else src[n - 1].rect
            rot = 90 if min(big.width / b.height, big.height / b.width) > min(big.width / b.width, big.height / b.height) * 1.001 else 0
            tw, th = (b.height, b.width) if rot else (b.width, b.height)
            f = min(big.width / tw, big.height / th) if scale != 'pct' else pct / 100
            tw, th = tw * f, th * f
            ox, oy = (big.width - tw) / 2, (big.height - th) / 2
            for row in range(rows):
                for col in range(cols):
                    page = out.new_page(width=pw, height=ph)
                    x0, y0 = col * (area.width - ov), row * (area.height - ov)
                    target = fitz.Rect(area.x0 - x0 + ox, area.y0 - y0 + oy, area.x0 - x0 + ox + tw, area.y0 - y0 + oy + th)
                    page.show_pdf_page(target, src, n - 1, clip=(b * src[n - 1].derotation_matrix) if clip else None, rotate=-rot if rot else 0)
                    # outside the printable part stays white; the overlap edge is marked to cut along
                    sh = page.new_shape()
                    if col: sh.draw_line((area.x0 + ov, area.y0), (area.x0 + ov, area.y1))
                    if row: sh.draw_line((area.x0, area.y0 + ov), (area.x1, area.y0 + ov))
                    sh.finish(color=(0.55, 0.55, 0.55), width=0.4, dashes='[3 3] 0'); sh.commit()
                    page.insert_text((area.x0 + 3, area.y1 - 3), f'{chr(65 + row)}{col + 1}', fontsize=7, color=(0.5, 0.5, 0.5))
                    plan.append((tag, clip))
    elif layout == 'sheet':   # a sheet of stickers (Avery...): one item per sticker, from the sticker chosen
        key = d.get('sheet') or '5163'
        _, lw_, lh_, cols, rows, ml, mt, sx, sy = LABEL_SHEETS.get(key, LABEL_SHEETS['5163'])
        pw, ph = 612, 792
        slots = [fitz.Rect((ml + c_ * sx) * 72, (mt + r_ * sy) * 72, (ml + c_ * sx + lw_) * 72, (mt + r_ * sy + lh_) * 72) for r_ in range(rows) for c_ in range(cols)]
        k = max(0, int(d.get('start') or 1) - 1) % len(slots)
        page = None
        for src, n, clip, tag in items:
            if page is None or k >= len(slots):
                if page is not None: k = 0
                page = out.new_page(width=pw, height=ph); plan.append((tag, clip))
            _place_in(page, src, n, clip, slots[k] + (2, 2, -2, -2), 'auto', 'fit', rot_min_gain=1.0)
            k += 1
    elif layout == 'receipt':   # a receipt roll: the full width, as long as the page needs
        m = 4
        for src, n, clip, tag in items:
            b = fitz.Rect(clip) if clip else _content_box(src[n - 1])
            f = (pw - 2 * m) / b.width
            page = out.new_page(width=pw, height=b.height * f + 2 * m)
            _place_in(page, src, n, b, fitz.Rect(m, m, pw - m, page.rect.height - m), 'portrait' if page.rect.width < page.rect.height else 'landscape', 'fit')
            plan.append((tag, b))
    else:
        for src, n, clip, tag in items:
            page = out.new_page(width=pw, height=ph)
            _place_in(page, src, n, clip, area, orient, scale, pct)
            plan.append((tag, clip))
    out.item_count = len(items)
    return out, plan, skipped

def _content_box(page):
    """The part of a page with something on it (white margins left off)."""
    rs, kind = _label_regions(page)
    if kind == 'trim' and rs: return rs[0]
    pix = page.get_pixmap(dpi=50, colorspace=fitz.csGRAY, annots=True)
    w, h, s = pix.width, pix.height, pix.samples
    rows = [y for y in range(h) if min(s[y * w:(y + 1) * w]) < 235]
    if not rows: return page.rect
    cols = [x for x in range(w) if min(s[x::w]) < 235]
    k = 72 / 50
    return fitz.Rect(cols[0] * k - 4, rows[0] * k - 4, (cols[-1] + 1) * k + 4, (rows[-1] + 1) * k + 4) & page.rect

@app.route('/print_prepare', methods=['POST'])
def print_prepare():
    """The sheets to print, as a PDF (the preview shows exactly these)."""
    d = request.json
    try:
        out, plan, skipped = _sheets(d)
        if not out.page_count: return jsonify({'error': 'Nothing to print on the pages chosen', 'skipped': skipped}), 400
        files = d.get('files') or [d['filepath']]
        path = os.path.join(TEMP_DIR, 'sheets_' + hashlib.md5((json.dumps(d, sort_keys=True) + ''.join(str(os.path.getmtime(f)) for f in files)).encode()).hexdigest() + '.pdf')
        out.save(path, garbage=3, deflate=True)
        sizes = [[round(p.rect.width, 2), round(p.rect.height, 2)] for p in out]
        return jsonify({'path': path, 'sheets': [{'page': n, 'rect': list(r) if r else None} for n, r in plan], 'skipped': skipped, 'sizes': sizes,
                        'count': getattr(out, 'item_count', len(plan))})
    except Exception as e: return jsonify({'error': str(e)}), 500

@app.route('/print_render', methods=['POST'])
def print_render():
    """Pictures of the sheets at the printer's own resolution (sharp barcodes on thermal printers), PNG."""
    d = request.json
    try:
        doc = fitz.open(d['path'])
        dpi = int(d.get('dpi') or 300)
        dpi = dpi if 150 <= dpi <= 300 else 300
        out = tempfile.mkdtemp(dir=TEMP_DIR, prefix='print_')
        first, count = int(d.get('first', 0)), int(d.get('count', doc.page_count))
        order = [int(i) for i in d['indices'] if 0 <= int(i) < doc.page_count] if d.get('indices') else range(first, min(doc.page_count, first + count))
        pages = []
        for i in order:
            p = doc[i]
            f = os.path.join(out, f'{i + 1}.png')
            pix = p.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY if d.get('bw') or d.get('toner') else fitz.csRGB)
            if d.get('toner'): pix.tint_with(0x6E6E6E, 0xFFFFFF)   # black prints as mid gray: about half the toner
            pix.save(f)
            pages.append({'path': f, 'w': p.rect.width, 'h': p.rect.height})
        return jsonify({'success': True, 'pages': pages, 'total': doc.page_count})
    except Exception as e: return jsonify({'error': str(e)}), 500

@app.route('/save_labels', methods=['POST'])
def save_labels():
    """Every label as its own 4x6 (or chosen size) page, as vectors, saved next to the original."""
    d = dict(request.json)
    try:
        size = d.get('size') or [288, 432]
        d.update({'paper': size, 'printable': None, 'margin': 0, 'orient': 'auto', 'scale': 'fit', 'skip_blank': True,
                  'mode': d.get('mode') or 'label'})
        out, plan, skipped = _sheets(d)
        if not out.page_count: return jsonify({'error': 'No labels found on those pages'}), 400
        orig = d.get('original_path') or d.get('filepath') or d['files'][0]
        base = os.path.splitext(os.path.basename(orig))[0]
        tag = f'{size[0] / 72:g}x{size[1] / 72:g}'
        path = d.get('out') or get_unique_filename(os.path.join(os.path.dirname(orig), f'{base} ({tag} labels).pdf'))
        out.save(path, garbage=3, deflate=True)
        return jsonify({'success': True, 'path': path, 'count': out.page_count, 'skipped': skipped})
    except Exception as e: return jsonify({'error': str(e)}), 500

@app.route('/area_picture', methods=['POST'])
def area_picture():
    """A picture of an area of a page (for copying), at 300 dpi."""
    d = request.json
    try:
        doc = _printable_copy(d['filepath']); page = doc[int(d['page']) - 1]
        clip = fitz.Rect(d['rect']) * page.derotation_matrix
        pix = page.get_pixmap(dpi=int(d.get('dpi', 300)), clip=clip)
        f = os.path.join(TEMP_DIR, 'area_' + hashlib.md5(json.dumps(d).encode()).hexdigest() + '.png'); pix.save(f)
        return jsonify({'path': f})
    except Exception as e: return jsonify({'error': str(e)}), 500

@app.route('/render_pages', methods=['POST'])
def render_pages():
    """Page pictures for printing: JPEGs at print quality, with each page's size in points."""
    fp = request.json['filepath']
    try:
        doc = fitz.open(fp)
        out = tempfile.mkdtemp(dir=TEMP_DIR, prefix='print_')
        pages = []
        wanted = request.json.get('pages') or range(1, doc.page_count + 1)   # 1-based; all when not given
        for n in wanted:
            i, p = n - 1, doc[n - 1]
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
    names = data.get('names') or []
    original_folder = os.path.dirname(data.get('first_original_path', ''))
    try:
        merged_doc = fitz.open(); toc = []
        for i, pdf_path in enumerate(files):
            if os.path.exists(pdf_path):
                pdf_to_insert = fitz.open(pdf_path)
                start = merged_doc.page_count
                merged_doc.insert_pdf(pdf_to_insert)
                if data.get('bookmarks'):   # each file starts a bookmark (its own bookmarks go under it)
                    toc.append([1, os.path.splitext(names[i] if i < len(names) else os.path.basename(pdf_path))[0], start + 1])
                    for lvl, title, pg in pdf_to_insert.get_toc(simple=True): toc.append([lvl + 1, title, pg + start])
                pdf_to_insert.close()
        if toc:
            try: merged_doc.set_toc(toc)
            except Exception as e: print('bookmarks skipped:', e, file=sys.stderr)
        name = re.sub(r'[\\/:*?"<>|]+', '-', (data.get('name') or 'Combined_Document').strip()) or 'Combined_Document'
        if not name.lower().endswith('.pdf'): name += '.pdf'
        out_path = get_unique_filename(os.path.join(original_folder, name))
        merged_doc.save(out_path, garbage=3, deflate=True)
        merged_doc.close()
        return jsonify({"success": True, "path": out_path, "message": f"Saved {os.path.basename(out_path)}"})
    except Exception as e: return jsonify({"error": str(e)}), 500

def _safe_copy(src, dst):
    """Write a PDF where the user keeps it without ever leaving a broken file: a temporary copy first,
    checked to open with all its pages, then given its real name."""
    with fitz.open(src) as d: pages = d.page_count
    part = dst + '.part'
    shutil.copyfile(src, part)
    try:
        with fitz.open(part) as d:
            if d.page_count != pages or d.is_repaired: raise RuntimeError("The saved copy didn't check out, so it wasn't kept")
        os.replace(part, dst)
    finally:
        if os.path.exists(part): os.remove(part)

@app.route('/save_modified', methods=['POST'])
def save_modified():
    data = request.json
    current_temp_pdf = data.get('filepath')
    original_path = data.get('original_path')
    try:
        original_folder = os.path.dirname(original_path)
        base_name = os.path.splitext(os.path.basename(original_path))[0]
        out_path = get_unique_filename(os.path.join(original_folder, f"{base_name}_Modified.pdf"))
        _safe_copy(current_temp_pdf, out_path)
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

def _exit_with_parent():
    """Close when the app that started us is gone, even if it never said goodbye (a crash, Task Manager),
    so no old backend is left holding Office and the port."""
    import ctypes
    pid = int(os.environ.get('MYPDFKIT_PARENT_PID') or 0)
    if not pid: return
    k = ctypes.WinDLL('kernel32', use_last_error=True)
    k.OpenProcess.restype = ctypes.c_void_p
    k.WaitForSingleObject.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
    h = k.OpenProcess(0x00100000, False, pid)   # SYNCHRONIZE: enough to wait for it
    if not h:
        if ctypes.get_last_error() == 87: os._exit(0)   # no such process: the app is already gone
        return
    k.WaitForSingleObject(h, 0xFFFFFFFF)
    try: OFFICE.shutdown()
    except Exception: pass
    os._exit(0)

def _bind_port_alone():
    """Werkzeug lets a second server share a port on Windows; then requests go to whichever one Windows
    picks (an old leftover backend kept answering the new app). Bind so nobody else can use it."""
    import socket
    import werkzeug.serving as ws
    bind = ws.BaseWSGIServer.server_bind
    def server_bind(self):
        self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        bind(self)
    ws.BaseWSGIServer.allow_reuse_address = False
    ws.BaseWSGIServer.server_bind = server_bind

if __name__ == '__main__':
    _threading.Thread(target=_exit_with_parent, daemon=True).start()
    _threading.Thread(target=_prewarm_office, daemon=True).start()
    _bind_port_alone()
    app.run(host='127.0.0.1', port=int(os.environ.get('MYPDFKIT_PORT') or 5000), threaded=True)