"""Quality checks from the market research (OCR, layout, speed, bloat):  python tests/qa_tests.py
- OCR accuracy: a scan of known text is read back word for word (Windows' own text recognition)
- Load speed: every test PDF opens and shows its first page quickly, including a 220-page file
- Summary: the free summary finds a document's main points and amounts
- Bloat: the built app starts fast and stays light (needs a build in release/win-unpacked)
Layout stability while editing is checked line by line by edit_fidelity.py (run_tests.py runs it)."""
import os, sys, time, re, tempfile, shutil, subprocess
T = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(T)
sys.path.insert(0, ROOT)
import server, pymupdf

c = server.app.test_client()
work = tempfile.mkdtemp(prefix='mypdfkit-qa-')
results = []
def check(name, ok, info=''):
    results.append(ok); print(('  PASS ' if ok else '  FAIL ') + name + (f'  ({info})' if info else ''))

print('OCR accuracy')
TRUTH = ['PACKING LIST 2026 Order 4471', 'Ship to: Maria Lopez, 8402 Main Street', 'Cartons 16  Units 249  Weight 523 kg',
         'Style DR3775 Color BLACK Size M', 'Thank you for your business']
for dpi, label in ((300, 'a clean 300 dpi scan'), (150, 'a rough 150 dpi scan')):
    src = pymupdf.open(); p = src.new_page(width=612, height=792)
    for i, line in enumerate(TRUTH): p.insert_text((72, 100 + i * 40), line, fontname='helv', fontsize=16)
    pix = p.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY)
    scan = pymupdf.open(); sp = scan.new_page(width=612, height=792); sp.insert_image(sp.rect, pixmap=pix)
    f = os.path.join(work, f'scan_{dpi}.pdf'); scan.save(f)
    r = c.post('/ocr', json={'filepath': f}).get_json()
    got = pymupdf.open(r['new_file'])[0].get_text() if r.get('new_file') else ''
    want = ' '.join(TRUTH).split()
    have = got.split()
    hits = sum(1 for w in want if w in have)
    check(f'reads {label}', hits / len(want) >= 0.95, f'{hits} of {len(want)} words right')
check('a scan is recognized as one (to offer making it searchable)', c.post('/is_scanned', json={'filepath': f}).get_json().get('scanned') is True)

print('\nLoad speed')
long = os.path.join(work, 'long.pdf'); d = pymupdf.open()
for i in range(220): d.new_page().insert_text((72, 72), f'Page {i + 1}', fontname='helv', fontsize=18)
d.save(long)
files = [long] + [os.path.join(folder, n) for folder in (os.path.join(T, 'pdfs'), os.path.join(T, 'samples', 'pdf'))
                  if os.path.isdir(folder) for n in sorted(os.listdir(folder)) if n.lower().endswith('.pdf')]
slowest = (0, '')
for f in files:
    t = time.time()
    r = c.post('/process_file', json={'filepath': f}).get_json()
    if r.get('pdf_path'): pymupdf.open(r['pdf_path'])[0].get_pixmap(dpi=96)
    slowest = max(slowest, (time.time() - t, os.path.basename(f)))
check(f'every one of {len(files)} PDFs opens and shows page 1 in under 1.5 s', slowest[0] < 1.5, f'slowest {slowest[0]:.2f}s: {slowest[1]}')

print('\nSummary')
rep = os.path.join(T, 'samples', 'pdf_from_word.pdf')
s = c.post('/summarize', json={'filepath': rep}).get_json()
check('finds the main points', any('Revenue grew' in x for x in s.get('sentences', [])), f"{len(s.get('sentences', []))} points")
check('finds the amounts', '$48,300' in s.get('facts', {}).get('money', []), str(s.get('facts', {}).get('money')))
s = c.post('/summarize', json={'filepath': f if 'scan' in f else os.path.join(work, 'scan_300.pdf')}).get_json()
check('says when a file is only pictures', s.get('empty') is True)

app = os.path.join(ROOT, 'release', 'win-unpacked', 'MyPDFKit.exe')
if os.path.exists(app):
    import psutil
    print('\nBloat (the built app)')
    prof = os.path.join(work, 'profile')
    t0 = time.time()
    proc = subprocess.Popen([app, f'--user-data-dir={prof}'])
    ready = None
    for _ in range(600):
        time.sleep(0.1)
        kids = [k for k in psutil.Process(proc.pid).children(recursive=True) if k.name().lower() == 'server.exe']
        conns = [x for k in kids for x in k.net_connections(kind='tcp') if x.status == 'LISTEN']
        if conns:
            import urllib.request
            try: urllib.request.urlopen(f'http://127.0.0.1:{conns[0].laddr.port}/ping', timeout=1); ready = time.time() - t0; break
            except Exception: pass
    check('starts and is ready to open files in under 10 s', ready is not None and ready < 10, f'{ready:.1f}s' if ready else 'never')
    time.sleep(6)   # settled, with a file-less window open
    procs = [psutil.Process(proc.pid)] + psutil.Process(proc.pid).children(recursive=True)
    mine = [p for p in procs if p.name().lower() in ('mypdfkit.exe', 'server.exe')]
    mb = sum(p.memory_info().private for p in mine) / 1048576
    check('uses under 700 MB of memory while idle', mb < 700, f'{mb:.0f} MB in {len(mine)} processes (Office it keeps ready not counted)')
    for p in reversed(procs):
        try: p.kill()
        except Exception: pass
    # its Office helpers belong to it only when started after it; leave any others alone
    time.sleep(1)
    for p in psutil.process_iter(['name', 'create_time', 'cmdline']):
        if (p.info['name'] or '').upper() in ('WINWORD.EXE', 'EXCEL.EXE') and p.info['create_time'] >= t0 and '/automation' in ' '.join(p.info['cmdline'] or []).lower():
            try: p.kill()
            except Exception: pass

shutil.rmtree(work, ignore_errors=True)
print(f'\n{sum(results)} of {len(results)} checks passed')
sys.exit(0 if all(results) else 1)
