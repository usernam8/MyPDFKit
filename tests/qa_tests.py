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
         'Style DR3775 Color BLACK Size M', 'Thank you for your business',
         'some soft cotton socks were sold out', 'Ocean View Suite with very warm water', 'call us on weekdays or visit our store']
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
    check(f'reads {label}, every word exactly', hits == len(want), f'{hits} of {len(want)} words right')
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

if server._local_model() and server._llama_exe():
    import psutil
    print('\nBuilt-in AI (runs on this PC)')
    inv = os.path.join(work, 'invoice.pdf'); d = pymupdf.open(); p = d.new_page()
    p.insert_text((72, 80), 'INVOICE 1042 - Acme Supplies Inc.', fontsize=18)
    p.insert_textbox(pymupdf.Rect(60, 100, 550, 800), 'Bill to: Blue River Market. Invoice date: February 1, 2026. Payment is due March 3, 2026. '
                     '10 boxes of copy paper at $18.00 each; 4 toner cartridges at $60.00 each. Total due: $420.00.', fontsize=11)
    d.save(inv); d.close()
    def ask(question, dense):
        t0 = time.time(); first = None
        j = c.post('/ai_start', json={'filepath': inv, 'question': question, 'dense': dense}).get_json()['id']
        while True:
            r = c.post('/ai_status', json={'id': j}).get_json()
            if r['text'] and first is None: first = time.time() - t0
            if r['finished']: return r, first or 0, time.time() - t0
            time.sleep(0.1)
    server._local_ai_up()
    r, first, total = ask('', True)
    check('a dense summary has the total and who it is from', '420' in r['text'] and 'Acme' in r['text'] and not r['error'], f'first words {first:.1f}s, done {total:.1f}s')
    r, first, total = ask('When is the payment due?', False)
    check('answers a question about the file', 'March 3' in r['text'], f'first words {first:.1f}s, done {total:.1f}s')
    check('starts answering a short file in under 5 s', first < 5, f'{first:.1f}s')
    try: os.remove(inv); ok_unlocked = True
    except OSError: ok_unlocked = False
    check('leaves the file free (not locked) after reading it', ok_unlocked)
    pid = server.LOCAL_AI['proc'].pid
    c.post('/shutdown')
    time.sleep(1)
    check('closes the AI when the app closes', not psutil.pid_exists(pid) or not psutil.Process(pid).is_running())

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
    # the app goes away suddenly (as in a crash): its converter must notice, close the Word and Excel it
    # kept ready, and exit by itself
    engine = [p for p in procs if p.name().lower() == 'server.exe']
    for p in procs:
        if p.name().lower() == 'mypdfkit.exe':
            try: p.kill()
            except Exception: pass
    mine = lambda: [p for p in psutil.process_iter(['name', 'create_time']) if (p.info['name'] or '').upper() in ('WINWORD.EXE', 'EXCEL.EXE') and p.info['create_time'] >= t0]
    for _ in range(40):
        if not mine() and not any(p.is_running() for p in engine): break
        time.sleep(0.5)
    left = mine()
    check('after a crash, the converter closes its Office and itself', not left and not any(p.is_running() for p in engine), f'{len(left)} Office left')
    for p in left + engine:   # never leave test leftovers running
        try: p.kill()
        except Exception: pass

shutil.rmtree(work, ignore_errors=True)
print(f'\n{sum(results)} of {len(results)} checks passed')
sys.exit(0 if all(results) else 1)
