"""Speed benchmark: how long each everyday job takes, next to Microsoft Word doing the same job when Word is installed.
  python tests/bench.py            (writes tests/bench-results.md)
Every job runs 3 times on copies in a temp folder; the best time counts (the first run also pays for loading
code from disk, which a person only pays once per session)."""
import os, sys, time, shutil, tempfile, statistics, json
T = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(T)
sys.path.insert(0, ROOT)
import server, pymupdf

c = server.app.test_client()
work = tempfile.mkdtemp(prefix='toast-bench-')
S = os.path.join(T, 'samples')
BOOK = os.path.join(T, 'pdfs', 'meditations-public-domain.pdf')
if not os.path.exists(BOOK):
    import subprocess; subprocess.run([sys.executable, os.path.join(T, 'make_book.py')], check=True)

def copy(src, name=None):
    d = tempfile.mkdtemp(dir=work); dst = os.path.join(d, name or os.path.basename(src)); shutil.copyfile(src, dst); return dst

def timed(fn, runs=3):
    times = []
    for _ in range(runs):
        t = time.perf_counter(); r = fn(); times.append(time.perf_counter() - t)
        if isinstance(r, dict) and r.get('error'): raise RuntimeError(r['error'])
    return min(times), r

def post(route, **kw):
    r = c.post(route, json=kw); j = r.get_json()
    return j if j is not None else {'error': f'HTTP {r.status_code}'}

rows = []
def job(name, fn, word=None):
    try:
        ours, _ = timed(fn)
    except Exception as e:
        rows.append((name, None, None, str(e)[:80])); print(f'  {name}: FAILED {e}'); return
    if word: word.insert(0, name); word_jobs.append(word)
    rows.append((name, ours, None, '')); print(f'  {name}: {ours:.2f}s')

# Microsoft Word doing the same job, the way a person would "Save as PDF" / open a PDF in Word.
# Timed afterwards in its own process by bench_word.py (Word's automation and the engine's don't mix in one process).
word_jobs = []
word_app = '--word' in sys.argv   # Word's PDF import can hang on a hidden prompt when automated, so it's opt-in
def word_to_pdf(src): return ['to_pdf', src]
def word_pdf_to_docx(src): return ['to_docx', src]

print('Toast Docs speed benchmark\n')
book = copy(BOOK)
n_book = pymupdf.open(book).page_count

def open_first():
    r = post('/process_file', filepath=copy(book))
    pymupdf.open(r['pdf_path'])[0].get_pixmap(dpi=96); return r
job(f'Open a {n_book}-page book and show page 1', open_first)

docx = copy(os.path.join(S, 'word_report.docx'))
job('Word file to PDF', lambda: post('/process_file', filepath=copy(docx)), word_to_pdf(docx) if word_app else None)
xlsx = copy(os.path.join(S, 'excel_items.xlsx'))
job('Excel file to PDF', lambda: post('/process_file', filepath=copy(xlsx)))
pptx = copy(os.path.join(S, 'slides_deck.pptx'))
job('PowerPoint file to PDF', lambda: post('/process_file', filepath=copy(pptx)))
jpg = copy(os.path.join(S, 'picture_invoice.jpg'))
job('Picture to PDF', lambda: post('/process_file', filepath=copy(jpg)))

rep = copy(os.path.join(S, 'pdf_from_word.pdf'))
job('PDF to Word (1 page)', lambda: post('/export', filepath=rep, original_path=rep, format='docx'),
    word_pdf_to_docx(rep) if word_app else None)
mixed = copy(os.path.join(S, 'mixed', 'report.pdf'))
n_mixed = pymupdf.open(mixed).page_count
job(f'PDF to Word ({n_mixed} pages, text and tables)', lambda: post('/export', filepath=mixed, original_path=mixed, format='docx'),
    word_pdf_to_docx(mixed) if word_app else None)
job(f'PDF to Word ({n_book}-page book)', lambda: post('/export', filepath=book, original_path=book, format='docx'),
    word_pdf_to_docx(book) if word_app else None)
xl = copy(os.path.join(S, 'pdf_from_excel_borderless.pdf'))
job('PDF to Excel', lambda: post('/export', filepath=xl, original_path=xl, format='xlsx'))

parts = [copy(os.path.join(S, 'pdf', n)) for n in sorted(os.listdir(os.path.join(S, 'pdf')))[:20]]
job(f'Combine {len(parts)} PDFs', lambda: post('/merge_queue', files=parts, first_original_path=parts[0]))
job(f'Split the {n_book}-page book into pages', lambda: post('/split_pdf', filepath=book, original_path=book))
job(f'Compress the {n_book}-page book', lambda: post('/compress_pdf', filepath=copy(book)))
job('Password-protect a PDF', lambda: post('/protect', filepath=rep, original_path=rep, password='bench-1234'))

scan = os.path.join(work, 'scan.pdf'); d = pymupdf.open(); p = d.new_page()
for i in range(25): p.insert_text((60, 60 + i * 28), f'Line {i + 1}: the quick brown fox jumps over the lazy dog 2026', fontsize=13)
pix = p.get_pixmap(dpi=200, colorspace=pymupdf.csGRAY); s2 = pymupdf.open(); s2.new_page().insert_image(s2[0].rect, pixmap=pix); s2.save(scan)
job('Recognize text on a scanned page (OCR)', lambda: post('/ocr', filepath=scan))
job(f'Quick summary of the {n_book}-page book', lambda: post('/summarize', filepath=book))

if word_jobs:
    import subprocess
    r = subprocess.run([sys.executable, os.path.join(T, 'bench_word.py'), json.dumps(word_jobs)], capture_output=True, text=True)
    try: wt = json.loads(r.stdout.strip().splitlines()[-1])
    except Exception: wt = {}; print('Word timing failed:', r.stderr[-400:])
    rows = [(n, o, wt.get(n), e) for n, o, _, e in rows]

lines = ['# Toast Docs speed benchmark', '', f'PC: {os.cpu_count()} logical cores. Best of 3 runs (Word: best of 2). Lower is better.', '',
         '| Job | Toast Docs | Microsoft Word |', '|---|---:|---:|']
for name, ours, w, err in rows:
    lines.append(f'| {name} | {"failed: " + err if ours is None else f"{ours:.2f} s"} | {f"{w:.2f} s" if w else "—"} |')
open(os.path.join(T, 'bench-results.md'), 'w', encoding='utf-8').write('\n'.join(lines) + '\n')
json.dump([r[:3] for r in rows], open(os.path.join(work, 'bench.json'), 'w'))
print('\n' + '\n'.join(lines))
shutil.rmtree(work, ignore_errors=True)
