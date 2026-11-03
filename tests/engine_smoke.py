"""Start a built server.exe the way a fresh PC would run it and use its main tools over HTTP.
usage: engine_smoke.py <server.exe> <work folder> [no-office]
The engine gets a bare PATH (only Windows), no Python variables, and a temp folder with spaces and accents,
so anything it borrows from this PC's Python or tools shows up as a failure."""
import sys, os, json, time, shutil, subprocess, urllib.request, glob

exe, work = sys.argv[1], sys.argv[2]
no_office = len(sys.argv) > 3
S = os.path.dirname(os.path.abspath(__file__))
SM = os.path.join(S, 'samples')
os.makedirs(work, exist_ok=True)
temp = os.path.join(work, 'Témp Földer ñ')
os.makedirs(temp, exist_ok=True)
files = os.path.join(work, 'My Fïles')
if os.path.exists(files): shutil.rmtree(files)
os.makedirs(files)
src = {'label.pdf': SM + '/private/label.pdf', 'packing list.xlsx': SM + '/private/packing list.xlsx',
       'report.docx': SM + '/word_report.docx', 'items.xls': SM + '/excel_items.xls', 'items.csv': SM + '/excel_items.csv',
       'deck.pptx': SM + '/slides_deck.pptx', 'scan.tif': SM + '/picture_scan.tif', 'invoice.png': SM + '/picture_invoice.png',
       'page.html': SM + '/web_newsletter.html', 'essay.pdf': SM + '/pdf_essay_word.pdf', 'notes.txt': SM + '/notes.txt'}
src = {k: v for k, v in src.items() if os.path.exists(v)}   # the private samples (real labels) may not be here
for k, v in src.items(): shutil.copy(v, os.path.join(files, k))
P = lambda n: os.path.join(files, n)

port = 5123
env = {'SystemRoot': os.environ['SystemRoot'], 'windir': os.environ['windir'], 'PATH': r'C:\Windows\System32;C:\Windows',
       'TEMP': temp, 'TMP': temp, 'USERPROFILE': os.environ['USERPROFILE'], 'LOCALAPPDATA': os.environ['LOCALAPPDATA'],
       'APPDATA': os.environ['APPDATA'], 'ProgramFiles': os.environ['ProgramFiles'], 'ProgramFiles(x86)': os.environ['ProgramFiles(x86)'],
       'ProgramData': os.environ['ProgramData'], 'SystemDrive': 'C:', 'COMPUTERNAME': os.environ.get('COMPUTERNAME', ''),
       'MYPDFKIT_PORT': str(port)}
if no_office: env['MYPDFKIT_NO_OFFICE'] = '1'
t0 = time.time()
proc = subprocess.Popen([exe], env=env, stderr=open(os.path.join(work, 'engine.log'), 'w'), creationflags=0x08000000)
base = f'http://127.0.0.1:{port}/'
for _ in range(600):
    try: urllib.request.urlopen(base + 'ping', timeout=1); break
    except Exception: time.sleep(0.2)
print(f'engine up in {time.time() - t0:.1f}s')

def call(route, body):
    req = urllib.request.Request(base + route, data=json.dumps(body).encode(), headers={'Content-Type': 'application/json'})
    try: return json.loads(urllib.request.urlopen(req, timeout=300).read())
    except urllib.error.HTTPError as e:
        raw = e.read()
        try: return json.loads(raw or b'{}')
        except ValueError: return {'error': f'HTTP {e.code}'}

results = []
def check(name, ok, info=''):
    results.append(ok); print(('PASS ' if ok else 'FAIL ') + name + (f'  ({info})' if info else ''))

opened = {}
for n in src:
    t = time.time(); d = call('process_file', {'filepath': P(n)})
    ok = bool(d.get('pdf_path')) and os.path.exists(d['pdf_path'])
    if ok: opened[n] = d['pdf_path']
    check(f'open {n}', ok, f"{time.time() - t:.1f}s" + (' ' + d['message'] if d.get('message') else '') + ('' if ok else ' ' + str(d)[:150]))

pdf = opened.get('essay.pdf')
d = call('page_info', {'filepath': pdf, 'page': 0}); check('page info (text to edit)', 'error' not in d, str(d)[:80] if 'error' in d else '')
d = call('render_pages', {'filepath': pdf, 'pages': [1]}); check('render page pictures', 'error' not in d)
d = call('search', {'filepath': pdf, 'query': 'the'}); check('search', d.get('count', 0) > 0, f"{d.get('count')} hits")
d = call('ocr', {'filepath': opened.get('scan.tif')}); check('recognize text (OCR, Windows built-in)', 'error' not in d, str(d)[:120])
d = call('compress_pdf', {'filepath': pdf}); check('compress', 'error' not in d, str(d)[:100])
d = call('number_pages', {'filepath': pdf}); check('page numbers', 'error' not in d, str(d)[:100])
d = call('header_footer', {'filepath': pdf, 'header': 'Head', 'footer': 'Foot'}); check('header and footer', 'error' not in d, str(d)[:100])
d = call('watermark', {'files': [pdf], 'kind': 'text', 'text': 'DRAFT', 'color': '#c0392b', 'opacity': 30}); check('watermark', 'error' not in d, str(d)[:100])
d = call('merge_queue', {'files': [opened['essay.pdf'], opened['invoice.png']], 'names': ['essay.pdf', 'invoice.png'], 'first_original_path': P('essay.pdf'), 'bookmarks': True, 'name': 'Combined'})
check('merge with bookmarks', bool(d.get('success') or d.get('path')), str(d)[:120])
d = call('find_patterns', {'filepath': pdf, 'kinds': ['email', 'phone', 'date'], 'words': ''}); check('find personal info', 'error' not in d, str(d)[:80])
d = call('privacy_check', {'filepath': pdf}); check('privacy check', 'error' not in d, str(d)[:80])
d = call('pdfa', {'filepath': pdf, 'original_path': P('essay.pdf')}); check('save as PDF/A', bool(d.get('success')), str(d)[:120])
d = call('protect', {'filepath': pdf, 'original_path': P('essay.pdf'), 'password': 'x1'}); check('password protect', 'error' not in d, str(d)[:120])
d = call('export', {'filepath': pdf, 'original_path': P('essay.pdf'), 'format': 'docx'}); check('export to Word', 'error' not in d, str(d)[:120])
d = call('export', {'filepath': opened['items.xls'], 'original_path': P('items.xls'), 'format': 'xlsx'}); check('export to Excel', 'error' not in d, str(d)[:120])
if 'label.pdf' in opened: d = call('box_labels', {'filepath': opened['label.pdf'], 'original_path': P('label.pdf'), 'packing_list': P('packing list.xlsx')})
if 'label.pdf' in opened: check('box labels from packing list', d.get('labels') == 16 and d.get('units') == 249, str({k: d.get(k) for k in ('labels', 'units', 'error')}))
d = json.loads(urllib.request.urlopen(base + 'fonts', timeout=60).read()); check('installed fonts list', bool(d), f'{len(d)} entries')
d = call('create_blank', {}); check('blank page', 'error' not in d, str(d)[:80])
d = call('cert_make', {'name': 'Test Person', 'email': 't@example.com', 'org': '', 'password': 'pw12345', 'out': P('my id.pfx')}); check('make a digital ID', 'error' not in d, str(d)[:120])

try: call('shutdown', {})
finally: proc.kill(); proc.wait()
log = open(os.path.join(work, 'engine.log'), encoding='utf-8', errors='replace').read()
bad = [l for l in log.splitlines() if 'Traceback' in l or 'Error' in l and 'GET' not in l and 'POST' not in l]
print('engine log problems:', bad[:6])
print(f'{sum(results)} of {len(results)} passed')
