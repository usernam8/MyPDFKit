# For each PDF: rewrite sample text runs with identical text and score how close the ink is to the original
import sys, os, glob, shutil, collections
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))   # the project folder
import server, fitz
work = os.path.join(os.environ['TEMP'], 'fid_work'); os.makedirs(work, exist_ok=True)
c = server.app.test_client()
DPI = 150

def ink(pix):
    w, h, n, s = pix.width, pix.height, pix.n, pix.samples
    return {(i % w, i // w) for i in range(w * h) if s[i * n] < 140}

def near(a, b):   # share of a's ink pixels with ink of b within 1px
    if not a: return 1.0
    hit = sum(1 for (x, y) in a if any((x + dx, y + dy) in b for dx in (-1, 0, 1) for dy in (-1, 0, 1)))
    return hit / len(a)

def score_span(fp, pno, s):
    tmp = os.path.join(work, 'x.pdf'); shutil.copy(fp, tmp)
    page0 = fitz.open(fp)[pno]
    clip = fitz.Rect(s['bbox']) + (-2, -2, 2, 2)
    clip_view = clip * page0.rotation_matrix
    a = ink(page0.get_pixmap(dpi=DPI, clip=clip))
    r = c.post('/burn_edits', json={'filepath': tmp, 'page': pno, 'text_edits': [
        {'rect': list(s['bbox']), 'base_y': s['origin'][1], 'new_text': s['text']}]}).get_json()
    if not r.get('success'): return 0, 'error'
    q = fitz.open(r['new_file'])[pno]
    b = ink(q.get_pixmap(dpi=DPI, clip=clip))
    new = [t for bl in q.get_text('dict')['blocks'] if bl['type'] == 0 for l in bl['lines'] for t in l['spans']
           if abs(t['origin'][1] - s['origin'][1]) < 2 and abs(t['origin'][0] - s['origin'][0]) < 3]
    return min(near(a, b), near(b, a)), (new[0]['font'] if new else '?')

files = [f for a in sys.argv[1:] for f in (sorted(glob.glob(os.path.join(a, "*.pdf"))) if os.path.isdir(a) else [a])]
summary = []
for fp in files:
    d = fitz.open(fp)
    for pno in range(min(2, len(d))):
        seen = collections.OrderedDict()
        for bl in d[pno].get_text('dict')['blocks']:
            if bl['type'] != 0: continue
            for l in bl['lines']:
                for s in l['spans']:
                    key = (s['font'], round(s['size'])) if not os.environ.get('ALL') else (s['text'], round(s['origin'][1]))
                    if len(s['text'].strip()) >= 2 and key not in seen: seen[key] = s
        for key, s in list(seen.items())[:6 if not os.environ.get('ALL') else 99]:
            font, size = s['font'], round(s['size'])
            sc, used = score_span(fp, pno, s)
            summary.append(sc)
            flag = '  ' if sc >= 0.9 else '!!'
            print(f'{flag} {sc:5.0%}  {os.path.basename(fp)[:30]:30s} p{pno + 1} {font[-28:]:28s} {size:3d}pt -> {used[-24:]:24s} {s["text"].strip()[:28]!r}')
print(f'\n{len(summary)} runs, {sum(1 for x in summary if x >= 0.9)} at 90%+, average {sum(summary) / max(1, len(summary)):.0%}')
