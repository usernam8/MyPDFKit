"""Makes the long test book (public domain): Marcus Aurelius, Meditations, from Project Gutenberg, laid out as a
book-size PDF of about 150 pages, for the AI and reading tests.  python tests/make_book.py
The PDF goes to tests/pdfs/ (not in the repository)."""
import os, re, sys, html, urllib.request
import pymupdf
T = os.path.dirname(os.path.abspath(__file__))
out = os.path.join(T, 'pdfs', 'meditations-public-domain.pdf')
if os.path.exists(out) and '--again' not in sys.argv: print('already there:', out); sys.exit(0)
t = urllib.request.urlopen('https://www.gutenberg.org/cache/epub/2680/pg2680.txt', timeout=60).read().decode('utf-8-sig')
a, b = t.find('*** START'), t.find('*** END')
t = t[t.find('\n', a) + 1:b].strip()
paras = [re.sub(r'\s*\n\s*', ' ', p).strip() for p in re.split(r'\n\s*\n', t) if p.strip()]
body = ''.join(f'<h2>{html.escape(p)}</h2>' if p.isupper() and len(p) < 60 else f'<p>{html.escape(p)}</p>' for p in paras)
story = pymupdf.Story(html=body, user_css='p { font-family: serif; font-size: 10.5pt; margin: 0 0 5pt 0; text-align: justify; } h2 { font-family: serif; font-size: 15pt; margin: 14pt 0 8pt 0; }')
os.makedirs(os.path.dirname(out), exist_ok=True)
w = pymupdf.DocumentWriter(out)
page = pymupdf.Rect(0, 0, 432, 648)
more = True
while more:
    dev = w.begin_page(page)
    more, _ = story.place(page + (48, 50, -48, -50))
    story.draw(dev)
    w.end_page()
w.close()
d = pymupdf.open(out); d.set_metadata({'title': 'Meditations', 'author': 'Marcus Aurelius (translated by Meric Casaubon)'}); d.saveIncr()
print(d.page_count, 'pages:', out)
