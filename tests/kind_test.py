# Does page sorting (table / text / slide) get every test page right?
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))   # the project folder
import server, fitz
S = sys.argv[1]
expect = {  # file: kind per page
    r'pdf\word_report_docx.pdf': ['docx'], r'pdf\pdf_essay_word_pdf.pdf': ['docx', 'docx'],
    r'pdf\pdf_from_excel_borderless_pdf.pdf': ['xlsx', 'xlsx'], r'pdf\excel_items_csv.pdf': ['xlsx'],
    r'pdf\slides_deck_pptx.pdf': ['pptx', 'pptx', 'pptx'], r'pdf\web_newsletter_html.pdf': ['docx'],
    r'pdf\notes_txt.pdf': ['docx', 'docx'], r'pdf\data_json.pdf': ['docx', 'docx', 'docx'],
    r'pdf\picture_invoice_png.pdf': ['docx'], r'pdf\pdf_base14_pdf.pdf': ['docx'],
    r'private\essay_orig.pdf': ['docx'], r'private\pl_dr4905.pdf': ['xlsx'], r'private\pl_dr3775.pdf': ['xlsx'], r'kinds\slides.pdf': ['pptx'] * 3,
    r'mixed\report.pdf': ['docx', 'xlsx', 'pptx', 'xlsx'],
}
bad = 0; total = 0
for rel, kinds in expect.items():
    if not os.path.exists(os.path.join(S, rel)): continue   # private samples may not be on this PC
    d = fitz.open(os.path.join(S, rel))
    got = [server._page_kind(p) for p in d][:len(kinds)]
    total += len(kinds); wrong = sum(a != b for a, b in zip(got, kinds)); bad += wrong
    print(('  ' if not wrong else '!!'), f'{rel:45s} expected {kinds} got {got}')
print(f'\n{total - bad} of {total} pages right')
