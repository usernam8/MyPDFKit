"""Microsoft Word doing the same jobs as bench.py, timed in its own process (Word's automation and the engine's
don't mix in one process). Prints JSON: {job: best seconds}. Used by bench.py."""
import os, sys, time, json, tempfile, shutil
import win32com.client, pythoncom
pythoncom.CoInitialize()
w = win32com.client.DispatchEx('Word.Application'); w.Visible = False; w.DisplayAlerts = 0
work = tempfile.mkdtemp(prefix='toast-bench-word-')
out = {}
def best(name, fn, runs=2):
    times = []
    for _ in range(runs):
        t = time.perf_counter(); fn(); times.append(time.perf_counter() - t)
    out[name] = min(times)
def to_pdf(src):
    d = w.Documents.Open(src, ReadOnly=True); d.ExportAsFixedFormat(os.path.join(work, f'{time.time_ns()}.pdf'), 17); d.Close(False)
def pdf_to_docx(src):
    d = w.Documents.Open(src, ConfirmConversions=False, ReadOnly=True); d.SaveAs2(os.path.join(work, f'{time.time_ns()}.docx'), 16); d.Close(False)
try:
    for name, kind, path in json.loads(sys.argv[1]):
        try: best(name, (lambda p=path: to_pdf(p)) if kind == 'to_pdf' else (lambda p=path: pdf_to_docx(p)))
        except Exception as e: out[name] = None; print(name, e, file=sys.stderr)
finally:
    try: w.Quit()
    except Exception: pass
    shutil.rmtree(work, ignore_errors=True)
print(json.dumps(out))
