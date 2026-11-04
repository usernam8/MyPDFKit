"""Everything to check before a release, in one go:  python tests/run_tests.py
1. Text editing keeps each line's look (every line of every sample PDF)          edit_fidelity.py
2. Pages are sorted right for export (tables, text, slides)                       kind_test.py
3. The built app carries every file it needs (nothing borrowed from this PC)      dllaudit.py
4. The built converter, started like on a new PC, with Office and without it      engine_smoke.py
Parts 3 and 4 need a build first (npm run build:server, then electron-builder). Samples are in tests/samples;
tests/samples/private holds real work files (box labels) and isn't uploaded to GitHub."""
import os, re, sys, subprocess, tempfile

T = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(T)
SAMPLES = os.path.join(T, 'samples')
results = []

def run(name, args, ok_pattern, env=None):
    print(f'\n== {name}')
    r = subprocess.run([sys.executable, '-W', 'ignore'] + args, capture_output=True, text=True, encoding='utf-8', errors='replace',
                       env=dict(os.environ, **(env or {})), cwd=ROOT)
    out = (r.stdout or '') + (r.stderr or '')
    last = [l for l in out.splitlines() if 'FAIL' in l][:10] + [l for l in out.splitlines() if l.strip()][-3:]   # what failed, then the totals
    print('\n'.join(last))
    ok = r.returncode == 0 and bool(re.search(ok_pattern, out, re.M))
    results.append((name, ok))
    return out

run('Text editing keeps the look', [os.path.join(T, 'edit_fidelity.py'), os.path.join(SAMPLES, 'pdf')],
    r'^(\d+) runs, \1 at 90%\+', {'ALL': '1'})
run('Pages sorted right for export', [os.path.join(T, 'kind_test.py'), SAMPLES], r'^(\d+) of \1 pages right')
run("PDF test library: open, show, fill, crop, dry-run print, against Chrome's engine", [os.path.join(T, 'print_library.py')], r'^(\d+) of \1 checks passed')
run('QA: OCR accuracy, load speed, summaries, startup time and memory', [os.path.join(T, 'qa_tests.py')], r'^(\d+) of \1 checks passed')

exe = os.path.join(ROOT, 'release', 'win-unpacked', 'resources', 'server', 'server.exe')
if os.path.exists(exe):
    run('Built app carries everything it needs', [os.path.join(T, 'dllaudit.py'), os.path.join(ROOT, 'dist', 'server'),
        os.path.join(ROOT, 'release', 'win-unpacked')], r'^problems: 0$')
    work = tempfile.mkdtemp(prefix='mypdfkit-test-')
    run('Converter on a new PC, no Office', [os.path.join(T, 'engine_smoke.py'), exe, os.path.join(work, 'a'), 'no-office'], r'^(\d+) of \1 passed')
    run('Converter on a new PC, with Office', [os.path.join(T, 'engine_smoke.py'), exe, os.path.join(work, 'b')], r'^(\d+) of \1 passed')
else:
    print('\n(no build found in release/win-unpacked: parts 3 and 4 skipped)')

print('\n' + '\n'.join(('PASS  ' if ok else 'FAIL  ') + n for n, ok in results))
sys.exit(0 if all(ok for _, ok in results) else 1)
