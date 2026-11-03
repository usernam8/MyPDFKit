import os, sys, pefile
roots = sys.argv[1:]
bundled, binaries = {}, []
for root in roots:
    for dp, dn, fn in os.walk(root):
        for f in fn:
            if f.lower().endswith(('.dll', '.pyd', '.exe', '.node')):
                bundled.setdefault(f.lower(), os.path.join(dp, f)); binaries.append(os.path.join(dp, f))
need = {}
for b in binaries:
    try:
        pe = pefile.PE(b, fast_load=True)
        pe.parse_data_directories(directories=[pefile.DIRECTORY_ENTRY['IMAGE_DIRECTORY_ENTRY_IMPORT'], pefile.DIRECTORY_ENTRY['IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT']])
        for attr in ('DIRECTORY_ENTRY_IMPORT', 'DIRECTORY_ENTRY_DELAY_IMPORT'):
            for e in getattr(pe, attr, []):
                need.setdefault(e.dll.decode(errors='replace').lower(), set()).add(os.path.basename(b))
        pe.close()
    except Exception as ex: pass
sys32 = r'C:\Windows\System32'
def product(path):
    try:
        pe = pefile.PE(path, fast_load=True); pe.parse_data_directories(directories=[pefile.DIRECTORY_ENTRY['IMAGE_DIRECTORY_ENTRY_RESOURCE']])
        for fi in getattr(pe, 'FileInfo', []) or []:
            for e in fi:
                for st in getattr(e, 'StringTable', []):
                    d = {k.decode(errors='replace'): v.decode(errors='replace') for k, v in st.entries.items()}
                    return d.get('ProductName', '')
    except Exception: return '?'
    return ''
problems = []
for dll, users in sorted(need.items()):
    if dll in bundled or dll.startswith(('api-ms-win-', 'ext-ms-')): continue
    p = os.path.join(sys32, dll)
    if not os.path.exists(p): problems.append((dll, 'NOT ON THIS PC', sorted(users)[:3])); continue
    prod = product(p)
    if 'Windows' not in prod and 'Internet Explorer' not in prod:
        problems.append((dll, f'in System32 but from: {prod!r}', sorted(users)[:3]))
print('binaries checked:', len(binaries), ' distinct imports:', len(need))
for x in problems: print('  ', x)
print('problems:', len(problems))
