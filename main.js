const { app, BrowserWindow, dialog, Menu, ipcMain, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const http = require('http');
const net = require('net');

// The backend gets a free port of its own every launch, so it never meets another copy's backend
// (a leftover one from an older version kept answering on 5000 and the app talked to it).
// A test copy can still pin one with MYPDFKIT_PORT.
let SERVER_PORT = +process.env.MYPDFKIT_PORT || 0;
function pickPort() {
  if (SERVER_PORT) return Promise.resolve(SERVER_PORT);
  return new Promise(resolve => {
    const s = net.createServer();
    s.once('error', () => resolve(SERVER_PORT = 5000));
    s.listen(0, '127.0.0.1', () => { SERVER_PORT = s.address().port; s.close(() => resolve(SERVER_PORT)); });
  });
}
let serverProc = null;
let win = null;

// Only allow one running copy; a second launch (e.g. double-clicking another file)
// hands its file to the first window.
const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
}

// everything the backend can turn into a PDF (keep in step with OPENABLE_EXT in server.py)
const OPENABLE = new RegExp('\\.(' +['pdf', 'doc', 'docx', 'docm', 'dot', 'dotx', 'dotm', 'rtf', 'odt', 'wps', 'wpd',
  'xls', 'xlsx', 'xlsm', 'xlsb', 'xlt', 'xltx', 'csv', 'tsv', 'ods', 'ppt', 'pptx', 'pptm', 'pps', 'ppsx', 'pot', 'potx', 'odp',
  'jpg', 'jpeg', 'jfif', 'png', 'bmp', 'gif', 'tif', 'tiff', 'webp', 'ico', 'pnm', 'jxr', 'xps', 'oxps', 'epub', 'fb2', 'cbz', 'mobi', 'svg',
  'html', 'htm', 'mhtml', 'mht', 'xhtml', 'txt', 'md', 'markdown', 'log', 'json', 'xml', 'yaml', 'yml', 'ini', 'cfg', 'conf',
  // design, CAD, 3D and photo files
  'dxf', 'dwg', 'dwf', 'dwfx', 'plt', 'hpgl', 'hpg', 'hgl', 'psd', 'psb', 'heic', 'heif', 'avif', 'tga', 'dds', 'pcx', 'jp2', 'j2k', 'emf', 'wmf', 'svgz',
  'ai', 'eps', 'ps', 'stl', 'obj', 'ply', 'off', 'skp', 'rvt', 'rfa', '3dm', 'max', 'blend', 'indd', 'idml', 'cdr', 'afdesign', 'afphoto', 'afpub',
  'sketch', 'fig', 'xd', 'pages', 'numbers', 'key', 'vsd', 'vsdx', 'pub', 'ifc', 'dgn', 'step', 'stp', 'iges', 'igs', 'sldprt', 'sldasm', 'slddrw',
  'ipt', 'iam', 'f3d', 'c4d', 'fbx', 'dae', '3ds', 'cr2', 'cr3', 'nef', 'arw', 'dng', 'raf', 'orf', 'rw2', 'pef', 'srw', 'clo', 'zprj', 'avt', 'pds', 'mdl', 'aama', 'rul', 'zpac', 'zfab', 'zsel', 'bw', 'vsp', 'u3m', 'iba', 'vet', 'mrk', 'tuk', 'pln', 'pla', 'bimx', 'vwx', 'nwd', 'nwc', 'nwf', 'lcf', 'tcp', 'dwfx3d', 'eml'].join('|') + ')$', 'i');
function filesFromArgv(argv) {
  return argv.slice(app.isPackaged ? 1 : 2).filter(a => OPENABLE.test(a) && fs.existsSync(a));
}

// a small log of what went wrong (the converter's errors, the page's errors), for "Send a problem report"
let logFile = null;
function logLine(src, text) {
  try {
    if (!logFile) { const dir = path.join(app.getPath('userData'), 'logs'); fs.mkdirSync(dir, { recursive: true }); logFile = path.join(dir, 'mypdfkit.log'); }
    const lines = String(text).split(/\r?\n/).filter(l => l.trim() && !/"(GET|POST) \/[^ ]* HTTP\/1\.1" (200|304)/.test(l) && !/^ \* |Press CTRL\+C|development server|fitz` API is deprecated/.test(l));
    if (!lines.length) return;
    if (fs.existsSync(logFile) && fs.statSync(logFile).size > 1_000_000) fs.renameSync(logFile, logFile + '.old');
    fs.appendFileSync(logFile, lines.map(l => `${new Date().toISOString()} [${src}] ${l}`).join('\n') + '\n');
  } catch (e) {}
}
process.on('uncaughtException', e => logLine('app', e.stack || e));

let serverReady = false, serverRestarts = 0, serverLog = '', serverExit = null;
function startServer() {
  // the backend closes itself if this app goes away without telling it (a crash or Task Manager)
  const env = { ...process.env, MYPDFKIT_PORT: String(SERVER_PORT), MYPDFKIT_PARENT_PID: String(process.pid) };
  // detached: outside the job Node puts children in, which kills them the instant the app dies; that left
  // no time to close the hidden Word/Excel (started by Windows, not by us), so they stayed running.
  // The backend watches this app's process and closes Office and itself when it's gone.
  const opts = { windowsHide: true, env, detached: true };
  if (app.isPackaged) {
    // a PyInstaller folder build: starts at once, unlike a one-file exe that unpacks itself every launch
    const exe = path.join(process.resourcesPath, 'server', 'server.exe');
    serverProc = spawn(exe, [], opts);
  } else {
    // Dev mode: run the Python script directly
    serverProc = spawn('python', [path.join(__dirname, 'server.py')], opts);
  }
  serverExit = null;
  const proc = serverProc;
  proc.on('error', err => { serverExit = String(err); });
  proc.stderr && proc.stderr.on('data', d => {
    console.error('[server]', d.toString());
    logLine('converter', d.toString());
    serverLog = (serverLog + d.toString()).slice(-3000);   // the last lines, for the error message
  });
  proc.on('exit', code => {
    if (proc !== serverProc || closing) return;
    serverExit = serverExit || `It closed with code ${code}.`;
    // it was working and then went away: start it again on the same port (the page reconnects by itself)
    if (serverReady && serverRestarts++ < 3) { console.error('[server] exited, restarting'); startServer(); }
  });
}

function backendError(reason) {
  logLine('app', 'converter did not start: ' + reason);
  const exe = path.join(process.resourcesPath, 'server', 'server.exe');
  if (app.isPackaged && !fs.existsSync(exe)) {
    // the usual reason: Windows Security mistook it for a threat and moved it away
    dialog.showErrorBox('Windows Security blocked part of MyPDFKit',
      'Windows Security removed a part of MyPDFKit it didn\'t recognize (it happens with apps from small makers). To bring it back:\n\n' +
      '1. Open the Start menu and type Windows Security, then open it.\n' +
      '2. Click Virus & threat protection, then Protection history.\n' +
      '3. Click the item that mentions MyPDFKit or server.exe, then Actions, then Restore.\n' +
      '4. Open MyPDFKit again.\n\nIf there\'s nothing to restore, install MyPDFKit again from the same download page.');
    return;
  }
  const tail = serverLog.trim().split(/\r?\n/).slice(-12).join('\n');
  dialog.showErrorBox('MyPDFKit could not start its converter',
    'Opening and converting files won\'t work until it starts. Try closing MyPDFKit and opening it again. ' +
    'If it keeps happening, reinstall MyPDFKit, or use All tools, then Send a problem report, and send that file to whoever helps you.\n\n' +
    reason + (tail ? '\n\nDetails:\n' + tail : ''));
}

// up to 3 minutes: the first start on a new PC can be slow while Windows checks the new files
function waitForServer(retries = 1200) {
  return new Promise((resolve, reject) => {
    const attempt = n => {
      if (serverExit) return reject(new Error(serverExit));
      http.get(`http://127.0.0.1:${SERVER_PORT}/ping`, res => { res.resume(); resolve(); })
        .on('error', () => (n <= 0 ? reject(new Error('It didn\'t answer within 3 minutes.')) : setTimeout(() => attempt(n - 1), 150)));
    };
    attempt(retries);
  });
}

let pageLoaded = false, serverUp = false;
function flushOpenFiles() {
  if (!pageLoaded || !serverUp || !win) return;
  const files = filesFromArgv(process.argv);
  if (files.length) win.webContents.send('open-files', files);
}

// The window opens where and how big it was last time (on a screen that's still there), else at about 85% of
// the screen the mouse is on, so "restore" never fills a small screen.
const winStateFile = () => path.join(app.getPath('userData'), 'window.json');
function windowBounds() {
  const { screen } = require('electron');
  let st = null;
  try { st = JSON.parse(fs.readFileSync(winStateFile(), 'utf8')); } catch (e) {}
  const area = screen.getDisplayNearestPoint(screen.getCursorScreenPoint()).workArea;
  if (st && st.bounds) {
    const b = st.bounds, d = screen.getDisplayMatching(b).workArea;
    const fits = b.width <= d.width + 8 && b.height <= d.height + 8 && b.x >= d.x - 8 && b.y >= d.y - 8 && b.x + b.width <= d.x + d.width + 8 && b.y + b.height <= d.y + d.height + 8;
    if (fits) return { bounds: b, maximized: !!st.maximized };
    return { bounds: centered(d, Math.min(b.width, d.width * 0.85), Math.min(b.height, d.height * 0.85)), maximized: !!st.maximized };
  }
  return { bounds: centered(area, area.width * 0.85, area.height * 0.85), maximized: false };
}
const centered = (d, w, h) => ({ x: Math.round(d.x + (d.width - w) / 2), y: Math.round(d.y + (d.height - h) / 2), width: Math.round(w), height: Math.round(h) });
function rememberWindow() {
  if (!win || win.isDestroyed()) return;
  try { fs.writeFileSync(winStateFile(), JSON.stringify({ bounds: win.getNormalBounds(), maximized: win.isMaximized() })); } catch (e) {}
}
function createWindow() {
  const place = windowBounds();
  win = new BrowserWindow({
    ...place.bounds, minWidth: 900, minHeight: 600, title: 'MyPDFKit',
    backgroundColor: '#25262a',
    show: false,
    autoHideMenuBar: true,
    // The app's own top bar acts as the title bar, with its own minimize / maximize / close buttons. (Windows'
    // overlay buttons misread clicks on a monitor above the main one, where the window's top is below zero,
    // so Restore did nothing there.) Dragging, double-clicking the bar and Windows snapping still work.
    titleBarStyle: 'hidden',
    webPreferences: {
      // index.html uses require('fs') / require('electron') directly.
      // Safe here because the page now loads only local files (pdf.js is bundled).
      nodeIntegration: true,
      contextIsolation: false,
    },
  });
  if (app.isPackaged) Menu.setApplicationMenu(null);
  win.once('ready-to-show', () => { if (place.maximized) win.maximize(); win.show(); });
  let saveTimer = null;
  const remember = () => { clearTimeout(saveTimer); saveTimer = setTimeout(rememberWindow, 400); };
  for (const ev of ['resize', 'move', 'maximize', 'unmaximize']) win.on(ev, remember);
  win.on('close', rememberWindow);
  // closing the main window ends the app, even if a hidden helper window (printing) is still open
  win.on('closed', () => { win = null; quitCleanly(); });
  // Safety net: never leave the window hidden if the first paint is delayed or fails
  setTimeout(() => { if (win && !win.isDestroyed() && !win.isVisible()) win.show(); }, 2500);

  // Print renderer problems in the terminal so they can be diagnosed
  win.webContents.on('did-fail-load', (_e, code, desc, url) => console.error('[load failed]', code, desc, url));
  win.webContents.on('render-process-gone', (_e, d) => console.error('[renderer gone]', d.reason));
  win.webContents.on('console-message', (...a) => {
    const e = a[0];
    const msg = e && e.message !== undefined ? e.message : a[2];
    console.log('[page]', msg);
    const level = e && e.level !== undefined ? e.level : a[1];
    if ((level === 'error' || level === 3 || level === 'warning' || level === 2) && !/Electron Security Warning/.test(msg)) logLine('page', msg);
  });
  win.loadFile('index.html', { query: { port: String(SERVER_PORT) } });
  win.webContents.on('did-finish-load', () => { pageLoaded = true; flushOpenFiles(); });
}

app.on('second-instance', (_e, argv) => {
  if (!win || win.isDestroyed()) return;
  if (win.isMinimized()) win.restore();
  win.focus();
  const files = filesFromArgv(argv);
  if (files.length) win.webContents.send('open-files', files);
});

if (gotLock) app.whenReady().then(async () => {
  await pickPort();
  startServer();
  createWindow();   // show the window immediately; the backend finishes starting in parallel
  try { await waitForServer(); serverReady = serverUp = true; flushOpenFiles(); }
  catch (e) { backendError(e.message); }
  setupAutoUpdate();
});

// ---------- updates ----------
// The installed app checks GitHub (the public MyPDFKit-releases repo) at startup and every few hours,
// downloads a newer version in the background, then offers a restart. If the user ignores it, the
// update is installed the next time the app closes.
function setupAutoUpdate() {
  if (!app.isPackaged || process.windowsStore) return;   // the Microsoft Store updates its own copy
  let autoUpdater;
  try { ({ autoUpdater } = require('electron-updater')); } catch (e) { console.error('[update] unavailable', e); return; }
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;
  autoUpdater.on('update-downloaded', info => {
    if (win && !win.isDestroyed()) win.webContents.send('update-ready', info.version);
  });
  autoUpdater.on('error', err => console.error('[update]', err && err.message));
  const check = () => autoUpdater.checkForUpdates().catch(err => console.error('[update]', err && err.message));
  setTimeout(check, 4000);   // after startup settles
  setInterval(check, 4 * 60 * 60 * 1000);
  ipcMain.on('install-update', () => { closing = true; closeBackend().then(() => autoUpdater.quitAndInstall(true, true)); });
}
ipcMain.handle('app-version', () => app.getVersion());
ipcMain.on('win-ctl', (e, what) => {
  const w = BrowserWindow.fromWebContents(e.sender); if (!w) return;
  if (what === 'min') w.minimize();
  else if (what === 'max') w.isMaximized() ? w.unmaximize() : w.maximize();
  else if (what === 'close') w.close();
});
ipcMain.handle('win-is-max', e => { const w = BrowserWindow.fromWebContents(e.sender); return !!(w && w.isMaximized()); });
// the built-in AI for the Lite version: downloaded once, checked by its fingerprint, then it works offline
const AI_MODEL = { file: 'Qwen3.5-2B-Q4_K_M.gguf', size: 1280835840, sha256: 'aaf42c8b7c3cab2bf3d69c355048d4a0ee9973d48f16c731c0520ee914699223',
  urls: ['https://github.com/usernam8/MyPDFKit-releases/releases/download/ai-model-1/Qwen3.5-2B-Q4_K_M.gguf',
         'https://huggingface.co/unsloth/Qwen3.5-2B-GGUF/resolve/main/Qwen3.5-2B-Q4_K_M.gguf'] };
ipcMain.handle('ai-download', async (e, dir) => {
  const fs = require('fs'), crypto = require('crypto'), { net } = require('electron');
  if (!dir) return { error: 'No place to save the AI' };
  fs.mkdirSync(dir, { recursive: true });
  const dest = path.join(dir, AI_MODEL.file), part = dest + '.part';
  let lastErr = 'Could not download the AI. Check the internet connection and try again.';
  for (const url of AI_MODEL.urls) {
    try {
      await new Promise((resolve, reject) => {
        const req = net.request({ url, redirect: 'follow' });
        req.on('response', res => {
          if (res.statusCode !== 200) return reject(new Error('HTTP ' + res.statusCode));
          const total = +res.headers['content-length'] || AI_MODEL.size, out = fs.createWriteStream(part), hash = crypto.createHash('sha256');
          let done = 0, shown = 0;
          res.on('data', c => { out.write(c); hash.update(c); done += c.length; if (done - shown > 4e6) { shown = done; e.sender.send('ai-download-progress', { done, total }); } });
          res.on('end', () => out.end(() => { e.sender.send('ai-download-progress', { done, total });
            hash.digest('hex') === AI_MODEL.sha256 ? resolve() : reject(new Error('The download was damaged')); }));
          res.on('error', reject);
        });
        req.on('error', reject);
        req.end();
      });
      fs.renameSync(part, dest);
      logLine('ai', 'downloaded ' + url);
      return { ok: true, path: dest };
    } catch (err) {
      lastErr = String(err.message || err) === 'The download was damaged' ? 'The download was damaged. Try again.' : lastErr;
      logLine('ai', 'download failed ' + url + ': ' + (err.message || err));
      try { fs.unlinkSync(part); } catch (x) {}
    }
  }
  return { error: lastErr };
});
ipcMain.handle('set-fullscreen', (e, on) => { const w = BrowserWindow.fromWebContents(e.sender); if (w) w.setFullScreen(!!on); });
ipcMain.handle('user-data', () => app.getPath('userData'));
// a text file on the desktop with the version, the PC and the recent problems, to send to whoever helps
ipcMain.handle('problem-report', async (_e, extra) => {
  const os = require('os');
  let log = '';
  try { log = fs.readFileSync(logFile || path.join(app.getPath('userData'), 'logs', 'mypdfkit.log'), 'utf8').split('\n').slice(-400).join('\n'); } catch (e) { log = '(no problems logged)'; }
  const text = [`MyPDFKit ${app.getVersion()} problem report, ${new Date().toString()}`,
    `Windows ${os.release()} ${os.arch()}, ${Math.round(os.totalmem() / 1073741824)} GB memory`,
    `Installed in: ${path.dirname(process.execPath)}`, `Converter: ${serverReady ? 'running' : 'not running'}${serverExit ? ' (' + serverExit + ')' : ''}`,
    extra ? `What the person wrote: ${extra}` : '', '', '--- Recent problems ---', log].join('\r\n');
  const out = path.join(app.getPath('desktop'), 'MyPDFKit problem report.txt');
  fs.writeFileSync(out, text);
  shell.showItemInFolder(out);
  return out;
});
ipcMain.on('user-data-sync', e => { e.returnValue = app.getPath('userData'); });
ipcMain.handle('pick-folder', async () => {
  const r = await dialog.showOpenDialog(win, { properties: ['openDirectory'], title: 'Pick a folder' });
  return r.canceled ? null : r.filePaths[0];
});
ipcMain.handle('save-copy', async (_e, from, suggested) => {
  const r = await dialog.showSaveDialog(win, { title: 'Save', defaultPath: suggested, filters: [{ name: 'PDF', extensions: ['pdf'] }] });
  if (r.canceled || !r.filePath) return null;
  fs.copyFileSync(from, r.filePath);
  return r.filePath;
});
ipcMain.handle('pick-files', async () => {
  const r = await dialog.showOpenDialog(win, { properties: ['openFile', 'multiSelections'], title: 'Pick files' });
  return r.canceled ? [] : r.filePaths;
});

// ---------- printing ----------
// The renderer's print window picks the pages and settings (and shows the preview); the page pictures
// (made by the backend) go into a hidden window, one per sheet, laid out the same way as the preview.
ipcMain.handle('printers', () => win.webContents.getPrintersAsync());
// Printing: each picture is one whole sheet, already the printer's paper size with the page placed on it
// (made by the backend), so the job's paper size is exactly the sheet's. opt.dryRun = a file path: save what
// would be sent to the printer as a PDF instead of printing (for testing without a printer).
ipcMain.handle('print-pages', async (_e, pages, opt = {}) => {
  if (!pages || !pages.length) return { error: 'Nothing to print' };
  const sw = pages[0].w, sh = pages[0].h;
  const landscape = sw > sh;
  const [nw, nh] = [Math.min(sw, sh), Math.max(sw, sh)];   // the paper, narrow side first
  const html = `<!doctype html><meta charset="utf-8"><title>MyPDFKit</title><style>
    @page { size: ${sw}pt ${sh}pt; margin: 0; }
    html, body { margin: 0; background: #fff; }
    .pg { width: ${sw}pt; height: ${sh}pt; page-break-after: always; break-after: page; overflow: hidden; }
    .pg:last-child { page-break-after: auto; break-after: auto; }
    img { display: block; width: ${sw}pt; height: ${sh}pt; }
  </style>` + pages.map(p => `<div class="pg"><img src="${require('url').pathToFileURL(p.path).href}"></div>`).join('');
  const file = path.join(path.dirname(pages[0].path), 'print.html');
  fs.writeFileSync(file, html);
  const pw = new BrowserWindow({ show: false, parent: win, webPreferences: { sandbox: true } });
  try {
    await pw.loadFile(file);
    await pw.webContents.executeJavaScript('Promise.all([...document.images].map(i => i.decode().catch(() => {})))');
    // named sizes go by name; others (4x6 labels...) as their size in microns. pageSize must be given:
    // Electron 34 otherwise sends an empty page size and no dialog ever opens
    const NAMED = { Letter: [612, 792], Legal: [612, 1008], Tabloid: [792, 1224], A3: [841.89, 1190.55], A4: [595.28, 841.89], A5: [419.53, 595.28] };
    const name = Object.keys(NAMED).find(k => Math.abs(NAMED[k][0] - nw) < 2 && Math.abs(NAMED[k][1] - nh) < 2);
    const pageSize = name || { width: Math.round(nw / 72 * 25400), height: Math.round(nh / 72 * 25400) };
    const settings = { silent: !opt.system, printBackground: true, landscape, pageSize, margins: { marginType: 'none' },
      copies: opt.copies || 1, color: opt.color !== false, duplexMode: opt.duplex || 'simplex' };
    if (opt.deviceName) settings.deviceName = opt.deviceName;
    if (opt.dryRun) {
      const pdf = await pw.webContents.printToPDF({ landscape, printBackground: true, preferCSSPageSize: true,
        pageSize: name || { width: nw / 72, height: nh / 72 }, margins: { top: 0, bottom: 0, left: 0, right: 0 } });
      fs.writeFileSync(opt.dryRun, pdf);
      fs.writeFileSync(opt.dryRun.replace(/\.pdf$/i, '') + '.json', JSON.stringify({ settings, sheets: pages.length, sheetPt: [sw, sh] }, null, 1));
      return { ok: true, dryRun: opt.dryRun, settings };
    }
    return await new Promise(resolve => pw.webContents.print(settings,
      (ok, reason) => resolve(ok || reason === 'cancelled' || reason === 'Print job canceled' ? { ok, settings } : { error: "Couldn't print: " + reason })));
  } finally {
    pw.destroy();
    fs.rm(path.dirname(file), { recursive: true, force: true }, () => {});
  }
});
// a save-as box for test prints and saved labels
ipcMain.handle('save-dialog', async (_e, suggested, kind) => {
  const r = await dialog.showSaveDialog(win, { title: 'Save', defaultPath: suggested, filters: kind === 'png' ? [{ name: 'Picture', extensions: ['png'] }] : [{ name: 'PDF', extensions: ['pdf'] }] });
  return r.canceled ? null : r.filePath;
});
ipcMain.handle('copy-image', (_e, file) => { require('electron').clipboard.writeImage(require('electron').nativeImage.createFromPath(file)); return true; });

function killServer() {
  if (serverProc && !serverProc.killed) {
    // /T kills the whole process tree
    spawn('taskkill', ['/pid', String(serverProc.pid), '/f', '/t'], { windowsHide: true });
  }
}
// the backend keeps a hidden Word ready for quick conversions: ask it to close that first (up to 15 seconds)
function closeBackend() {
  return new Promise(resolve => {
    if (!serverProc || serverProc.killed) return resolve();
    const done = () => { killServer(); resolve(); };
    const req = http.request({ host: '127.0.0.1', port: SERVER_PORT, path: '/shutdown', method: 'POST', timeout: 15000 }, res => { res.resume(); res.on('end', done); });
    req.on('error', done); req.on('timeout', () => { req.destroy(); done(); });
    req.end();
  });
}
let closing = false;
// every way of closing comes here: the windows close, the backend closes its Word, then the app exits
function quitCleanly() {
  if (closing) return;
  closing = true;
  // windows close normally, so the page still saves the open tabs
  const shut = BrowserWindow.getAllWindows().filter(w => !w.isDestroyed()).map(w => new Promise(r => { w.once('closed', r); w.close(); }));
  const waitWindows = Promise.race([Promise.all(shut), new Promise(r => setTimeout(r, 3000))]);
  Promise.all([waitWindows, closeBackend()]).then(() => app.exit(0));
}
app.on('before-quit', e => { if (!closing) { e.preventDefault(); quitCleanly(); } });
app.on('window-all-closed', quitCleanly);
