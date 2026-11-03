const { app, BrowserWindow, dialog, Menu, ipcMain } = require('electron');
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
  const tail = serverLog.trim().split(/\r?\n/).slice(-12).join('\n');
  dialog.showErrorBox('MyPDFKit could not start its converter',
    'Opening and converting files won\'t work until it starts. Try closing MyPDFKit and opening it again. ' +
    'If it keeps happening, reinstall MyPDFKit.\n\n' + reason + (tail ? '\n\nDetails:\n' + tail : ''));
}

function waitForServer(retries = 200) {
  return new Promise((resolve, reject) => {
    const attempt = n => {
      if (serverExit) return reject(new Error(serverExit));
      http.get(`http://127.0.0.1:${SERVER_PORT}/ping`, res => { res.resume(); resolve(); })
        .on('error', () => (n <= 0 ? reject(new Error('It didn\'t answer within 30 seconds.')) : setTimeout(() => attempt(n - 1), 150)));
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

function createWindow() {
  win = new BrowserWindow({
    width: 1400, height: 900, minWidth: 980, minHeight: 620, title: 'MyPDFKit',
    backgroundColor: '#25262a',
    show: false,
    autoHideMenuBar: true,
    // The app's own top bar acts as the title bar; Windows draws the min/max/close buttons over it.
    titleBarStyle: 'hidden',
    titleBarOverlay: { color: '#1c1d20', symbolColor: '#e6e5e1', height: 44 },
    webPreferences: {
      // index.html uses require('fs') / require('electron') directly.
      // Safe here because the page now loads only local files (pdf.js is bundled).
      nodeIntegration: true,
      contextIsolation: false,
    },
  });
  if (app.isPackaged) Menu.setApplicationMenu(null);
  win.once('ready-to-show', () => win.show());
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
  if (!app.isPackaged) return;
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
ipcMain.handle('user-data', () => app.getPath('userData'));
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
ipcMain.handle('print-pages', async (_e, pages, opt = {}) => {
  if (!pages || !pages.length) return { error: 'Nothing to print' };
  const landscape = opt.landscape !== undefined ? opt.landscape : pages.filter(p => p.w > p.h).length > pages.length / 2;
  let [sw, sh] = opt.paperSize || [612, 792];
  if (landscape) [sw, sh] = [sh, sw];
  const size = p => {   // the picture's size on the sheet, in points (as in the preview)
    const fit = Math.min(sw / p.w, sh / p.h);
    const f = opt.scale === 'actual' ? 1 : opt.scale === 'shrink' ? Math.min(1, fit) : fit;
    return `width:${p.w * f}pt;height:${p.h * f}pt`;
  };
  const html = `<!doctype html><meta charset="utf-8"><title>MyPDFKit</title><style>
    @page { size: ${sw}pt ${sh}pt; margin: 0; }
    html, body { margin: 0; background: #fff; }
    .pg { width: ${sw}pt; height: ${sh}pt; display: flex; align-items: center; justify-content: center; page-break-after: always; break-after: page; overflow: hidden; }
    .pg:last-child { page-break-after: auto; break-after: auto; }
    img { flex: none; ${opt.color === false ? 'filter: grayscale(1);' : ''} }
  </style>` + pages.map(p => `<div class="pg"><img style="${size(p)}" src="${require('url').pathToFileURL(p.path).href}"></div>`).join('');
  const file = path.join(path.dirname(pages[0].path), 'print.html');
  fs.writeFileSync(file, html);
  const pw = new BrowserWindow({ show: false, parent: win, webPreferences: { sandbox: true } });
  try {
    await pw.loadFile(file);
    await pw.webContents.executeJavaScript('Promise.all([...document.images].map(i => i.decode().catch(() => {})))');
    // pageSize must be given: Electron 34 otherwise sends an empty page size and no dialog ever opens
    const settings = { silent: !opt.system, printBackground: true, landscape, pageSize: opt.paper || 'Letter', margins: { marginType: 'none' },
      copies: opt.copies || 1, color: opt.color !== false, duplexMode: opt.duplex || 'simplex' };
    if (opt.deviceName) settings.deviceName = opt.deviceName;
    return await new Promise(resolve => pw.webContents.print(settings,
      (ok, reason) => resolve(ok || reason === 'cancelled' || reason === 'Print job canceled' ? { ok } : { error: "Couldn't print: " + reason })));
  } finally {
    pw.destroy();
    fs.rm(path.dirname(file), { recursive: true, force: true }, () => {});
  }
});

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
