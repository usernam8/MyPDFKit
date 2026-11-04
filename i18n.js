// The app in another language. English is written in the page itself; for another language, text that
// exactly matches an entry here is swapped (also in windows that open later), and put back for English.
const I18N = {
  es: {
    'All tools': 'Todas las herramientas', 'Edit': 'Editar', 'Convert': 'Convertir', 'E-Sign': 'Firmar', 'Settings': 'Configuración',
    'Files': 'Archivos', 'Add': 'Agregar', 'Clear': 'Borrar', 'Recent': 'Recientes', 'Document': 'Documento', 'More tools': 'Más herramientas',
    'Drop files': 'Suelta archivos aquí', 'or choose files': 'o elige archivos', 'Find any tool': 'Buscar una herramienta',
    'Find, or go to page #': 'Buscar, o ir a la página #', 'No files yet': 'Aún no hay archivos', 'Backend connected': 'Listo',
    // tools
    'Compress': 'Comprimir', 'Compress a PDF': 'Comprimir un PDF', 'Protect with password': 'Proteger con contraseña', 'Protect a PDF': 'Proteger un PDF',
    'Add page numbers': 'Numerar páginas', 'Header and footer': 'Encabezado y pie de página', 'Blank page': 'Página en blanco',
    'Watermark': 'Marca de agua', 'Page size and margins': 'Tamaño de página y márgenes', 'Insert pages': 'Insertar páginas',
    'Privacy check': 'Revisión de privacidad', 'Lock and finalize': 'Bloquear y finalizar', 'Find and replace': 'Buscar y reemplazar',
    'Bates numbering': 'Numeración Bates', 'Redact': 'Tachar (redactar)', 'Recognize text (OCR)': 'Reconocer texto (OCR)',
    'Create a PDF': 'Crear un PDF', 'Combine files': 'Combinar archivos', 'Edit a PDF': 'Editar un PDF', 'Convert to PDF': 'Convertir a PDF',
    'Export a PDF': 'Exportar un PDF', 'PDF to Word': 'PDF a Word', 'PDF to Excel': 'PDF a Excel', 'PDF to PowerPoint': 'PDF a PowerPoint',
    'Split a PDF': 'Dividir un PDF', 'Delete pages': 'Eliminar páginas', 'Fill a form': 'Llenar un formulario', 'Compare files': 'Comparar archivos',
    'Comment': 'Comentar', 'Fill from spreadsheet': 'Llenar desde hoja de cálculo', 'Box labels from packing list': 'Etiquetas de cajas desde la lista de empaque',
    'Side by side': 'Lado a lado', 'Bookmarks': 'Marcadores', 'Sign with a digital ID': 'Firmar con un ID digital', 'Check signatures': 'Verificar firmas',
    'Many files at once': 'Muchos archivos a la vez', 'Reading view': 'Vista de lectura', 'Measure': 'Medir', 'Save as PDF/A': 'Guardar como PDF/A',
    'Save as PDF/A (archive)': 'Guardar como PDF/A (archivo)', 'Shapes': 'Formas', 'Stamps and initials': 'Sellos e iniciales',
    'Prepare for signing': 'Preparar para firmar', 'Translate': 'Traducir', 'Read aloud': 'Leer en voz alta', 'Export all files': 'Exportar todos los archivos',
    'Extract pages': 'Extraer páginas', 'Extract page': 'Extraer página', 'Print': 'Imprimir', 'Print an area': 'Imprimir un área',
    'Make it my PDF app': 'Usar como mi app de PDF', 'Unlock a file I locked': 'Desbloquear un archivo que bloqueé', 'Summarize': 'Resumir',
    'Ask Claude about this PDF': 'Preguntarle a Claude sobre este PDF', 'Send a problem report': 'Enviar un reporte de problema', 'Fill and sign': 'Llenar y firmar',
    'Convert all to PDF': 'Convertir todo a PDF', 'Merge into one PDF': 'Unir en un PDF', 'Split into pages': 'Dividir en páginas',
    'Convert and combine': 'Convertir y combinar', 'Sign': 'Firmar', 'Signature': 'Firma', 'Quick marks': 'Marcas rápidas', 'Keep adding': 'Seguir agregando',
    'click one, then the page': 'elige una y luego la página', 'Date': 'Fecha', 'Name': 'Nombre', 'Initials': 'Iniciales',
    'Add text': 'Agregar texto', 'Edit text': 'Editar texto', 'Add image': 'Agregar imagen', 'Draw': 'Dibujar', 'Highlight': 'Resaltar',
    'Select': 'Seleccionar', 'Undo': 'Deshacer', 'Save': 'Guardar', 'Open': 'Abrir', 'Close': 'Cerrar', 'Cancel': 'Cancelar', 'OK': 'Aceptar',
    'Apply': 'Aplicar', 'Done': 'Listo', 'Copy': 'Copiar', 'Fill': 'Llenar', 'Finalize': 'Finalizar', 'Delete page': 'Eliminar página', 'Crop page': 'Recortar página',
    'Rotate this page 90 degrees': 'Girar esta página 90 grados', 'Go to page': 'Ir a la página', 'Next page': 'Página siguiente', 'Previous page': 'Página anterior',
    'Zoom in (Ctrl + wheel)': 'Acercar (Ctrl + rueda)', 'Zoom out (Ctrl + wheel)': 'Alejar (Ctrl + rueda)', 'Fit page': 'Ajustar página',
    'Page thumbnails': 'Miniaturas', 'All bookmarks': 'Todos los marcadores', 'Dark pages (easier on the eyes at night)': 'Páginas oscuras (mejor de noche)',
    'Bookmark this page (Ctrl+B)': 'Marcar esta página (Ctrl+B)', 'Print (Ctrl+P)': 'Imprimir (Ctrl+P)', 'Undo (Ctrl+Z)': 'Deshacer (Ctrl+Z)',
    'Save as a new PDF next to the original (Ctrl+S)': 'Guardar como PDF nuevo junto al original (Ctrl+S)', 'Show or hide the left panel': 'Mostrar u ocultar el panel izquierdo',
    // print window
    'Printer': 'Impresora', 'Paper': 'Papel', 'Pages': 'Páginas', 'Copies': 'Copias', 'Orientation': 'Orientación', 'Size on paper': 'Tamaño en el papel',
    'Color': 'Color', 'Two-sided': 'Doble cara', 'Layout': 'Distribución', 'Whole page': 'Página completa', 'Label only': 'Solo la etiqueta',
    'Automatic': 'Automática', 'Portrait': 'Vertical', 'Landscape': 'Horizontal', 'Fit to paper': 'Ajustar al papel', 'Shrink large pages only': 'Reducir solo páginas grandes',
    'Actual size': 'Tamaño real', 'Exact percent...': 'Porcentaje exacto...', 'Black and white': 'Blanco y negro', 'Toner saver (lighter)': 'Ahorro de tóner (más claro)',
    'Off': 'No', 'Flip on long edge': 'Voltear por el lado largo', 'Flip on short edge': 'Voltear por el lado corto', 'By hand: fronts, then backs': 'A mano: frentes y luego reversos',
    'Skip blank pages': 'Omitir páginas en blanco', 'Print comments and markup': 'Imprimir comentarios y marcas', 'One per sheet': 'Una por hoja',
    '2 per sheet': '2 por hoja', '4 per sheet': '4 por hoja', '6 per sheet': '6 por hoja', '9 per sheet': '9 por hoja', 'Booklet (fold in half)': 'Folleto (doblar a la mitad)',
    'Poster (across several sheets)': 'Póster (en varias hojas)', 'Sticker sheet (Avery and similar)': 'Hoja de etiquetas (Avery y similares)', 'Poster size': 'Tamaño del póster',
    'Sticker sheet': 'Hoja de etiquetas', 'Start at sticker': 'Empezar en la etiqueta', 'Adjust area...': 'Ajustar área...', 'Save as PDF': 'Guardar como PDF',
    'More options...': 'Más opciones...', 'Custom...': 'Personalizado...', 'Test print: save as PDF (no printer)': 'Prueba: guardar como PDF (sin impresora)',
    'Use this area': 'Usar esta área', 'All pages': 'Todas las páginas', 'Drag the corners': 'Arrastra las esquinas', 'Getting the preview ready…': 'Preparando la vista previa…',
    'Now the backs': 'Ahora los reversos', 'Print the backs': 'Imprimir los reversos', 'Stop': 'Detener',
    // lock, forms, summaries
    'The text': 'El texto', 'Lock it': 'Bloquearlo', 'Allow printing': 'Permitir imprimir', 'Allow copying text': 'Permitir copiar texto',
    'Password to make changes later': 'Contraseña para hacer cambios después', 'Password again': 'Repite la contraseña', 'Digital ID': 'ID digital',
    'Make every edit part of the page (comments, shapes, signatures, form fields)': 'Integrar todos los cambios a la página (comentarios, formas, firmas, campos)',
    'Remove hidden things: covered-up text, earlier versions, author details, attachments, scripts': 'Quitar lo oculto: texto tapado, versiones anteriores, datos del autor, adjuntos, scripts',
    'Summary': 'Resumen', 'Sections': 'Secciones', 'Main points': 'Puntos principales', 'Dates': 'Fechas', 'Amounts': 'Montos', 'Emails': 'Correos',
    'Phone numbers': 'Teléfonos', 'Comes up most': 'Lo que más aparece', 'Ask Claude…': 'Preguntarle a Claude…', 'Ask': 'Preguntar', 'Copy answer': 'Copiar respuesta',
    'Summarize it': 'Resumirlo', 'Dates and amounts': 'Fechas y montos', 'Translate to Spanish': 'Traducir al español', 'Explain simply': 'Explicar fácil',
    'Your Claude API key': 'Tu clave de API de Claude', 'Make it searchable': 'Hacerlo buscable', 'Nothing to read yet': 'Aún no hay nada que leer',
    'Welcome to MyPDFKit': 'Bienvenido a MyPDFKit', 'Get started': 'Empezar', 'Open PDFs with MyPDFKit when I double-click them': 'Abrir los PDF con MyPDFKit al hacer doble clic',
    // settings
    'Language': 'Idioma', 'Tools shown in the side panel': 'Herramientas en el panel lateral', 'Show all': 'Mostrar todas', 'Claude (optional)': 'Claude (opcional)',
    'Remove my key': 'Quitar mi clave', 'Save key': 'Guardar clave', 'PDF app': 'App de PDF', 'Help': 'Ayuda', 'Version': 'Versión',
    'Add your Claude API key to ask Claude about your files (console.anthropic.com).': 'Agrega tu clave de API de Claude para preguntarle sobre tus archivos (console.anthropic.com).',
    'Your Claude API key is saved on this computer. Ask Claude about this PDF uses it.': 'Tu clave de API de Claude está guardada en esta computadora. Preguntarle a Claude la usa.',
    'Summarize, or ask Claude about this file': 'Resumir, o preguntarle a Claude sobre este archivo', 'Find any tool (Ctrl+K)': 'Buscar una herramienta (Ctrl+K)',
    'Uncheck what you never use: it disappears from the side panel (it stays in All tools).': 'Desmarca lo que nunca usas: desaparece del panel lateral (sigue en Todas las herramientas).',
  },
};
const I18N_ATTRS = ['title', 'placeholder', 'aria-label', 'data-empty'];
let uiLang = 'en';
const _orig = new WeakMap();   // node -> its English text (to switch back)
function _tr(str) {
  const dict = I18N[uiLang]; if (!dict) return null;
  const key = str.trim(); if (!key || !(key in dict)) return null;
  return str.replace(key, dict[key]);
}
function translateNode(node) {
  if (uiLang === 'en') return;
  if (node.nodeType === 3) {
    const t = _tr(node.nodeValue);
    if (t !== null && t !== node.nodeValue) { if (!_orig.has(node)) _orig.set(node, node.nodeValue); node.nodeValue = t; }
    return;
  }
  if (node.nodeType !== 1 || node.tagName === 'SCRIPT' || node.tagName === 'STYLE' || node.tagName === 'TEXTAREA' || node.isContentEditable) return;
  if (node.closest && node.closest('#canvas-wrapper .pdf-text, .pdf-field')) return;   // the document's own words are never translated
  for (const a of I18N_ATTRS) {
    const v = node.getAttribute && node.getAttribute(a);
    if (v) { const t = _tr(v); if (t !== null) { node.dataset['en' + a.replace('-', '')] = node.dataset['en' + a.replace('-', '')] || v; node.setAttribute(a, t); } }
  }
  if (node.tagName === 'OPTION' || node.tagName === 'BUTTON' || node.childNodes.length) node.childNodes.forEach(translateNode);
}
let _observer = null;
function setUiLanguage(lang) {
  if (lang !== 'en' && !I18N[lang]) lang = 'en';
  if (_observer) { _observer.disconnect(); _observer = null; }
  if (uiLang !== 'en' && lang !== uiLang) {   // another language: first put every changed text back to English
    const walk = n => {
      if (n.nodeType === 3 && _orig.has(n)) { n.nodeValue = _orig.get(n); _orig.delete(n); }
      if (n.nodeType === 1) {
        for (const a of I18N_ATTRS) { const k = 'en' + a.replace('-', ''); if (n.dataset && n.dataset[k]) { n.setAttribute(a, n.dataset[k]); delete n.dataset[k]; } }
        n.childNodes.forEach(walk);
      }
    };
    walk(document.body);
  }
  uiLang = lang;
  document.documentElement.lang = lang;
  if (_observer) _observer.disconnect();
  if (lang === 'en') return;
  translateNode(document.body);
  _observer = new MutationObserver(list => {
    for (const m of list) {
      if (m.type === 'characterData') translateNode(m.target);
      else if (m.type === 'attributes') translateNode(m.target);
      else m.addedNodes.forEach(translateNode);
    }
  });
  _observer.observe(document.body, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: I18N_ATTRS });
}
