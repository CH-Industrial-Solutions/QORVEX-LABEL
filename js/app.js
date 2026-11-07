/* QORVEX_Label · lógica de interfaz y de archivo Excel.
   Una sola estación. El Excel es la fuente de trazabilidad.
   Pura lógica de IDs en js/ids.js (QorvexIds). */
(function () {
  'use strict';

  const IDS = window.QorvexIds;
  const OPERACIONES = ['Validación', 'Secuencia'];
  const MAX_QTY = 500;
  const SHEETS = {
    comp: 'Componentes', line: 'Lineas', cfg: 'Config', users: 'Usuarios', cons: 'Consecutivos', reg: 'Registro'
  };
  const CONS_HEADERS = ['Operacion', 'Componente', 'Linea', 'Prefijo', 'Ultimo'];
  const REG_HEADERS = ['Fecha', 'Hora', 'Usuario', 'Operacion', 'Componente', 'Linea', 'ID', 'Posicion', 'Lote', 'Estado', 'Impresora'];

  const S = {
    wb: null,            // libro SheetJS en memoria
    handle: null,        // FileSystemFileHandle del .xlsx
    user: null,          // fila de Usuarios
    lastActivity: Date.now(),
    printing: false,
    adminUnlocked: false,
    batch: null,         // lote en vista previa
    timer: null
  };

  // ---------- Utilidades ----------
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const today = () => { const d = new Date(); return String(d.getDate()).padStart(2, '0') + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + d.getFullYear(); };
  const nowHM = () => { const d = new Date(); return String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'); };

  function toast(msg) {
    const t = $('toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(() => (t.hidden = true), 2800);
  }

  function show(viewId) {
    ['view-open', 'view-login', 'view-printer', 'view-main'].forEach((v) => ($(v).hidden = v !== viewId));
    $('topbar').hidden = viewId === 'view-open' || viewId === 'view-login';
  }

  function showTab(name) {
    document.querySelectorAll('.nav-item').forEach((b) => b.classList.toggle('active', b.dataset.tab === name));
    ['home', 'generate', 'history', 'config', 'about'].forEach((n) => ($('tab-' + n).hidden = n !== name));
    if (name === 'home') renderHome();
    if (name === 'history') renderHistory();
    if (name === 'config') { $('config-locked').hidden = S.adminUnlocked; $('config-unlocked').hidden = !S.adminUnlocked; if (S.adminUnlocked) renderConfig(); }
  }

  // ---------- Tema (oscuro / claro, con colores QORVEX) ----------
  function applyTheme(t) {
    document.documentElement.setAttribute('data-theme', t);
    try { localStorage.setItem('qorvex-theme', t); } catch (e) { /* sin almacenamiento: sólo visual */ }
  }
  function toggleTheme() {
    applyTheme(document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
  }
  function initTheme() {
    let t = 'dark';
    try { t = localStorage.getItem('qorvex-theme') || 'dark'; } catch (e) { /* ignorar */ }
    applyTheme(t);
  }

  // ---------- Inicio: último y próximo ID por contador ----------
  function renderHome() {
    const cons = normCons(sheetRows(SHEETS.cons));
    const line = (op, comp, linea, r) => {
      const prefix = r ? r.Prefijo : basePrefix(comp);
      const ult = r ? r.Ultimo : 0;
      const last = ult > 0 ? IDS.formatId(prefix, ult) : '—';
      const next = IDS.allocate(prefix, ult, 1).ids[0];
      return `<tr><td>${esc(op)}</td><td>${esc(comp)}</td><td>${esc(linea || '—')}</td><td>${esc(prefix)}</td><td><b>${last}</b></td><td>${next}</td></tr>`;
    };
    const rows = sheetRows(SHEETS.comp).map((c) => line('Validación', c.Nombre, '', findCounter(cons, 'Validación', c.Nombre, ''))).join('');
    const seq = cons.filter((r) => r.Operacion === 'Secuencia').map((r) => line('Secuencia', r.Componente, r.Linea, r)).join('');
    const tb = document.querySelector('#home-table tbody');
    tb.innerHTML = (rows + seq) || '<tr><td colspan="6" class="muted">Sin componentes configurados.</td></tr>';
  }

  // ---------- Excel ----------
  function sheetRows(name) {
    const ws = S.wb.Sheets[name];
    return ws ? XLSX.utils.sheet_to_json(ws, { defval: '' }) : [];
  }

  function setRows(name, rows, headers) {
    S.wb.Sheets[name] = XLSX.utils.json_to_sheet(rows, { header: headers });
  }

  async function saveWorkbook() {
    const out = XLSX.write(S.wb, { bookType: 'xlsx', type: 'array' });
    if (S.handle && S.handle.createWritable) {
      const w = await S.handle.createWritable();
      await w.write(out);
      await w.close();
    } else {
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'QORVEX_Label.xlsx';
      a.click();
    }
  }

  // Re-lee el archivo antes de escribir para no pisar cambios.
  async function reloadWorkbook() {
    if (!S.handle) return;
    const file = await S.handle.getFile();
    S.wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
  }

  function defaultWorkbook() {
    const wb = XLSX.utils.book_new();
    const add = (name, rows, headers) => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows, { header: headers }), name);
    add(SHEETS.comp, [
      { Nombre: 'Arnés', Prefijo: 'AA', Activo: 'SI' },
      { Nombre: 'BCM', Prefijo: 'BA', Activo: 'SI' },
      { Nombre: 'Cluster', Prefijo: 'CA', Activo: 'SI' },
      { Nombre: 'Switch Gde', Prefijo: 'DA', Activo: 'SI' },
      { Nombre: 'Switch CH', Prefijo: 'EA', Activo: 'SI' },
      { Nombre: 'Llaves', Prefijo: 'FA', Activo: 'SI' }
    ], ['Nombre', 'Prefijo', 'Activo']);
    add(SHEETS.line, ['011', '065', '014', '044'].map((l) => ({ Linea: l, Activa: 'SI' })), ['Linea', 'Activa']);
    add(SHEETS.cfg, [
      { Clave: 'MINUTOS_INACTIVIDAD', Valor: 5 },
      { Clave: 'ADMIN_PASS', Valor: 'admin2026' }
    ], ['Clave', 'Valor']);
    add(SHEETS.users, [{ Usuario: 'admin', Nombre: 'Administrador', Contrasena: 'admin', Impresora: '' }], ['Usuario', 'Nombre', 'Contrasena', 'Impresora']);
    add(SHEETS.cons, [['Arnés', 'AA'], ['BCM', 'BA'], ['Cluster', 'CA'], ['Switch Gde', 'DA'], ['Switch CH', 'EA'], ['Llaves', 'FA']].map((c) => ({ Operacion: 'Validación', Componente: c[0], Linea: '', Prefijo: c[1], Ultimo: 0 })), CONS_HEADERS);
    add(SHEETS.reg, [], REG_HEADERS);
    return wb;
  }

  // ---------- Config leída ----------
  function cfgValue(clave, def) {
    const r = sheetRows(SHEETS.cfg).find((x) => x.Clave === clave);
    return r ? r.Valor : def;
  }
  function timeoutMs() { return Number(cfgValue('MINUTOS_INACTIVIDAD', 5)) * 60 * 1000; }
  function activeComponents() { return sheetRows(SHEETS.comp).filter((c) => String(c.Activo).toUpperCase() === 'SI'); }
  // ---------- Contadores de consecutivo ----------
  // Validación: un contador por componente (sin importar la línea).
  // Secuencia: un contador por componente y línea.
  function normCons(rows) {
    return rows.map((r) => ({
      Operacion: r.Operacion || 'Validación',
      Componente: String(r.Componente),
      Linea: r.Linea === undefined || r.Linea === null ? '' : String(r.Linea),
      Prefijo: r.Prefijo,
      Ultimo: Number(r.Ultimo) || 0
    }));
  }
  const lineaKey = IDS.lineaKey;
  const findCounter = IDS.findCounter;
  function basePrefix(comp) {
    const c = sheetRows(SHEETS.comp).find((x) => x.Nombre === comp);
    return c ? c.Prefijo : null;
  }
  // Devuelve el contador; si no existe, lo crea (no se guarda hasta reservar).
  function getCounter(rows, op, comp, linea) {
    let r = findCounter(rows, op, comp, linea);
    if (!r) {
      const pre = basePrefix(comp);
      if (!IDS.validPrefix(pre)) throw new Error('El componente no tiene prefijo válido.');
      r = { Operacion: op, Componente: comp, Linea: lineaKey(op, linea), Prefijo: pre, Ultimo: 0 };
      rows.push(r);
    }
    return r;
  }

  function activeLines() { return sheetRows(SHEETS.line).filter((l) => String(l.Activa).toUpperCase() === 'SI'); }

  // ---------- Apertura del libro: se elige la CARPETA, no el archivo ----------
  const FILE_NAME = 'QORVEX_Label.xlsx';

  // Guarda la carpeta elegida en IndexedDB para recordarla en la siguiente sesión.
  function idb() {
    return new Promise((res, rej) => {
      const r = indexedDB.open('qorvex-label', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('kv');
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  }
  async function idbGet(k) {
    const db = await idb();
    return new Promise((res, rej) => {
      const q = db.transaction('kv').objectStore('kv').get(k);
      q.onsuccess = () => res(q.result);
      q.onerror = () => rej(q.error);
    });
  }
  async function idbSet(k, v) {
    const db = await idb();
    return new Promise((res, rej) => {
      const t = db.transaction('kv', 'readwrite');
      t.objectStore('kv').put(v, k);
      t.oncomplete = () => res();
      t.onerror = () => rej(t.error);
    });
  }

  // Abre la carpeta: carga el libro o, si no existe, ofrece crearlo.
  async function openFolder(dir) {
    S.dir = dir;
    try {
      S.handle = await dir.getFileHandle(FILE_NAME);
      await reloadWorkbook();
      toast('Libro cargado: ' + dir.name);
    } catch (e) {
      if (e.name !== 'NotFoundError') throw e;
      if (!confirm('La carpeta no contiene ' + FILE_NAME + '.\n\n¿Crear el libro con datos iniciales en esta carpeta?')) return;
      S.handle = await dir.getFileHandle(FILE_NAME, { create: true });
      S.wb = defaultWorkbook();
      await saveWorkbook();
      toast('Libro creado. Usuario inicial: admin / admin. Cámbialo en Configuración.');
    }
    goLogin();
  }

  async function openWorkbook() {
    $('open-error').textContent = '';
    if (!window.showDirectoryPicker) {
      $('open-error').textContent = 'Usa Microsoft Edge o Google Chrome para seleccionar la carpeta.';
      return;
    }
    try {
      const dir = await window.showDirectoryPicker({ mode: 'readwrite' });
      try { await idbSet('lastDir', dir); } catch (e) { /* sin memoria de carpeta */ }
      await openFolder(dir);
    } catch (e) {
      if (e.name !== 'AbortError') $('open-error').textContent = 'No se pudo abrir la carpeta: ' + e.message;
    }
  }

  // Reabre la última carpeta. Pide permiso de lectura y escritura de nuevo si el navegador lo requiere.
  async function openLastFolder() {
    $('open-error').textContent = '';
    try {
      const dir = await idbGet('lastDir');
      if (!dir) { $('open-error').textContent = 'No hay una carpeta recordada.'; return; }
      let perm = await dir.queryPermission({ mode: 'readwrite' });
      if (perm !== 'granted') perm = await dir.requestPermission({ mode: 'readwrite' });
      if (perm !== 'granted') { $('open-error').textContent = 'Sin permiso para escribir en la carpeta.'; return; }
      await openFolder(dir);
    } catch (e) {
      $('open-error').textContent = 'No se pudo reabrir la carpeta: ' + e.message;
    }
  }

  async function showLastFolderButton() {
    try {
      const dir = await idbGet('lastDir');
      if (dir) {
        $('btn-last').textContent = 'Abrir última carpeta: ' + dir.name;
        $('btn-last').hidden = false;
      }
    } catch (e) { /* sin memoria: no se muestra el botón */ }
  }

  function goLogin() {
    $('login-user').value = '';
    $('login-pass').value = '';
    $('login-error').textContent = '';
    show('view-login');
    $('login-user').focus();
  }

  // ---------- Sesión ----------
  function login(userVal, passVal) {
    const u = sheetRows(SHEETS.users).find((x) => String(x.Usuario) === userVal && String(x.Contrasena) === passVal);
    return u || null;
  }

  function startSession(u) {
    S.user = u;
    S.adminUnlocked = false;
    S.lastActivity = Date.now();
    $('whoami').textContent = 'Usuario: ' + (u.Nombre || u.Usuario);
    if (!u.Impresora) { goPrinter(); } else { enterMain(); }
  }

  function goPrinter() {
    $('printer-name').value = S.user.Impresora || '';
    $('printer-cancel').hidden = !S.user.Impresora;
    show('view-printer');
    $('printer-name').focus();
  }

  async function savePrinter(name) {
    await reloadWorkbook();
    const rows = sheetRows(SHEETS.users);
    const r = rows.find((x) => String(x.Usuario) === String(S.user.Usuario));
    if (r) r.Impresora = name;
    setRows(SHEETS.users, rows, ['Usuario', 'Nombre', 'Contrasena', 'Impresora']);
    await saveWorkbook();
    S.user.Impresora = name;
    toast('Impresora guardada: ' + name);
  }

  function enterMain() {
    fillSelectors();
    show('view-main');
    showTab('home');
    updateRange();
    startIdleWatch();
  }

  function logout(reason) {
    S.user = null;
    S.batch = null;
    S.adminUnlocked = false;
    stopIdleWatch();
    $('preview-card').hidden = true;
    if (reason) toast(reason);
    show('view-open');
  }

  function startIdleWatch() {
    stopIdleWatch();
    const bump = () => { S.lastActivity = Date.now(); };
    ['pointerdown', 'keydown', 'input'].forEach((ev) => document.addEventListener(ev, bump, true));
    S.timer = setInterval(() => {
      if (S.printing) { S.lastActivity = Date.now(); return; }
      if (IDS.isIdleExpired(S.lastActivity, Date.now(), timeoutMs())) logout('Sesión cerrada por inactividad.');
    }, 15000);
  }
  function stopIdleWatch() { clearInterval(S.timer); S.timer = null; }

  // ---------- Generar ----------
  function fillSelectors() {
    $('gen-op').innerHTML = OPERACIONES.map((o) => `<option>${esc(o)}</option>`).join('');
    $('gen-comp').innerHTML = activeComponents().map((c) => `<option value="${esc(c.Nombre)}">${esc(c.Nombre)} (${esc(c.Prefijo)})</option>`).join('');
    $('gen-line').innerHTML = activeLines().map((l) => `<option>${esc(l.Linea)}</option>`).join('');
  }

  function updateRange() {
    const qty = Number($('gen-qty').value);
    const comp = $('gen-comp').value, op = $('gen-op').value, linea = $('gen-line').value;
    if (!comp || !qty || qty < 1 || qty > MAX_QTY) { $('gen-range').textContent = 'Cantidad inválida (1 a ' + MAX_QTY + ').'; return null; }
    const cons = normCons(sheetRows(SHEETS.cons));
    const row = findCounter(cons, op, comp, linea) || { Prefijo: basePrefix(comp), Ultimo: 0 };
    if (!IDS.validPrefix(row.Prefijo)) { $('gen-range').textContent = 'El componente no tiene prefijo válido.'; return null; }
    const alloc = IDS.allocate(row.Prefijo, row.Ultimo, qty);
    const first = alloc.ids[0], last = alloc.ids[alloc.ids.length - 1];
    const cambio = alloc.prefix !== row.Prefijo ? ' (incluye cambio de prefijo a ' + alloc.prefix + ')' : '';
    const contexto = op === 'Secuencia' ? ' · línea ' + linea : '';
    $('gen-range').textContent = `Se generará de ${first} a ${last} (${qty} etiquetas)${contexto}${cambio}.`;
    return alloc;
  }

  async function generate() {
    $('gen-error').textContent = '';
    const op = $('gen-op').value, comp = $('gen-comp').value, linea = $('gen-line').value;
    const qty = Number($('gen-qty').value);
    if (!op || !comp || !linea) { $('gen-error').textContent = 'Selecciona operación, componente y línea.'; return; }
    if (!Number.isInteger(qty) || qty < 1 || qty > MAX_QTY) { $('gen-error').textContent = 'La cantidad debe ser entera entre 1 y ' + MAX_QTY + '.'; return; }
    if (!confirm(`¿Confirmas ${qty} etiqueta(s) de ${comp} para la línea ${linea}?`)) return;

    try {
      // 1. Leer el registro central justo antes de reservar.
      await reloadWorkbook();
      const cons = normCons(sheetRows(SHEETS.cons));
      const row = getCounter(cons, op, comp, linea);

      // 2. Continuar la numeración del contador (límite 999 y luego el siguiente prefijo).
      const alloc = IDS.allocate(row.Prefijo, row.Ultimo, qty);
      const ids = alloc.ids;
      const lote = `1/${qty}`;

      // 3. Reservar: cada etiqueta como Reservado y mover el consecutivo.
      const reg = sheetRows(SHEETS.reg);
      const fecha = today(), hora = nowHM();
      ids.forEach((id, i) => {
        reg.push({ Fecha: fecha, Hora: hora, Usuario: S.user.Usuario, Operacion: op, Componente: comp, Linea: linea, ID: id, Posicion: `${i + 1}/${qty}`, Lote: lote, Estado: 'Reservado', Impresora: S.user.Impresora });
      });
      row.Prefijo = alloc.prefix;
      row.Ultimo = alloc.last;
      setRows(SHEETS.cons, cons, CONS_HEADERS);
      setRows(SHEETS.reg, reg, REG_HEADERS);
      await saveWorkbook();

      // 4. Vista previa del lote.
      S.batch = { op, comp, linea, ids, qty, fecha, hora };
      renderPreview();
      $('preview-card').hidden = false;
      $('preview-card').scrollIntoView({ behavior: 'smooth' });
      updateRange();
    } catch (e) {
      $('gen-error').textContent = 'Error al reservar: ' + e.message + '. No se imprimió nada.';
    }
  }

  function labelHTML(d) {
    return `<div class="grid">
      <div class="l-logo"><img src="assets/qorvex-logo-light.png" alt="QORVEX"></div>
      <div class="l-op">${esc(d.op.toUpperCase())}</div>
      <div class="l-linea"><span>LÍNEA</span><b>${esc(d.linea)}</b></div>
      <div class="l-comp"><span>COMPONENTE</span><b>${esc(d.comp.toUpperCase())}</b></div>
      <div class="l-id">${esc(d.id)}</div>
      <div class="l-cons"><span>CONSECUTIVO</span><b>${d.pos} / ${d.total}</b></div>
      <div class="l-foot"><div class="fecha"><span>FECHA Y HORA</span><b>${esc(d.fecha)} ${esc(d.hora)}</b></div><svg class="barcode"></svg></div>
    </div>`;
  }

  function buildLabels(list) {
    return list.map((d) => {
      const el = document.createElement('div');
      el.className = 'label';
      el.innerHTML = labelHTML(d);
      return el;
    });
  }

  function paintBarcodes(root) {
    root.querySelectorAll('svg.barcode').forEach((svg) => {
      const id = svg.closest('.label').querySelector('.l-id').textContent;
      try { JsBarcode(svg, id, { format: 'CODE128', displayValue: false, margin: 0, height: 50, width: 2 }); } catch (e) { /* ID inválido: no se dibuja */ }
    });
  }

  function renderPreview() {
    const b = S.batch;
    const data = b.ids.map((id, i) => ({ op: b.op, comp: b.comp, linea: b.linea, id: id, pos: i + 1, total: b.qty, fecha: b.fecha, hora: b.hora }));
    const preview = $('preview');
    preview.innerHTML = '';
    buildLabels(data.slice(0, 6)).forEach((el) => preview.appendChild(el));
    if (data.length > 6) {
      const n = document.createElement('p');
      n.className = 'muted';
      n.textContent = `Se muestran 6 de ${data.length}. Se imprimirán todas.`;
      preview.parentNode.appendChild(n);
      const old = preview.parentNode.querySelector('p.more');
      if (old) old.remove();
      n.className = 'muted more';
    }
    paintBarcodes(preview);
    $('btn-print').disabled = false;
  }

  function setEstadoLote(ids, estado) {
    const reg = sheetRows(SHEETS.reg);
    const b = S.batch;
    const set = new Set(ids);
    reg.forEach((r) => {
      if (set.has(r.ID) && r.Componente === b.comp && r.Fecha === b.fecha && r.Hora === b.hora && r.Estado !== 'Anulado') {
        r.Estado = estado;
      }
    });
    setRows(SHEETS.reg, reg, REG_HEADERS);
  }

  async function printBatch() {
    const b = S.batch;
    if (!b) return;
    if (!S.user.Impresora) { goPrinter(); return; }
    try {
      await reloadWorkbook();
      setEstadoLote(b.ids, 'Enviado');
      await saveWorkbook();
    } catch (e) {
      toast('No se pudo registrar el envío: ' + e.message);
      return;
    }
    await runPrint(buildLabelsFor(b));
    const ok = confirm(`¿Se imprimieron correctamente las ${b.qty} etiquetas en ${S.user.Impresora}?\n\nSi cancelas, quedarán marcadas como "Enviado" para revisión.`);
    try {
      await reloadWorkbook();
      setEstadoLote(b.ids, ok ? 'Impreso' : 'Enviado');
      await saveWorkbook();
      toast(ok ? 'Impresión confirmada.' : 'Lote marcado como Enviado.');
    } catch (e) {
      toast('No se pudo actualizar el estado: ' + e.message);
    }
    S.batch = null;
    $('preview-card').hidden = true;
    updateRange();
  }

  function buildLabelsFor(b) {
    return b.ids.map((id, i) => ({ op: b.op, comp: b.comp, linea: b.linea, id: id, pos: i + 1, total: b.qty, fecha: b.fecha, hora: b.hora }));
  }

  // window.print() bloquea hasta cerrar el diálogo; no depende de 'afterprint'.
  function runPrint(list) {
    return new Promise((resolve) => {
      const area = $('print-area');
      area.innerHTML = '';
      buildLabels(list).forEach((el) => area.appendChild(el));
      paintBarcodes(area);
      setTimeout(() => {
        S.printing = true;
        try { window.print(); } finally {
          S.printing = false;
          S.lastActivity = Date.now();
          resolve();
        }
      }, 150);
    });
  }

  async function cancelBatch() {
    const b = S.batch;
    if (!b) return;
    if (!confirm('¿Cancelar el lote? Los IDs ya reservados quedarán anulados y no se volverán a usar.')) return;
    try {
      await reloadWorkbook();
      setEstadoLote(b.ids, 'Anulado');
      await saveWorkbook();
      toast('Lote anulado. Los IDs no se reutilizarán.');
    } catch (e) { toast('Error al anular: ' + e.message); return; }
    S.batch = null;
    $('preview-card').hidden = true;
    updateRange();
  }

  // ---------- Historial ----------
  function renderHistory() {
    const q = $('hist-filter').value.trim().toLowerCase();
    const rows = sheetRows(SHEETS.reg).slice().reverse().filter((r) => {
      if (!q) return true;
      return [r.ID, r.Componente, r.Usuario].some((v) => String(v).toLowerCase().includes(q));
    }).slice(0, 200);
    const tb = document.querySelector('#hist-table tbody');
    tb.innerHTML = rows.map((r, i) => `<tr>
      <td>${esc(r.Fecha)}</td><td>${esc(r.Hora)}</td><td>${esc(r.Usuario)}</td><td>${esc(r.Operacion)}</td>
      <td>${esc(r.Componente)}</td><td>${esc(r.Linea)}</td><td><b>${esc(r.ID)}</b></td><td>${esc(r.Posicion)}</td>
      <td>${esc(r.Estado)}</td>
      <td><button class="btn small" data-reprint="${esc(r.ID)}" data-idx="${i}">Reimprimir</button>
          <button class="btn small" data-annul="${esc(r.ID)}">Anular</button></td>
    </tr>`).join('') || '<tr><td colspan="10" class="muted">Sin registros.</td></tr>';
  }

  async function reprint(id) {
    await reloadWorkbook();
    const reg = sheetRows(SHEETS.reg);
    const orig = reg.find((r) => r.ID === id && r.Estado !== 'Anulado');
    if (!orig) { toast('ID no encontrado o anulado.'); return; }
    if (!confirm(`¿Reimprimir ${id}? Conservará su ID original y quedará registrado como Reimpreso.`)) return;
    const reimp = { ...orig, Fecha: today(), Hora: nowHM(), Usuario: S.user.Usuario, Estado: 'Reimpreso', Impresora: S.user.Impresora };
    reg.push(reimp);
    setRows(SHEETS.reg, reg, REG_HEADERS);
    await saveWorkbook();
    const [pos, total] = String(orig.Posicion).split('/').map(Number);
    await runPrint([{ op: orig.Operacion, comp: orig.Componente, linea: orig.Linea, id: id, pos: pos || 1, total: total || 1, fecha: orig.Fecha, hora: orig.Hora }]);
    renderHistory();
  }

  async function annul(id) {
    if (!S.adminUnlocked && !(await askAdmin())) return;
    if (!confirm(`¿Anular ${id}? Quedará registrado y el ID no se reutilizará.`)) return;
    await reloadWorkbook();
    const reg = sheetRows(SHEETS.reg);
    reg.forEach((r) => { if (r.ID === id) r.Estado = 'Anulado'; });
    setRows(SHEETS.reg, reg, REG_HEADERS);
    await saveWorkbook();
    toast(id + ' anulado.');
    renderHistory();
  }

  // ---------- Administración ----------
  async function askAdmin() {
    const pass = prompt('Contraseña de administrador:');
    if (pass === null) return false;
    if (String(cfgValue('ADMIN_PASS', '')) === pass) { S.adminUnlocked = true; return true; }
    toast('Contraseña incorrecta.');
    return false;
  }

  function renderConfig() {
    $('cfg-comp-table').querySelector('tbody').innerHTML = sheetRows(SHEETS.comp).map((c, i) =>
      `<tr><td>${esc(c.Nombre)}</td><td>${esc(c.Prefijo)}</td><td><button class="btn small" data-toggle-comp="${i}">${esc(c.Activo)}</button></td></tr>`).join('');
    $('cfg-line-table').querySelector('tbody').innerHTML = sheetRows(SHEETS.line).map((l, i) =>
      `<tr><td>${esc(l.Linea)}</td><td><button class="btn small" data-toggle-line="${i}">${esc(l.Activa)}</button></td></tr>`).join('');
    $('cfg-user-table').querySelector('tbody').innerHTML = sheetRows(SHEETS.users).map((u) =>
      `<tr><td>${esc(u.Usuario)}</td><td>${esc(u.Nombre)}</td><td>••••••</td><td>${esc(u.Impresora || '—')}</td></tr>`).join('');
    $('cfg-timeout').value = cfgValue('MINUTOS_INACTIVIDAD', 5);
  }

  async function adminWrite(mutator) {
    try {
      await reloadWorkbook();
      mutator();
      await saveWorkbook();
    } catch (e) {
      toast(e.message || 'No se pudo guardar el cambio.');
      return false;
    }
    renderConfig();
    fillSelectors();
    updateRange();
    return true;
  }

  function bindEvents() {
    $('btn-open').onclick = openWorkbook;
    $('btn-last').onclick = openLastFolder;

    $('form-login').onsubmit = (e) => {
      e.preventDefault();
      const u = login($('login-user').value.trim(), $('login-pass').value);
      if (!u) { $('login-error').textContent = 'Usuario o contraseña incorrectos.'; return; }
      startSession(u);
    };

    $('form-printer').onsubmit = async (e) => {
      e.preventDefault();
      const name = $('printer-name').value.trim();
      if (!IDS.validSecret(name)) { toast('Escribe el nombre de la impresora.'); return; }
      await savePrinter(name);
      enterMain();
    };
    $('printer-cancel').onclick = () => enterMain();
    $('btn-printer').onclick = () => goPrinter();
    $('btn-logout').onclick = () => logout('Sesión cerrada.');

    document.querySelectorAll('.nav-item').forEach((b) => (b.onclick = () => showTab(b.dataset.tab)));
    $('btn-theme').onclick = toggleTheme;
    $('btn-theme-open').onclick = toggleTheme;

    ['gen-qty', 'gen-comp'].forEach((id) => ($(id).oninput = updateRange));
    $('gen-comp').onchange = updateRange;
    $('gen-op').onchange = updateRange;
    $('gen-line').onchange = updateRange;
    $('btn-preview-range').onclick = updateRange;
    $('btn-generate').onclick = generate;
    $('btn-print').onclick = printBatch;
    $('btn-cancel-batch').onclick = cancelBatch;

    $('hist-filter').oninput = renderHistory;
    document.querySelector('#hist-table').onclick = (e) => {
      const rp = e.target.dataset.reprint, an = e.target.dataset.annul;
      if (rp) reprint(rp);
      if (an) annul(an);
    };

    $('form-admin').onsubmit = (e) => {
      e.preventDefault();
      if (String(cfgValue('ADMIN_PASS', '')) === $('admin-pass').value) {
        S.adminUnlocked = true;
        $('admin-pass').value = '';
        $('admin-error').textContent = '';
        showTab('config');
      } else {
        $('admin-error').textContent = 'Contraseña incorrecta.';
      }
    };

    $('form-comp').onsubmit = async (e) => {
      e.preventDefault();
      const nombre = $('cfg-comp-name').value.trim();
      const prefijo = $('cfg-comp-prefix').value.trim().toUpperCase();
      if (!IDS.validPrefix(prefijo)) { toast('El prefijo debe ser de 1 a 4 letras.'); return; }
      await adminWrite(() => {
        const comp = sheetRows(SHEETS.comp);
        if (comp.some((c) => c.Nombre === nombre || c.Prefijo === prefijo)) throw new Error('Nombre o prefijo repetido.');
        comp.push({ Nombre: nombre, Prefijo: prefijo, Activo: 'SI' });
        setRows(SHEETS.comp, comp, ['Nombre', 'Prefijo', 'Activo']);
      });
      e.target.reset();
      toast('Componente agregado.');
    };

    $('form-line').onsubmit = async (e) => {
      e.preventDefault();
      const l = $('cfg-line-id').value.trim();
      await adminWrite(() => {
        const lines = sheetRows(SHEETS.line);
        if (lines.some((x) => String(x.Linea) === l)) throw new Error('Línea repetida.');
        lines.push({ Linea: l, Activa: 'SI' });
        setRows(SHEETS.line, lines, ['Linea', 'Activa']);
      });
      e.target.reset();
      toast('Línea agregada.');
    };

    $('form-user').onsubmit = async (e) => {
      e.preventDefault();
      const u = $('cfg-user-id').value.trim(), n = $('cfg-user-name').value.trim(), p = $('cfg-user-pass').value;
      if (!IDS.validSecret(p)) { toast('La contraseña no puede empezar ni terminar con espacios.'); return; }
      await adminWrite(() => {
        const users = sheetRows(SHEETS.users);
        if (users.some((x) => x.Usuario === u)) throw new Error('Usuario repetido.');
        users.push({ Usuario: u, Nombre: n, Contrasena: p, Impresora: '' });
        setRows(SHEETS.users, users, ['Usuario', 'Nombre', 'Contrasena', 'Impresora']);
      });
      e.target.reset();
      toast('Usuario agregado.');
    };

    $('form-params').onsubmit = async (e) => {
      e.preventDefault();
      const min = Number($('cfg-timeout').value);
      if (!Number.isInteger(min) || min < 1 || min > 60) { toast('Minutos entre 1 y 60.'); return; }
      await adminWrite(() => {
        const cfg = sheetRows(SHEETS.cfg);
        const r = cfg.find((x) => x.Clave === 'MINUTOS_INACTIVIDAD');
        if (r) r.Valor = min; else cfg.push({ Clave: 'MINUTOS_INACTIVIDAD', Valor: min });
        setRows(SHEETS.cfg, cfg, ['Clave', 'Valor']);
      });
      toast('Parámetros guardados.');
    };

    $('form-admin-pass').onsubmit = async (e) => {
      e.preventDefault();
      const nueva = $('cfg-admin-new').value;
      if (!IDS.validSecret(nueva)) { toast('Contraseña no válida.'); return; }
      await adminWrite(() => {
        const cfg = sheetRows(SHEETS.cfg);
        const r = cfg.find((x) => x.Clave === 'ADMIN_PASS');
        if (r) r.Valor = nueva; else cfg.push({ Clave: 'ADMIN_PASS', Valor: nueva });
        setRows(SHEETS.cfg, cfg, ['Clave', 'Valor']);
      });
      e.target.reset();
      toast('Contraseña de administrador actualizada.');
    };

    // Activar / desactivar componentes y líneas
    document.body.addEventListener('click', async (e) => {
      const tc = e.target.dataset.toggleComp, tl = e.target.dataset.toggleLine;
      if (tc !== undefined) {
        await adminWrite(() => {
          const comp = sheetRows(SHEETS.comp);
          comp[Number(tc)].Activo = comp[Number(tc)].Activo === 'SI' ? 'NO' : 'SI';
          setRows(SHEETS.comp, comp, ['Nombre', 'Prefijo', 'Activo']);
        });
      }
      if (tl !== undefined) {
        await adminWrite(() => {
          const lines = sheetRows(SHEETS.line);
          lines[Number(tl)].Activa = lines[Number(tl)].Activa === 'SI' ? 'NO' : 'SI';
          setRows(SHEETS.line, lines, ['Linea', 'Activa']);
        });
      }
    });

    // Al cambiar contraseña de administrador, se bloquea de nuevo.
    $('gen-qty').addEventListener('change', updateRange);
  }

  function init() {
    initTheme();
    bindEvents();
    show('view-open');
    showLastFolderButton();
    if (!window.showDirectoryPicker) {
      $('open-error').textContent = 'Usa Microsoft Edge o Google Chrome. Este navegador no permite seleccionar la carpeta.';
    }
  }

  init();
})();
