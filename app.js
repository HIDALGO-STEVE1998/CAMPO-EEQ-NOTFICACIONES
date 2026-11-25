'use strict';
/* ===== Configuración ===== */
const SLOTS_LEC = [
  { id: 'antes', label: 'ANTES', req: true },
  { id: 'despues', label: 'DESPUÉS', req: true },
  { id: 'adic1', label: 'ADICIONAL', req: false },
];
const SLOTS_CORTE = [
  { id: 'antes', label: 'ANTES', req: true },
  { id: 'despues', label: 'DESPUÉS', req: true },
  { id: 'corr1', label: 'CORRECCIÓN 1', req: true },
  { id: 'corr2', label: 'CORRECCIÓN 2', req: true },
  { id: 'adic1', label: 'ADICIONAL 1', req: true },
  { id: 'adic2', label: 'ADICIONAL 2', req: true },
];
const MODS = {
  lect: { name: 'Lecturas', slots: SLOTS_LEC },
  notif: { name: 'Notificaciones', slots: SLOTS_LEC },
  corte: { name: 'Cortes y reconexiones', slots: SLOTS_CORTE },
};
const MAX_PX = 1280, JPG_Q = 0.8, ZIP_MAX = 90 * 1048576;
// Candado de cortesía: la clave NO está escrita en el código, solo su huella (hash).
const AUTH_HASH = '776edea2335468632fbdf3b3e6fafacc6f1d2cfeaac6b5a6ec89ca7f0f306207';
const $ = (s) => document.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
let MOD = null;                                   // módulo abierto: 'lect' | 'notif' | 'corte'
const SL = () => MODS[MOD].slots;
const kOf = (mod, orden) => mod + '|' + orden;

/* ===== IndexedDB ===== */
let db;
function openDB() {
  return new Promise((res, rej) => {
    const r = indexedDB.open('campo-eeq', 1);
    r.onupgradeneeded = () => {
      const d = r.result;
      d.createObjectStore('orders', { keyPath: 'k' });
      d.createObjectStore('works', { keyPath: 'k' });
      d.createObjectStore('photos', { keyPath: 'key' });
      d.createObjectStore('samples', { keyPath: 'key' });
      d.createObjectStore('kv');
    };
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
const rq = (r) => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
const dbGet = (s, k) => rq(db.transaction(s).objectStore(s).get(k));
const dbAll = (s) => rq(db.transaction(s).objectStore(s).getAll());
const dbPut = (s, v, k) => rq(db.transaction(s, 'readwrite').objectStore(s).put(v, k));
const dbDel = (s, k) => rq(db.transaction(s, 'readwrite').objectStore(s).delete(k));
const dbByPrefix = (s, p) => rq(db.transaction(s).objectStore(s).getAll(IDBKeyRange.bound(p, p + '￿')));
const dbPhotosOf = (k) => dbByPrefix('photos', k + '|');
const dbKeysByPrefix = (s, p) => rq(db.transaction(s).objectStore(s).getAllKeys(IDBKeyRange.bound(p, p + '￿')));

/* ===== Lectura de CSV / Excel ===== */
function detectDelim(line) { const c = [',', ';', '\t'].map((d) => [d, line.split(d).length]); c.sort((a, b) => b[1] - a[1]); return c[0][0]; }
function parseCSV(text) {
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  const delim = detectDelim(text.slice(0, text.indexOf('\n') > 0 ? text.indexOf('\n') : text.length));
  const rows = []; let row = [], f = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
    else if (c === '"') q = true;
    else if (c === delim) { row.push(f); f = ''; }
    else if (c === '\n') { row.push(f); rows.push(row); row = []; f = ''; }
    else if (c !== '\r') f += c;
  }
  if (f !== '' || row.length) { row.push(f); rows.push(row); }
  return rows;
}
async function parseXLSX(buf) {
  const z = await JSZip.loadAsync(buf); const dp = new DOMParser();
  const xml = async (n) => { const f = z.file(n); return f ? dp.parseFromString(await f.async('string'), 'application/xml') : null; };
  const ss = []; const sx = await xml('xl/sharedStrings.xml');
  if (sx) sx.querySelectorAll('si').forEach((si) => ss.push([...si.querySelectorAll('t')].map((t) => t.textContent).join('')));
  let path = 'xl/worksheets/sheet1.xml';
  const wb = await xml('xl/workbook.xml'), rels = await xml('xl/_rels/workbook.xml.rels');
  if (wb && rels) { const sh = wb.querySelector('sheet'); const rid = sh && (sh.getAttribute('r:id') || sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')); const rel = rid && [...rels.querySelectorAll('Relationship')].find((r) => r.getAttribute('Id') === rid); if (rel) { let t = rel.getAttribute('Target'); path = t.startsWith('/') ? t.slice(1) : 'xl/' + t; } }
  const sd = await xml(path); if (!sd) throw new Error('No se encontró la hoja del Excel.');
  const col = (ref) => { let n = 0; for (const ch of ref.replace(/\d/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  const rows = [];
  sd.querySelectorAll('sheetData > row').forEach((r) => {
    const arr = [];
    r.querySelectorAll('c').forEach((c) => {
      const i = col(c.getAttribute('r')), t = c.getAttribute('t'); const v = c.querySelector('v');
      let val = '';
      if (t === 's') val = v ? ss[parseInt(v.textContent, 10)] ?? '' : '';
      else if (t === 'inlineStr') val = [...c.querySelectorAll('t')].map((x) => x.textContent).join('');
      else val = v ? v.textContent : '';
      arr[i] = val;
    });
    for (let i = 0; i < arr.length; i++) if (arr[i] === undefined) arr[i] = '';
    rows.push(arr);
  });
  return rows;
}
const norm = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
function rowsToOrders(rows, mod) {
  const h = rows[0].map(norm); const ix = (...ns) => { for (const n of ns) { const i = h.indexOf(norm(n)); if (i >= 0) return i; } return -1; };
  const iO = ix('Orden'); if (iO < 0) throw new Error('El archivo no tiene la columna "Orden". Columnas encontradas: ' + rows[0].slice(0, 8).join(', ') + '…');
  const g = (r, i) => (i >= 0 && r[i] != null ? String(r[i]).trim() : '');
  const num = (s) => { const n = parseFloat(String(s).replace(',', '.')); return isFinite(n) ? n : null; };
  const I = { med: ix('Medidor'), cta: ix('Cuenta'), cli: ix('Nombre cliente'), tec: ix('Nombre tecnico'), la: ix('Latitud orden'), lo: ix('Longitud orden'), act: ix('Actividades'), fo: ix('Fotos'), mru: ix('MRU'), fe: ix('Fecha de ejecucion'), dir: ix('Direccion') };
  const out = [];
  for (const r of rows.slice(1)) {
    const orden = g(r, iO); if (!orden) continue;
    out.push({
      k: kOf(mod, orden), mod, orden, medidor: g(r, I.med), cuenta: g(r, I.cta), cliente: g(r, I.cli), tecnico: g(r, I.tec), mru: g(r, I.mru),
      lat: num(g(r, I.la)), lon: num(g(r, I.lo)), actividades: g(r, I.act), fechaEj: g(r, I.fe), direccion: g(r, I.dir),
      fotos: g(r, I.fo).split(',').map((s) => s.trim()).filter((s) => /^https?:\/\//.test(s)).sort((x, y) => nombreFoto(x).localeCompare(nombreFoto(y))),
    });
  }
  return out;
}

/* ===== Utilidades ===== */
const nombreFoto = (u) => decodeURIComponent(u.split('/').pop().replace(/\.[a-z]+$/i, ''));
const hoy = () => { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
const hora = (ts) => { const d = new Date(ts); return String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'); };
const slotLabel = (id) => (SL().find((s) => s.id === id) || {}).label || id;
function modal(title, html, btns) {
  return new Promise((res) => {
    $('#mTitle').textContent = title; $('#mBody').innerHTML = html; const b = $('#mBtns'); b.innerHTML = '';
    btns.forEach(([txt, val, cls]) => { const x = document.createElement('button'); x.className = 'btn ' + (cls || ''); x.textContent = txt; x.onclick = () => { $('#modal').hidden = true; res(val); }; b.appendChild(x); });
    $('#modal').hidden = false;
  });
}
const aviso = (t, lines) => modal(t, Array.isArray(lines) ? '<ul>' + lines.map((l) => '<li>' + esc(l) + '</li>').join('') + '</ul>' : esc(lines), [['Entendido', true, 'primary']]);
let busyT1, busyT2;
function busy(on, txt) {
  clearTimeout(busyT1); clearTimeout(busyT2);
  $('#busy').hidden = !on; $('#busyX').hidden = true; $('#busyTxt').textContent = txt || 'Procesando…';
  if (on) {
    busyT1 = setTimeout(() => { $('#busyX').hidden = false; }, 6000);
    busyT2 = setTimeout(() => { busy(false); aviso('Tardó demasiado', 'El mensaje "Procesando" se cerró solo. Revisa si la foto quedó guardada en su recuadro; si no, vuelve a tomarla.'); }, 40000);
  }
}
$('#busyX').onclick = () => busy(false);
function download(blob, name) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 30000);
}
const sha256hex = async (s) => [...new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s)))].map((b) => b.toString(16).padStart(2, '0')).join('');
function distM(a, b, c, d) { const R = 6371000, r = Math.PI / 180, x = (c - a) * r, y = (d - b) * r * Math.cos(((a + c) / 2) * r); return Math.sqrt(x * x + y * y) * R; }

/* ===== GPS ===== */
let gps = null, gpsWatch = null, gpsErr = null; const gpsUse = new Set();
function updGps() {
  const c = $('#gpsChip'); c.hidden = false;
  const cg = $('#camGps'); if (cg) cg.textContent = gps ? 'GPS ±' + Math.round(gps.acc) + ' m' : (gpsErr === 1 ? 'GPS sin permiso' : 'Buscando GPS…');
  if (gpsErr === 1) { c.className = 'chip bad'; c.textContent = 'GPS sin permiso'; return; }
  if (!gps) { c.className = 'chip warn'; c.textContent = 'Buscando GPS…'; return; }
  c.className = 'chip ' + (gps.acc <= 25 ? 'ok' : 'warn'); c.textContent = 'GPS ±' + Math.round(gps.acc) + ' m';
}
function startGPS(who = 'orden') {
  gpsUse.add(who);
  if (!navigator.geolocation || gpsWatch != null) return; updGps();
  gpsWatch = navigator.geolocation.watchPosition(
    (p) => { gpsErr = null; gps = { lat: p.coords.latitude, lon: p.coords.longitude, acc: p.coords.accuracy, ts: Date.now() }; updGps(); if (typeof updMapMe === 'function') updMapMe(); },
    (e) => { gpsErr = e.code; updGps(); }, { enableHighAccuracy: true, maximumAge: 0, timeout: 20000 });
}
function stopGPS(who = 'orden') { gpsUse.delete(who); if (gpsUse.size) return; if (gpsWatch != null) navigator.geolocation.clearWatch(gpsWatch); gpsWatch = null; gps = null; $('#gpsChip').hidden = true; }
function getFix() {
  return new Promise((res) => {
    let done = false; const fin = (v) => { if (!done) { done = true; res(v); } };
    if (gps && Date.now() - gps.ts < 60000) return fin(gps);          // posición reciente: no esperar
    if (!navigator.geolocation) return fin(gps && Date.now() - gps.ts < 300000 ? gps : null);
    const viejo = () => fin(gps && Date.now() - gps.ts < 300000 ? gps : null);
    setTimeout(viejo, 6000);                                            // tope duro: nunca más de 6 s
    navigator.geolocation.getCurrentPosition(
      (p) => { gps = { lat: p.coords.latitude, lon: p.coords.longitude, acc: p.coords.accuracy, ts: Date.now() }; updGps(); fin(gps); },
      viejo, { enableHighAccuracy: true, timeout: 5000, maximumAge: 20000 });
  });
}
const fmtCoord = (f) => (f ? f.lat.toFixed(6) + ', ' + f.lon.toFixed(6) : 'SIN GPS');


/* ===== Coordenadas desde la foto de galería (EXIF) ===== */
function parseTiffGps(v, t) {
  const le = v.getUint16(t) === 0x4949;
  const u16 = (p) => v.getUint16(p, le), u32 = (p) => v.getUint32(p, le);
  if (u16(t + 2) !== 42) return null;
  const ifd0 = t + u32(t + 4); let gpsOff = null;
  for (let i = 0, n = u16(ifd0); i < n; i++) { const e = ifd0 + 2 + i * 12; if (u16(e) === 0x8825) gpsOff = t + u32(e + 8); }
  if (gpsOff == null) return null;
  const g = {};
  for (let i = 0, n = u16(gpsOff); i < n; i++) {
    const e = gpsOff + 2 + i * 12, tag = u16(e);
    if (tag === 1 || tag === 3) g[tag] = String.fromCharCode(v.getUint8(e + 8));
    if (tag === 2 || tag === 4) {
      const o = t + u32(e + 8); const r = (k) => u32(o + k * 8) / (u32(o + k * 8 + 4) || 1);
      g[tag] = r(0) + r(1) / 60 + r(2) / 3600;
    }
  }
  if (g[2] == null || g[4] == null) return null;
  const lat = g[1] === 'S' ? -g[2] : g[2], lon = g[3] === 'W' ? -g[4] : g[4];
  if (!isFinite(lat) || !isFinite(lon) || (lat === 0 && lon === 0) || Math.abs(lat) > 90 || Math.abs(lon) > 180) return null;
  return { lat, lon, acc: null, ts: Date.now() };
}
async function readExifGps(file) {
  try {
    const v = new DataView(await file.slice(0, 262144).arrayBuffer());
    if (v.getUint16(0) !== 0xFFD8) return null;
    let o = 2;
    while (o + 4 < v.byteLength) {
      const m = v.getUint16(o), len = v.getUint16(o + 2);
      if (m === 0xFFE1 && v.getUint32(o + 4) === 0x45786966) return parseTiffGps(v, o + 10);
      if ((m & 0xFF00) !== 0xFF00 || m === 0xFFDA) break;
      o += 2 + len;
    }
  } catch (_) { /* sin EXIF */ }
  return null;
}

async function loadBitmap(blob) {
  try { return await createImageBitmap(blob, { imageOrientation: 'from-image' }); }
  catch (e) {
    return await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = URL.createObjectURL(blob); });
  }
}

/* ===== Imagen + marca de agua ===== */
async function stampPhoto(file, rotulo, o, fix) {
  const bmp = (file instanceof HTMLCanvasElement) ? file : await loadBitmap(file);
  const sw = bmp.width, sh = bmp.height, k = Math.min(1, MAX_PX / Math.max(sw, sh));
  const w = Math.round(sw * k), h = Math.round(sh * k);
  const c = document.createElement('canvas'); c.width = w; c.height = h;
  const x = c.getContext('2d'); x.drawImage(bmp, 0, 0, w, h);
  const d = new Date(); const fh = d.toLocaleDateString('es-EC') + ' ' + hora(d.getTime());
  const lines = [rotulo + ' · ' + MODS[o.mod].name, 'Número de orden: ' + o.orden, 'Número de medidor: ' + o.medidor, 'Coordenadas: ' + fmtCoord(fix) + (fix && fix.note ? ' (' + fix.note + ')' : '')];
  if (o.cuenta) lines.push('Cuenta contrato: ' + o.cuenta);
  lines.push('Fecha: ' + fh);
  const fs = Math.max(14, Math.round(Math.min(w, h) / 24)), pad = Math.round(fs * 0.5), lh = Math.round(fs * 1.3);
  const bh = lines.length * lh + pad * 2;
  x.fillStyle = 'rgba(0,0,0,0.62)'; x.fillRect(0, h - bh, w, bh);
  x.textBaseline = 'top';
  lines.forEach((t, i) => {
    x.font = (i === 0 ? 'bold ' : '') + fs + 'px sans-serif';
    x.fillStyle = i === 0 ? '#ffd54a' : '#ffffff';
    x.fillText(t, pad, h - bh + pad + i * lh, w - pad * 2);
  });
  if (bmp.close) bmp.close();
  const out = await new Promise((res) => c.toBlob(res, 'image/jpeg', JPG_Q));
  if (!out) throw new Error('El teléfono no pudo crear la imagen (poca memoria). Cierra otras apps e inténtalo de nuevo.');
  return out;
}

/* ===== Estado ===== */
const DATA = { lect: { orders: [], works: {} }, notif: { orders: [], works: {} }, corte: { orders: [], works: {} } };
let filtro = 'todas', cur = null; // cur = {o, w}
const ORD = () => DATA[MOD].orders, WKS = () => DATA[MOD].works;
const orderMap = () => Object.fromEntries(ORD().map((o) => [o.orden, o]));
async function loadAll() {
  for (const m of Object.keys(DATA)) { DATA[m].orders = []; DATA[m].works = {}; }
  (await dbAll('orders')).forEach((o) => DATA[o.mod]?.orders.push(o));
  (await dbAll('works')).forEach((w) => { if (DATA[w.mod]) DATA[w.mod].works[w.orden] = w; });
  for (const m of Object.keys(DATA)) DATA[m].orders.sort((a, b) => Number(a.orden) - Number(b.orden) || String(a.orden).localeCompare(String(b.orden)));
}
const saveWork = (w) => { DATA[w.mod].works[w.orden] = w; return dbPut('works', w); };
const nGrid = (o) => o.fotos.length;
function faltantes(o, w) {
  const m = [];
  MODS[o.mod].slots.filter((s) => s.req && !w.fotos[s.id]).forEach((s) => m.push('Falta la foto: ' + s.label));
  Object.entries(w.fotos).forEach(([id, f]) => { if (!f.ok) m.push('Foto sin coordenadas, vuelve a tomarla: ' + slotLabel(id)); });
  for (let i = 0; i < nGrid(o); i++) if (!w.conf?.[i]) m.push('Falta confirmar que revisaste la foto ' + (i + 1) + ' del Grid');
  return m;
}
const estOf = (o) => { const w = DATA[o.mod].works[o.orden]; return w ? (w.estado === 'completa' ? 'ok' : 'curso') : 'pend'; };
const cuentasDe = (mod) => { const c = { pend: 0, curso: 0, ok: 0 }; DATA[mod].orders.forEach((o) => c[estOf(o)]++); return c; };

/* ===== Menú de módulos ===== */
function renderMenu() {
  const t = { lect: '#cLect', notif: '#cNotif', corte: '#cCorte' }; let sinDesc = 0;
  for (const m of Object.keys(t)) {
    const c = cuentasDe(m), n = DATA[m].orders.length;
    $(t[m]).textContent = n ? `${n} órdenes · Atendidas ${c.ok} · En curso ${c.curso} · Pendientes ${c.pend}` : 'Sin órdenes cargadas';
    sinDesc += Object.values(DATA[m].works).filter((w) => w.estado === 'completa' && !w.descargada).length;
  }
  $('#riesgoTodo').innerHTML = sinDesc ? `<div class="alertBox"><b>${sinDesc} orden(es) completadas aún sin descargar.</b> Descarga el reporte al terminar la jornada.</div>` : '<p class="muted">Todo lo completado ya fue descargado.</p>';
  storInfo();
}
async function storInfo() {
  try {
    if (navigator.storage?.estimate) {
      const e = await navigator.storage.estimate(); const libre = (e.quota - e.usage) / 1048576;
      $('#storInfo').textContent = `Usado: ${(e.usage / 1048576).toFixed(0)} MB · Libre: ${libre.toFixed(0)} MB` + (libre < 300 ? ' ⚠ POCO ESPACIO: descarga y libera.' : '');
    }
  } catch (_) { /* nada */ }
}
function showMenu() {
  MOD = null; tabActual = 'ordenes'; stopGPS('mapa');
  $('#viewMenu').hidden = false; $('#viewHome').hidden = true; $('#viewOrden').hidden = true; $('#brandSub').textContent = 'Fiscalización · Quito';
  renderMenu(); window.scrollTo(0, 0);
}
function enterMod(m, push = true) {
  MOD = m; filtro = 'todas'; mapaCentrado = false; tabActual = 'ordenes';
  $('#viewMenu').hidden = true; $('#viewHome').hidden = false; $('#viewOrden').hidden = true;
  $('#modName').textContent = MODS[m].name; $('#brandSub').textContent = MODS[m].name;
  document.querySelectorAll('.tabs button').forEach((x) => x.classList.toggle('on', x.dataset.tab === 'ordenes'));
  document.querySelectorAll('.filters button').forEach((x) => x.classList.toggle('on', x.dataset.f === 'todas'));
  $('#tabOrdenes').hidden = false; $('#tabMapa').hidden = true; $('#tabReporte').hidden = true; $('#q').value = '';
  if (MAP) { ptLayer.clearLayers(); }
  if (push) history.pushState({ mod: 1 }, '');
  renderList(); renderReporte(); window.scrollTo(0, 0);
}
document.querySelectorAll('.modBtn').forEach((b) => (b.onclick = () => enterMod(b.dataset.mod)));
$('#btnMenu').onclick = () => { if (history.state?.mod) history.back(); else showMenu(); };

/* ===== Vista: lista ===== */
function renderCsvBox() {
  const n = ORD().length, c = cuentasDe(MOD);
  $('#csvBox').innerHTML = n
    ? `<b>${n} órdenes cargadas.</b> <span class="muted">Atendidas ${c.ok} · En curso ${c.curso} · Pendientes ${c.pend}</span> <button class="btn small" id="btnCsv">Cargar otro CSV</button>`
    : `<b>Aún no hay órdenes.</b><p class="muted">Carga el CSV (o Excel) con las órdenes para empezar.</p><button class="btn primary" id="btnCsv">Cargar CSV</button>`;
  $('#btnCsv').onclick = () => $('#csvIn').click();
  $('#listBox').hidden = !n;
}
function renderChips() {
  const c = cuentasDe(MOD); const t = { todas: `Todas (${c.pend + c.curso})`, pend: `Pendientes (${c.pend})`, curso: `En curso (${c.curso})`, ok: `Atendidas (${c.ok})` };
  document.querySelectorAll('.filters button').forEach((b) => { b.textContent = t[b.dataset.f]; });
}
function renderList() {
  if (!MOD) return;
  const q = $('#q').value.trim().toLowerCase(); renderChips(); renderCsvBox();
  const items = ORD().filter((o) => {
    const e = estOf(o);
    if (q) return (o.orden + ' ' + o.medidor + ' ' + o.cuenta).toLowerCase().includes(q);
    return filtro === 'todas' ? e !== 'ok' : e === filtro;
  });
  const lbl = { ok: 'Atendida', curso: 'En curso', pend: 'Pendiente' };
  $('#list').innerHTML = (q ? '<p class="muted small">Resultado de buscar en todas las órdenes (incluye atendidas).</p>' : '') + items.slice(0, 150).map((o) => {
    const e = estOf(o); const w = WKS()[o.orden];
    const extra = e === 'ok' && w?.fecha ? ' · ' + w.fecha.split('-').reverse().join('/') : '';
    return `<div class="item ${e === 'ok' ? 'atend' : ''}" data-o="${esc(o.orden)}"><div><b>Orden ${esc(o.orden)}</b><div class="muted">Medidor ${esc(o.medidor)}${extra}</div></div><span class="badge ${e}">${lbl[e]}</span></div>`;
  }).join('') + (items.length > 150 ? `<p class="muted">Mostrando 150 de ${items.length}. Usa el buscador.</p>` : '') + (items.length ? '' : '<p class="muted">Sin resultados.</p>');
  $('#list').querySelectorAll('.item').forEach((el) => (el.onclick = () => openOrder(el.dataset.o)));
}

/* ===== Vista: orden ===== */
let urlsVivas = [];
async function openOrder(orden) {
  const o = orderMap()[orden]; if (!o) return; const W = WKS();
  if (W[orden]?.estado === 'completa') {
    const f = W[orden].fecha ? ' el ' + W[orden].fecha.split('-').reverse().join('/') : '';
    const r = await modal('Esta orden ya está atendida', 'La orden ' + esc(o.orden) + ' (medidor ' + esc(o.medidor) + ') se completó' + f + '. ¿Quieres abrirla para corregirla? Tendrás que volver a guardarla.', [['No', false], ['Sí, abrir', true, 'primary']]);
    if (!r) return;
  }
  let w = W[orden] || { k: o.k, mod: o.mod, orden, estado: 'en_curso', ejecutor: o.tecnico || '', fotos: {}, conf: [], obs: '', fecha: null, descargada: false, abierta: Date.now() };
  if (o.tecnico) w.ejecutor = o.tecnico; if (!w.conf) w.conf = []; if (w.obs == null) w.obs = '';
  const ph = await dbPhotosOf(o.k);
  ph.forEach((p) => { if (!w.fotos[p.slot]) w.fotos[p.slot] = { ts: p.ts, lat: p.lat, lon: p.lon, ok: p.ok }; });
  if (w.estado === 'completa') { w = { ...w, estado: 'en_curso' }; }
  await saveWork(w); await dbPut('kv', o.k, 'activa');
  cur = { o, w };
  $('#viewHome').hidden = true; $('#viewOrden').hidden = false; window.scrollTo(0, 0);
  history.pushState({ orden: 1 }, '');
  stopGPS('mapa'); startGPS('orden'); renderOrder(); await renderMuestras(); await renderPhotos();
}
function renderOrder() {
  const { o, w } = cur;
  $('#oNum').textContent = o.orden; $('#oMed').textContent = o.medidor || '(sin medidor)';
  $('#oExtra').innerHTML = [o.cuenta && 'Cuenta contrato: <b>' + esc(o.cuenta) + '</b>', o.cliente && esc(o.cliente), o.actividades && 'Actividad: ' + esc(o.actividades), o.direccion && esc(o.direccion)].filter(Boolean).join('<br>');
  $('#ejecutorTxt').textContent = o.tecnico || '(el CSV no trae el nombre)';
  $('#coordTxt').innerHTML = o.lat != null ? `${o.lat.toFixed(6)}, ${o.lon.toFixed(6)} · <a href="https://www.google.com/maps?q=${o.lat},${o.lon}" target="_blank" rel="noopener">Ver en Maps</a>` : 'La orden no trae coordenadas';
  $('#obs').value = w.obs || '';
  $('#gridBox').hidden = !nGrid(o);
  $('#faltan').hidden = true;
}
$('#obs').oninput = () => { if (!cur) return; cur.w.obs = $('#obs').value; clearTimeout(tObs); const wk = cur.w; tObs = setTimeout(() => saveWork(wk), 400); };
let tObs;
async function renderMuestras() {
  urlsVivas.forEach((u) => URL.revokeObjectURL(u)); urlsVivas = [];
  const { o, w } = cur; const box = $('#muestras'); box.innerHTML = '';
  for (let i = 0; i < o.fotos.length; i++) {
    const nm = nombreFoto(o.fotos[i]); const s = await dbGet('samples', o.orden + '|' + nm); let src = '', pie = '';
    if (s) { src = URL.createObjectURL(s.blob); urlsVivas.push(src); pie = 'Guardada en el teléfono'; }
    else if (navigator.onLine) { src = o.fotos[i]; pie = 'Desde internet (no guardada)'; }
    const on = !!w.conf[i];
    const d = document.createElement('div'); d.className = 'mTile';
    d.innerHTML = `<div class="im">${src ? `<img src="${esc(src)}" alt="">` : 'Sin foto guardada.<br>Carga el ZIP de fotos del Grid o conéctate a internet.'}</div><div class="muted small">${i + 1}. ${esc(nm)} · ${pie}</div><button class="okbtn ${on ? 'on' : ''}" data-i="${i}">${on ? '✔ Revisada' : 'Confirmo que la revisé'}</button>`;
    const im = d.querySelector('img'); if (im) { im.onclick = () => { $('#viewerImg').src = src; $('#viewer').hidden = false; }; im.onerror = () => { im.replaceWith(document.createTextNode('No se pudo cargar')); }; }
    d.querySelector('button').onclick = (e) => { const j = +e.target.dataset.i; cur.w.conf[j] = !cur.w.conf[j]; e.target.classList.toggle('on', cur.w.conf[j]); e.target.textContent = cur.w.conf[j] ? '✔ Revisada' : 'Confirmo que la revisé'; saveWork(cur.w); };
    box.appendChild(d);
  }
}
$('#viewerX').onclick = () => { $('#viewer').hidden = true; $('#viewerImg').src = ''; };
async function renderPhotos() {
  const { o, w } = cur; const ph = Object.fromEntries((await dbPhotosOf(o.k)).map((p) => [p.slot, p]));
  const sl = MODS[o.mod].slots, nr = sl.filter((s) => s.req).length;
  const n = sl.filter((s) => s.req && w.fotos[s.id]).length;
  $('#fotoCount').textContent = `(${n}/${nr} obligatorias · ${Object.keys(w.fotos).length} en total)`;
  $('#grid').innerHTML = sl.map((s) => {
    const p = ph[s.id];
    if (!p) return `<button class="tile ${s.req ? 'req' : 'opt'}" data-s="${s.id}">📷<br>${s.label}${s.req ? '' : '<br><small>(opcional)</small>'}</button>`;
    const url = URL.createObjectURL(p.blob);
    return `<button class="tile ${p.ok ? 'done' : 'nogps'}" data-s="${s.id}"><img src="${url}" onload="window.URL.revokeObjectURL(this.src)"><span class="tag ${p.ok ? '' : 'w'}">${s.label}${p.ok ? ' ✓' : ' · SIN GPS'}</span></button>`;
  }).join('');
  $('#grid').querySelectorAll('.tile').forEach((t) => (t.onclick = () => takePhoto(t.dataset.s)));
}
let pendingSlot = null, camStream = null, camSlot = null;
function takePhoto(slot) {
  if (cur.w.fotos[slot] && !confirm('Ya hay una foto en ' + slotLabel(slot) + '. ¿Reemplazarla?')) return;
  openCamera(slot);
}
function nativeCam(slot) { pendingSlot = slot; $('#cam').value = ''; $('#cam').click(); }
async function openCamera(slot) {
  camSlot = slot;
  if (!navigator.mediaDevices?.getUserMedia) return nativeCam(slot);
  try {
    camStream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' }, width: { ideal: 1920 }, height: { ideal: 1440 } }, audio: false });
  } catch (e) {
    await aviso('Cámara integrada no disponible', 'No se pudo abrir la cámara dentro de la app (' + (e.name || e.message) + '). Se usará la cámara del teléfono.');
    return nativeCam(slot);
  }
  camStream.getVideoTracks().forEach((t) => (t.onended = () => closeCam()));
  const v = $('#camVideo'); v.srcObject = camStream; v.play().catch(() => {});
  $('#camLabel').textContent = slotLabel(slot) + ' · Orden ' + cur.o.orden; $('#camOv').hidden = false; updGps();
}
function closeCam() {
  if (camStream) camStream.getTracks().forEach((t) => t.stop());
  camStream = null; $('#camVideo').srcObject = null; $('#camOv').hidden = true;
}
$('#camCancel').onclick = closeCam;
$('#camNative').onclick = () => { const s = camSlot; closeCam(); nativeCam(s); };
$('#camGal').onclick = () => { const s = camSlot; closeCam(); pendingSlot = s; $('#gal').value = ''; $('#gal').click(); };
let disparando = false;
$('#camShot').onclick = async () => {
  if (disparando) return;
  const v = $('#camVideo'); if (!v.videoWidth || !camSlot) return;
  disparando = true; $('#camShot').disabled = true;
  try {
    const k = Math.min(1, MAX_PX / Math.max(v.videoWidth, v.videoHeight));
    const c = document.createElement('canvas'); c.width = Math.round(v.videoWidth * k); c.height = Math.round(v.videoHeight * k);
    c.getContext('2d').drawImage(v, 0, 0, c.width, c.height); const slot = camSlot;
    closeCam(); await procesarFoto(slot, c);
  } finally { disparando = false; $('#camShot').disabled = false; }
};
async function procesarFoto(slot, file, galeria) {
  if (!cur || !file) return; busy(true); let sinGps = false;
  try {
    let fix;
    if (galeria) {
      fix = await readExifGps(file);
      if (!fix) { const f = await getFix(); fix = f ? { lat: f.lat, lon: f.lon, acc: f.acc, ts: f.ts, note: 'ubicación actual' } : null; }
    } else fix = await getFix();
    const blob = await stampPhoto(file, slotLabel(slot) + (galeria ? ' · GALERÍA' : ''), cur.o, fix);
    const ts = Date.now();
    await dbPut('photos', { key: cur.o.k + '|' + slot, k: cur.o.k, orden: cur.o.orden, mod: cur.o.mod, slot, blob, lat: fix?.lat ?? null, lon: fix?.lon ?? null, acc: fix?.acc ?? null, ts, ok: !!fix, origen: galeria ? (fix?.note ? 'galeria-ubicacion-actual' : 'galeria-exif') : 'camara' });
    cur.w.fotos[slot] = { ts, lat: fix?.lat ?? null, lon: fix?.lon ?? null, ok: !!fix };
    await saveWork(cur.w); await renderPhotos(); sinGps = !fix;
  } catch (err) { busy(false); await aviso('No se pudo guardar la foto', String(err.message || err)); return; }
  finally { busy(false); }
  if (sinGps) await aviso('Sin coordenadas', 'No se pudo obtener el GPS. La foto quedó marcada "SIN GPS". Actívalo y vuelve a tomarla antes de guardar la orden.');
}
$('#gal').addEventListener('change', async (e) => { const file = e.target.files[0]; const slot = pendingSlot; pendingSlot = null; if (file && slot) await procesarFoto(slot, file, true); });
$('#cam').addEventListener('change', async (e) => { const file = e.target.files[0]; const slot = pendingSlot; pendingSlot = null; if (file && slot) await procesarFoto(slot, file); });

async function guardarOrden() {
  clearTimeout(tObs); const { o, w } = cur; w.obs = $('#obs').value; const f = faltantes(o, w);
  if (f.length) {
    await saveWork(w);
    $('#faltan').hidden = false; $('#faltan').innerHTML = '<b>No se puede guardar. Falta:</b><ul>' + f.map((x) => '<li>' + esc(x) + '</li>').join('') + '</ul>';
    $('#faltan').scrollIntoView({ behavior: 'smooth', block: 'center' }); return false;
  }
  w.estado = 'completa'; w.fecha = hoy(); w.completada = Date.now(); w.descargada = false;
  await saveWork(w); await dbDel('kv', 'activa'); closeOrder(); return true;
}
function closeOrder() {
  urlsVivas.forEach((u) => URL.revokeObjectURL(u)); urlsVivas = [];
  cur = null; stopGPS('orden'); $('#viewOrden').hidden = true; $('#viewHome').hidden = false; renderList(); renderReporte(); renderMapPts(); if (tabActual === 'mapa') mostrarMapa(); window.scrollTo(0, 0);
}
$('#btnSave').onclick = guardarOrden;
$('#btnDiscard').onclick = async () => {
  const r = await modal('¿Descartar esta orden?', 'Se borrarán las fotos y los datos de la orden ' + esc(cur.o.orden) + '. No se puede deshacer.', [['Cancelar', false], ['Sí, descartar', true, 'danger']]);
  if (!r) return;
  const ph = await dbPhotosOf(cur.o.k); for (const p of ph) await dbDel('photos', p.key);
  delete WKS()[cur.o.orden]; await dbDel('works', cur.o.k); await dbDel('kv', 'activa'); closeOrder();
};
window.addEventListener('popstate', async () => {
  if (cur) {
    history.pushState({ orden: 1 }, '');
    if (!$('#viewer').hidden) { $('#viewerX').click(); return; }
    if (!$('#camOv').hidden) { closeCam(); return; }
    const f = faltantes(cur.o, cur.w);
    await aviso('Termina la orden antes de salir', f.length ? f : ['La orden está completa: toca "Guardar orden".']);
    return;
  }
  if (MOD) showMenu();
});

/* ===== Reportes (Excel + fotos) ===== */
function worksDe(mod, fecha) { return Object.values(DATA[mod].works).filter((w) => w.estado === 'completa' && (!fecha || w.fecha === fecha)).sort((a, b) => (a.completada || 0) - (b.completada || 0)); }
async function renderReporte() {
  if (!MOD) return;
  const f = $('#repFecha').value || hoy(); if (!$('#repFecha').value) $('#repFecha').value = f;
  const ws = worksDe(MOD, f); const todas = worksDe(MOD, null); const sin = todas.filter((w) => !w.descargada);
  const encurso = Object.values(WKS()).filter((w) => w.estado !== 'completa').length;
  $('#repResumen').innerHTML = `<p>${ws.length} orden(es) completada(s) en esta fecha.${encurso ? ' <b>' + encurso + ' en curso (no se incluyen).</b>' : ''}</p>`;
  $('#riesgo').innerHTML = sin.length ? `<div class="alertBox"><b>${sin.length} orden(es) completadas aún sin descargar.</b> Descarga el reporte al terminar la jornada.</div>` : '<p class="muted">Todo lo completado ya fue descargado.</p>';
  const ns = (await dbAll('samples')).length;
  $('#muestraInfo').textContent = ns ? `${ns} fotos del Grid guardadas en el teléfono.` : 'Aún no hay fotos del Grid guardadas. Mientras tengas internet, la app las muestra directamente.';
}
const colL = (i) => { let s = ''; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
const xe = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
async function buildXlsx(headers, widths, rows, sheetName) {
  const z = new JSZip();
  z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  z.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${xe(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  z.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  z.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0B4F8A"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFBBBBBB"/></left><right style="thin"><color rgb="FFBBBBBB"/></right><top style="thin"><color rgb="FFBBBBBB"/></top><bottom style="thin"><color rgb="FFBBBBBB"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs></styleSheet>');
  const cell = (v, r, c, s) => {
    const ref = colL(c) + r;
    if (typeof v === 'number' && isFinite(v)) return `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
    if (v == null || v === '') return `<c r="${ref}" s="${s}"/>`;
    return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xe(v)}</t></is></c>`;
  };
  let sd = `<row r="1" ht="42" customHeight="1">${headers.map((h, i) => cell(h, 1, i, 1)).join('')}</row>`;
  rows.forEach((r, ri) => { sd += `<row r="${ri + 2}">${r.map((v, i) => cell(v, ri + 2, i, 2)).join('')}</row>`; });
  const last = colL(headers.length - 1) + (rows.length + 1);
  z.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols><sheetData>${sd}</sheetData><autoFilter ref="A1:${last}"/></worksheet>`);
  return await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
const XH = ['N°', 'Módulo', 'Fecha fiscalización', 'Hora', 'Orden', 'Medidor', 'Cuenta contrato', 'Cliente', 'Actividad', 'Ejecutado por (técnico)', 'Latitud orden', 'Longitud orden', 'Latitud fiscalización', 'Longitud fiscalización', 'Distancia al punto de la orden (m)', 'Fotos obligatorias', 'Detalle de fotos', 'Fotos sin GPS', 'Foto 1 del Grid confirmada', 'Foto 2 del Grid confirmada', 'Observación'];
const XW = [5, 20, 12, 7, 12, 13, 15, 28, 18, 20, 11, 11, 12, 12, 13, 11, 34, 9, 12, 12, 36];
function rowDe(n, w, o, ph) {
  const sl = MODS[o.mod].slots; const c1 = ph.find((p) => p.slot === 'antes' && p.ok) || ph.find((p) => p.ok);
  const req = sl.filter((s) => s.req); const okReq = req.filter((s) => ph.find((p) => p.slot === s.id)).length;
  const det = sl.map((s) => s.label + (ph.find((p) => p.slot === s.id) ? ' ✓' : ' —')).join(', ');
  const d = c1 && o.lat != null ? Math.round(distM(c1.lat, c1.lon, o.lat, o.lon)) : '';
  const cf = (i) => (nGrid(o) > i ? (w.conf?.[i] ? 'SÍ' : 'NO') : 'N/A');
  return [n, MODS[o.mod].name, w.fecha ? w.fecha.split('-').reverse().join('/') : '', w.completada ? hora(w.completada) : '', o.orden, o.medidor, o.cuenta, o.cliente, o.actividades, o.tecnico || w.ejecutor, o.lat ?? '', o.lon ?? '', c1 ? +c1.lat.toFixed(6) : '', c1 ? +c1.lon.toFixed(6) : '', d, `${okReq}/${req.length}`, det, ph.filter((p) => !p.ok).length, cf(0), cf(1), w.obs || ''];
}
const safe = (s) => String(s || 'SIN_MEDIDOR').replace(/[\\/:*?"<>|]+/g, '_').trim() || 'SIN_MEDIDOR';
// Reparte los trabajos en paquetes de ~90 MB para que el teléfono no se quede sin memoria.
async function paquetes(items) {
  const parts = []; let cur = [], sz = 0;
  for (const it of items) {
    const ph = await dbPhotosOf(it.o.k); const s = ph.reduce((a, p) => a + p.blob.size, 0);
    if (cur.length && sz + s > ZIP_MAX) { parts.push(cur); cur = []; sz = 0; }
    cur.push({ ...it, ph }); sz += s;
  }
  if (cur.length) parts.push(cur); return parts;
}
async function generar(mods, modo, fecha) {
  const items = [];
  for (const m of mods) { const om = Object.fromEntries(DATA[m].orders.map((o) => [o.orden, o])); worksDe(m, modo === 'todo' ? null : fecha).forEach((w) => { if (om[w.orden]) items.push({ w, o: om[w.orden], m }); }); }
  if (!items.length) { await aviso('Nada que descargar', 'No hay órdenes completadas para esa selección.'); return; }
  busy(true, 'Preparando el reporte…');
  try {
    const multi = mods.length > 1; const tag = modo === 'todo' ? 'RESPALDO_' + hoy() : fecha;
    const nom = multi ? 'TODOS' : MODS[mods[0]].name.replace(/ /g, '_');
    if (modo === 'xls') {
      const rows = []; let n = 0; for (const it of items) rows.push(rowDe(++n, it.w, it.o, await dbPhotosOf(it.o.k)));
      download(new Blob([await buildXlsx(XH, XW, rows, 'Fiscalización')], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `Reporte_${nom}_${tag}.xlsx`);
    } else {
      const parts = await paquetes(items); let pi = 0;
      for (const part of parts) {
        pi++; busy(true, `Armando paquete ${pi} de ${parts.length}…`);
        const zip = new JSZip(); const rows = []; let n = 0;
        for (const it of part) {
          const { w, o, ph } = it; rows.push(rowDe(++n, w, o, ph));
          const dir = zip.folder((multi ? MODS[o.mod].name + '/' : '') + safe(o.medidor || o.orden));
          const sl = MODS[o.mod].slots;
          for (const p of ph) { const i = sl.findIndex((s) => s.id === p.slot); if (i >= 0) dir.file(`${String(i + 1).padStart(2, '0')}_${sl[i].label.replace(/ /g, '_').replace('Ó', 'O').replace('É', 'E')}_orden${o.orden}.jpg`, p.blob); }
        }
        zip.file(`Reporte_${nom}_${tag}${parts.length > 1 ? '_p' + pi : ''}.xlsx`, await buildXlsx(XH, XW, rows, 'Fiscalización'));
        download(await zip.generateAsync({ type: 'blob', compression: 'STORE' }), `Trabajos_${nom}_${tag}${parts.length > 1 ? '_p' + pi + 'de' + parts.length : ''}.zip`);
        for (const it of part) { it.w.descargada = true; await saveWork(it.w); }
        await new Promise((r) => setTimeout(r, 1200));
      }
    }
  } catch (err) { busy(false); await aviso('Error al generar el reporte', String(err.message || err)); return; }
  busy(false); if (MOD) renderReporte(); else renderMenu();
}
$('#btnZip').onclick = () => generar([MOD], 'zip', $('#repFecha').value);
$('#btnXls').onclick = () => generar([MOD], 'xls', $('#repFecha').value);
$('#btnBackup').onclick = () => generar([MOD], 'todo');
$('#btnBackupAll').onclick = () => generar(['lect', 'notif', 'corte'], 'todo');
$('#repFecha').onchange = renderReporte;

/* ===== Fotos de muestra del Grid (ZIP) ===== */
$('#btnMuestras').onclick = () => $('#zipIn').click();
$('#zipIn').onchange = async (e) => {
  const f = e.target.files[0]; if (!f) return; busy(true, 'Leyendo el ZIP de fotos…'); clearTimeout(busyT1); clearTimeout(busyT2);
  try {
    const z = await JSZip.loadAsync(f); const ents = Object.values(z.files).filter((x) => !x.dir && /([^/]+)_([A-Za-z0-9]+)\.jpe?g$/i.test(x.name)); let n = 0;
    if (!ents.length) throw new Error('El ZIP no tiene fotos con nombre "orden_Nombre.jpg".');
    for (const x of ents) {
      const m = x.name.match(/([^/]+)_([A-Za-z0-9]+)\.jpe?g$/i); const blob = await x.async('blob');
      await dbPut('samples', { key: m[1] + '|' + m[2], orden: m[1], name: m[2], blob }); n++;
      if (n % 25 === 0) $('#busyTxt').textContent = `Guardando fotos ${n} de ${ents.length}…`;
    }
    busy(false); await aviso('Fotos del Grid guardadas', `${n} fotos guardadas en el teléfono. Ya se ven sin internet.`); renderReporte();
  } catch (err) { busy(false); await aviso('No se pudo leer el ZIP', String(err.message || err)); }
  e.target.value = '';
};

/* ===== Cargar CSV / Excel ===== */
$('#csvIn').onchange = async (e) => {
  const f = e.target.files[0]; if (!f) return; const mod = MOD;
  try {
    busy(true, 'Leyendo el archivo…');
    const rows = /\.xlsx$/i.test(f.name) ? await parseXLSX(await f.arrayBuffer()) : parseCSV(await f.text());
    const nuevas = rowsToOrders(rows, mod); if (!nuevas.length) throw new Error('El archivo no tiene órdenes.');
    const tx = db.transaction('orders', 'readwrite'); const st = tx.objectStore('orders'); nuevas.forEach((o) => st.put(o));
    await new Promise((res, rej) => { tx.oncomplete = res; tx.onerror = () => rej(tx.error); tx.onabort = () => rej(tx.error); });
    await loadAll(); busy(false); renderList(); mapaCentrado = false; renderMapPts();
    const sinCoord = nuevas.filter((o) => o.lat == null).length, sinMed = nuevas.filter((o) => !o.medidor).length;
    await aviso('Archivo cargado', [`${nuevas.length} órdenes cargadas en ${MODS[mod].name}. Lo ya trabajado se conserva.`, sinCoord ? sinCoord + ' sin coordenadas (no saldrán en el mapa).' : 'Todas con coordenadas.', sinMed ? sinMed + ' sin número de medidor.' : 'Todas con medidor.'].filter(Boolean));
  } catch (err) { busy(false); await aviso('No se pudo leer el archivo', String(err.message || err)); }
  e.target.value = '';
};

/* ===== Eliminar CSV anterior ===== */
$('#btnDelCsv').onclick = async () => {
  const ords = ORD(), W = WKS(); if (!ords.length) { await aviso('No hay nada que eliminar', 'Este módulo no tiene órdenes cargadas.'); return; }
  const pend = ords.filter((o) => estOf(o) === 'pend'), sinDesc = Object.values(W).filter((w) => w.estado === 'completa' && !w.descargada).length, curso = Object.values(W).filter((w) => w.estado !== 'completa').length;
  const r = await modal('Eliminar CSV anterior · ' + MODS[MOD].name,
    `<p>Órdenes cargadas: <b>${ords.length}</b> · Pendientes sin trabajar: <b>${pend.length}</b> · Atendidas: <b>${ords.length - pend.length - curso}</b> · En curso: <b>${curso}</b>.</p><p>Atendidas sin descargar: <b>${sinDesc}</b>.</p><p class="muted small">Opción 1 quita solo las órdenes sin trabajar y conserva todo lo atendido. Opción 2 borra TODO este módulo (órdenes, fotos y trabajos).</p>`,
    [['Cancelar', 0], ['1. Quitar solo las pendientes', 1, 'primary'], ['2. Borrar TODO el módulo', 2, 'danger']]);
  if (!r) return;
  if (r === 1) {
    for (const o of pend) await dbDel('orders', o.k);
    await loadAll(); renderList(); renderMapPts(); await aviso('Listo', `${pend.length} órdenes pendientes eliminadas. Lo atendido y en curso se conserva.`); return;
  }
  if (sinDesc || curso) {
    const ir = await modal('Antes de borrar todo', `<p>Hay <b>${sinDesc}</b> orden(es) atendida(s) sin descargar y <b>${curso}</b> en curso. Si borras todo se pierden.</p><p>Primero descarga el respaldo.</p>`, [['Cancelar', 0], ['Descargar respaldo ahora', 1, 'primary']]);
    if (ir === 1) await generar([MOD], 'todo'); return;
  }
  const t = await modal('Confirmar borrado total', `<p>Se borrarán las ${ords.length} órdenes de <b>${MODS[MOD].name}</b> con todas sus fotos y trabajos. Todo lo atendido ya fue descargado. No se puede deshacer.</p>`, [['Cancelar', 0], ['Sí, borrar todo', 1, 'danger']]);
  if (t !== 1) return;
  busy(true, 'Borrando…');
  try {
    for (const o of ords) {
      for (const k of await dbKeysByPrefix('photos', o.k + '|')) await dbDel('photos', k);
      await dbDel('works', o.k); await dbDel('orders', o.k);
      for (const k of await dbKeysByPrefix('samples', o.orden + '|')) await dbDel('samples', k);
    }
    await dbDel('kv', 'activa'); await loadAll();
  } finally { busy(false); }
  renderList(); renderReporte(); renderMapPts(); await aviso('Listo', 'El módulo quedó vacío. Ya puedes cargar un CSV nuevo.');
};

/* ===== Mapa ===== */
let MAP = null, ptLayer = null, meMk = null, meCirc = null, roadsMajor = null, roadsInfo = '', minorCells = null, minorOn = new Map(), mapaCentrado = false;
const COL = { pend: '#c0392b', curso: '#e67e22', ok: '#1b8a3f' }; const CELL = 0.02;
function decodeRoads(arr) {
  return (arr || []).map((w) => { let la = w[0], lo = w[1]; const pts = [[la / 1e5, lo / 1e5]]; for (let i = 2; i < w.length; i += 2) { la += w[i]; lo += w[i + 1]; pts.push([la / 1e5, lo / 1e5]); } return pts; });
}
async function cargarCalles() {
  try {
    const r = await fetch('calles.json'); if (!r.ok) throw new Error('sin archivo'); const d = await r.json();
    const rr = L.canvas({ pane: 'calles', padding: 0.3 });
    const mayor = decodeRoads(d.major), menor = decodeRoads(d.minor);
    roadsMajor = mayor.length ? L.polyline(mayor, { renderer: rr, color: '#8a9ba5', weight: 3, interactive: false, smoothFactor: 1.5, lineCap: 'round' }) : null;
    minorCells = new Map();   // las calles menores se dibujan por cuadrículas, solo las que están en pantalla
    for (const p of menor) { const key = Math.floor(p[0][0] / CELL) + ',' + Math.floor(p[0][1] / CELL); if (!minorCells.has(key)) minorCells.set(key, []); minorCells.get(key).push(p); }
    minorCells.renderer = rr;
    roadsInfo = (mayor.length + menor.length) ? '' : 'Todavía no hay calles cargadas: el mapa muestra solo los puntos.';
    ajustarCalles();
  } catch (_) { roadsInfo = 'Todavía no hay calles cargadas: el mapa muestra solo los puntos.'; }
  const n = $('#mapNota'); if (n) n.textContent = roadsInfo;
}
function ajustarCalles() {
  if (!MAP) return; const z = MAP.getZoom();
  if (roadsMajor && !MAP.hasLayer(roadsMajor)) roadsMajor.addTo(MAP);
  if (!minorCells) return;
  if (z < 14) { minorOn.forEach((l) => MAP.removeLayer(l)); minorOn.clear(); return; }
  const b = MAP.getBounds().pad(0.4); const want = new Set();
  for (let i = Math.floor(b.getSouth() / CELL); i <= Math.floor(b.getNorth() / CELL); i++) for (let j = Math.floor(b.getWest() / CELL); j <= Math.floor(b.getEast() / CELL); j++) { const k = i + ',' + j; if (minorCells.has(k)) want.add(k); }
  minorOn.forEach((l, k) => { if (!want.has(k)) { MAP.removeLayer(l); minorOn.delete(k); } });
  want.forEach((k) => { if (!minorOn.has(k)) { const l = L.polyline(minorCells.get(k), { renderer: minorCells.renderer, color: '#b9c5cb', weight: 1.6, interactive: false, smoothFactor: 2, lineCap: 'round' }).addTo(MAP); minorOn.set(k, l); } });
}
function initMap() {
  if (MAP || typeof L === 'undefined') return;
  MAP = L.map('mapBox', { zoomControl: true, preferCanvas: true, zoomSnap: 0.5, minZoom: 10, maxZoom: 19, attributionControl: false });
  MAP.createPane('calles').style.zIndex = 350;
  L.control.attribution({ prefix: false }).addAttribution('Calles © colaboradores de OpenStreetMap').addTo(MAP);
  L.control.scale({ imperial: false }).addTo(MAP);
  MAP.setView([-0.15, -78.49], 12);
  ptLayer = L.layerGroup().addTo(MAP);
  MAP.on('zoomend moveend', ajustarCalles);
  MAP.on('popupopen', (e) => { const b = e.popup.getElement().querySelector('.popBtn'); if (b) b.onclick = () => { MAP.closePopup(); openOrder(b.dataset.o); }; });
  cargarCalles();
}
function renderMapPts() {
  if (!MAP || !MOD) return; ptLayer.clearLayers();
  const verOk = $('#mapAtend').checked; let sinCoord = 0, np = 0, nc = 0;
  ORD().forEach((o) => {
    const e = estOf(o); if (e === 'ok' && !verOk) return;
    if (o.lat == null || o.lon == null) { if (e !== 'ok') sinCoord++; return; }
    if (e === 'pend') np++; if (e === 'curso') nc++;
    const m = L.circleMarker([o.lat, o.lon], { radius: 9, color: '#fff', weight: 2, fillColor: COL[e], fillOpacity: 0.95 });
    m.bindPopup(`<div class="popT"><b>Orden ${esc(o.orden)}</b><br>Medidor ${esc(o.medidor)}${o.cuenta ? '<br>Cuenta ' + esc(o.cuenta) : ''}${o.tecnico ? '<br>Ejecutor: ' + esc(o.tecnico) : ''}<br><small>${o.lat.toFixed(6)}, ${o.lon.toFixed(6)}</small>${e === 'ok' ? '<br><i>Ya atendida</i>' : ''}<button class="popBtn" data-o="${esc(o.orden)}">Abrir orden</button></div>`, { closeButton: false, maxWidth: 240, autoPanPadding: [30, 30] });
    ptLayer.addLayer(m);
  });
  $('#mapInfo').innerHTML = `<span style="color:${COL.pend}">●</span> Pendientes ${np} &nbsp;<span style="color:${COL.curso}">●</span> En curso ${nc}` + (verOk ? ` &nbsp;<span style="color:${COL.ok}">●</span> Atendidas` : '');
  $('#mapNota').textContent = [roadsInfo, sinCoord ? sinCoord + ' orden(es) sin coordenadas en el CSV no se pueden dibujar.' : ''].filter(Boolean).join(' ');
}
function ajustarATodas() {
  if (!MAP) return; const pts = ORD().filter((o) => estOf(o) !== 'ok' && o.lat != null).map((o) => [o.lat, o.lon]);
  if (pts.length) MAP.fitBounds(pts, { padding: [24, 24], maxZoom: 16 }); else MAP.setView([-0.15, -78.49], 12);
}
function mostrarMapa() {
  initMap(); if (!MAP) { $('#mapNota').textContent = 'No se pudo cargar el mapa.'; return; }
  startGPS('mapa'); renderMapPts();
  setTimeout(() => { MAP.invalidateSize(); if (!mapaCentrado) { ajustarATodas(); mapaCentrado = true; } updMapMe(); ajustarCalles(); }, 60);
}
function updMapMe() {
  if (!MAP || !gps) return; const ll = [gps.lat, gps.lon];
  if (!meMk) { meCirc = L.circle(ll, { radius: gps.acc || 20, color: '#1976d2', weight: 1, fillOpacity: 0.12, interactive: false }).addTo(MAP); meMk = L.circleMarker(ll, { radius: 8, color: '#fff', weight: 3, fillColor: '#1976d2', fillOpacity: 1, interactive: false }).addTo(MAP); }
  else { meMk.setLatLng(ll); meCirc.setLatLng(ll); meCirc.setRadius(gps.acc || 20); }
}
$('#mapMe').onclick = () => { if (!gps) return aviso('Buscando tu ubicación', 'Aún no hay GPS. Espera unos segundos al aire libre y vuelve a tocar.'); MAP.setView([gps.lat, gps.lon], Math.max(MAP.getZoom(), 17)); updMapMe(); };
$('#mapAll').onclick = () => { ajustarATodas(); };
$('#mapAtend').onchange = renderMapPts;
/* ===== Sin internet: revisión ===== */
const OFFLINE_FILES = ['./', 'index.html', 'style.css', 'app.js', 'manifest.json', 'icon-192.png', 'icon-512.png', 'lib/jszip.min.js', 'lib/leaflet.js', 'lib/leaflet.css', 'calles.json'];
async function estadoOffline() {
  const r = { lineas: [], ok: false };
  const add = (ok, t) => r.lineas.push([ok, t]);
  const tiene = 'serviceWorker' in navigator;
  add(tiene, tiene ? 'Este navegador permite trabajar sin internet' : 'Este navegador NO permite trabajar sin internet (usa Chrome)');
  let reg = null, ctrl = false, faltan = OFFLINE_FILES.slice();
  if (tiene) { try { reg = await navigator.serviceWorker.getRegistration(); } catch (_) { } ctrl = !!navigator.serviceWorker.controller; }
  add(!!reg, reg ? 'El guardado sin internet está registrado' : 'El guardado sin internet NO está registrado');
  add(ctrl, ctrl ? 'La app abierta ahora está protegida para abrir sin internet' : 'La app abierta ahora NO está protegida: ciérrala por completo y ábrela de nuevo con internet');
  try {
    const hay = [];
    for (const f of OFFLINE_FILES) { if (await caches.match(f, { ignoreSearch: true })) hay.push(f); }
    faltan = OFFLINE_FILES.filter((f) => !hay.includes(f));
    add(!faltan.length, faltan.length ? 'Faltan archivos guardados: ' + faltan.join(', ') : 'Todos los archivos están guardados (' + OFFLINE_FILES.length + ')');
  } catch (e) { add(false, 'No se pudo revisar lo guardado (' + (e.name || e.message) + ')'); }
  const p = location.pathname; const barra = p.endsWith('/') || p.endsWith('.html');
  add(barra, barra ? 'La dirección es correcta' : 'La dirección no termina en "/" (' + p + '): ábrela desde el ícono instalado o agrega "/" al final');
  const ua = navigator.userAgent; const inapp = /\bwv\b|FBAN|FBAV|Instagram|WhatsApp|Line\//i.test(ua); const mi = /MiuiBrowser|XiaoMi|Mi Browser/i.test(ua); const chrome = /Chrome\//.test(ua) && !/SamsungBrowser|OPR|EdgA|MiuiBrowser/.test(ua);
  add(chrome && !inapp && !mi, chrome && !inapp && !mi ? 'Navegador: Chrome' : 'Navegador no recomendado (' + (mi ? 'Mi Browser' : inapp ? 'navegador dentro de otra app' : 'otro') + '). Usa Chrome.');
  const standalone = matchMedia('(display-mode: standalone)').matches || navigator.standalone;
  r.lineas.push([null, standalone ? 'Abierta como app instalada' : 'Abierta en una pestaña de Chrome (se recomienda instalarla)']);
  try { if (navigator.storage?.persisted) { const pe = await navigator.storage.persisted(); add(pe, pe ? 'Almacenamiento protegido contra limpiezas automáticas' : 'Almacenamiento NO protegido: el teléfono podría borrarlo. Instala la app y evita limpiadores'); } } catch (_) { }
  r.ok = tiene && !!reg && ctrl && !faltan.length;
  return r;
}
async function revisarChip() {
  const c = $('#offChip'); try { const r = await estadoOffline(); c.hidden = false; c.className = 'chip ' + (r.ok ? 'ok' : 'warn'); c.textContent = r.ok ? 'Sin internet: lista ✓' : 'Sin internet: NO lista'; } catch (_) { }
}
$('#btnDiag').onclick = async () => {
  const r = await estadoOffline();
  $('#diagOut').innerHTML = r.lineas.map(([ok, t]) => `<div class="diagL">${ok === null ? 'ℹ️' : ok ? '✅' : '❌'} ${esc(t)}</div>`).join('') + `<div class="diagV ${r.ok ? 'ok' : 'bad'}">${r.ok ? 'Lista para trabajar sin internet.' : 'Todavía NO está lista para trabajar sin internet.'}</div>`;
  revisarChip();
};

/* ===== Entrada (usuario supervisor) ===== */
$('#loginForm').onsubmit = async (e) => {
  e.preventDefault();
  const u = $('#lgU').value.trim().toUpperCase().replace(/\s+/g, ''), p = $('#lgP').value.trim().toUpperCase();
  let ok = false;
  try { ok = (await sha256hex('campo-eeq|' + u + '|' + p)) === AUTH_HASH; } catch (_) { await aviso('No se pudo validar', 'Abre la app desde el enlace seguro (https).'); return; }
  if (ok) { await dbPut('kv', { u: 'supervisor 02', ts: Date.now() }, 'auth'); $('#lgErr').hidden = true; $('#lgP').value = ''; $('#login').hidden = true; await arrancar(); }
  else { $('#lgErr').hidden = false; $('#lgP').value = ''; }
};
$('#btnLogout').onclick = async () => { await dbDel('kv', 'auth'); $('#lgU').value = ''; $('#login').hidden = false; };
async function arrancar() {
  const activa = await dbGet('kv', 'activa'); const o = activa && DATA[activa.split('|')[0]]?.orders.find((x) => x.k === activa);
  if (o && !cur) { MOD = o.mod; enterMod(o.mod); await openOrder(o.orden); await aviso('Orden en curso recuperada', 'Se cerró la app con la orden ' + o.orden + ' sin guardar. Continúa donde la dejaste.'); }
  else showMenu();
}

/* ===== Navegación y arranque ===== */
let tabActual = 'ordenes';
document.querySelectorAll('.tabs button').forEach((b) => (b.onclick = () => {
  const t = b.dataset.tab; const antes = tabActual; tabActual = t;
  document.querySelectorAll('.tabs button').forEach((x) => x.classList.toggle('on', x === b));
  $('#tabOrdenes').hidden = t !== 'ordenes'; $('#tabMapa').hidden = t !== 'mapa'; $('#tabReporte').hidden = t !== 'reporte';
  if (t === 'reporte') renderReporte();
  if (t === 'mapa') mostrarMapa(); else if (antes === 'mapa') stopGPS('mapa');
}));
document.querySelectorAll('.filters button').forEach((b) => (b.onclick = () => { filtro = b.dataset.f; document.querySelectorAll('.filters button').forEach((x) => x.classList.toggle('on', x === b)); renderList(); }));
$('#q').oninput = renderList;
let deferredInstall;
window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); deferredInstall = e; $('#btnInstall').hidden = false; });
$('#btnInstall').onclick = async () => { if (deferredInstall) { deferredInstall.prompt(); deferredInstall = null; $('#btnInstall').hidden = true; } };

(async function init() {
  db = await openDB();
  try { if (navigator.storage?.persist) await navigator.storage.persist(); } catch (_) { /* nada */ }
  await loadAll();
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.addEventListener('controllerchange', () => setTimeout(revisarChip, 800));
    navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).catch(() => {});
    setTimeout(revisarChip, 2500); setTimeout(revisarChip, 9000);
  }
  const auth = await dbGet('kv', 'auth');
  if (!auth) { $('#login').hidden = false; $('#viewMenu').hidden = true; return; }
  await arrancar();
})();
