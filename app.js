// Finanças — app pessoal ligada ao Excel no OneDrive.
// Lê o Excel (só leitura) e guarda os registos do iPhone em Financas_iPhone.csv,
// na mesma pasta. O Excel importa esse CSV pela folha "Do iPhone" (Power Query).
'use strict';

/* ------------------------------------------------------------------ config */
const AUTH = 'https://login.microsoftonline.com/consumers/oauth2/v2.0';
const GRAPH = 'https://graph.microsoft.com/v1.0';
const SCOPES = 'openid profile offline_access User.Read Files.ReadWrite';
const CSV_NAME = 'Financas_iPhone.csv';
const CSV_HEAD = ['ID', 'Registado', 'Registo', 'Data', 'Descricao', 'Tipo', 'Categoria', 'Conta', 'Valor', 'Divida', 'Nota'];
const MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
const MES3 = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
const GROUP_COLORS = ['#1F6F66', '#E0A34A', '#6C8FD1', '#E76F51', '#8E7CC3', '#2A9D8F', '#F4A261', '#A3B18A', '#B9C3BF', '#C0392B', '#7F8C8D', '#34495E'];

// Mapa de células do Excel (tem de bater com a estrutura do livro)
const MAP = {
  resumo: 'Resumo', mov: 'Movimentos', dias: 'Dias de Trabalho', div: 'Dívidas',
  cat: 'Categorias', cfg: 'Configuração', iph: 'Do iPhone',
  groupsFirstRow: 65, groupsLastRow: 76,
  diasFirstRow: 7, diasLastRow: 66, diasFirstCol: 10, // J
  divFirstRow: 6, divLastRow: 105,
};

const LS = {
  get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* cheio ou bloqueado */ } },
  del(k) { try { localStorage.removeItem(k); } catch (e) { /* */ } },
};
const cfg = () => LS.get('fin_cfg', {});
const setCfg = (patch) => LS.set('fin_cfg', Object.assign(cfg(), patch));
const redirectUri = () => location.origin + location.pathname.replace(/index\.html$/, '');

/* ------------------------------------------------------------------ utils */
const $app = () => document.getElementById('app');
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const eur = (v) => (typeof v === 'number' && isFinite(v) ? v : 0).toLocaleString('pt-PT', { style: 'currency', currency: 'EUR' });
const eur0 = (v) => (typeof v === 'number' && isFinite(v) ? v : 0).toLocaleString('pt-PT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });
const pct = (v) => (typeof v === 'number' && isFinite(v) ? Math.round(v * 100) + '%' : '—');
const num = (v) => (typeof v === 'number' && isFinite(v) ? v : 0);
const pad2 = (n) => String(n).padStart(2, '0');
const iso = (y, m, d) => `${y}-${pad2(m)}-${pad2(d)}`;
const todayIso = () => { const t = new Date(); return iso(t.getFullYear(), t.getMonth() + 1, t.getDate()); };
const parseIso = (s) => { const [y, m, d] = String(s).slice(0, 10).split('-').map(Number); return { y, m, d }; };
function serialToIso(n) { // número de série do Excel → AAAA-MM-DD
  if (typeof n !== 'number') return null;
  const ms = Math.round((n - 25569) * 86400000);
  const dt = new Date(ms);
  return iso(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
}
const daysIn = (y, m) => new Date(y, m, 0).getDate();
const weekday = (y, m, d) => new Date(y, m - 1, d).getDay(); // 0 dom … 6 sáb
const isWeekend = (y, m, d) => { const w = weekday(y, m, d); return w === 0 || w === 6; };
function parseMoney(s) {
  s = String(s || '').replace(/[€\s]/g, '');
  if (s.includes(',') && s.includes('.')) s = s.replace(/\./g, '').replace(',', '.');
  else s = s.replace(',', '.');
  const v = Number(s);
  return isFinite(v) ? Math.round(v * 100) / 100 : NaN;
}
function colLetter(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
function uid() { return 'i' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }
function toast(msg) { const t = document.getElementById('toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove('show'), 2600); }
const ICON = {
  plus: '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.6" aria-hidden="true"><path d="M12 5v14M5 12h14"/></svg>',
  grid: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><rect x="3" y="3" width="7" height="9" rx="1.5"/><rect x="14" y="3" width="7" height="5" rx="1.5"/><rect x="14" y="12" width="7" height="9" rx="1.5"/><rect x="3" y="16" width="7" height="5" rx="1.5"/></svg>',
  list: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M4 6h16M4 12h16M4 18h10"/></svg>',
  check: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M9 11l3 3 8-8"/><path d="M20 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/></svg>',
  swap: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M16 3h5v5"/><path d="M8 21H3v-5"/><path d="M21 3l-7 7M3 21l7-7"/></svg>',
  gear: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M4.2 4.2l2.1 2.1M17.7 17.7l2.1 2.1M2 12h3M19 12h3M4.2 19.8l2.1-2.1M17.7 6.3l2.1-2.1"/></svg>',
  sync: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M21 12a9 9 0 01-15.5 6.2L3 16"/><path d="M3 12a9 9 0 0115.5-6.2L21 8"/><path d="M21 3v5h-5M3 21v-5h5"/></svg>',
  left: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" aria-hidden="true"><path d="M15 18l-6-6 6-6"/></svg>',
  right: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" aria-hidden="true"><path d="M9 18l6-6-6-6"/></svg>',
  tick: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="3" aria-hidden="true"><path d="M5 12l5 5 9-10"/></svg>',
  warn: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true" style="flex-shrink:0"><path d="M12 9v4M12 17h.01"/><path d="M10.3 3.9L1.8 18a2 2 0 001.7 3h17a2 2 0 001.7-3L13.7 3.9a2 2 0 00-3.4 0z"/></svg>',
  chart: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M3 3v18h18"/><path d="M7 15l4-4 3 3 5-6"/></svg>',
  ext: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M14 3h7v7"/><path d="M10 14L21 3"/><path d="M21 14v5a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h5"/></svg>',
};

/* ------------------------------------------------------------------ auth (OAuth 2.0 + PKCE, sem bibliotecas) */
function b64url(bytes) { return btoa(String.fromCharCode(...new Uint8Array(bytes))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''); }
async function login(prompt) {
  const verifier = b64url(crypto.getRandomValues(new Uint8Array(48)));
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const state = b64url(crypto.getRandomValues(new Uint8Array(16)));
  LS.set('fin_pkce', { verifier, state });
  const q = new URLSearchParams({
    client_id: cfg().clientId, response_type: 'code', redirect_uri: redirectUri(), response_mode: 'query',
    scope: SCOPES, code_challenge: challenge, code_challenge_method: 'S256', state,
  });
  if (prompt) q.set('prompt', prompt);
  location.assign(`${AUTH}/authorize?${q}`);
}
async function tokenRequest(params) {
  const r = await fetch(`${AUTH}/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams(params) });
  const j = await r.json();
  if (!r.ok) throw new Error(j.error_description || j.error || 'Falha no login');
  LS.set('fin_tok', { access: j.access_token, refresh: j.refresh_token, exp: Date.now() + (j.expires_in - 60) * 1000 });
}
async function handleRedirect() {
  const u = new URL(location.href);
  const code = u.searchParams.get('code'); const err = u.searchParams.get('error');
  if (!code && !err) return;
  history.replaceState(null, '', redirectUri());
  if (err) throw new Error(u.searchParams.get('error_description') || err);
  const pk = LS.get('fin_pkce', {});
  if (!pk.state || pk.state !== u.searchParams.get('state')) throw new Error('Resposta de login inválida — tenta outra vez.');
  await tokenRequest({ client_id: cfg().clientId, grant_type: 'authorization_code', code, redirect_uri: redirectUri(), code_verifier: pk.verifier, scope: SCOPES });
  LS.del('fin_pkce');
}
async function accessToken() {
  const t = LS.get('fin_tok', null);
  if (!t) return null;
  if (t.access && Date.now() < t.exp) return t.access;
  if (t.refresh) {
    try { await tokenRequest({ client_id: cfg().clientId, grant_type: 'refresh_token', refresh_token: t.refresh, scope: SCOPES }); return LS.get('fin_tok').access; }
    catch (e) { if (!navigator.onLine) throw e; }
  }
  LS.del('fin_tok');
  await login(); // sessão Microsoft normalmente ainda ativa: volta sem pedir nada
  return new Promise(() => {});
}
async function graph(path, opts = {}) {
  const tok = await accessToken();
  const r = await fetch(path.startsWith('http') ? path : GRAPH + path, Object.assign({}, opts, { headers: Object.assign({ Authorization: 'Bearer ' + tok }, opts.headers || {}) }));
  if (r.status === 401) { LS.del('fin_tok'); await login(); }
  return r;
}

/* ------------------------------------------------------------------ CSV */
function csvParse(text) {
  const rows = []; let row = []; let f = ''; let q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
    else if (c === '"') q = true;
    else if (c === ',') { row.push(f); f = ''; }
    else if (c === '\n' || c === '\r') { if (c === '\r' && text[i + 1] === '\n') i++; row.push(f); rows.push(row); row = []; f = ''; }
    else f += c;
  }
  if (f !== '' || row.length) { row.push(f); rows.push(row); }
  const head = rows.shift() || [];
  return rows.filter((r) => r.some((x) => x !== '')).map((r) => {
    const o = {}; head.forEach((h, i) => { o[h.replace(/^﻿/, '')] = r[i] ?? ''; });
    o.Valor = o.Valor === '' ? null : Number(o.Valor);
    return o;
  });
}
function csvWrite(recs) {
  const cell = (v) => { v = v == null ? '' : String(v); return /[",\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
  const lines = [CSV_HEAD.join(',')];
  for (const r of recs) lines.push(CSV_HEAD.map((h) => cell(h === 'Valor' && r.Valor != null ? Number(r.Valor).toFixed(2) : r[h])).join(','));
  return lines.join('\r\n') + '\r\n';
}

/* ------------------------------------------------------------------ OneDrive */
async function csvMeta() {
  const r = await graph(`/me/drive/items/${cfg().folderId}:/${encodeURIComponent(CSV_NAME)}`);
  if (r.status === 404) return null;
  if (!r.ok) throw new Error('Não consegui ler o ficheiro do iPhone (' + r.status + ')');
  return r.json();
}
async function csvLoad() {
  const meta = await csvMeta();
  if (!meta) return { recs: [], eTag: null };
  const txt = await (await fetch(meta['@microsoft.graph.downloadUrl'], { cache: 'no-store' })).text();
  return { recs: csvParse(txt), eTag: meta.eTag };
}
async function csvSave(mutate) {
  for (let attempt = 0; attempt < 3; attempt++) {
    const { recs, eTag } = await csvLoad();
    const next = mutate(recs.slice());
    const url = `/me/drive/items/${cfg().folderId}:/${encodeURIComponent(CSV_NAME)}:/content` + (eTag ? '' : '?@microsoft.graph.conflictBehavior=fail');
    const r = await graph(url, { method: 'PUT', headers: Object.assign({ 'Content-Type': 'text/csv' }, eTag ? { 'If-Match': eTag } : {}), body: csvWrite(next) });
    if (r.ok) return next;
    if (r.status !== 412 && r.status !== 409) throw new Error('Não consegui guardar (' + r.status + ')');
  }
  throw new Error('O ficheiro estava a ser alterado — tenta outra vez.');
}
async function excelLoad() {
  const r = await graph(`/me/drive/items/${cfg().excelId}`);
  if (!r.ok) throw new Error('Não encontrei o Excel no OneDrive (' + r.status + ')');
  const meta = await r.json();
  const buf = await (await fetch(meta['@microsoft.graph.downloadUrl'], { cache: 'no-store' })).arrayBuffer();
  return { meta, book: await XlsxLite.open(buf) };
}

/* ------------------------------------------------------------------ outbox (registos feitos sem rede) */
const outbox = () => LS.get('fin_outbox', []);
function queue(op) { const o = outbox(); o.push(op); LS.set('fin_outbox', o); }
function applyOps(recs, ops) {
  for (const op of ops) {
    if (op.op === 'add' && !recs.some((r) => r.ID === op.rec.ID)) recs.push(op.rec);
    if (op.op === 'del') recs = recs.filter((r) => r.ID !== op.id);
  }
  return recs;
}
async function flushOutbox() {
  const ops = outbox();
  if (!ops.length || !navigator.onLine) return false;
  await csvSave((recs) => applyOps(recs, ops));
  LS.set('fin_outbox', []);
  return true;
}

/* ------------------------------------------------------------------ modelo a partir do Excel */
async function extract(book) {
  const R = await book.sheet(MAP.resumo), C = await book.sheet(MAP.cfg), K = await book.sheet(MAP.cat);
  const D = await book.sheet(MAP.div), T = await book.sheet(MAP.dias), I = await book.sheet(MAP.iph);
  const v = (s, a) => s[a];
  const m = {
    ano: num(v(R, 'C5')) || new Date().getFullYear(),
    mes: MESES.indexOf(v(R, 'E5')) + 1 || new Date().getMonth() + 1,
    card: { rec: num(v(R, 'B8')), desp: num(v(R, 'C8')), inv: num(v(R, 'E8')), taxa: v(R, 'G8'), saldo: num(v(R, 'I8')), patr: num(v(R, 'K8')), fixas: num(v(R, 'C9')) },
    contas: [], etf: { investido: num(v(R, 'G13')), valor: num(v(R, 'G14')), ganho: num(v(R, 'G15')), rent: v(R, 'G16') },
    div: { devo: num(v(R, 'C23')), devem: num(v(R, 'C24')) },
    trab: { dias: num(v(R, 'G23')), fds: num(v(R, 'G24')), sal: num(v(R, 'G25')), recebido: v(R, 'G27') },
    todo: {}, fluxo: [], planeadas: [], grupos: [],
    cfg: { contas: [], vdia: num(v(C, 'C28')), vfds: num(v(C, 'C29')), csal: v(C, 'C30') || '', cpoup: v(C, 'I19') || '', metaInv: num(v(C, 'C16')) },
    cats: { Receita: [], Despesa: [], Investimento: [] }, catGrupo: {}, catCartao: {},
    dividas: [], dias: {}, imported: new Set(),
  };
  for (let r = 13; r <= 17; r++) if (v(R, 'B' + r)) m.contas.push({ nome: v(R, 'B' + r), saldo: num(v(R, 'C' + r)) });
  for (let r = 7; r <= 11; r++) if (v(C, 'B' + r)) m.cfg.contas.push(v(C, 'B' + r));
  [[23, 'inv'], [24, 'poup'], [25, 'carro'], [26, 'plan'], [27, 'tot'], [28, 'reserva']].forEach(([r, k]) => {
    m.todo[k] = { obj: num(v(R, 'K' + r)), feito: num(v(R, 'L' + r)), falta: num(v(R, 'M' + r)), estado: v(R, 'N' + r) || '' };
  });
  for (let r = 33; r <= 42; r++) m.fluxo.push({ k: String(v(R, 'B' + r) || '').trim(), v: num(v(R, 'C' + r)) });
  for (let r = 33; r <= 40; r++) if (v(R, 'F' + r)) m.planeadas.push({ data: serialToIso(v(R, 'E' + r)), desc: v(R, 'F' + r), valor: num(v(R, 'H' + r)), faltam: v(R, 'I' + r), estado: v(R, 'J' + r) || '' });
  for (let r = MAP.groupsFirstRow; r <= MAP.groupsLastRow; r++) if (v(R, 'B' + r)) m.grupos.push({ nome: v(R, 'B' + r), mes: num(v(R, 'C' + r)), ano: num(v(R, 'E' + r)) });
  for (let r = 6; r <= 25; r++) if (v(K, 'B' + r)) m.cats.Receita.push(v(K, 'B' + r));
  for (let r = 6; r <= 75; r++) if (v(K, 'D' + r)) { const c = v(K, 'D' + r); m.cats.Despesa.push(c); m.catGrupo[c] = v(K, 'E' + r) || 'Outros'; m.catCartao[c] = v(K, 'G' + r) === 'Sim'; }
  for (let r = 6; r <= 15; r++) if (v(K, 'I' + r)) m.cats.Investimento.push(v(K, 'I' + r));
  for (let r = MAP.divFirstRow; r <= MAP.divLastRow; r++) {
    const id = v(D, 'B' + r); if (!id) continue;
    const dsc = v(D, 'E' + r); const descTxt = typeof dsc === 'number' && dsc > 20000 && dsc < 80000 ? serialToIso(dsc).split('-').reverse().join('/') : (dsc ?? '');
    m.dividas.push({ id: String(id), pessoa: String(v(D, 'D' + r) || '').trim(), desc: String(descTxt), tipo: v(D, 'F' + r) || '', valor: num(v(D, 'G' + r)), pago: num(v(D, 'I' + r)), falta: num(v(D, 'J' + r)) });
  }
  for (let r = MAP.diasFirstRow; r <= MAP.diasLastRow; r++) {
    const mi = serialToIso(v(T, 'B' + r)); if (!mi) continue;
    const marks = [];
    for (let d = 1; d <= 31; d++) if (v(T, colLetter(MAP.diasFirstCol - 1 + d) + r) != null) marks.push(d);
    m.dias[mi.slice(0, 7)] = { marks, vdia: v(T, 'C' + r), vfds: v(T, 'D' + r), recebido: v(T, 'I' + r) != null };
  }
  for (const [addr, val] of Object.entries(I)) if (/^A\d+$/.test(addr) && addr !== 'A1' && val != null) m.imported.add(String(val));
  m.imported = Array.from(m.imported);
  return m;
}

/* ------------------------------------------------------------------ estado */
const S = { tab: 'resumo', model: LS.get('fin_model', null), recs: LS.get('fin_recs', []), excelUrl: LS.get('fin_excelUrl', ''), loadedAt: LS.get('fin_loadedAt', null), busy: false, err: null, sheet: null, diasMes: null };
const allRecs = () => applyOps(S.recs.slice(), outbox());
const pendingRecs = () => { const imp = new Set(S.model ? S.model.imported : []); return allRecs().filter((r) => !imp.has(r.ID)); };

async function refresh(silent) {
  if (S.busy) return;
  S.busy = true; S.err = null; if (!silent) render();
  try {
    try { await flushOutbox(); } catch (e) { /* tenta mais tarde */ }
    const [{ meta, book }, csv] = await Promise.all([excelLoad(), csvLoad()]);
    S.model = await extract(book); S.recs = csv.recs; S.excelUrl = meta.webUrl; S.loadedAt = meta.lastModifiedDateTime;
    LS.set('fin_model', S.model); LS.set('fin_recs', S.recs); LS.set('fin_excelUrl', S.excelUrl); LS.set('fin_loadedAt', S.loadedAt);
  } catch (e) {
    S.err = navigator.onLine ? e.message : 'Sem rede — a mostrar os últimos dados guardados.';
  }
  S.busy = false; render();
}

async function addRec(rec) {
  rec = Object.assign({ ID: uid(), Registado: new Date().toISOString().slice(0, 19), Descricao: '', Tipo: '', Categoria: '', Conta: '', Valor: null, Divida: '', Nota: '' }, rec);
  try {
    if (!navigator.onLine) throw new TypeError('offline');
    S.recs = await csvSave((recs) => { recs.push(rec); return recs; });
    LS.set('fin_recs', S.recs);
    toast('Guardado ✓');
  } catch (e) {
    if (e instanceof TypeError) { queue({ op: 'add', rec }); toast('Sem rede — guardado no telemóvel, envio depois.'); }
    else { toast(e.message); throw e; }
  }
  render();
}
async function delRec(id) {
  try {
    if (!navigator.onLine) throw new TypeError('offline');
    S.recs = await csvSave((recs) => recs.filter((r) => r.ID !== id));
    LS.set('fin_recs', S.recs);
  } catch (e) {
    if (e instanceof TypeError) queue({ op: 'del', id }); else { toast(e.message); return; }
  }
  render();
}

/* ------------------------------------------------------------------ cálculos com registos por sincronizar */
function adjusted() {
  const m = JSON.parse(JSON.stringify(S.model));
  const pend = pendingRecs();
  const d = { rec: 0, desp: 0, inv: 0, poup: 0, grupo: {} };
  const acc = (name, delta) => { const c = m.contas.find((x) => x.nome === name); if (c) c.saldo += delta; };
  for (const r of pend) {
    const v = num(r.Valor); const dt = parseIso(r.Data);
    const inMonth = dt.y === m.ano && dt.m === m.mes;
    if (r.Registo === 'Movimento') {
      if (r.Tipo === 'Receita') { acc(r.Conta, v); if (inMonth) d.rec += v; }
      if (r.Tipo === 'Despesa') { acc(r.Conta, -v); if (inMonth) { d.desp += v; const g = m.catGrupo[r.Categoria] || 'Outros'; d.grupo[g] = (d.grupo[g] || 0) + v; } }
      if (r.Tipo === 'Investimento') { acc(r.Conta, -v); if (inMonth) d.inv += v; }
      if (r.Tipo === 'Transferência') { acc(r.Conta, -v); acc(r.Categoria, v); if (inMonth && r.Categoria === m.cfg.cpoup) d.poup += v; }
    }
    if (r.Registo === 'Pagamento') {
      const dv = m.dividas.find((x) => x.id === r.Divida);
      if (dv) { dv.pago += v; dv.falta = Math.max(0, dv.falta - v); }
      if (r.Tipo === 'Devem-me') { acc(r.Conta, v); m.div.devem = Math.max(0, m.div.devem - v); }
      if (r.Tipo === 'Eu devo') { acc(r.Conta, -v); m.div.devo = Math.max(0, m.div.devo - v); }
    }
  }
  // salário: dias marcados no iPhone ainda não importados
  const st = diasState(m.ano, m.mes);
  const salDelta = st.sal - num(S.model.trab.sal);
  if (Math.abs(salDelta) > 0.005) { d.rec += salDelta; if (m.fluxo[0]) m.fluxo[0].v += salDelta; m.trab.dias = st.uteis + st.fds; m.trab.sal = st.sal; }
  const imp = new Set(S.model.imported);
  if (st.salRec && !imp.has(st.salRec.ID) && !st.recXl) { acc(m.cfg.csal, st.sal); m.trab.recebido = '✔ Sim'; }
  else if ((st.recXl || st.salRec) && Math.abs(salDelta) > 0.005) acc(m.cfg.csal, salDelta);
  m.card.rec += d.rec; m.card.desp += d.desp; m.card.inv += d.inv;
  m.card.taxa = m.card.rec ? (m.card.rec - m.card.desp) / m.card.rec : null;
  m.card.saldo = m.contas.reduce((a, c) => a + c.saldo, 0);
  m.card.patr = m.card.saldo + m.etf.valor + m.div.devem - m.div.devo;
  m.sobra = (m.fluxo[9] ? m.fluxo[9].v : 0) + d.rec - d.desp - d.inv - d.poup;
  if (m.fluxo.length >= 10) {
    m.fluxo[1].v += d.rec - salDelta; m.fluxo[2].v += d.rec; m.fluxo[4].v += d.desp; m.fluxo[6].v += d.desp;
    m.fluxo[7].v += d.inv; m.fluxo[8].v += d.poup; m.fluxo[9].v = m.sobra;
  }
  m.todo.inv.feito += d.inv; m.todo.inv.falta = Math.max(0, m.todo.inv.obj - m.todo.inv.feito);
  m.todo.poup.feito += d.poup; m.todo.poup.falta = Math.max(0, m.todo.poup.obj - m.todo.poup.feito);
  for (const g of m.grupos) g.mes += d.grupo[g.nome] || 0;
  m.pendentes = pend.length;
  return m;
}

/* ------------------------------------------------------------------ ecrãs */
function render() {
  const c = cfg();
  if (!c.clientId) return viewSetup();
  if (!LS.get('fin_tok', null)) return viewLogin();
  if (!c.excelId) return viewPick();
  if (!S.model) { $app().innerHTML = `<div class="splash"><div class="logo-mark"></div><p>${S.err ? esc(S.err) : 'A ler o teu Excel…'}</p>${S.err ? '<button class="btn small" data-act="refresh">Tentar outra vez</button>' : ''}</div>`; return; }
  const views = { resumo: viewResumo, registos: viewRegistos, dias: viewDias, dividas: viewDividas };
  $app().innerHTML = views[S.tab]() + tabbar() + (S.sheet ? S.sheet() : '');
}
function tabbar() {
  const t = (id, label, ic) => `<button class="${S.tab === id ? 'on' : ''}" data-tab="${id}" aria-current="${S.tab === id ? 'page' : 'false'}">${ic}${label}</button>`;
  return `<button class="fab mobile-only" data-act="novo" aria-label="Novo movimento">${ICON.plus}</button>
  <nav class="tabbar mobile-only" aria-label="Separadores">${t('resumo', 'Resumo', ICON.grid)}${t('registos', 'Registos', ICON.list)}${t('dias', 'Dias', ICON.check)}${t('dividas', 'Dívidas', ICON.swap)}</nav>`;
}
function statusBanner(m) {
  let h = '';
  if (S.err) h += `<div class="banner err">${ICON.warn}<span>${esc(S.err)}</span></div>`;
  const ob = outbox().length;
  if (ob) h += `<div class="banner warn">${ICON.warn}<span>${ob} registo(s) guardado(s) só no telemóvel — são enviados quando houver rede.</span></div>`;
  if (m && m.pendentes) h += `<div class="banner info">${ICON.sync}<span>${m.pendentes} registo(s) ainda não estão no Excel. Já contam aqui; entram no Excel quando o abrires no PC e atualizares.</span></div>`;
  return h;
}

function viewResumo() {
  const m = adjusted();
  const sobra = m.sobra;
  const tot = m.grupos.reduce((a, g) => a + Math.max(0, g.mes), 0);
  const gs = m.grupos.map((g, i) => ({ ...g, cor: GROUP_COLORS[i % GROUP_COLORS.length], p: tot ? g.mes / tot : 0 })).filter((g) => g.mes > 0).sort((a, b) => b.mes - a.mes);
  const when = S.loadedAt ? new Date(S.loadedAt).toLocaleString('pt-PT', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) : '';
  const todo = (k, label, sub, color, bg) => {
    const t = m.todo[k]; const done = t.obj > 0 && t.falta <= 0.005;
    return `<div class="todo"><div class="ic" style="background:${done ? 'var(--green-l)' : bg};color:${done ? 'var(--green)' : color}">${done ? '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" aria-hidden="true"><path d="M5 12l5 5 9-10"/></svg>' : ICON.chart}</div>
      <div class="t"><div>${esc(label)}</div><div>${eur(t.feito)} de ${eur(t.obj)}${sub ? ' · ' + esc(sub) : ''}</div></div>
      <span class="s" style="color:${done ? 'var(--green)' : 'var(--amber)'}">${t.obj <= 0 ? '—' : done ? 'Feito' : 'Falta ' + eur(t.falta)}</span></div>`;
  };
  const fluxoRows = m.fluxo.map((f, i) => {
    const big = i === 2 || i === 6 || i === 9;
    const val = i === 9 ? sobra : f.v;
    const cls = i === 9 ? (val < 0 ? 'neg' : 'pos') : i <= 1 ? 'pos' : i <= 6 ? 'neg' : 'inv';
    return `<div class="row" style="${big ? 'font-weight:700' : ''}"><span class="${big ? '' : 'muted'}">${esc(f.k)}</span><b class="${big && i !== 9 ? '' : cls}">${eur(val)}</b></div>`;
  }).join('');
  return `<main class="page resumo">
    <div class="top"><div><div class="sub">${MESES[m.mes - 1]} ${m.ano} · mês escolhido no Excel</div><h1>Resumo</h1></div>
      <div style="display:flex;gap:8px">
        <button class="icon-btn" data-act="refresh" aria-label="Atualizar">${S.busy ? '…' : ICON.sync}</button>
        <a class="icon-btn desk-only" href="${esc(S.excelUrl)}" target="_blank" rel="noopener" aria-label="Abrir no Excel">${ICON.ext}</a>
        <button class="icon-btn" data-act="settings" aria-label="Definições">${ICON.gear}</button></div></div>
    ${statusBanner(m)}
    <section class="hero mobile-hero"><div class="lbl">Sobra do mês</div><div class="big">${eur(sobra)}</div>
      <div class="grid"><div class="mini"><div>Receitas</div><div style="color:#8FE0B0">${eur(m.card.rec)}</div></div><div class="mini"><div>Despesas</div><div style="color:#F4A99F">${eur(m.card.desp)}</div></div></div></section>
    <section class="desk" aria-label="Indicadores">
      <div class="k"><div class="l">Receitas</div><div class="v pos">${eur(m.card.rec)}</div><div class="s">${m.trab.dias} dias trabalhados · ${esc(m.trab.recebido || '')}</div></div>
      <div class="k"><div class="l">Despesas</div><div class="v neg">${eur(m.card.desp)}</div><div class="s">das quais fixas: ${eur(m.card.fixas)}</div></div>
      <div class="k"><div class="l">Sobra do mês</div><div class="v">${eur(sobra)}</div><div class="s">taxa de poupança ${pct(m.card.taxa)}</div></div>
      <div class="k dark"><div class="l">Património total</div><div class="v">${eur(m.card.patr)}</div><div class="s">contas + ETFs ± dívidas</div></div>
    </section>
    <div class="cols"><div>
      <section class="card"><h2>A fazer este mês</h2>
        ${todo('inv', 'Investir em ETFs', '', 'var(--amber)', 'var(--amber-l)')}
        ${todo('poup', 'Passar para a ' + (m.cfg.cpoup || 'Conta Poupança'), 'carro ' + eur(m.todo.carro.obj) + ' + planeadas ' + eur(m.todo.plan.obj), 'var(--amber)', 'var(--amber-l)')}
      </section>
      <section class="card"><h2>O mês em números</h2>${fluxoRows}</section>
      <section class="card"><h2>Próximas despesas planeadas</h2>
        ${m.planeadas.length ? m.planeadas.map((p) => { const d = parseIso(p.data || ''); const cls = String(p.estado).startsWith('⚠') ? 'late' : p.estado === 'Em breve' ? 'soon' : '';
          return `<div class="pl ${cls}"><div class="d"><div>${d.d || ''}</div><div>${d.m ? MES3[d.m - 1] : ''}</div></div><span class="x">${esc(p.desc)}</span><b>${eur(p.valor)}</b></div>`; }).join('') : '<p class="muted" style="font-size:14px;margin:0">Nada planeado.</p>'}
      </section>
    </div><div>
      <section class="card"><h2>Para onde vai o dinheiro</h2>
        ${gs.length ? `<div class="stack" role="img" aria-label="${esc(gs.map((g) => g.nome + ' ' + pct(g.p)).join(', '))}">${gs.map((g) => `<div style="width:${(g.p * 100).toFixed(1)}%;background:${g.cor}"></div>`).join('')}</div>
        <div class="legend">${gs.map((g) => `<div><span class="sw" style="background:${g.cor}"></span><span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(g.nome)}</span><b>${pct(g.p)}</b></div>`).join('')}</div>` : '<p class="muted" style="font-size:14px;margin:0">Ainda sem despesas este mês.</p>'}
      </section>
      <section class="card"><h2>Contas</h2>
        ${m.contas.map((c) => `<div class="row"><span>${esc(c.nome)}</span><b class="${c.saldo < 0 ? 'neg' : ''}">${eur(c.saldo)}</b></div>`).join('')}
        <div class="row"><span>ETFs</span><b class="${m.etf.ganho < 0 ? 'neg' : 'pos'}">${eur(m.etf.valor)}${typeof m.etf.rent === 'number' ? ' <span style="font-size:12px">' + (m.etf.rent >= 0 ? '+' : '') + (m.etf.rent * 100).toFixed(1) + '%</span>' : ''}</b></div>
      </section>
      <section class="card"><h2>Dívidas</h2>
        <div class="row"><span>Eu devo</span><b class="${m.div.devo > 0 ? 'neg' : ''}">${eur(m.div.devo)}</b></div>
        <div class="row"><span>Devem-me</span><b class="${m.div.devem > 0 ? 'warn' : ''}">${eur(m.div.devem)}</b></div>
      </section>
      <p class="sub" style="text-align:center">Excel guardado ${esc(when)}</p>
    </div></div>
  </main>`;
}

function viewRegistos() {
  const imp = new Set(S.model.imported); const ob = new Set(outbox().filter((o) => o.op === 'add').map((o) => o.rec.ID));
  const recs = allRecs().slice().sort((a, b) => String(b.Registado).localeCompare(String(a.Registado)));
  const line = (r) => {
    const d = parseIso(r.Data); const when = `${d.d} ${MES3[d.m - 1] || ''}`;
    let title = '', val = '', cls = '';
    if (r.Registo === 'Movimento') { title = r.Descricao || r.Categoria; val = eur(num(r.Valor)); cls = r.Tipo === 'Receita' ? 'pos' : r.Tipo === 'Despesa' ? 'neg' : 'inv'; }
    if (r.Registo === 'Pagamento') { title = 'Pagamento · ' + (r.Descricao || r.Divida); val = eur(num(r.Valor)); }
    if (r.Registo === 'Dia') { title = 'Dia trabalhado'; }
    if (r.Registo === 'Salario') { title = 'Salário recebido · ' + MESES[d.m - 1]; }
    const st = ob.has(r.ID) ? '<span class="warn">no telemóvel</span>' : imp.has(r.ID) ? '<span class="pos">no Excel ✓</span>' : '<span class="muted">por sincronizar</span>';
    return `<div class="row"><div style="min-width:0"><div style="font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(title)}</div>
      <div class="sub">${esc(when)} · ${esc(r.Registo === 'Movimento' ? r.Tipo + ' · ' + r.Conta : r.Conta || '')} · ${st}</div></div>
      <div style="display:flex;align-items:center;gap:8px"><b class="${cls}">${val}</b>
      <button class="icon-btn" style="width:36px;height:36px" data-del="${esc(r.ID)}" aria-label="Apagar registo">×</button></div></div>`;
  };
  return `<main class="page"><div class="top"><div><div class="sub">Feitos nesta app</div><h1>Registos</h1></div>
    <button class="icon-btn" data-act="refresh" aria-label="Atualizar">${S.busy ? '…' : ICON.sync}</button></div>
    ${statusBanner(adjusted())}
    <section class="card">${recs.length ? recs.slice(0, 200).map(line).join('') : '<p class="muted" style="margin:0;font-size:14px">Ainda não registaste nada. Toca no + para começar.</p>'}</section>
    <p class="sub">Apagar aqui também tira o registo do Excel na próxima atualização.</p></main>`;
}

function diasState(y, m) {
  const key = `${y}-${pad2(m)}`;
  const xl = S.model.dias[key] || { marks: [], vdia: null, vfds: null, recebido: false };
  const recs = allRecs();
  const phone = new Set(recs.filter((r) => r.Registo === 'Dia' && String(r.Data).startsWith(key)).map((r) => parseIso(r.Data).d));
  const xlSet = new Set(xl.marks.filter((d) => d <= daysIn(y, m)));
  const all = new Set([...xlSet, ...phone]);
  let uteis = 0, fds = 0; all.forEach((d) => (isWeekend(y, m, d) ? fds++ : uteis++));
  const vd = typeof xl.vdia === 'number' ? xl.vdia : S.model.cfg.vdia;
  const vf = typeof xl.vfds === 'number' ? xl.vfds : S.model.cfg.vfds;
  const salRec = recs.find((r) => r.Registo === 'Salario' && String(r.Data).startsWith(key));
  return { key, xlSet, phone, uteis, fds, vd, vf, sal: uteis * vd + fds * vf, recXl: xl.recebido, salRec };
}
function viewDias() {
  const now = new Date();
  if (!S.diasMes) S.diasMes = { y: now.getFullYear(), m: now.getMonth() + 1 };
  const { y, m } = S.diasMes; const st = diasState(y, m);
  const pend = new Set(pendingRecs().filter((r) => r.Registo === 'Dia').map((r) => r.Data));
  const offset = (weekday(y, m, 1) + 6) % 7;
  let cells = '';
  for (let i = 0; i < offset; i++) cells += '<span></span>';
  for (let d = 1; d <= daysIn(y, m); d++) {
    const on = st.xlSet.has(d) || st.phone.has(d);
    const cls = [st.xlSet.has(d) ? 'lock' : on ? 'on' : isWeekend(y, m, d) ? 'we' : '', (y === now.getFullYear() && m === now.getMonth() + 1 && d === now.getDate()) ? 'today' : '', pend.has(iso(y, m, d)) ? 'pend' : ''].join(' ');
    cells += `<button class="${cls}" data-day="${d}" aria-pressed="${on}" aria-label="${d} de ${MESES[m - 1]}${on ? ', trabalhado' : ''}${st.xlSet.has(d) ? ' (marcado no Excel)' : ''}">${d}</button>`;
  }
  const rec = st.recXl || !!st.salRec;
  return `<main class="page">
    <div class="top"><div><div class="sub">Toca nos dias em que trabalhaste</div><h1>${MESES[m - 1]}${y !== now.getFullYear() ? ' ' + y : ''}</h1></div>
      <div style="display:flex;gap:6px"><button class="icon-btn" data-act="mesAnt" aria-label="Mês anterior">${ICON.left}</button><button class="icon-btn" data-act="mesSeg" aria-label="Mês seguinte">${ICON.right}</button></div></div>
    <section class="card" style="padding:14px 12px"><div class="cal">${['S', 'T', 'Q', 'Q', 'S', 'S', 'D'].map((x) => `<span class="wd">${x}</span>`).join('')}${cells}</div>
      <p class="sub" style="margin:10px 4px 0">Verde escuro = marcado no Excel (muda lá). Ponto amarelo = ainda não está no Excel.</p></section>
    <section class="card">
      <div class="row"><span>Dias úteis</span><span><b>${st.uteis}</b> × ${eur(st.vd)}</span></div>
      <div class="row"><span>Fim de semana</span><span><b>${st.fds}</b> × ${eur(st.vf)}</span></div>
      <div class="row"><span style="font-weight:600">Salário do mês</span><b class="num pos" style="font-size:24px">${eur(st.sal)}</b></div>
    </section>
    <button class="check ${rec ? 'on' : ''}" data-act="recebido" aria-pressed="${rec}">
      <span class="box">${rec ? ICON.tick : ''}</span>
      <span style="flex-grow:1"><span style="display:block;font-weight:700;font-size:15px">Salário já entrou na conta</span>
      <span class="sub" style="display:block">${st.recXl ? 'Marcado no Excel' : rec ? 'Somado ao saldo da ' + esc(S.model.cfg.csal) : 'Marca quando aparecer no banco'}</span></span></button>
  </main>`;
}

function viewDividas() {
  const m = adjusted();
  const abertas = m.dividas.filter((d) => d.falta > 0.005);
  const porPessoa = {};
  for (const d of abertas) {
    const k = String(d.pessoa || '').trim(); if (!k) continue;
    const key = k.toLocaleLowerCase('pt-PT');
    if (!porPessoa[key]) porPessoa[key] = { nome: k, saldo: 0 };
    porPessoa[key].saldo += d.tipo === 'Devem-me' ? d.falta : d.tipo === 'Eu devo' ? -d.falta : 0;
  }
  const pessoas = Object.values(porPessoa).filter((p) => Math.abs(p.saldo) > 0.005).sort((a, b) => Math.abs(b.saldo) - Math.abs(a.saldo));
  const item = (d) => {
    const p = d.valor ? Math.min(1, d.pago / d.valor) : 0; const devo = d.tipo === 'Eu devo';
    return `<div style="padding:12px 0;border-bottom:1px solid var(--line2)">
      <div style="display:flex;justify-content:space-between;gap:10px;align-items:baseline"><div style="min-width:0"><div style="font-weight:600">${esc(d.pessoa)}${d.desc ? ' · ' + esc(d.desc) : ''}</div>
      <div class="sub">${esc(d.tipo)} · ${esc(d.id)}</div></div><b style="color:${devo ? 'var(--red-d)' : '#8A520F'}">${eur(d.falta)}</b></div>
      <div class="prog" style="margin:8px 0 4px"><div style="width:${(p * 100).toFixed(0)}%;background:${devo ? '#C2554A' : '#E0A34A'}"></div></div>
      <div class="sub">Pago ${eur(d.pago)} de ${eur(d.valor)}</div></div>`;
  };
  return `<main class="page"><div class="top"><h1>Dívidas</h1></div>
    ${statusBanner(m)}
    <div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px">
      <div style="background:var(--red-l);border-radius:16px;padding:14px"><div class="sub" style="color:var(--red-d)">Eu devo</div><div class="num" style="font-weight:700;font-size:24px;color:var(--red-d)">${eur(m.div.devo)}</div></div>
      <div style="background:var(--amber-l);border-radius:16px;padding:14px"><div class="sub" style="color:#8A520F">Devem-me</div><div class="num" style="font-weight:700;font-size:24px;color:#8A520F">${eur(m.div.devem)}</div></div></div>
    ${pessoas.length ? `<section class="card"><h2>Por pessoa</h2>${pessoas.map((p) => `<div class="row"><span style="font-weight:600">${esc(p.nome)}</span>
      <b style="color:${p.saldo > 0 ? '#8A520F' : 'var(--red-d)'}">${p.saldo > 0 ? 'deve-te ' + eur(p.saldo) : 'deves ' + eur(-p.saldo)}</b></div>`).join('')}
      <p class="sub" style="margin:8px 0 0">Junta tudo o que deves e te devem a cada pessoa.</p></section>` : ''}
    <section class="card" style="padding:6px 16px"><h2 style="padding-top:10px">Dívidas em aberto</h2>${abertas.length ? abertas.map(item).join('') : '<p class="muted" style="font-size:14px">Sem dívidas em aberto.</p>'}</section>
    ${abertas.length ? '<button class="btn" data-act="pagar">Registar pagamento</button>' : ''}
    <p class="sub">Dívidas novas criam-se no Excel (folha Dívidas).</p></main>`;
}

/* ------------------------------------------------------------------ folhas (novo movimento, pagar, definições) */
const F = { tipo: 'Despesa', valor: '', cat: '', conta: '', data: '', desc: '', div: '', pconta: '' };
function sheetNovo() {
  const m = S.model; const tipos = [['Despesa', 'Despesa'], ['Receita', 'Receita'], ['Investimento', 'Investir'], ['Transferência', 'Transferir']];
  if (!F.conta) F.conta = m.cfg.contas[0] || '';
  const all = F.tipo === 'Transferência' ? m.cfg.contas.filter((c) => c !== F.conta) : (m.cats[F.tipo] || []);
  // atalhos: as mais usadas nesta app; se ainda não há histórico, as primeiras da lista
  const freq = {}; allRecs().forEach((r) => { if (r.Registo === 'Movimento' && r.Tipo === F.tipo) freq[r.Categoria] = (freq[r.Categoria] || 0) + 1; });
  const fav = { Despesa: ['Supermercado', 'Restaurantes / Cafés', 'Combustível', 'Portagens / Estacionamento', 'Saídas / Lazer', 'Farmácia'] }[F.tipo] || [];
  let cats = all.slice().sort((a, b) => (freq[b] || 0) - (freq[a] || 0) || ((fav.indexOf(a) + 1 || 99) - (fav.indexOf(b) + 1 || 99)) || all.indexOf(a) - all.indexOf(b)).slice(0, 8);
  if (!all.includes(F.cat)) F.cat = cats[0] || '';
  if (F.cat && !cats.includes(F.cat)) cats = cats.slice(0, 7).concat(F.cat);
  const cor = { Despesa: 'var(--red)', Receita: 'var(--green)', Investimento: 'var(--blue)', Transferência: 'var(--ink)' }[F.tipo];
  const cartaoMau = F.conta === 'Cartão Refeição' && ((F.tipo === 'Despesa' && !m.catCartao[F.cat]) || F.tipo === 'Investimento' || F.tipo === 'Transferência');
  return `<div class="full" role="dialog" aria-modal="true" aria-label="Novo movimento"><div class="page">
    <div class="navrow"><button class="link" data-act="fechar">Cancelar</button><h1>Novo movimento</h1><span style="width:70px"></span></div>
    <div class="seg" role="group" aria-label="Tipo">${tipos.map(([t, l]) => `<button class="${F.tipo === t ? 'on' : ''}" data-tipo="${t}">${l}</button>`).join('')}</div>
    <div class="amount" style="color:${cor}"><label class="sub" for="valor">Valor (€)</label><input id="valor" inputmode="decimal" autocomplete="off" placeholder="0,00" value="${esc(F.valor)}" data-f="valor"></div>
    <div><div class="sub" style="font-weight:600;margin-bottom:8px">${F.tipo === 'Transferência' ? 'Para a conta' : 'Categoria'}</div>
      <div class="chips">${cats.map((c) => `<button class="chip ${c === F.cat ? 'on' : ''}" data-cat="${esc(c)}">${esc(c)}</button>`).join('')}</div>
      ${all.length > cats.length ? `<label class="field" style="margin-top:10px">Outra categoria<select data-f="catSel"><option value="">Escolher…</option>${all.map((c) => `<option ${c === F.cat ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></label>` : ''}</div>
    <div class="form">
      <label><span>${F.tipo === 'Transferência' ? 'Da conta' : 'Conta'}</span><select data-f="conta">${m.cfg.contas.map((c) => `<option ${c === F.conta ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></label>
      <label><span>Data</span><input type="date" data-f="data" value="${esc(F.data || todayIso())}"></label>
      <label><span>Descrição</span><input type="text" data-f="desc" placeholder="Opcional" value="${esc(F.desc)}"></label>
    </div>
    ${cartaoMau ? `<div class="banner warn" role="status">${ICON.warn}<span>O Cartão Refeição não aceita esta ${F.tipo === 'Despesa' ? 'categoria' : 'operação'}. Confirma a conta.</span></div>` : ''}
    <button class="btn" data-act="guardarMov">Guardar</button>
  </div></div>`;
}
function sheetPagar() {
  const m = adjusted(); const abertas = m.dividas.filter((d) => d.falta > 0.005);
  if (!abertas.some((d) => d.id === F.div)) F.div = abertas[0] ? abertas[0].id : '';
  const sel = abertas.find((d) => d.id === F.div);
  if (F.valor === '' && sel) F.valor = sel.falta.toFixed(2).replace('.', ',');
  return `<div class="sheet-bg" data-act="fechar"></div><section class="sheet" role="dialog" aria-modal="true" aria-label="Registar pagamento">
    <div class="grab"></div><h2 style="font-size:18px">Registar pagamento</h2>
    ${abertas.map((d) => `<label class="radio ${d.id === F.div ? 'on' : ''}"><input type="radio" name="div" data-div="${esc(d.id)}" ${d.id === F.div ? 'checked' : ''}>
      <span style="flex-grow:1"><b>${esc(d.pessoa)}</b>${d.desc ? ' · ' + esc(d.desc) : ''}<br><span class="sub">${esc(d.tipo)}</span></span><b>${eur(d.falta)}</b></label>`).join('')}
    <div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px">
      <label class="field">Valor<input inputmode="decimal" data-f="valor" value="${esc(F.valor)}"></label>
      <label class="field">Conta<select data-f="pconta"><option value="">Nenhuma</option>${m.cfg.contas.map((c) => `<option ${c === F.pconta ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></label>
    </div>
    <button class="btn" data-act="guardarPag">Guardar pagamento</button></section>`;
}
function sheetSettings() {
  const c = cfg();
  return `<div class="sheet-bg" data-act="fechar"></div><section class="sheet" role="dialog" aria-modal="true" aria-label="Definições">
    <div class="grab"></div><h2 style="font-size:18px">Definições</h2>
    <div class="form" style="border:1px solid var(--line)">
      <label><span>Excel</span><span class="sub" style="text-align:right">${esc(c.excelName || '')}</span></label>
      <label><span>Registos do iPhone</span><span class="sub" style="text-align:right">${CSV_NAME} (mesma pasta)</span></label>
    </div>
    <button class="btn ghost" data-act="trocarExcel">Escolher outro Excel</button>
    <button class="btn ghost" data-act="sair">Terminar sessão</button>
    <p class="sub">ID da app: ${esc(c.clientId || '')}<br>Endereço de retorno: ${esc(redirectUri())}</p></section>`;
}

/* ------------------------------------------------------------------ primeira configuração */
function viewSetup() {
  $app().innerHTML = `<main class="page"><div class="top"><h1>Configurar</h1></div>
    <section class="card"><h2>1. ID da aplicação (Microsoft)</h2>
      <p class="sub">Cola o “Application (client) ID” do registo que fizeste no portal da Microsoft.</p>
      <label class="field">Client ID<input id="cid" autocomplete="off" placeholder="00000000-0000-0000-0000-000000000000"></label>
      <p class="sub">No registo, o endereço de retorno (Single-page application) tem de ser exatamente:<br><b style="user-select:all">${esc(redirectUri())}</b></p>
      <button class="btn" data-act="saveCid">Guardar</button></section></main>`;
}
function viewLogin() {
  $app().innerHTML = `<main class="page" style="min-height:90vh;justify-content:center">
    <div style="text-align:center;display:flex;flex-direction:column;align-items:center;gap:12px">
      <div class="logo-mark"></div><h1 class="num" style="font-size:30px">Finanças</h1>
      <p class="sub">Entra com a conta Microsoft onde está o teu Excel.</p></div>
    ${S.err ? `<div class="banner err">${ICON.warn}<span>${esc(S.err)}</span></div>` : ''}
    <button class="btn" data-act="login">Entrar com a conta Microsoft</button></main>`;
}
let pickResults = null; let pickErr = ''; let browse = null; // { id, path, items }
function viewPick() {
  const fileRow = (f) => `<div class="row"><div style="min-width:0"><div style="font-weight:600">${esc(f.name)}</div><div class="sub">${esc((f.parentReference && f.parentReference.path || '').replace(/^\/drive\/root:?/, '') || '/')}</div></div>
      <button class="btn small" data-pick="${esc(f.id)}">Usar</button></div>`;
  $app().innerHTML = `<main class="page"><div class="top"><h1>Escolher o Excel</h1></div>
    ${pickErr ? `<div class="banner err">${ICON.warn}<span>${esc(pickErr)}</span></div>` : ''}
    <section class="card"><h2>Pelas pastas do OneDrive</h2>
      ${browse ? `<div class="sub" style="margin-bottom:6px">📁 ${esc(browse.path || 'OneDrive')}</div>
        ${browse.path ? '<div class="row"><button class="btn small ghost" data-act="browseUp">← Voltar</button><span></span></div>' : ''}
        ${browse.items.map((f) => f.folder
          ? `<div class="row"><span style="font-weight:600">📁 ${esc(f.name)}</span><button class="btn small ghost" data-open="${esc(f.id)}" data-name="${esc(f.name)}">Abrir</button></div>`
          : fileRow(f)).join('') || '<p class="muted" style="font-size:14px">Pasta sem ficheiros .xlsx.</p>'}`
        : '<button class="btn" data-act="browseRoot">Ver pastas</button>'}
    </section>
    <section class="card"><h2>Pelo caminho</h2>
      <label class="field">Caminho dentro do OneDrive<input id="caminho" placeholder="FTP/Financas_Pessoais.xlsx" autocomplete="off"></label>
      <div style="height:10px"></div><button class="btn ghost" data-act="byPath">Abrir</button></section>
    <section class="card"><h2>Pesquisa</h2><label class="field">Nome<input id="q" value="Financas_Pessoais" autocomplete="off"></label>
      <div style="height:10px"></div><button class="btn ghost" data-act="search">Procurar</button></section>
    ${pickResults ? `<section class="card">${pickResults.length ? pickResults.map(fileRow).join('') : '<p class="muted">Nenhum .xlsx encontrado. Usa as pastas ou o caminho.</p>'}</section>` : ''}</main>`;
}
async function openFolder(id, path) {
  pickErr = '';
  const url = id ? `/me/drive/items/${id}/children?$top=200` : '/me/drive/root/children?$top=200';
  const r = await graph(url);
  if (!r.ok) { pickErr = 'Não consegui abrir a pasta (' + r.status + ').'; render(); return; }
  const j = await r.json();
  const items = (j.value || []).filter((f) => f.folder || (f.file && /\.xlsx$/i.test(f.name)))
    .sort((a, b) => (b.folder ? 1 : 0) - (a.folder ? 1 : 0) || a.name.localeCompare(b.name));
  browse = { id, path, items, parent: browse };
  render();
}

/* ------------------------------------------------------------------ eventos */
document.addEventListener('input', (e) => { const f = e.target.dataset && e.target.dataset.f; if (f) F[f] = e.target.value; });
document.addEventListener('change', (e) => {
  const t = e.target;
  if (t.dataset.f === 'catSel') { if (t.value) { F.cat = t.value; render(); } return; }
  if (t.dataset.f) { F[t.dataset.f] = t.value; if (t.dataset.f === 'conta' && S.sheet) render(); }
  if (t.dataset.div) { F.div = t.dataset.div; F.valor = ''; render(); }
});
document.addEventListener('click', async (e) => {
  const el = e.target.closest('[data-act],[data-tab],[data-tipo],[data-cat],[data-day],[data-del],[data-pick],[data-open]');
  if (!el) return;
  const d = el.dataset;
  if (d.tab) { S.tab = d.tab; S.sheet = null; render(); window.scrollTo(0, 0); return; }
  if (d.tipo) { F.tipo = d.tipo; render(); return; }
  if (d.cat) { F.cat = d.cat; render(); return; }
  if (d.del) { if (confirm('Apagar este registo?')) await delRec(d.del); return; }
  if (d.pick) { await choose(d.pick); return; }
  if (d.open) { await openFolder(d.open, (browse && browse.path ? browse.path + '/' : '') + d.name); return; }
  if (d.day) { await toggleDay(+d.day); return; }
  switch (d.act) {
    case 'refresh': await refresh(); break;
    case 'novo': Object.assign(F, { tipo: 'Despesa', valor: '', cat: '', conta: S.model.cfg.contas[0] || '', data: todayIso(), desc: '' }); S.sheet = sheetNovo; render(); setTimeout(() => { const i = document.getElementById('valor'); if (i) i.focus(); }, 50); break;
    case 'pagar': Object.assign(F, { valor: '', div: '', pconta: S.model.cfg.contas[0] || '' }); S.sheet = sheetPagar; render(); break;
    case 'settings': S.sheet = sheetSettings; render(); break;
    case 'fechar': S.sheet = null; render(); break;
    case 'guardarMov': await saveMov(el); break;
    case 'guardarPag': await savePag(el); break;
    case 'mesAnt': case 'mesSeg': { const s = S.diasMes; let m = s.m + (d.act === 'mesAnt' ? -1 : 1), y = s.y; if (m < 1) { m = 12; y--; } if (m > 12) { m = 1; y++; } S.diasMes = { y, m }; render(); break; }
    case 'recebido': await toggleRecebido(); break;
    case 'saveCid': { const v = document.getElementById('cid').value.trim(); if (!/^[0-9a-f-]{36}$/i.test(v)) { toast('Isso não parece um Client ID.'); return; } setCfg({ clientId: v }); render(); break; }
    case 'login': await login('select_account'); break;
    case 'search': await search(document.getElementById('q').value); break;
    case 'browseRoot': await openFolder(null, ''); break;
    case 'browseUp': { const p = browse && browse.parent; browse = p ? p.parent : null; if (p) await openFolder(p.id, p.path); else render(); break; }
    case 'byPath': {
      const p = document.getElementById('caminho').value.trim().replace(/\\/g, '/').replace(/^\/+/, '');
      if (!p) return;
      const r = await graph(`/me/drive/root:/${p.split('/').map(encodeURIComponent).join('/')}`);
      if (!r.ok) { pickErr = 'Não encontrei esse caminho no OneDrive. Confirma o nome da pasta e do ficheiro.'; render(); return; }
      const f = await r.json(); await choose(f.id); break;
    }
    case 'trocarExcel': setCfg({ excelId: null }); S.model = null; S.sheet = null; pickResults = null; render(); break;
    case 'sair': ['fin_tok', 'fin_model', 'fin_recs', 'fin_outbox'].forEach(LS.del); setCfg({ excelId: null }); location.reload(); break;
  }
});

async function saveMov(btn) {
  const v = parseMoney(F.valor);
  if (!(v > 0)) { toast('Escreve um valor maior que zero.'); return; }
  if (!F.cat) { toast('Escolhe a categoria.'); return; }
  btn.disabled = true;
  try {
    await addRec({ Registo: 'Movimento', Data: F.data || todayIso(), Descricao: F.desc.trim(), Tipo: F.tipo, Categoria: F.cat, Conta: F.conta, Valor: v });
    S.sheet = null; render();
  } catch (e) { btn.disabled = false; }
}
async function savePag(btn) {
  const v = parseMoney(F.valor); const dv = S.model.dividas.find((x) => x.id === F.div);
  if (!dv) { toast('Escolhe a dívida.'); return; }
  if (!(v > 0)) { toast('Escreve um valor maior que zero.'); return; }
  btn.disabled = true;
  try {
    await addRec({ Registo: 'Pagamento', Data: todayIso(), Descricao: dv.pessoa + (dv.desc ? ' · ' + dv.desc : ''), Tipo: dv.tipo, Conta: F.pconta, Valor: v, Divida: dv.id });
    S.sheet = null; render();
  } catch (e) { btn.disabled = false; }
}
async function toggleDay(day) {
  const { y, m } = S.diasMes; const st = diasState(y, m);
  if (st.xlSet.has(day)) { toast('Este dia está marcado no Excel — muda-o lá.'); return; }
  const date = iso(y, m, day);
  const rec = allRecs().find((r) => r.Registo === 'Dia' && r.Data === date);
  if (rec) await delRec(rec.ID); else await addRec({ Registo: 'Dia', Data: date, Valor: 1 });
}
async function toggleRecebido() {
  const { y, m } = S.diasMes; const st = diasState(y, m);
  if (st.recXl) { toast('Está marcado no Excel — muda-o lá.'); return; }
  if (st.salRec) await delRec(st.salRec.ID); else await addRec({ Registo: 'Salario', Data: iso(y, m, 1), Conta: S.model.cfg.csal, Valor: st.sal });
}
async function search(q) {
  pickErr = '';
  const r = await graph(`/me/drive/root/search(q='${encodeURIComponent(String(q || 'xlsx').replace(/'/g, "''"))}')?$top=50`);
  if (!r.ok) { pickErr = 'A pesquisa do OneDrive falhou (' + r.status + '). Usa as pastas ou o caminho.'; pickResults = null; render(); return; }
  const j = await r.json();
  pickResults = (j.value || []).filter((f) => f.file && /\.xlsx$/i.test(f.name));
  render();
}
async function choose(id) {
  const r = await graph(`/me/drive/items/${id}`); const f = await r.json();
  setCfg({ excelId: f.id, excelName: f.name, folderId: f.parentReference.id });
  S.model = null; render(); await refresh(true);
}

/* ------------------------------------------------------------------ arranque */
(async function start() {
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
  window.addEventListener('online', () => refresh(true));
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && cfg().excelId && LS.get('fin_tok', null)) refresh(true); });
  try { await handleRedirect(); } catch (e) { S.err = e.message; }
  render();
  if (cfg().clientId && LS.get('fin_tok', null) && cfg().excelId) refresh(!!S.model);
})();
