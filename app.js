// Finanças — app pessoal ligada ao Excel no OneDrive.
// Todos os registos (movimentos, planeadas, dívidas, dias) vivem em Financas_iPhone.csv, na pasta do Excel.
// Do Excel a app só lê a configuração (contas, categorias, fixos, dias marcados, ETFs) e faz as contas sozinha.
// O Excel importa o CSV pela folha "Dados" (Power Query) e faz as mesmas contas no PC.
'use strict';

/* ------------------------------------------------------------------ config */
const AUTH = 'https://login.microsoftonline.com/consumers/oauth2/v2.0';
const GRAPH = 'https://graph.microsoft.com/v1.0';
const SCOPES = 'openid profile offline_access User.Read Files.ReadWrite';
const CSV_NAME = 'Financas_iPhone.csv';
const CSV_BACKUP = 'Financas_iPhone_backup.csv';
const CSV_HEAD = ['ID', 'Registado', 'Registo', 'Data', 'Descricao', 'Tipo', 'Categoria', 'Conta', 'Valor', 'Divida', 'Nota'];
const MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
const MES3 = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
const GROUP_COLORS = ['#1F6F66', '#E0A34A', '#6C8FD1', '#E76F51', '#8E7CC3', '#2A9D8F', '#F4A261', '#A3B18A', '#B9C3BF', '#C0392B', '#7F8C8D', '#34495E'];
const SHEETS = { cfg: 'Configuração', cat: 'Categorias', fix: 'Fixos', dias: 'Dias de Trabalho', inv: 'Investimentos' };

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
const pct = (v) => (typeof v === 'number' && isFinite(v) ? Math.round(v * 100) + '%' : '—');
const num = (v) => (typeof v === 'number' && isFinite(v) ? v : 0);
const r2 = (v) => Math.round(v * 100) / 100;
const pad2 = (n) => String(n).padStart(2, '0');
const iso = (y, m, d) => `${y}-${pad2(m)}-${pad2(d)}`;
const todayIso = () => { const t = new Date(); return iso(t.getFullYear(), t.getMonth() + 1, t.getDate()); };
const parseIso = (s) => { const [y, m, d] = String(s || '').slice(0, 10).split('-').map(Number); return { y, m, d }; };
const isIso = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || '').slice(0, 10));
const idxOf = (y, m) => y * 12 + (m - 1);                       // mês → número sequencial
const idxIso = (s) => { const p = parseIso(s); return idxOf(p.y, p.m); };
const idxStart = (i) => iso(Math.floor(i / 12), (i % 12) + 1, 1);
const curIdx = () => { const t = new Date(); return idxOf(t.getFullYear(), t.getMonth() + 1); };
const fmtDate = (s) => { const p = parseIso(s); return p.d ? `${p.d} ${MES3[p.m - 1]}` : ''; };
function serialToIso(n) { // número de série do Excel → AAAA-MM-DD
  if (typeof n !== 'number') return null;
  const dt = new Date(Math.round((n - 25569) * 86400000));
  return iso(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
}
const daysIn = (y, m) => new Date(y, m, 0).getDate();
const weekday = (y, m, d) => new Date(y, m - 1, d).getDay(); // 0 dom … 6 sáb
const isWeekend = (y, m, d) => { const w = weekday(y, m, d); return w === 0 || w === 6; };
const daysUntil = (s) => { const p = parseIso(s); const t = new Date(); return Math.round((new Date(p.y, p.m - 1, p.d) - new Date(t.getFullYear(), t.getMonth(), t.getDate())) / 86400000); };
function parseMoney(s) {
  s = String(s || '').replace(/[€\s]/g, '');
  if (s.includes(',') && s.includes('.')) s = s.replace(/\./g, '').replace(',', '.');
  else s = s.replace(',', '.');
  const v = Number(s);
  return isFinite(v) ? Math.round(v * 100) / 100 : NaN;
}
const money = (v) => (v == null || v === '' ? '' : Number(v).toFixed(2).replace('.', ','));
function colLetter(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
function uid() { return 'i' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }
const pkey = (s) => String(s || '').trim().toLocaleLowerCase('pt-PT');
function toast(msg) { const t = document.getElementById('toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove('show'), 2600); }
const ICON = {
  plus: '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.6" aria-hidden="true"><path d="M12 5v14M5 12h14"/></svg>',
  grid: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><rect x="3" y="3" width="7" height="9" rx="1.5"/><rect x="14" y="3" width="7" height="5" rx="1.5"/><rect x="14" y="12" width="7" height="9" rx="1.5"/><rect x="3" y="16" width="7" height="5" rx="1.5"/></svg>',
  list: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M4 6h16M4 12h16M4 18h10"/></svg>',
  cal: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/></svg>',
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
const itemPath = (name) => `/me/drive/items/${cfg().folderId}:/${encodeURIComponent(name)}`;
async function csvMeta() {
  const r = await graph(itemPath(CSV_NAME));
  if (r.status === 404) return null;
  if (!r.ok) throw new Error('Não consegui ler o ficheiro de registos (' + r.status + ')');
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
    const url = itemPath(CSV_NAME) + ':/content' + (eTag ? '' : '?@microsoft.graph.conflictBehavior=fail');
    const r = await graph(url, { method: 'PUT', headers: Object.assign({ 'Content-Type': 'text/csv' }, eTag ? { 'If-Match': eTag } : {}), body: csvWrite(next) });
    if (r.ok) return next;
    if (r.status !== 412 && r.status !== 409) throw new Error('Não consegui guardar (' + r.status + ')');
  }
  throw new Error('O ficheiro estava a ser alterado — tenta outra vez.');
}
async function backupDaily(recs) { // uma cópia por dia, ao lado do CSV
  if (!recs.length || LS.get('fin_backup', '') === todayIso()) return;
  const r = await graph(itemPath(CSV_BACKUP) + ':/content', { method: 'PUT', headers: { 'Content-Type': 'text/csv' }, body: csvWrite(recs) });
  if (r.ok) LS.set('fin_backup', todayIso());
}

/* ------------------------------------------------------------------ alterações (com fila para quando não há rede) */
const outbox = () => LS.get('fin_outbox', []);
function cascadeIds(recs, id) {
  // apagar uma despesa dividida apaga a dívida ligada (e vice-versa) e os pagamentos dessa dívida
  const ids = new Set([id]);
  const r = recs.find((x) => x.ID === id);
  if (r && r.Registo === 'Movimento' && r.Divida) ids.add(r.Divida);
  if (r && r.Registo === 'Partilha') recs.filter((x) => x.Registo === 'Movimento' && x.Divida === id).forEach((x) => ids.add(x.ID));
  return ids;
}
function applyOps(recs, ops) {
  for (const op of ops) {
    if (op.op === 'add' && !recs.some((r) => r.ID === op.rec.ID)) recs.push(op.rec);
    if (op.op === 'upd') recs = recs.map((r) => (r.ID === op.id ? Object.assign({}, r, op.patch) : r));
    if (op.op === 'del') {
      const ids = cascadeIds(recs, op.id);
      recs = recs.filter((r) => !ids.has(r.ID) && !(r.Registo === 'Pagamento' && ids.has(r.Divida)));
    }
  }
  return recs;
}
async function flushOutbox() {
  const ops = outbox();
  if (!ops.length || !navigator.onLine) return false;
  S.recs = await csvSave((recs) => applyOps(recs, ops));
  LS.set('fin_recs', S.recs); LS.set('fin_outbox', []);
  return true;
}
function newRec(rec) {
  return Object.assign({ ID: uid(), Registado: new Date().toISOString().slice(0, 19), Data: todayIso(), Descricao: '', Tipo: '', Categoria: '', Conta: '', Valor: null, Divida: '', Nota: '' }, rec);
}
async function commit(ops, okMsg) {
  try {
    if (!navigator.onLine) throw new TypeError('offline');
    S.recs = await csvSave((recs) => applyOps(recs, ops));
    LS.set('fin_recs', S.recs);
    if (okMsg !== false) toast(okMsg || 'Guardado ✓');
  } catch (e) {
    if (e instanceof TypeError) { LS.set('fin_outbox', outbox().concat(ops)); toast('Sem rede — guardado no telemóvel, envio depois.'); }
    else { toast(e.message); throw e; }
  }
  render();
}
const allRecs = () => applyOps(S.recs.slice(), outbox());

/* ------------------------------------------------------------------ configuração lida do Excel */
async function excelLoad(force) {
  const r = await graph(`/me/drive/items/${cfg().excelId}`);
  if (!r.ok) throw new Error('Não encontrei o Excel no OneDrive (' + r.status + ')');
  const meta = await r.json();
  if (!force && S.conf && S.conf.eTag === meta.eTag) return { meta, conf: S.conf };
  const buf = await (await fetch(meta['@microsoft.graph.downloadUrl'], { cache: 'no-store' })).arrayBuffer();
  const conf = await readConf(await XlsxLite.open(buf));
  conf.eTag = meta.eTag; conf.savedAt = meta.lastModifiedDateTime;
  return { meta, conf };
}
async function readConf(book) {
  const C = await book.sheet(SHEETS.cfg), K = await book.sheet(SHEETS.cat), X = await book.sheet(SHEETS.fix);
  const T = await book.sheet(SHEETS.dias), I = await book.sheet(SHEETS.inv);
  const s = (o, a) => (o[a] == null ? '' : String(o[a]).trim());
  const c = {
    contas: [], vdia: num(C.C28), vfds: num(C.C29), csal: s(C, 'C30'), subDia: num(C.C31), csub: s(C, 'C32'),
    carro: { total: 0, mensal: 0, conta: s(C, 'I19'), inicio: serialToIso(C.I20) || '' }, metaInv: 0, grupos: [],
    cats: { Receita: [], Despesa: [], Investimento: [] }, catGrupo: {}, catCartao: {},
    fixos: [], dias: {}, etf: { lista: [], ops: [] },
  };
  for (let r = 7; r <= 11; r++) if (s(C, 'B' + r)) c.contas.push({ nome: s(C, 'B' + r), saldo0: num(C['C' + r]), data0: serialToIso(C['D' + r]) || '' });
  for (let r = 7; r <= 16; r++) c.carro.total += num(C['I' + r]);
  c.carro.mensal = r2(c.carro.total / 12);
  for (let r = 16; r <= 23; r++) if (s(C, 'E' + r) === 'Investido no mês') { c.metaInv = num(C['C' + r]); break; }
  for (let r = 36; r <= 47; r++) if (s(C, 'B' + r)) c.grupos.push(s(C, 'B' + r));
  for (let r = 6; r <= 25; r++) if (s(K, 'B' + r)) c.cats.Receita.push(s(K, 'B' + r));
  for (let r = 6; r <= 75; r++) if (s(K, 'D' + r)) { const n = s(K, 'D' + r); c.cats.Despesa.push(n); c.catGrupo[n] = s(K, 'E' + r) || 'Outros'; c.catCartao[n] = s(K, 'G' + r) === 'Sim'; }
  for (let r = 6; r <= 15; r++) if (s(K, 'I' + r)) c.cats.Investimento.push(s(K, 'I' + r));
  for (let r = 6; r <= 55; r++) {
    if (!s(X, 'B' + r) || !num(X['F' + r]) || typeof X['I' + r] !== 'number') continue;
    c.fixos.push({ desc: s(X, 'B' + r), tipo: s(X, 'C' + r), cat: s(X, 'D' + r), conta: s(X, 'E' + r), valor: num(X['F' + r]), varia: num(X['G' + r]),
      dia: Math.min(28, Math.max(1, num(X['H' + r]) || 1)), ini: serialToIso(X['I' + r]), fim: serialToIso(X['J' + r]) || '',
      w: { Anual: 12, Semestral: 6, Trimestral: 3 }[s(X, 'L' + r)] || 1 });
  }
  for (let r = 7; r <= 66; r++) {
    const mi = serialToIso(T['B' + r]); if (!mi) continue;
    const marks = [];
    for (let d = 1; d <= 31; d++) if (T[colLetter(9 + d) + r] != null) marks.push(d);
    c.dias[mi.slice(0, 7)] = { marks, vdia: typeof T['C' + r] === 'number' ? T['C' + r] : null, vfds: typeof T['D' + r] === 'number' ? T['D' + r] : null };
  }
  for (let r = 6; r <= 15; r++) if (s(I, 'B' + r)) c.etf.lista.push({ t: s(I, 'B' + r), preco: typeof I['E' + r] === 'number' ? I['E' + r] : null });
  for (let r = 21; r <= 1020; r++) {
    if (!s(I, 'C' + r) || typeof I['E' + r] !== 'number' || typeof I['F' + r] !== 'number') continue;
    c.etf.ops.push({ t: s(I, 'C' + r), op: s(I, 'D' + r), u: I['E' + r], p: I['F' + r], fee: num(I['G' + r]) });
  }
  return c;
}

/* ------------------------------------------------------------------ contas (as mesmas regras do Excel) */
function fixoAt(f, i) { // valor do fixo no mês i (0 se não calhar)
  const ini = parseIso(f.ini);
  const m1 = idxOf(ini.y, ini.m) + (ini.d > f.dia ? 1 : 0);
  const last = f.fim ? idxIso(f.fim) : Infinity;
  if (i < m1 || i > last || (i - m1) % f.w) return 0;
  return f.valor + f.varia * (i - m1) / f.w;
}
function fixoDebitado(f, desde) { // total já debitado até hoje, depois da data do saldo inicial
  const ini = parseIso(f.ini); const m1 = idxOf(ini.y, ini.m) + (ini.d > f.dia ? 1 : 0);
  const today = todayIso(); let tot = 0;
  for (let i = m1; i <= curIdx(); i += f.w) {
    const d = iso(Math.floor(i / 12), (i % 12) + 1, f.dia);
    if (f.fim && d > f.fim) break;
    if (d > (desde || '') && d <= today) tot += fixoAt(f, i);
  }
  return tot;
}
function diasMes(key, recs) { // dias de trabalho de um mês (Excel + app)
  const { y, m } = parseIso(key + '-01'); const conf = S.conf;
  const xl = conf.dias[key] || { marks: [], vdia: null, vfds: null };
  const phone = new Set(recs.filter((r) => r.Registo === 'Dia' && String(r.Data).startsWith(key)).map((r) => parseIso(r.Data).d));
  const xlSet = new Set(xl.marks.filter((d) => d <= daysIn(y, m)));
  const all = new Set([...xlSet, ...phone]);
  let uteis = 0, fds = 0; all.forEach((d) => (isWeekend(y, m, d) ? fds++ : uteis++));
  const vd = xl.vdia != null ? xl.vdia : conf.vdia; const vf = xl.vfds != null ? xl.vfds : conf.vfds;
  const salRec = recs.find((r) => r.Registo === 'Salario' && String(r.Data).startsWith(key));
  const nextIdx = idxOf(y, m) + 1; const payday = idxStart(nextIdx);
  return { key, y, m, xlSet, phone, uteis, fds, vd, vf, sal: uteis * vd + fds * vf, sub: (uteis + fds) * conf.subDia, payday, salRec, recebido: payday <= todayIso() || !!salRec };
}
function diasKeys(recs) {
  const ks = new Set(Object.keys(S.conf.dias));
  recs.forEach((r) => { if (r.Registo === 'Dia' || r.Registo === 'Salario') ks.add(String(r.Data).slice(0, 7)); });
  return Array.from(ks);
}

function dividas(recs) {
  const pagos = {};
  recs.filter((r) => r.Registo === 'Pagamento' && r.Divida).forEach((p) => { pagos[p.Divida] = (pagos[p.Divida] || 0) + num(p.Valor); });
  return recs.filter((r) => r.Registo === 'Divida' || r.Registo === 'Partilha').map((r) => {
    const pago = r2(pagos[r.ID] || 0);
    return { id: r.ID, rec: r, pessoa: String(r.Categoria || '').trim(), desc: r.Descricao || '', tipo: r.Tipo, valor: num(r.Valor), pago,
      falta: Math.max(0, r2(num(r.Valor) - pago)), prazo: r.Registo === 'Divida' && isIso(r.Divida) ? r.Divida : '', data: r.Data, partilha: r.Registo === 'Partilha', conta: r.Conta };
  });
}
function porPessoa(divs) {
  const P = {};
  for (const d of divs) {
    if (d.falta <= 0.005 || !d.pessoa) continue;
    const k = pkey(d.pessoa);
    if (!P[k]) P[k] = { nome: d.pessoa, devo: 0, devem: 0 };
    if (d.tipo === 'Eu devo') P[k].devo += d.falta; else P[k].devem += d.falta;
  }
  return Object.values(P).map((p) => Object.assign(p, { saldo: r2(p.devem - p.devo) })).sort((a, b) => Math.abs(b.saldo) - Math.abs(a.saldo));
}

function calc(Y, M) {
  const conf = S.conf; const recs = allRecs(); const today = todayIso();
  const i = idxOf(Y, M); const mS = idxStart(i), mE = idxStart(i + 1); const ci = curIdx();
  const inM = (r) => r.Data >= mS && r.Data < mE;
  const caixa = conf.carro.conta;
  const carroM = !conf.carro.inicio || mS >= conf.carro.inicio.slice(0, 8) + '01' ? conf.carro.mensal : 0; // reserva só a partir do mês de início
  const sum = (arr) => arr.reduce((a, r) => a + num(r.Valor), 0);
  const is = (reg, tipo) => (r) => r.Registo === reg && (!tipo || r.Tipo === tipo);
  const fixSum = (tipo, idx = i, pred = () => true) => conf.fixos.filter((f) => f.tipo === tipo && pred(f)).reduce((a, f) => a + fixoAt(f, idx), 0);

  // saldo das contas
  const dk = diasKeys(recs).map((k) => diasMes(k, recs));
  const contas = conf.contas.map((c) => {
    const d0 = c.data0 || ''; let s = c.saldo0;
    for (const r of recs) {
      if (!(r.Data >= d0)) continue;
      const v = num(r.Valor);
      if (r.Registo === 'Movimento') {
        if (r.Conta === c.nome) s += r.Tipo === 'Receita' ? v : -v;
        if (r.Tipo === 'Transferência' && r.Categoria === c.nome) s += v;
      }
      if ((r.Registo === 'Divida' || r.Registo === 'Partilha') && r.Conta === c.nome) s += r.Tipo === 'Eu devo' ? v : -v;
      if (r.Registo === 'Pagamento' && r.Conta === c.nome) s += r.Tipo === 'Devem-me' ? v : -v;
      if (r.Registo === 'Planeada' && r.Tipo === 'Paga' && r.Conta === c.nome) s -= v;
    }
    for (const f of conf.fixos) {
      if (f.conta === c.nome) s += (f.tipo === 'Receita' ? 1 : -1) * fixoDebitado(f, d0);
      if (f.tipo === 'Transferência' && f.cat === c.nome) { const o = conf.contas.find((x) => x.nome === f.conta); s += fixoDebitado(f, o ? o.data0 : ''); }
    }
    for (const d of dk) {
      if (!d.recebido || !(d.payday > d0)) continue;
      if (c.nome === conf.csal) s += d.sal;
      if (c.nome === conf.csub) s += d.sub;
    }
    return { nome: c.nome, saldo: r2(s) };
  });

  // mês escolhido
  const prev = diasMes(idxStart(i - 1).slice(0, 7), recs);
  const salPrev = prev.sal + prev.sub;
  const movM = recs.filter((r) => r.Registo === 'Movimento' && inM(r));
  const recMov = sum(movM.filter((r) => r.Tipo === 'Receita'));
  const despMov = sum(movM.filter((r) => r.Tipo === 'Despesa'));
  const invMov = sum(movM.filter((r) => r.Tipo === 'Investimento'));
  const planM = recs.filter((r) => r.Registo === 'Planeada' && inM(r));
  const fxRec = fixSum('Receita'), fxDesp = fixSum('Despesa'), fxInv = fixSum('Investimento');
  const naoCaixa = (r) => !caixa || r.Conta !== caixa;

  // dívidas no mês: pagamentos, dívidas em que o dinheiro mexeu, partilhas e prazos que vencem
  const divs = dividas(recs);
  const proj = (tipo) => {
    if (i < ci) return 0;
    const lo = i === ci ? '' : mS;
    return divs.filter((d) => d.tipo === tipo && d.prazo && d.prazo >= lo && d.prazo < mE).reduce((a, d) => a + d.falta, 0);
  };
  const criadas = (tipo) => sum(recs.filter((r) => inM(r) && r.Tipo === tipo && ((r.Registo === 'Divida' && r.Conta) || r.Registo === 'Partilha')));
  const pagM = (tipo) => sum(recs.filter((r) => is('Pagamento', tipo)(r) && inM(r)));
  const divSai = pagM('Eu devo') + criadas('Devem-me') + proj('Eu devo');
  const divEntra = pagM('Devem-me') + criadas('Eu devo') + proj('Devem-me');

  const f = {
    salPrev, outras: recMov + fxRec, fixas: fxDesp,
    variaveis: sum(movM.filter((r) => r.Tipo === 'Despesa' && naoCaixa(r))),
    planeadas: sum(planM.filter(naoCaixa)),
    dividas: divSai - divEntra, carro: carroM,
  };
  f.receitas = f.salPrev + f.outras;
  f.sai = f.fixas + f.variaveis + f.planeadas + f.dividas + f.carro;
  f.sobra = f.receitas - f.sai;

  const card = { rec: f.receitas, desp: despMov + fxDesp + sum(planM.filter(is('Planeada', 'Paga'))), inv: invMov + fxInv };
  card.taxa = card.rec ? (card.rec - card.desp) / card.rec : null;
  card.saldo = contas.reduce((a, c) => a + c.saldo, 0);

  // para onde vai o dinheiro
  const g = {}; const add = (cat, v) => { const k = conf.catGrupo[cat] || 'Outros'; g[k] = (g[k] || 0) + v; };
  movM.filter((r) => r.Tipo === 'Despesa').forEach((r) => add(r.Categoria, num(r.Valor)));
  planM.filter((r) => r.Tipo === 'Paga').forEach((r) => add(r.Categoria, num(r.Valor)));
  conf.fixos.filter((x) => x.tipo === 'Despesa').forEach((x) => add(x.cat, fixoAt(x, i)));
  const nomes = conf.grupos.length ? conf.grupos.concat(Object.keys(g).filter((k) => !conf.grupos.includes(k))) : Object.keys(g);
  const grupos = nomes.map((n) => ({ nome: n, mes: g[n] || 0 }));

  // a pôr de parte
  const paraCaixa = sum(movM.filter((r) => r.Tipo === 'Transferência' && r.Categoria === caixa)) + fixSum('Transferência', i, (x) => x.cat === caixa);
  const invFeito = card.inv;
  const todo = {
    inv: { obj: conf.metaInv, feito: invFeito, falta: Math.max(0, conf.metaInv - invFeito) },
    caixa: { obj: carroM, feito: paraCaixa, falta: Math.max(0, carroM - paraCaixa) },
  };
  const depoisEtf = f.sobra - Math.max(conf.metaInv, invFeito);

  // planeadas por pagar (todas, por data)
  const porPagar = recs.filter(is('Planeada', 'Por pagar')).sort((a, b) => String(a.Data).localeCompare(String(b.Data)));

  // ETFs (custo médio)
  const etf = { custo: 0, valor: 0, ganho: 0, realizado: 0 };
  for (const e of conf.etf.lista) {
    const ops = conf.etf.ops.filter((o) => o.t === e.t);
    const bU = ops.filter((o) => o.op === 'Compra').reduce((a, o) => a + o.u, 0);
    const bT = ops.filter((o) => o.op === 'Compra').reduce((a, o) => a + o.u * o.p + o.fee, 0);
    const sU = ops.filter((o) => o.op === 'Venda').reduce((a, o) => a + o.u, 0);
    const sT = ops.filter((o) => o.op === 'Venda').reduce((a, o) => a + o.u * o.p - o.fee, 0);
    const custo = bU ? bT - sU * bT / bU : 0;
    etf.custo += custo; etf.realizado += bU ? sT - sU * bT / bU : 0;
    if (e.preco != null) etf.valor += (bU - sU) * e.preco;
  }
  etf.ganho = etf.valor - etf.custo; etf.rent = etf.custo > 0 ? etf.ganho / etf.custo : null;

  const divTot = { devo: divs.filter((d) => d.tipo === 'Eu devo').reduce((a, d) => a + d.falta, 0), devem: divs.filter((d) => d.tipo === 'Devem-me').reduce((a, d) => a + d.falta, 0) };
  return { Y, M, contas, f, card, grupos, todo, depoisEtf, porPagar, etf, divs, divTot, pessoas: porPessoa(divs), today, trab: prev };
}

/* ------------------------------------------------------------------ estado */
const now0 = new Date();
const S = {
  tab: 'resumo', conf: LS.get('fin_conf', null), recs: LS.get('fin_recs', []), excelUrl: LS.get('fin_excelUrl', ''),
  busy: false, err: null, sheet: null, mes: { y: now0.getFullYear(), m: now0.getMonth() + 1 }, diasMes: null, planVer: 'abertas',
};

async function refresh(silent) {
  if (S.busy) return;
  S.busy = true; S.err = null; if (!silent) render();
  try {
    try { await flushOutbox(); } catch (e) { /* tenta mais tarde */ }
    const [{ meta, conf }, csv] = await Promise.all([excelLoad(), csvLoad()]);
    S.conf = conf; S.recs = csv.recs; S.excelUrl = meta.webUrl;
    LS.set('fin_conf', S.conf); LS.set('fin_recs', S.recs); LS.set('fin_excelUrl', S.excelUrl);
    backupDaily(S.recs).catch(() => {});
  } catch (e) {
    S.err = navigator.onLine ? e.message : 'Sem rede — a mostrar os últimos dados guardados.';
  }
  S.busy = false; render();
}

/* ------------------------------------------------------------------ ecrãs */
function render() {
  const c = cfg();
  if (!c.clientId) return viewSetup();
  if (!LS.get('fin_tok', null)) return viewLogin();
  if (!c.excelId) return viewPick();
  if (!S.conf) { $app().innerHTML = `<div class="splash"><div class="logo-mark"></div><p>${S.err ? esc(S.err) : 'A ler o teu Excel…'}</p>${S.err ? '<button class="btn small" data-act="refresh">Tentar outra vez</button>' : ''}</div>`; return; }
  const views = { resumo: viewResumo, registos: viewRegistos, planeadas: viewPlaneadas, dividas: viewDividas, dias: viewDias };
  const y = window.scrollY;
  $app().innerHTML = views[S.tab]() + tabbar() + (S.sheet ? S.sheet() : '');
  if (S.sheet) window.scrollTo(0, y);
}
function tabbar() {
  const t = (id, label, ic) => `<button class="${S.tab === id ? 'on' : ''}" data-tab="${id}" aria-current="${S.tab === id ? 'page' : 'false'}">${ic}${label}</button>`;
  return `<button class="fab" data-act="novo" aria-label="Novo movimento">${ICON.plus}</button>
  <nav class="tabbar" aria-label="Separadores">${t('resumo', 'Resumo', ICON.grid)}${t('registos', 'Registos', ICON.list)}${t('planeadas', 'Planeadas', ICON.cal)}${t('dividas', 'Dívidas', ICON.swap)}${t('dias', 'Dias', ICON.check)}</nav>`;
}
function statusBanner() {
  let h = '';
  if (S.err) h += `<div class="banner err">${ICON.warn}<span>${esc(S.err)}</span></div>`;
  const ob = outbox().length;
  if (ob) h += `<div class="banner warn">${ICON.warn}<span>${ob} alteração(ões) guardada(s) só no telemóvel — são enviadas quando houver rede.</span></div>`;
  return h;
}
const syncBtn = () => `<button class="icon-btn" data-act="refresh" aria-label="Atualizar">${S.busy ? '…' : ICON.sync}</button>`;
function planRow(r) {
  const dd = daysUntil(r.Data); const cls = r.Tipo === 'Paga' ? '' : dd < 0 ? 'late' : dd <= 30 ? 'soon' : '';
  const p = parseIso(r.Data);
  const sub = r.Tipo === 'Paga' ? 'paga' : dd < 0 ? `atrasada ${-dd} dia(s)` : dd === 0 ? 'hoje' : `daqui a ${dd} dia(s)`;
  return `<button class="pl ${cls}" data-edit="${esc(r.ID)}"><div class="d"><div>${p.d || ''}</div><div>${p.m ? MES3[p.m - 1] : ''}</div></div>
    <span class="x"><span style="display:block;overflow:hidden;text-overflow:ellipsis">${esc(r.Descricao || r.Categoria)}</span><span class="sub">${esc(sub)}</span></span><b>${eur(num(r.Valor))}</b></button>`;
}

function viewResumo() {
  const m = calc(S.mes.y, S.mes.m);
  const tot = m.grupos.reduce((a, g) => a + Math.max(0, g.mes), 0);
  const gs = m.grupos.map((g, i) => ({ ...g, cor: GROUP_COLORS[i % GROUP_COLORS.length], p: tot ? g.mes / tot : 0 })).filter((g) => g.mes > 0).sort((a, b) => b.mes - a.mes);
  const todo = (t, label, sub) => {
    const done = t.obj > 0 && t.falta <= 0.005;
    return `<div class="todo"><div class="ic" style="background:${done ? 'var(--green-l)' : 'var(--amber-l)'};color:${done ? 'var(--green)' : 'var(--amber)'}">${done ? '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" aria-hidden="true"><path d="M5 12l5 5 9-10"/></svg>' : ICON.chart}</div>
      <div class="t"><div>${esc(label)}</div><div>${eur(t.feito)} de ${eur(t.obj)}${sub ? ' · ' + esc(sub) : ''}</div></div>
      <span class="s" style="color:${done ? 'var(--green)' : 'var(--amber)'}">${t.obj <= 0 ? '—' : done ? 'Feito' : 'Falta ' + eur(t.falta)}</span></div>`;
  };
  const f = m.f;
  const rows = [
    ['Salário (dias do mês anterior)', f.salPrev, 'pos'], ['Outras receitas', f.outras, 'pos'], ['Receitas', f.receitas, '', 1],
    ['Despesas fixas', f.fixas, 'neg'], ['Despesas variáveis', f.variaveis, 'neg'], ['Despesas planeadas', f.planeadas, 'neg'],
    [f.dividas < 0 ? 'Dívidas (entra mais do que sai)' : 'Dívidas (sai − entra)', f.dividas, f.dividas < 0 ? 'pos' : 'neg'],
    ['Reserva do carro', f.carro, 'neg'], ['Total que sai', f.sai, '', 1], ['O que me sobra', f.sobra, f.sobra < 0 ? 'neg' : 'pos', 1],
  ].map(([k, v, cls, big]) => `<div class="row" style="${big ? 'font-weight:700' : ''}"><span class="${big ? '' : 'muted'}">${esc(k)}</span><b class="${cls}">${eur(v)}</b></div>`).join('');
  const prox = m.porPagar.slice(0, 6);
  const isCur = S.mes.y === now0.getFullYear() && S.mes.m === now0.getMonth() + 1;
  return `<main class="page resumo">
    <div class="top"><div><div class="sub">Resumo</div>
      <div class="mnav"><button class="icon-btn" data-act="resAnt" aria-label="Mês anterior">${ICON.left}</button><h1>${MESES[S.mes.m - 1]}${S.mes.y !== now0.getFullYear() ? ' ' + S.mes.y : ''}</h1><button class="icon-btn" data-act="resSeg" aria-label="Mês seguinte">${ICON.right}</button></div></div>
      <div style="display:flex;gap:8px">${syncBtn()}
        <a class="icon-btn desk-only" href="${esc(S.excelUrl)}" target="_blank" rel="noopener" aria-label="Abrir no Excel">${ICON.ext}</a>
        <button class="icon-btn" data-act="settings" aria-label="Definições">${ICON.gear}</button></div></div>
    ${!isCur ? '<button class="btn small ghost" data-act="resHoje" style="align-self:flex-start">Voltar ao mês atual</button>' : ''}
    ${statusBanner()}
    <section class="hero mobile-hero"><div class="lbl">O que me sobra</div><div class="big">${eur(f.sobra)}</div>
      <div class="grid"><div class="mini"><div>Entra</div><div style="color:#8FE0B0">${eur(f.receitas)}</div></div><div class="mini"><div>Sai</div><div style="color:#F4A99F">${eur(f.sai)}</div></div></div>
      <div class="lbl" style="margin-top:12px">Depois dos ETFs: <b style="color:#fff">${eur(m.depoisEtf)}</b></div></section>
    <section class="desk" aria-label="Indicadores">
      <div class="k"><div class="l">Entra</div><div class="v pos">${eur(f.receitas)}</div><div class="s">${m.trab.uteis + m.trab.fds} dias trabalhados no mês anterior</div></div>
      <div class="k"><div class="l">Sai</div><div class="v neg">${eur(f.sai)}</div><div class="s">fixas ${eur(f.fixas)} · planeadas ${eur(f.planeadas)}</div></div>
      <div class="k"><div class="l">O que me sobra</div><div class="v">${eur(f.sobra)}</div><div class="s">depois dos ETFs: ${eur(m.depoisEtf)}</div></div>
      <div class="k dark"><div class="l">Saldo nas contas</div><div class="v">${eur(m.card.saldo)}</div><div class="s">hoje</div></div>
    </section>
    <div class="cols"><div>
      <section class="card"><h2>O mês em números</h2>${rows}</section>
      <section class="card"><h2>A pôr de parte</h2>
        ${todo(m.todo.inv, 'Investir em ETFs', 'fica fora do que sobra')}
        ${todo(m.todo.caixa, 'Passar para a ' + (S.conf.carro.conta || 'Conta Caixa'), 'reserva do carro')}
      </section>
      <section class="card"><h2>Próximas despesas planeadas</h2>
        ${prox.length ? prox.map(planRow).join('') : '<p class="muted" style="font-size:14px;margin:0">Nada planeado.</p>'}
      </section>
    </div><div>
      <section class="card"><h2>Contas (saldo hoje)</h2>
        ${m.contas.map((c) => `<div class="row"><span>${esc(c.nome)}</span><b class="${c.saldo < 0 ? 'neg' : ''}">${eur(c.saldo)}</b></div>`).join('')}
        ${S.conf.etf.lista.length ? `<div class="row"><span>ETFs</span><b class="${m.etf.ganho < 0 ? 'neg' : 'pos'}">${eur(m.etf.valor)}${typeof m.etf.rent === 'number' ? ' <span style="font-size:12px">' + (m.etf.rent >= 0 ? '+' : '') + (m.etf.rent * 100).toFixed(1) + '%</span>' : ''}</b></div>` : ''}
      </section>
      <section class="card"><h2>Para onde vai o dinheiro</h2>
        ${gs.length ? `<div class="stack" role="img" aria-label="${esc(gs.map((g) => g.nome + ' ' + pct(g.p)).join(', '))}">${gs.map((g) => `<div style="width:${(g.p * 100).toFixed(1)}%;background:${g.cor}"></div>`).join('')}</div>
        <div class="legend">${gs.map((g) => `<div><span class="sw" style="background:${g.cor}"></span><span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(g.nome)}</span><b>${eur(g.mes)}</b></div>`).join('')}</div>` : '<p class="muted" style="font-size:14px;margin:0">Ainda sem despesas neste mês.</p>'}
      </section>
      <section class="card"><h2>Dívidas</h2>
        <div class="row"><span>Eu devo</span><b class="${m.divTot.devo > 0 ? 'neg' : ''}">${eur(m.divTot.devo)}</b></div>
        <div class="row"><span>Devem-me</span><b class="${m.divTot.devem > 0 ? 'warn' : ''}">${eur(m.divTot.devem)}</b></div>
      </section>
      <p class="sub" style="text-align:center">Configuração lida do Excel guardado a ${esc(S.conf.savedAt ? new Date(S.conf.savedAt).toLocaleString('pt-PT', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) : '—')}</p>
    </div></div>
  </main>`;
}

function recTitle(r) {
  if (r.Registo === 'Movimento') return (r.Descricao || r.Categoria) + (r.Divida ? ' · dividida' : '');
  if (r.Registo === 'Planeada') return 'Planeada · ' + (r.Descricao || r.Categoria);
  if (r.Registo === 'Pagamento') return (r.Tipo === 'Eu devo' ? 'Paguei a ' : 'Recebi de ') + (r.Categoria || r.Descricao || '');
  if (r.Registo === 'Divida') return (r.Tipo === 'Eu devo' ? 'Devo a ' : 'Deve-me: ') + (r.Categoria || '') + (r.Descricao ? ' · ' + r.Descricao : '');
  if (r.Registo === 'Partilha') return (r.Tipo === 'Eu devo' ? 'Devo a ' : 'Deve-me: ') + (r.Categoria || '') + ' · parte de ' + (r.Descricao || 'despesa');
  if (r.Registo === 'Dia') return 'Dia trabalhado';
  if (r.Registo === 'Salario') return 'Salário recebido · ' + MESES[parseIso(r.Data).m - 1];
  return r.Registo;
}
function viewRegistos() {
  const ob = new Set(outbox().filter((o) => o.op === 'add').map((o) => o.rec.ID));
  const all = allRecs().filter((r) => S.regFiltro === 'todos' || !['Dia', 'Salario'].includes(r.Registo));
  const recs = all.slice().sort((a, b) => String(b.Data).localeCompare(String(a.Data)) || String(b.Registado).localeCompare(String(a.Registado)));
  const line = (r) => {
    let val = '', cls = '';
    if (r.Valor != null && !['Dia', 'Salario'].includes(r.Registo)) val = eur(num(r.Valor));
    if (r.Registo === 'Movimento') cls = r.Tipo === 'Receita' ? 'pos' : r.Tipo === 'Despesa' ? 'neg' : 'inv';
    if (r.Registo === 'Divida' || r.Registo === 'Partilha') cls = r.Tipo === 'Eu devo' ? 'neg' : 'warn';
    if (r.Registo === 'Planeada') cls = r.Tipo === 'Paga' ? 'neg' : 'muted';
    const det = r.Registo === 'Movimento' ? r.Tipo + (r.Conta ? ' · ' + r.Conta : '') : r.Registo === 'Planeada' ? r.Tipo : r.Conta || '';
    return `<div class="row"><button class="rowbtn" data-edit="${esc(r.ID)}"><div style="font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(recTitle(r))}</div>
      <div class="sub">${esc(fmtDate(r.Data))}${det ? ' · ' + esc(det) : ''}${ob.has(r.ID) ? ' · <span class="warn">no telemóvel</span>' : ''}</div></button>
      <div style="display:flex;align-items:center;gap:8px"><b class="${cls}">${val}</b>
      <button class="icon-btn" style="width:36px;height:36px" data-del="${esc(r.ID)}" aria-label="Apagar registo">×</button></div></div>`;
  };
  return `<main class="page"><div class="top"><div><div class="sub">Toca num registo para o editar</div><h1>Registos</h1></div>${syncBtn()}</div>
    ${statusBanner()}
    <div class="seg" style="grid-template-columns:repeat(2,minmax(0,1fr))"><button class="${S.regFiltro !== 'todos' ? 'on' : ''}" data-act="regDinheiro">Dinheiro</button><button class="${S.regFiltro === 'todos' ? 'on' : ''}" data-act="regTodos">Tudo (com dias)</button></div>
    <section class="card">${recs.length ? recs.slice(0, 300).map(line).join('') : '<p class="muted" style="margin:0;font-size:14px">Ainda não registaste nada. Toca no + para começar.</p>'}</section>
    <p class="sub">Apagar aqui também tira o registo do Excel na próxima atualização. Apagar uma despesa dividida apaga também a dívida ligada.</p></main>`;
}

function viewPlaneadas() {
  const recs = allRecs().filter((r) => r.Registo === 'Planeada');
  const abertas = recs.filter((r) => r.Tipo !== 'Paga').sort((a, b) => String(a.Data).localeCompare(String(b.Data)));
  const pagas = recs.filter((r) => r.Tipo === 'Paga').sort((a, b) => String(b.Data).localeCompare(String(a.Data)));
  const total = abertas.reduce((a, r) => a + num(r.Valor), 0);
  const lista = S.planVer === 'pagas' ? pagas : abertas;
  return `<main class="page"><div class="top"><div><div class="sub">Despesas futuras e pontuais</div><h1>Planeadas</h1></div>${syncBtn()}</div>
    ${statusBanner()}
    <div style="background:var(--amber-l);border-radius:16px;padding:14px"><div class="sub" style="color:#8A520F">Por pagar</div><div class="num" style="font-weight:700;font-size:24px;color:#8A520F">${eur(total)}</div><div class="sub" style="color:#8A520F">${abertas.length} planeada(s)</div></div>
    <div class="seg" style="grid-template-columns:repeat(2,minmax(0,1fr))"><button class="${S.planVer !== 'pagas' ? 'on' : ''}" data-act="planAbertas">Por pagar</button><button class="${S.planVer === 'pagas' ? 'on' : ''}" data-act="planPagas">Pagas</button></div>
    <section class="card">${lista.length ? lista.map(planRow).join('') : `<p class="muted" style="margin:0;font-size:14px">${S.planVer === 'pagas' ? 'Ainda nenhuma paga.' : 'Nada por pagar.'}</p>`}</section>
    <button class="btn" data-act="novaPlan">Nova planeada</button></main>`;
}

function diasState(y, m) { return diasMes(`${y}-${pad2(m)}`, allRecs()); }
function viewDias() {
  if (!S.diasMes) S.diasMes = { y: now0.getFullYear(), m: now0.getMonth() + 1 };
  const { y, m } = S.diasMes; const st = diasState(y, m); const nw = new Date();
  const offset = (weekday(y, m, 1) + 6) % 7;
  let cells = '';
  for (let i = 0; i < offset; i++) cells += '<span></span>';
  for (let d = 1; d <= daysIn(y, m); d++) {
    const on = st.xlSet.has(d) || st.phone.has(d);
    const cls = [st.xlSet.has(d) ? 'lock' : on ? 'on' : isWeekend(y, m, d) ? 'we' : '', (y === nw.getFullYear() && m === nw.getMonth() + 1 && d === nw.getDate()) ? 'today' : ''].join(' ');
    cells += `<button class="${cls}" data-day="${d}" aria-pressed="${on}" aria-label="${d} de ${MESES[m - 1]}${on ? ', trabalhado' : ''}${st.xlSet.has(d) ? ' (marcado no Excel)' : ''}">${d}</button>`;
  }
  const rec = !!st.salRec;
  return `<main class="page">
    <div class="top"><div><div class="sub">Toca nos dias em que trabalhaste</div><h1>${MESES[m - 1]}${y !== nw.getFullYear() ? ' ' + y : ''}</h1></div>
      <div style="display:flex;gap:6px"><button class="icon-btn" data-act="mesAnt" aria-label="Mês anterior">${ICON.left}</button><button class="icon-btn" data-act="mesSeg" aria-label="Mês seguinte">${ICON.right}</button></div></div>
    <section class="card" style="padding:14px 12px"><div class="cal">${['S', 'T', 'Q', 'Q', 'S', 'S', 'D'].map((x) => `<span class="wd">${x}</span>`).join('')}${cells}</div>
      <p class="sub" style="margin:10px 4px 0">Verde escuro = marcado no Excel (muda lá).</p></section>
    <section class="card">
      <div class="row"><span>Dias úteis</span><span><b>${st.uteis}</b> × ${eur(st.vd)}</span></div>
      <div class="row"><span>Fim de semana</span><span><b>${st.fds}</b> × ${eur(st.vf)}</span></div>
      <div class="row"><span style="font-weight:600">Salário destes dias</span><b class="num pos" style="font-size:24px">${eur(st.sal)}</b></div>
      <p class="sub" style="margin:8px 0 0">Entra no saldo a ${esc(fmtDate(st.payday))}${st.payday <= todayIso() ? ' (já entrou)' : ''}. Conta no «O que me sobra» de ${MESES[parseIso(st.payday).m - 1]}.</p>
    </section>
    <button class="check ${rec ? 'on' : ''}" data-act="recebido" aria-pressed="${rec}">
      <span class="box">${rec ? ICON.tick : ''}</span>
      <span style="flex-grow:1"><span style="display:block;font-weight:700;font-size:15px">Salário já entrou na conta</span>
      <span class="sub" style="display:block">${rec ? 'Somado ao saldo da ' + esc(S.conf.csal) : 'Marca se entrar antes de ' + esc(fmtDate(st.payday))}</span></span></button>
  </main>`;
}

function viewDividas() {
  const m = calc(S.mes.y, S.mes.m);
  const abertas = m.divs.filter((d) => d.falta > 0.005).sort((a, b) => String(a.prazo || '9999').localeCompare(String(b.prazo || '9999')) || String(a.data).localeCompare(String(b.data)));
  const fechadas = m.divs.filter((d) => d.falta <= 0.005).sort((a, b) => String(b.data).localeCompare(String(a.data))).slice(0, 15);
  const item = (d) => {
    const p = d.valor ? Math.min(1, d.pago / d.valor) : 0; const devo = d.tipo === 'Eu devo';
    const prazo = d.prazo ? (daysUntil(d.prazo) < 0 && d.falta > 0 ? ` · <span class="neg">prazo ${esc(fmtDate(d.prazo))} (passou)</span>` : ` · prazo ${esc(fmtDate(d.prazo))}`) : '';
    return `<button class="debt" data-edit="${esc(d.id)}">
      <div style="display:flex;justify-content:space-between;gap:10px;align-items:baseline"><div style="min-width:0;text-align:left"><div style="font-weight:600">${esc(d.pessoa)}${d.desc ? ' · ' + esc(d.desc) : ''}</div>
      <div class="sub">${devo ? 'Eu devo' : 'Deve-me'}${d.partilha ? ' · despesa dividida' : ''} · ${esc(fmtDate(d.data))}${prazo}</div></div><b style="color:${devo ? 'var(--red-d)' : '#8A520F'}">${eur(d.falta)}</b></div>
      <div class="prog" style="margin:8px 0 4px"><div style="width:${(p * 100).toFixed(0)}%;background:${devo ? '#C2554A' : '#E0A34A'}"></div></div>
      <div class="sub" style="text-align:left">Pago ${eur(d.pago)} de ${eur(d.valor)}</div></button>`;
  };
  return `<main class="page"><div class="top"><div><div class="sub">Como no Splitwise</div><h1>Dívidas</h1></div>${syncBtn()}</div>
    ${statusBanner()}
    <div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px">
      <div style="background:var(--red-l);border-radius:16px;padding:14px"><div class="sub" style="color:var(--red-d)">Eu devo</div><div class="num" style="font-weight:700;font-size:24px;color:var(--red-d)">${eur(m.divTot.devo)}</div></div>
      <div style="background:var(--amber-l);border-radius:16px;padding:14px"><div class="sub" style="color:#8A520F">Devem-me</div><div class="num" style="font-weight:700;font-size:24px;color:#8A520F">${eur(m.divTot.devem)}</div></div></div>
    ${m.pessoas.length ? `<section class="card"><h2>Por pessoa</h2>${m.pessoas.map((p) => `<div class="row"><div><div style="font-weight:600">${esc(p.nome)}</div>
      <div class="sub">${p.devo && p.devem ? `deves ${eur(p.devo)} · deve-te ${eur(p.devem)}` : ''}</div></div>
      <div style="display:flex;align-items:center;gap:8px"><b style="color:${p.saldo > 0 ? '#8A520F' : 'var(--red-d)'}">${p.saldo > 0 ? 'deve-te ' + eur(p.saldo) : p.saldo < 0 ? 'deves ' + eur(-p.saldo) : 'acertado'}</b>
      <button class="btn small" data-pagar="${esc(p.nome)}">Pagar</button></div></div>`).join('')}</section>` : ''}
    <section class="card" style="padding:6px 16px"><h2 style="padding-top:10px">Em aberto</h2>${abertas.length ? abertas.map(item).join('') : '<p class="muted" style="font-size:14px">Sem dívidas em aberto.</p>'}</section>
    <button class="btn ghost" data-act="novaDiv">Nova dívida</button>
    ${fechadas.length ? `<section class="card" style="padding:6px 16px"><h2 style="padding-top:10px">Já pagas</h2>${fechadas.map(item).join('')}</section>` : ''}
  </main>`;
}

/* ------------------------------------------------------------------ folhas */
const F = {};
const pessoasConhecidas = () => Array.from(new Set(allRecs().filter((r) => ['Divida', 'Partilha', 'Pagamento'].includes(r.Registo)).map((r) => String(r.Categoria || '').trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'pt'));
const datalist = () => `<datalist id="pessoas">${pessoasConhecidas().map((n) => `<option value="${esc(n)}">`).join('')}</datalist>`;
const contaOpts = (sel, nenhuma) => (nenhuma ? `<option value="">${esc(nenhuma)}</option>` : '') + S.conf.contas.map((c) => `<option ${c.nome === sel ? 'selected' : ''}>${esc(c.nome)}</option>`).join('');
const sheetWrap = (title, body, del) => `<div class="sheet-bg" data-act="fechar"></div><section class="sheet" role="dialog" aria-modal="true" aria-label="${esc(title)}">
  <div class="grab"></div><div class="sheet-head"><h2 style="font-size:18px">${esc(title)}</h2><button class="link-btn" data-act="fechar">Cancelar</button></div>${body}
  ${del ? `<button class="btn ghost" style="color:var(--red)" data-del="${esc(del)}">Apagar</button>` : ''}</section>`;

function splitResumo() {
  if (!(F.tipo === 'Despesa' && F.split)) return '';
  const tot = parseMoney(F.valor); const minha = parseMoney(F.minha); const quem = esc(String(F.spessoa || '').trim() || '…');
  if (!(tot > 0) || !(minha >= 0)) return '';
  if (minha > tot) return '<span class="neg">A tua parte é maior do que o total.</span>';
  return F.squem === 'outro' ? `Fica: despesa tua ${eur(minha)} e deves ${eur(minha)} a ${quem}. A conta não mexe.`
    : `Fica: saem ${eur(tot)} da ${esc(F.conta)}, despesa tua ${eur(minha)} e ${quem} deve-te ${eur(r2(tot - minha))}.`;
}
function sheetNovo() {
  const conf = S.conf; const tipos = [['Despesa', 'Despesa'], ['Receita', 'Receita'], ['Investimento', 'Investir'], ['Transferência', 'Transferir']];
  const contas = conf.contas.map((c) => c.nome);
  if (!F.conta) F.conta = contas[0] || '';
  const all = F.tipo === 'Transferência' ? contas.filter((c) => c !== F.conta) : (conf.cats[F.tipo] || []);
  const freq = {}; allRecs().forEach((r) => { if (r.Registo === 'Movimento' && r.Tipo === F.tipo) freq[r.Categoria] = (freq[r.Categoria] || 0) + 1; });
  const fav = { Despesa: ['Supermercado', 'Restaurantes / Cafés', 'Combustível', 'Portagens / Estacionamento', 'Saídas / Lazer', 'Farmácia'] }[F.tipo] || [];
  let cats = all.slice().sort((a, b) => (freq[b] || 0) - (freq[a] || 0) || ((fav.indexOf(a) + 1 || 99) - (fav.indexOf(b) + 1 || 99)) || all.indexOf(a) - all.indexOf(b)).slice(0, 8);
  if (!all.includes(F.cat)) F.cat = cats[0] || '';
  if (F.cat && !cats.includes(F.cat)) cats = cats.slice(0, 7).concat(F.cat);
  const cor = { Despesa: 'var(--red)', Receita: 'var(--green)', Investimento: 'var(--blue)', Transferência: 'var(--ink)' }[F.tipo];
  const split = F.tipo === 'Despesa' && F.split;
  const outroPagou = split && F.squem === 'outro';
  const cartaoMau = !outroPagou && F.conta === 'Cartão Refeição' && ((F.tipo === 'Despesa' && !conf.catCartao[F.cat]) || F.tipo === 'Investimento' || F.tipo === 'Transferência');
  return `<div class="full" role="dialog" aria-modal="true" aria-label="${F.editId ? 'Editar movimento' : 'Novo movimento'}"><div class="page">
    <div class="navrow"><button class="link" data-act="fechar">Cancelar</button><h1>${F.editId ? 'Editar movimento' : 'Novo movimento'}</h1><span style="width:70px"></span></div>
    <div class="seg" role="group" aria-label="Tipo">${tipos.map(([t, l]) => `<button class="${F.tipo === t ? 'on' : ''}" data-tipo="${t}">${l}</button>`).join('')}</div>
    <div class="amount" style="color:${cor}"><label class="sub" for="valor">${split ? 'Valor total da conta (€)' : 'Valor (€)'}</label><input id="valor" inputmode="decimal" autocomplete="off" placeholder="0,00" value="${esc(F.valor)}" data-f="valor"></div>
    <div><div class="sub" style="font-weight:600;margin-bottom:8px">${F.tipo === 'Transferência' ? 'Para a conta' : 'Categoria'}</div>
      <div class="chips">${cats.map((c) => `<button class="chip ${c === F.cat ? 'on' : ''}" data-cat="${esc(c)}">${esc(c)}</button>`).join('')}</div>
      ${all.length > cats.length ? `<label class="field" style="margin-top:10px">Outra categoria<select data-f="catSel"><option value="">Escolher…</option>${all.map((c) => `<option ${c === F.cat ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></label>` : ''}</div>
    ${F.tipo === 'Despesa' ? `<button class="check ${split ? 'on' : ''}" data-act="toggleSplit" aria-pressed="${!!split}"><span class="box">${split ? ICON.tick : ''}</span>
      <span style="flex-grow:1"><span style="display:block;font-weight:700;font-size:15px">Dividir com alguém</span><span class="sub" style="display:block">A tua parte conta como despesa; a outra fica como dívida</span></span></button>` : ''}
    ${split ? `<div class="splitbox">
      <label class="field">Com quem<input data-f="spessoa" list="pessoas" autocomplete="off" value="${esc(F.spessoa)}" placeholder="Nome"></label>${datalist()}
      <div class="seg" style="grid-template-columns:repeat(2,minmax(0,1fr))"><button class="${!outroPagou ? 'on' : ''}" data-squem="eu">Paguei eu</button><button class="${outroPagou ? 'on' : ''}" data-squem="outro">Pagou a outra pessoa</button></div>
      <div style="display:grid;grid-template-columns:minmax(0,1fr) auto;gap:10px;align-items:end"><label class="field" style="min-width:0">A minha parte (€)<input inputmode="decimal" data-f="minha" value="${esc(F.minha)}" placeholder="0,00"></label>
      <button class="btn small ghost" data-act="metade">Metade</button></div>
      <p class="sub" id="splitResumo" style="margin:0">${splitResumo()}</p></div>` : ''}
    <div class="form">
      ${outroPagou ? '' : `<label><span>${F.tipo === 'Transferência' ? 'Da conta' : 'Conta'}</span><select data-f="conta">${contaOpts(F.conta)}</select></label>`}
      <label><span>Data</span><input type="date" data-f="data" value="${esc(F.data || todayIso())}"></label>
      <label><span>Descrição</span><input type="text" data-f="desc" placeholder="Opcional" value="${esc(F.desc)}"></label>
    </div>
    ${cartaoMau ? `<div class="banner warn" role="status">${ICON.warn}<span>O Cartão Refeição não aceita esta ${F.tipo === 'Despesa' ? 'categoria' : 'operação'}. Confirma a conta.</span></div>` : ''}
    <button class="btn" data-act="guardarMov">Guardar</button>
    ${F.editId ? `<button class="btn ghost" style="color:var(--red)" data-del="${esc(F.editId)}">Apagar</button>` : ''}
  </div></div>`;
}

function sheetPlan() {
  const conf = S.conf; const paga = F.ptipo === 'Paga';
  return sheetWrap(F.editId ? 'Despesa planeada' : 'Nova planeada', `
    <label class="field">Descrição<input data-f="desc" autocomplete="off" value="${esc(F.desc)}" placeholder="Ex.: revisão do carro"></label>
    <div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px">
      <label class="field">Valor (€)<input inputmode="decimal" data-f="valor" value="${esc(F.valor)}" placeholder="0,00"></label>
      <label class="field">${paga ? 'Data do pagamento' : 'Data prevista'}<input type="date" data-f="data" value="${esc(F.data)}"></label></div>
    <label class="field">Categoria<select data-f="cat">${conf.cats.Despesa.map((c) => `<option ${c === F.cat ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></label>
    <label class="field">Conta<select data-f="conta">${contaOpts(F.conta)}</select></label>
    ${paga && F.prevista ? `<p class="sub" style="margin:0">Estava prevista para ${esc(fmtDate(F.prevista))}.</p>` : ''}
    <button class="btn" data-act="guardarPlan">Guardar</button>
    ${F.editId ? `<button class="btn ghost" data-act="${paga ? 'planDesmarcar' : 'planPagar'}">${paga ? 'Voltar a «por pagar»' : 'Marcar como paga hoje'}</button>` : ''}`, F.editId);
}

function sheetNovaDiv() {
  const part = F.dreg === 'Partilha';
  return sheetWrap(F.editId ? (part ? 'Dívida de despesa dividida' : 'Editar dívida') : 'Nova dívida', `
    ${part ? '<p class="sub" style="margin:0">Esta dívida veio de uma despesa dividida. Para mudar valores, edita a despesa em Registos.</p>' : `
    <div class="seg" style="grid-template-columns:repeat(2,minmax(0,1fr))" role="group" aria-label="Tipo">
      <button class="${F.dtipo === 'Eu devo' ? 'on' : ''}" data-dtipo="Eu devo">Eu devo</button>
      <button class="${F.dtipo === 'Devem-me' ? 'on' : ''}" data-dtipo="Devem-me">Devem-me</button></div>`}
    <label class="field">Pessoa<input data-f="pessoa" list="pessoas" autocomplete="off" value="${esc(F.pessoa)}" placeholder="Nome"></label>${datalist()}
    <label class="field">Descrição<input data-f="desc" autocomplete="off" value="${esc(F.desc)}" placeholder="Opcional (ex.: jantar)"></label>
    ${part ? '' : `<div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px">
      <label class="field">Valor<input inputmode="decimal" data-f="valor" value="${esc(F.valor)}" placeholder="0,00"></label>
      <label class="field">Conta<select data-f="pconta">${contaOpts(F.pconta, 'Nenhuma')}</select></label></div>
    <div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px">
      <label class="field">Data<input type="date" data-f="data" value="${esc(F.data)}"></label>
      <label class="field">Prazo (opcional)<input type="date" data-f="prazo" value="${esc(F.prazo)}"></label></div>
    <p class="sub" style="margin:0">Conta: só se o dinheiro ${F.dtipo === 'Eu devo' ? 'entrou nela (alguém te emprestou)' : 'saiu dela (emprestaste)'}; se alguém pagou por ti, deixa «Nenhuma». Prazo: a parte que faltar entra no «O que me sobra» desse mês.</p>`}
    <button class="btn" data-act="guardarDiv">Guardar</button>`, F.editId && !part ? F.editId : null);
}

function sheetPagar() {
  const m = calc(S.mes.y, S.mes.m);
  const p = m.pessoas.find((x) => pkey(x.nome) === pkey(F.ppessoa));
  if (!p) return sheetWrap('Registar pagamento', '<p class="muted">Sem dívidas em aberto com esta pessoa.</p>');
  if (!F.ptipo || (F.ptipo === 'Eu devo' && !p.devo) || (F.ptipo === 'Devem-me' && !p.devem)) F.ptipo = p.devo >= p.devem ? 'Eu devo' : 'Devem-me';
  const max = F.ptipo === 'Eu devo' ? p.devo : p.devem;
  if (F.valor === '') F.valor = money(max);
  return sheetWrap('Pagamento · ' + p.nome, `
    ${p.devo && p.devem ? `<div class="seg" style="grid-template-columns:repeat(2,minmax(0,1fr))"><button class="${F.ptipo === 'Eu devo' ? 'on' : ''}" data-ptipo="Eu devo">Eu pago (devo ${eur(p.devo)})</button><button class="${F.ptipo === 'Devem-me' ? 'on' : ''}" data-ptipo="Devem-me">Recebo (${eur(p.devem)})</button></div>` : ''}
    <p class="sub" style="margin:0">${F.ptipo === 'Eu devo' ? `Deves ${eur(p.devo)} a ${esc(p.nome)}.` : `${esc(p.nome)} deve-te ${eur(p.devem)}.`} Podes pagar só uma parte — o resto fica em dívida.</p>
    <div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px">
      <label class="field">Valor<input inputmode="decimal" data-f="valor" value="${esc(F.valor)}"></label>
      <label class="field">Conta<select data-f="pconta">${contaOpts(F.pconta, 'Nenhuma')}</select></label></div>
    <label class="field">Data<input type="date" data-f="data" value="${esc(F.data)}"></label>
    <button class="btn" data-act="guardarPag">Guardar pagamento</button>`);
}

function sheetEditPag() {
  return sheetWrap('Pagamento', `
    <p class="sub" style="margin:0">${esc(F.titulo)}</p>
    <div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px">
      <label class="field">Valor<input inputmode="decimal" data-f="valor" value="${esc(F.valor)}"></label>
      <label class="field">Conta<select data-f="pconta">${contaOpts(F.pconta, 'Nenhuma')}</select></label></div>
    <label class="field">Data<input type="date" data-f="data" value="${esc(F.data)}"></label>
    <button class="btn" data-act="guardarEditPag">Guardar</button>`, F.editId);
}

function sheetSettings() {
  const c = cfg();
  return sheetWrap('Definições', `
    <div class="form" style="border:1px solid var(--line)">
      <label><span>Excel</span><span class="sub" style="text-align:right">${esc(c.excelName || '')}</span></label>
      <label><span>Registos</span><span class="sub" style="text-align:right">${CSV_NAME}</span></label>
      <label><span>Cópia de segurança</span><span class="sub" style="text-align:right">${CSV_BACKUP} · ${esc(fmtDate(LS.get('fin_backup', '')) || 'ainda não')}</span></label>
    </div>
    <button class="btn ghost" data-act="relerExcel">Voltar a ler o Excel</button>
    <button class="btn ghost" data-act="trocarExcel">Escolher outro Excel</button>
    <button class="btn ghost" data-act="sair">Terminar sessão</button>
    <p class="sub">ID da app: ${esc(c.clientId || '')}<br>Endereço de retorno: ${esc(redirectUri())}</p>`);
}

/* ------------------------------------------------------------------ abrir registos para editar */
function openEdit(id) {
  const recs = allRecs(); const r = recs.find((x) => x.ID === id); if (!r) return;
  const base = { editId: r.ID, valor: money(r.Valor), data: r.Data || todayIso(), desc: r.Descricao || '' };
  if (r.Registo === 'Movimento') {
    const part = r.Divida ? recs.find((x) => x.ID === r.Divida) : null;
    Object.keys(F).forEach((k) => delete F[k]);
    Object.assign(F, base, { tipo: r.Tipo, cat: r.Categoria, conta: r.Conta || (S.conf.contas[0] || {}).nome || '', split: !!part, partId: part ? part.ID : '', spessoa: part ? part.Categoria : '',
      squem: part && part.Tipo === 'Eu devo' ? 'outro' : 'eu', minha: money(r.Valor) });
    if (part && part.Tipo === 'Devem-me') F.valor = money(r2(num(r.Valor) + num(part.Valor)));
    S.sheet = sheetNovo;
  } else if (r.Registo === 'Planeada') {
    Object.assign(F, base, { ptipo: r.Tipo, cat: r.Categoria, conta: r.Conta, prevista: r.Tipo === 'Paga' ? r.Divida : '' });
    S.sheet = sheetPlan;
  } else if (r.Registo === 'Divida' || r.Registo === 'Partilha') {
    if (r.Registo === 'Partilha') { const mv = recs.find((x) => x.Registo === 'Movimento' && x.Divida === r.ID); if (mv) return openEdit(mv.ID); }
    Object.assign(F, base, { dreg: r.Registo, dtipo: r.Tipo, pessoa: r.Categoria, pconta: r.Conta, prazo: r.Registo === 'Divida' && isIso(r.Divida) ? r.Divida : '' });
    S.sheet = sheetNovaDiv;
  } else if (r.Registo === 'Pagamento') {
    Object.assign(F, base, { pconta: r.Conta, titulo: recTitle(r) });
    S.sheet = sheetEditPag;
  } else { toast('Este registo só se pode apagar.'); return; }
  render();
}

/* ------------------------------------------------------------------ eventos */
document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && S.sheet) { S.sheet = null; render(); } });
let touchY = null;
document.addEventListener('touchstart', (e) => { const sh = e.target.closest('.sheet'); touchY = sh && sh.scrollTop <= 0 ? e.touches[0].clientY : null; }, { passive: true });
document.addEventListener('touchend', (e) => { if (touchY != null && S.sheet && e.changedTouches[0].clientY - touchY > 90) { S.sheet = null; render(); } touchY = null; }, { passive: true });
document.addEventListener('input', (e) => {
  const f = e.target.dataset && e.target.dataset.f; if (!f) return;
  F[f] = e.target.value;
  const sr = document.getElementById('splitResumo'); if (sr) sr.innerHTML = splitResumo();
});
document.addEventListener('change', (e) => {
  const t = e.target;
  if (t.dataset.f === 'catSel') { if (t.value) { F.cat = t.value; render(); } return; }
  if (t.dataset.f) { F[t.dataset.f] = t.value; if (t.tagName === 'SELECT' && t.dataset.f === 'conta' && S.sheet) render(); }
});
document.addEventListener('click', async (e) => {
  const el = e.target.closest('[data-act],[data-tab],[data-tipo],[data-dtipo],[data-ptipo],[data-squem],[data-cat],[data-day],[data-del],[data-edit],[data-pagar],[data-pick],[data-open]');
  if (!el) return;
  const d = el.dataset;
  if (d.tab) { S.tab = d.tab; S.sheet = null; render(); window.scrollTo(0, 0); return; }
  if (d.tipo) { F.tipo = d.tipo; render(); return; }
  if (d.dtipo) { F.dtipo = d.dtipo; render(); return; }
  if (d.ptipo) { F.ptipo = d.ptipo; F.valor = ''; render(); return; }
  if (d.squem) { F.squem = d.squem; render(); return; }
  if (d.cat) { F.cat = d.cat; render(); return; }
  if (d.edit) { openEdit(d.edit); return; }
  if (d.pagar) { Object.keys(F).forEach((k) => delete F[k]); Object.assign(F, { ppessoa: d.pagar, ptipo: '', valor: '', pconta: (S.conf.contas[0] || {}).nome || '', data: todayIso() }); S.sheet = sheetPagar; render(); return; }
  if (d.del) {
    const recs = allRecs(); const rr = recs.find((x) => x.ID === d.del); if (!rr) return;
    const n = cascadeIds(recs, rr.ID).size; const pags = recs.filter((x) => x.Registo === 'Pagamento' && cascadeIds(recs, rr.ID).has(x.Divida)).length;
    const msg = n > 1 || pags ? `Apagar este registo${n > 1 ? ' e a dívida/despesa ligada' : ''}${pags ? ` (e ${pags} pagamento(s))` : ''}?` : 'Apagar este registo?';
    if (confirm(msg)) { S.sheet = null; await commit([{ op: 'del', id: rr.ID }], 'Apagado ✓'); }
    return;
  }
  if (d.pick) { await choose(d.pick); return; }
  if (d.open) { await openFolder(d.open, (browse && browse.path ? browse.path + '/' : '') + d.name); return; }
  if (d.day) { await toggleDay(+d.day); return; }
  const contaDef = (S.conf && S.conf.contas[0] || {}).nome || '';
  switch (d.act) {
    case 'refresh': await refresh(); break;
    case 'novo': Object.keys(F).forEach((k) => delete F[k]); Object.assign(F, { tipo: 'Despesa', valor: '', cat: '', conta: contaDef, data: todayIso(), desc: '', split: false, spessoa: '', squem: 'eu', minha: '' }); S.sheet = sheetNovo; render(); setTimeout(() => { const i = document.getElementById('valor'); if (i) i.focus(); }, 50); break;
    case 'toggleSplit': F.split = !F.split; if (F.split && !F.minha) { const v = parseMoney(F.valor); if (v > 0) F.minha = money(r2(v / 2)); } render(); break;
    case 'metade': { const v = parseMoney(F.valor); if (v > 0) { F.minha = money(r2(v / 2)); render(); } break; }
    case 'novaPlan': Object.keys(F).forEach((k) => delete F[k]); Object.assign(F, { ptipo: 'Por pagar', desc: '', valor: '', data: todayIso(), cat: S.conf.cats.Despesa[0] || '', conta: contaDef }); S.sheet = sheetPlan; render(); break;
    case 'guardarPlan': await savePlan(el); break;
    case 'planPagar': F.ptipo = 'Paga'; F.prevista = F.data; F.data = todayIso(); await savePlan(el); break;
    case 'planDesmarcar': F.ptipo = 'Por pagar'; F.data = F.prevista || F.data; F.prevista = ''; await savePlan(el); break;
    case 'planAbertas': S.planVer = 'abertas'; render(); break;
    case 'planPagas': S.planVer = 'pagas'; render(); break;
    case 'regDinheiro': S.regFiltro = ''; render(); break;
    case 'regTodos': S.regFiltro = 'todos'; render(); break;
    case 'settings': S.sheet = sheetSettings; render(); break;
    case 'novaDiv': Object.keys(F).forEach((k) => delete F[k]); Object.assign(F, { dreg: 'Divida', dtipo: 'Eu devo', pessoa: '', desc: '', valor: '', pconta: '', data: todayIso(), prazo: '' }); S.sheet = sheetNovaDiv; render(); break;
    case 'guardarDiv': await saveDiv(el); break;
    case 'fechar': S.sheet = null; render(); break;
    case 'guardarMov': await saveMov(el); break;
    case 'guardarPag': await savePag(el); break;
    case 'guardarEditPag': await saveEditPag(el); break;
    case 'resAnt': case 'resSeg': { let m = S.mes.m + (d.act === 'resAnt' ? -1 : 1), y = S.mes.y; if (m < 1) { m = 12; y--; } if (m > 12) { m = 1; y++; } S.mes = { y, m }; render(); break; }
    case 'resHoje': S.mes = { y: new Date().getFullYear(), m: new Date().getMonth() + 1 }; render(); break;
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
    case 'relerExcel': S.sheet = null; S.busy = true; render(); try { const { conf } = await excelLoad(true); S.conf = conf; LS.set('fin_conf', conf); toast('Excel lido ✓'); } catch (er) { toast(er.message); } S.busy = false; render(); break;
    case 'trocarExcel': setCfg({ excelId: null }); S.conf = null; S.sheet = null; pickResults = null; render(); break;
    case 'sair': ['fin_tok', 'fin_conf', 'fin_recs', 'fin_outbox'].forEach(LS.del); setCfg({ excelId: null }); location.reload(); break;
  }
});

/* ------------------------------------------------------------------ guardar */
async function run(btn, fn) { btn.disabled = true; try { await fn(); S.sheet = null; render(); } catch (e) { btn.disabled = false; } }

async function saveMov(btn) {
  const v = parseMoney(F.valor);
  if (!(v > 0)) { toast('Escreve um valor maior que zero.'); return; }
  if (!F.cat) { toast('Escolhe a categoria.'); return; }
  const split = F.tipo === 'Despesa' && F.split;
  const ops = []; const recs = allRecs();
  const old = F.editId ? recs.find((x) => x.ID === F.editId) : null;
  const oldPart = old && old.Divida ? old.Divida : '';
  let movId = F.editId || uid(); let partId = oldPart || uid();
  const mov = { Registo: 'Movimento', Data: F.data || todayIso(), Descricao: F.desc.trim(), Tipo: F.tipo, Categoria: F.cat, Conta: F.conta, Valor: v, Divida: '' };
  if (split) {
    const p = String(F.spessoa || '').trim(); const minha = parseMoney(F.minha);
    if (!p) { toast('Escreve com quem dividiste.'); return; }
    if (!(minha >= 0) || minha > v) { toast('A tua parte tem de estar entre 0 e o valor total.'); return; }
    const outro = F.squem === 'outro';
    const partVal = outro ? minha : r2(v - minha);
    Object.assign(mov, { Valor: minha, Conta: outro ? '' : F.conta, Divida: partId });
    const part = { Registo: 'Partilha', Data: mov.Data, Descricao: F.desc.trim() || F.cat, Tipo: outro ? 'Eu devo' : 'Devem-me', Categoria: p, Conta: outro ? '' : F.conta, Valor: partVal, Divida: '' };
    if (partVal <= 0) { mov.Divida = ''; if (oldPart) ops.push({ op: 'del', id: oldPart }); }
    else if (oldPart) ops.push({ op: 'upd', id: oldPart, patch: part });
    else ops.push({ op: 'add', rec: newRec(Object.assign({ ID: partId }, part)) });
  } else if (oldPart) {
    const pags = recs.filter((x) => x.Registo === 'Pagamento' && x.Divida === oldPart).length;
    if (pags && typeof confirm === 'function' && !confirm(`A dívida desta despesa já tem ${pags} pagamento(s). Deixar de dividir apaga a dívida e esses pagamentos. Continuar?`)) return;
    ops.push({ op: 'upd', id: movId, patch: { Divida: '' } }, { op: 'del', id: oldPart });
  }
  if (old) ops.unshift({ op: 'upd', id: movId, patch: mov }); else ops.unshift({ op: 'add', rec: newRec(Object.assign({ ID: movId }, mov)) });
  // apagar a partilha antiga não pode levar o movimento: tira a ligação antes
  const ordered = ops.filter((o) => o.op !== 'del').concat(ops.filter((o) => o.op === 'del'));
  await run(btn, () => commit(ordered));
}
async function savePlan(btn) {
  const v = parseMoney(F.valor);
  if (!(v > 0)) { toast('Escreve um valor maior que zero.'); return; }
  if (!isIso(F.data)) { toast('Escolhe a data.'); return; }
  const rec = { Registo: 'Planeada', Data: F.data, Descricao: String(F.desc || '').trim(), Tipo: F.ptipo === 'Paga' ? 'Paga' : 'Por pagar', Categoria: F.cat, Conta: F.conta, Valor: v, Divida: F.ptipo === 'Paga' ? (F.prevista || F.data) : '' };
  await run(btn, () => commit([F.editId ? { op: 'upd', id: F.editId, patch: rec } : { op: 'add', rec: newRec(rec) }], rec.Tipo === 'Paga' && F.editId ? 'Marcada como paga ✓' : undefined));
}
async function saveDiv(btn) {
  const p = String(F.pessoa || '').trim();
  if (!p) { toast('Escreve o nome da pessoa.'); return; }
  if (F.dreg === 'Partilha') { await run(btn, () => commit([{ op: 'upd', id: F.editId, patch: { Categoria: p, Descricao: String(F.desc || '').trim() } }])); return; }
  const v = parseMoney(F.valor);
  if (!(v > 0)) { toast('Escreve um valor maior que zero.'); return; }
  const rec = { Registo: 'Divida', Data: isIso(F.data) ? F.data : todayIso(), Categoria: p, Descricao: String(F.desc || '').trim(), Tipo: F.dtipo, Conta: F.pconta, Valor: v, Divida: isIso(F.prazo) ? F.prazo : '' };
  await run(btn, () => commit([F.editId ? { op: 'upd', id: F.editId, patch: rec } : { op: 'add', rec: newRec(rec) }]));
}
async function savePag(btn) {
  const v = parseMoney(F.valor);
  if (!(v > 0)) { toast('Escreve um valor maior que zero.'); return; }
  const m = calc(S.mes.y, S.mes.m);
  const abertas = m.divs.filter((d) => pkey(d.pessoa) === pkey(F.ppessoa) && d.tipo === F.ptipo && d.falta > 0.005)
    .sort((a, b) => String(a.prazo || '9999').localeCompare(String(b.prazo || '9999')) || String(a.data).localeCompare(String(b.data)));
  const max = r2(abertas.reduce((a, d) => a + d.falta, 0));
  if (v > max + 0.005) { toast('É mais do que está em dívida (' + eur(max) + ').'); return; }
  // reparte o pagamento pelas dívidas em aberto, primeiro as que têm prazo mais cedo
  let rest = v; const ops = [];
  for (const d of abertas) {
    if (rest <= 0.005) break;
    const x = r2(Math.min(rest, d.falta)); rest = r2(rest - x);
    ops.push({ op: 'add', rec: newRec({ Registo: 'Pagamento', Data: F.data || todayIso(), Descricao: d.desc, Tipo: d.tipo, Categoria: d.pessoa, Conta: F.pconta, Valor: x, Divida: d.id }) });
  }
  await run(btn, () => commit(ops));
}
async function saveEditPag(btn) {
  const v = parseMoney(F.valor);
  if (!(v > 0)) { toast('Escreve um valor maior que zero.'); return; }
  await run(btn, () => commit([{ op: 'upd', id: F.editId, patch: { Valor: v, Conta: F.pconta, Data: isIso(F.data) ? F.data : todayIso() } }]));
}
async function toggleDay(day) {
  const { y, m } = S.diasMes; const st = diasState(y, m);
  if (st.xlSet.has(day)) { toast('Este dia está marcado no Excel — muda-o lá.'); return; }
  const date = iso(y, m, day);
  const rec = allRecs().find((r) => r.Registo === 'Dia' && r.Data === date);
  await commit([rec ? { op: 'del', id: rec.ID } : { op: 'add', rec: newRec({ Registo: 'Dia', Data: date, Valor: 1 }) }], false);
}
async function toggleRecebido() {
  const { y, m } = S.diasMes; const st = diasState(y, m);
  await commit([st.salRec ? { op: 'del', id: st.salRec.ID } : { op: 'add', rec: newRec({ Registo: 'Salario', Data: iso(y, m, 1), Conta: S.conf.csal, Valor: r2(st.sal) }) }]);
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
  S.conf = null; render(); await refresh(true);
}

/* ------------------------------------------------------------------ arranque */
(async function start() {
  LS.del('fin_model'); // formato antigo
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
  window.addEventListener('online', () => refresh(true));
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && cfg().excelId && LS.get('fin_tok', null) && !S.sheet) refresh(true); });
  try { await handleRedirect(); } catch (e) { S.err = e.message; }
  render();
  if (cfg().clientId && LS.get('fin_tok', null) && cfg().excelId) refresh(!!S.conf);
})();
