// Leitor mínimo de .xlsx (só leitura de valores já calculados pelo Excel).
// Usa apenas funcionalidades do browser: DecompressionStream e DOMParser.
(function (global) {
  'use strict';

  async function inflateRaw(bytes) {
    const ds = new DecompressionStream('deflate-raw');
    const stream = new Blob([bytes]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  function readZipIndex(buf) {
    const dv = new DataView(buf);
    let eocd = -1;
    for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 70000); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('Ficheiro não é um .xlsx válido');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const dec = new TextDecoder();
    const entries = {};
    for (let n = 0; n < count; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true);
      const elen = dv.getUint16(p + 30, true);
      const clen = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      const name = dec.decode(new Uint8Array(buf, p + 46, nlen));
      entries[name] = { method, csize, off };
      p += 46 + nlen + elen + clen;
    }
    return entries;
  }

  async function readEntry(buf, entries, name) {
    const e = entries[name];
    if (!e) return null;
    const dv = new DataView(buf);
    const nlen = dv.getUint16(e.off + 26, true);
    const elen = dv.getUint16(e.off + 28, true);
    const start = e.off + 30 + nlen + elen;
    const raw = new Uint8Array(buf, start, e.csize);
    const out = e.method === 8 ? await inflateRaw(raw) : raw;
    return new TextDecoder().decode(out);
  }

  function xml(text) { return new DOMParser().parseFromString(text, 'application/xml'); }
  function byTag(node, tag) { return node.getElementsByTagNameNS('*', tag); }

  async function open(buf) {
    const entries = readZipIndex(buf);
    const wbXml = xml(await readEntry(buf, entries, 'xl/workbook.xml'));
    const relsXml = xml(await readEntry(buf, entries, 'xl/_rels/workbook.xml.rels'));
    const rels = {};
    for (const r of byTag(relsXml, 'Relationship')) {
      let t = r.getAttribute('Target');
      if (t.startsWith('/')) t = t.slice(1); else t = 'xl/' + t;
      rels[r.getAttribute('Id')] = t;
    }
    const sheets = {};
    for (const s of byTag(wbXml, 'sheet')) {
      const rid = s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      sheets[s.getAttribute('name')] = rels[rid];
    }
    let shared = [];
    const ss = await readEntry(buf, entries, 'xl/sharedStrings.xml');
    if (ss) {
      const doc = xml(ss);
      shared = Array.from(byTag(doc, 'si')).map((si) => Array.from(byTag(si, 't')).map((t) => t.textContent).join(''));
    }
    const cache = {};
    async function sheet(name) {
      if (cache[name]) return cache[name];
      const path = sheets[name];
      if (!path) return (cache[name] = {});
      const doc = xml(await readEntry(buf, entries, path));
      const cells = {};
      for (const c of byTag(doc, 'c')) {
        const r = c.getAttribute('r');
        const t = c.getAttribute('t');
        const vEl = byTag(c, 'v')[0];
        let v = null;
        if (t === 'inlineStr') v = Array.from(byTag(c, 't')).map((x) => x.textContent).join('');
        else if (vEl) {
          const raw = vEl.textContent;
          if (t === 's') v = shared[+raw];
          else if (t === 'str' || t === 'e') v = raw;
          else if (t === 'b') v = raw === '1';
          else v = raw === '' ? null : Number(raw);
        }
        if (v !== null && v !== '') cells[r] = v;
      }
      return (cache[name] = cells);
    }
    return { sheetNames: Object.keys(sheets), sheet };
  }

  global.XlsxLite = { open };
})(window);
