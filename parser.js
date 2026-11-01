/* Robô de Extrato — leitura e classificação de extratos bancários.
 * Sem dependência de DOM: funciona no navegador (painel do Excel) e no Node (testes).
 * Formatos: CSV/TXT, XLSX/XLS (via SheetJS), OFX, PDF (via pdf.js).
 */
(function (root) {
  'use strict';

  const MESES = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'];
  const MES_ABREV = { jan: 1, fev: 2, feb: 2, mar: 3, abr: 4, apr: 4, mai: 5, may: 5, jun: 6, jul: 7, ago: 8, aug: 8, set: 9, sep: 9, out: 10, oct: 10, nov: 11, dez: 12, dec: 12 };

  // ---------------------------------------------------------------- números e datas
  function excelSerial(y, m, d) {
    return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  }
  function serialToYMD(s) {
    const dt = new Date(Date.UTC(1899, 11, 30) + Math.round(s) * 86400000);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
  }
  function validYMD(y, m, d) {
    if (!(y >= 1990 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return false;
    const dt = new Date(Date.UTC(y, m - 1, d));
    return dt.getUTCMonth() === m - 1;
  }
  function fixYear(y) { y = +y; return y < 100 ? 2000 + y : y; }

  /** devolve serial Excel (inteiro) ou null. defaultYear usado quando a data vem só com dia/mês */
  function parseDate(v, defaultYear) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date && !isNaN(v)) {
      return excelSerial(v.getFullYear(), v.getMonth() + 1, v.getDate());
    }
    if (typeof v === 'number') {
      if (v > 20000 && v < 80000) return Math.floor(v);   // serial Excel
      return null;
    }
    let s = String(v).trim().toLowerCase();
    if (!s) return null;
    let m;
    // yyyy-mm-dd (ISO, com ou sem hora)
    if ((m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/))) {
      const y = +m[1], mo = +m[2], d = +m[3];
      return validYMD(y, mo, d) ? excelSerial(y, mo, d) : null;
    }
    // dd/mm/yyyy, dd-mm-yy, dd.mm.yyyy
    if ((m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/))) {
      const d = +m[1], mo = +m[2], y = fixYear(m[3]);
      return validYMD(y, mo, d) ? excelSerial(y, mo, d) : null;
    }
    // dd/mm (sem ano)
    if ((m = s.match(/^(\d{1,2})\/(\d{1,2})(?!\/|\d)/)) && defaultYear) {
      const d = +m[1], mo = +m[2];
      return validYMD(defaultYear, mo, d) ? excelSerial(defaultYear, mo, d) : null;
    }
    // 05 jan 2026 / 05 de janeiro de 2026 / 05-jan-26
    if ((m = s.match(/^(\d{1,2})(?:\s+de)?[\s\-\/]+([a-zç]{3})[a-zç]*\.?(?:\s+de)?(?:[\s\-\/]+(\d{2,4}))?/))) {
      const d = +m[1], mo = MES_ABREV[m[2]];
      const y = m[3] ? fixYear(m[3]) : defaultYear;
      if (mo && y && validYMD(y, mo, d)) return excelSerial(y, mo, d);
    }
    return null;
  }

  /** número em formato BR (1.234,56), US (1,234.56), com R$, parênteses, sinal no fim, D/C */
  function parseMoney(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    let s = String(v).trim();
    if (!s) return null;
    let neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    const up = s.toUpperCase();
    if (/\s*D$/.test(up) && !/[A-CE-Z]D$/.test(up)) { neg = true; s = s.replace(/\s*[dD]$/, ''); }
    else if (/\s*C$/.test(up) && !/[A-BD-Z]C$/.test(up)) { s = s.replace(/\s*[cC]$/, ''); }
    s = s.replace(/R\$|BRL|US\$|\$/gi, '').replace(/[\s ]/g, '');
    if (/-$/.test(s)) { neg = !neg; s = s.slice(0, -1); }
    if (/^[-−–]/.test(s)) { neg = !neg; s = s.slice(1); }
    if (/^\+/.test(s)) s = s.slice(1);
    if (!/^[\d.,]+$/.test(s) || !/\d/.test(s)) return null;
    const lastComma = s.lastIndexOf(','), lastDot = s.lastIndexOf('.');
    let num;
    if (lastComma > lastDot) {                 // vírgula decimal (BR)
      num = parseFloat(s.replace(/\./g, '').replace(',', '.'));
    } else if (lastDot > lastComma) {
      const decimals = s.length - lastDot - 1;
      if (lastComma >= 0 || decimals !== 3 || (s.match(/\./g) || []).length === 1 && decimals !== 3) {
        num = parseFloat(s.replace(/,/g, ''));   // ponto decimal (US)
      } else {
        num = parseFloat(s.replace(/\./g, '')); // 1.234 = milhar BR sem centavos
      }
    } else {
      num = parseFloat(s);
    }
    if (!isFinite(num)) return null;
    return Math.round((neg ? -num : num) * 100) / 100;
  }

  // ---------------------------------------------------------------- texto
  function norm(s) {
    return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
  }
  function cleanText(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }

  // ---------------------------------------------------------------- CSV
  function decodeBytes(buf) {
    const bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, ''); }
    catch (e) { return new TextDecoder('windows-1252').decode(bytes); }
  }
  function parseCSV(text) {
    const lines = text.split(/\r?\n/).slice(0, 20).join('\n');
    const cands = [';', ',', '\t', '|'];
    let best = ',', bestScore = -1;
    for (const c of cands) {
      const counts = lines.split('\n').filter(l => l.trim()).map(l => splitCSVLine(l, c).length);
      if (!counts.length) continue;
      const mode = counts.sort((a, b) => counts.filter(x => x === b).length - counts.filter(x => x === a).length)[0];
      const score = mode > 1 ? counts.filter(x => x === mode).length * mode : 0;
      if (score > bestScore) { bestScore = score; best = c; }
    }
    const rows = [];
    let cur = [], field = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; }
        else field += ch;
      } else if (ch === '"') q = true;
      else if (ch === best) { cur.push(field); field = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        cur.push(field); field = '';
        if (cur.some(x => x.trim() !== '')) rows.push(cur);
        cur = [];
      } else field += ch;
    }
    cur.push(field);
    if (cur.some(x => x.trim() !== '')) rows.push(cur);
    return rows;
  }
  function splitCSVLine(line, d) {
    const out = []; let f = '', q = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (q) { if (ch === '"') { if (line[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += ch; }
      else if (ch === '"') q = true; else if (ch === d) { out.push(f); f = ''; } else f += ch;
    }
    out.push(f); return out;
  }

  // ---------------------------------------------------------------- tabelas (CSV/XLSX)
  const H = {
    date: /^(data|dt|date|dia)\b|data (do |da )?(lan|mov|trans|opera|pag|compen)|^data$/,
    desc: /descri|historico|lancamento|detalhe|estabelecimento|transac|identifica|^nome|favorecido|beneficiario|pagador|origem|destino|memo|complemento|^titulo/,
    value: /^valor|^vlr|montante|quantia|amount|^total$|valor liquido|valor \(r\$\)/,
    credit: /^credito|^entrada|^recebid|^credit/,
    debit: /^debito|^saida|^pago|^enviad|^debit/,
    type: /^tipo|natureza|^d\/c$|^c\/d$|^dc$|^cd$|^operacao$|^movimento$/,
    balance: /saldo|balance/,
  };
  function looksSign(v) {
    const s = norm(v);
    return /^(c|d|cr|db|credito|debito|entrada|saida|\+|-)$/.test(s);
  }
  function signOf(v) {
    const s = norm(v);
    if (/^(d|db|debito|saida|-)$/.test(s)) return -1;
    if (/^(c|cr|credito|entrada|\+)$/.test(s)) return 1;
    return 0;
  }

  function detectTable(rows, defaultYear) {
    rows = rows.map(r => r.map(c => (typeof c === 'string' ? c.trim() : c)));
    // 1) linha de cabeçalho
    let hdrIdx = -1, hdrScore = 0;
    for (let i = 0; i < Math.min(rows.length, 40); i++) {
      const cells = rows[i].map(norm);
      let sc = 0;
      if (cells.some(c => H.date.test(c))) sc += 2;
      if (cells.some(c => H.value.test(c) || H.credit.test(c) || H.debit.test(c))) sc += 2;
      if (cells.some(c => H.desc.test(c))) sc += 1;
      if (sc > hdrScore && rows[i].filter(c => c !== '' && c != null).length >= 2) { hdrScore = sc; hdrIdx = i; }
    }
    if (hdrScore < 3) hdrIdx = -1;
    const header = hdrIdx >= 0 ? rows[hdrIdx].map(c => cleanText(c)) : [];
    const body = rows.slice(hdrIdx + 1).filter(r => r.some(c => c !== '' && c != null));
    const ncol = Math.max(header.length, ...body.slice(0, 200).map(r => r.length), 0);
    const sample = body.slice(0, 300);
    const stat = [];
    for (let c = 0; c < ncol; c++) {
      let dates = 0, money = 0, neg = 0, filled = 0, txtLen = 0, signs = 0;
      for (const r of sample) {
        const v = r[c];
        if (v === '' || v == null) continue;
        filled++;
        if (parseDate(v, defaultYear) !== null) dates++;
        const mv = parseMoney(v);
        if (mv !== null && (typeof v === 'number' || /\d[.,]\d{2}\b/.test(String(v)) || /^-?\d+$/.test(String(v)))) { money++; if (mv < 0) neg++; }
        if (typeof v === 'string' && !/^[\d\s.,R$()+\-\/:]+$/.test(v)) txtLen += v.length;
        if (looksSign(v)) signs++;
      }
      stat.push({ c, dates, money, neg, filled, txtLen, signs, name: norm(header[c] || '') });
    }
    const map = { date: -1, desc: -1, desc2: -1, value: -1, credit: -1, debit: -1, type: -1 };
    const byName = (re, excl) => stat.find(s => re.test(s.name) && !(excl || []).includes(s.c));
    let s;
    if ((s = byName(H.date))) map.date = s.c;
    if ((s = stat.find(x => H.value.test(x.name) && !H.balance.test(x.name)))) map.value = s.c;
    if (map.value < 0) {
      const cr = byName(H.credit), db = byName(H.debit);
      if (cr && db && cr.c !== db.c) { map.credit = cr.c; map.debit = db.c; }
    }
    const typ = stat.find(x => H.type.test(x.name) && x.c !== map.date && x.c !== map.value);
    if (typ && typ.signs >= typ.filled * 0.8 && typ.filled) map.type = typ.c;
    const descs = stat.filter(x => H.desc.test(x.name) && ![map.date, map.value, map.credit, map.debit, map.type].includes(x.c));
    if (descs.length) {
      descs.sort((a, b) => b.txtLen - a.txtLen);
      map.desc = descs[0].c;
    }
    // tipo descritivo (ex.: "Tipo de transação" = "Pix recebido") vira complemento da descrição
    if (typ && map.type < 0 && typ.txtLen > 0 && typ.c !== map.desc) map.desc2 = typ.c;
    else if (descs.length > 1 && descs[1].txtLen > 0) map.desc2 = descs[1].c;
    // 2) inferência por conteúdo quando faltou algo
    if (map.date < 0) {
      const cand = stat.filter(x => x.filled && x.dates / x.filled > 0.7).sort((a, b) => b.dates - a.dates)[0];
      if (cand) map.date = cand.c;
    }
    if (map.value < 0 && map.credit < 0) {
      const cand = stat.filter(x => x.c !== map.date && x.filled && x.money / x.filled > 0.7 && !H.balance.test(x.name))
        .sort((a, b) => (b.neg > 0) - (a.neg > 0) || a.c - b.c)[0];
      if (cand) map.value = cand.c;
    }
    if (map.desc < 0) {
      const cand = stat.filter(x => ![map.date, map.value, map.credit, map.debit, map.type].includes(x.c)).sort((a, b) => b.txtLen - a.txtLen)[0];
      if (cand && cand.txtLen > 0) map.desc = cand.c;
    }
    return { header, body, map, ncol, headerRow: hdrIdx };
  }

  function rowsFromTable(t, map, defaultYear) {
    const out = [];
    let lastDate = null;
    for (const r of t.body) {
      let dt = map.date >= 0 ? parseDate(r[map.date], defaultYear) : null;
      let val = null;
      if (map.value >= 0) val = parseMoney(r[map.value]);
      else if (map.credit >= 0 || map.debit >= 0) {
        const cr = map.credit >= 0 ? parseMoney(r[map.credit]) : null;
        const db = map.debit >= 0 ? parseMoney(r[map.debit]) : null;
        if (cr !== null && cr !== 0) val = Math.abs(cr);
        else if (db !== null && db !== 0) val = -Math.abs(db);
      }
      if (val !== null && map.type >= 0) {
        const sg = signOf(r[map.type]);
        if (sg) val = sg * Math.abs(val);
      }
      let desc = map.desc >= 0 ? cleanText(r[map.desc]) : '';
      if (map.desc2 >= 0) {
        const d2 = cleanText(r[map.desc2]);
        if (d2 && norm(d2) !== norm(desc)) desc = desc ? d2 + ' - ' + desc : d2;
      }
      if (dt === null && val !== null && lastDate !== null && map.date >= 0 && (r[map.date] === '' || r[map.date] == null)) dt = lastDate;
      if (dt !== null) lastDate = dt;
      if (dt === null || val === null || val === 0) continue;
      if (/^saldo|saldo (anterior|do dia|final|inicial|disponivel|em conta)|^total/i.test(norm(desc))) continue;
      out.push({ date: dt, desc: desc || '(sem descrição)', value: val });
    }
    return out;
  }

  // ---------------------------------------------------------------- OFX
  function parseOFX(text) {
    const out = [];
    const re = /<STMTTRN>([\s\S]*?)(?:<\/STMTTRN>|(?=<STMTTRN>)|(?=<\/BANKTRANLIST>))/gi;
    let m;
    const tag = (blk, t) => { const x = blk.match(new RegExp('<' + t + '>([^<\\r\\n]*)', 'i')); return x ? x[1].trim() : ''; };
    while ((m = re.exec(text))) {
      const blk = m[1];
      const d = tag(blk, 'DTPOSTED');
      const amt = tag(blk, 'TRNAMT');
      const memo = tag(blk, 'MEMO'), name = tag(blk, 'NAME');
      if (!d || !amt) continue;
      const y = +d.slice(0, 4), mo = +d.slice(4, 6), dd = +d.slice(6, 8);
      if (!validYMD(y, mo, dd)) continue;
      const v = parseFloat(amt.replace(',', '.'));
      if (!isFinite(v) || v === 0) continue;
      let desc = memo && name && norm(memo) !== norm(name) ? name + ' - ' + memo : (memo || name);
      out.push({ date: excelSerial(y, mo, dd), desc: cleanText(desc) || '(sem descrição)', value: Math.round(v * 100) / 100, fitid: tag(blk, 'FITID') });
    }
    return out;
  }

  // ---------------------------------------------------------------- PDF (texto por linhas)
  const MONEY_RE = /(?:^|\s)(\(?[-−–+]?\s?(?:R\$\s?)?[-−–]?\d{1,3}(?:\.\d{3})*,\d{2}\)?(?:\s?[-DC](?![a-zA-Z]))?)(?=\s|$)/g;
  const DATE_START = /^(\d{1,2}\/\d{1,2}(?:\/\d{2,4})?|\d{4}-\d{2}-\d{2}|\d{1,2}(?:\s+de)?\s+[a-zç]{3}[a-zç]*\.?(?:(?:\s+de)?\s+\d{4})?)(?=\s|$)/i;

  function pdfLinesToRows(lines, defaultYear) {
    const all = lines.join('\n');
    let docYear = defaultYear;
    const yrs = (all.match(/\b(20\d{2})\b/g) || []).map(Number).filter(y => y >= 2000 && y <= 2100);
    if (yrs.length) {
      const freq = {}; yrs.forEach(y => freq[y] = (freq[y] || 0) + 1);
      docYear = +Object.keys(freq).sort((a, b) => freq[b] - freq[a])[0];
    }
    const hasSaldoCol = /\bsaldo\b/i.test(all);
    const out = [];
    let curDate = null, last = null;
    for (let raw of lines) {
      const line = cleanText(raw);
      if (!line) continue;
      const nl = norm(line);
      let dm = line.match(DATE_START), dt = null, rest = line;
      if (dm) { dt = parseDate(dm[1], docYear); if (dt !== null) rest = line.slice(dm[0].length).trim(); }
      const moneys = [];
      let mm; MONEY_RE.lastIndex = 0;
      while ((mm = MONEY_RE.exec(rest))) moneys.push({ txt: mm[1], idx: mm.index + (mm[0].length - mm[1].length) });
      if (dt !== null && !moneys.length) { curDate = dt; last = null; continue; }   // linha só com data (agrupador)
      if (!moneys.length) {
        // continuação de descrição
        if (last && line.length < 80 && !/saldo|total|pagina|página/.test(nl)) last.desc = cleanText(last.desc + ' ' + line);
        continue;
      }
      if (/saldo (anterior|do dia|final|inicial|disponivel|em conta|bloqueado)|^saldo\b|^total\b|^resumo|^subtotal/.test(norm(rest))) { last = null; continue; }
      const useDate = dt !== null ? dt : curDate;
      if (useDate === null) continue;
      if (dt !== null) curDate = dt;
      let pick = moneys[0];
      if (moneys.length >= 2 && hasSaldoCol) pick = moneys[moneys.length - 2];
      let val = parseMoney(pick.txt);
      if (val === null || val === 0) continue;
      // palavras que indicam saída quando o PDF não mostra sinal
      let desc = cleanText(rest.slice(0, moneys[0].idx));
      if (!desc) desc = cleanText(rest.replace(pick.txt, ''));
      out.push({ date: useDate, desc: desc || '(sem descrição)', value: val });
      last = out[out.length - 1];
    }
    return out;
  }

  async function pdfToLines(pdfjsLib, data) {
    const doc = await pdfjsLib.getDocument({ data: new Uint8Array(data), isEvalSupported: false }).promise;
    const lines = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      const tc = await page.getTextContent();
      const rows = [];
      for (const it of tc.items) {
        if (!it.str || !it.str.trim()) continue;
        const y = it.transform[5], x = it.transform[4];
        let row = rows.find(r => Math.abs(r.y - y) <= 2.5);
        if (!row) { row = { y, items: [] }; rows.push(row); }
        row.items.push({ x, s: it.str, w: it.width || 0 });
      }
      rows.sort((a, b) => b.y - a.y);
      for (const r of rows) {
        r.items.sort((a, b) => a.x - b.x);
        let s = '', endX = null;
        for (const it of r.items) {
          if (endX !== null) s += (it.x - endX > 1.5 ? '  ' : '');
          s += it.s; endX = it.x + it.w;
        }
        lines.push(s);
      }
    }
    return lines;
  }

  // ---------------------------------------------------------------- entrada única
  /**
   * lê um arquivo e devolve {kind, rows:[{date,desc,value}], table?, map?, warnings:[]}
   * libs = {XLSX, pdfjsLib}; mapOverride opcional para tabelas
   */
  async function readStatement(fileName, buffer, libs, opts) {
    opts = opts || {};
    const defaultYear = opts.year || new Date().getFullYear();
    const ext = (fileName.split('.').pop() || '').toLowerCase();
    const bytes = new Uint8Array(buffer);
    const head = decodeBytes(bytes.slice(0, 600));
    const warnings = [];
    if (ext === 'ofx' || /OFXHEADER|<OFX>/i.test(head)) {
      const rows = parseOFX(decodeBytes(bytes));
      return { kind: 'OFX', rows, warnings };
    }
    if (ext === 'pdf' || head.startsWith('%PDF')) {
      if (!libs.pdfjsLib) throw new Error('Leitor de PDF indisponível.');
      const lines = await pdfToLines(libs.pdfjsLib, bytes);
      if (!lines.length) throw new Error('Este PDF não tem texto (parece imagem escaneada). Baixe o extrato em CSV, Excel ou OFX.');
      const rows = pdfLinesToRows(lines, defaultYear);
      if (rows.length && !rows.some(r => r.value < 0)) warnings.push('Não encontrei valores negativos neste PDF: confira se as saídas estão com o sinal certo (use o botão ± da linha).');
      return { kind: 'PDF', rows, lines, warnings };
    }
    let table;
    if (ext === 'xlsx' || ext === 'xls' || ext === 'xlsm' || ext === 'ods' || (bytes[0] === 0x50 && bytes[1] === 0x4b) || (bytes[0] === 0xd0 && bytes[1] === 0xcf)) {
      if (!libs.XLSX) throw new Error('Leitor de Excel indisponível.');
      const wb = libs.XLSX.read(bytes, { type: 'array', cellDates: true, raw: false });
      // escolhe a aba com mais linhas que parecem lançamentos
      let best = null;
      for (const sn of wb.SheetNames) {
        const rows = libs.XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: '' });
        const t = detectTable(rows, defaultYear);
        const n = rowsFromTable(t, t.map, defaultYear).length;
        if (!best || n > best.n) best = { t, n, sn };
      }
      table = best.t;
      table.sheet = best.sn;
    } else {
      table = detectTable(parseCSV(decodeBytes(bytes)), defaultYear);
    }
    const map = Object.assign({}, table.map, opts.mapOverride || {});
    const rows = rowsFromTable(table, map, defaultYear);
    return { kind: ext === 'csv' || ext === 'txt' ? 'CSV' : 'Excel', rows, table, map, warnings };
  }

  // ---------------------------------------------------------------- classificação (espelha as fórmulas da planilha)
  function wildcardToRegex(kw) {
    let r = '';
    for (const ch of String(kw)) {
      if (ch === '?') r += '.';
      else if (ch === '*') r += '.*';
      else r += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    return new RegExp(r, 'i');
  }
  function buildClassifier(rulesRows, importRows) {
    const rules = [];
    (rulesRows || []).forEach((r, i) => {
      const kw = String(r[0] == null ? '' : r[0]).trim();
      if (!kw) return;
      rules.push({ i, kw, re: wildcardToRegex(kw), tipo: String(r[1] || '').trim().toUpperCase(), cat: String(r[2] || '').trim(), serv: String(r[3] || '').trim(), desc: String(r[4] || '').trim() });
    });
    const memCat = new Map(), memServ = new Map(), memDesc = new Map();
    (importRows || []).forEach(r => {
      const b = r.descExtrato;
      if (!b) return;
      if (r.cat && !memCat.has(b)) memCat.set(b, r.cat);
      if (r.serv && !memServ.has(b)) memServ.set(b, r.serv);
      if (r.descPadrao && !memDesc.has(b)) memDesc.set(b, r.descPadrao);
    });
    // mesma ordem das fórmulas da planilha: escolha manual > memória (mesma descrição) > regra
    return function classify(desc, value) {
      const tipo = value < 0 ? 'DESPESA' : 'RECEITA';
      const rule = rules.find(r => (r.tipo === '' || r.tipo === tipo) && r.re.test(desc));
      const res = { tipo, cat: '', serv: '', descFinal: desc, origem: '' };
      if (memCat.has(desc)) { res.cat = memCat.get(desc); res.origem = 'APRENDIDO'; }
      else if (rule && rule.cat) { res.cat = rule.cat; res.origem = 'REGRA'; res.rule = rule.kw; }
      if (memServ.has(desc)) res.serv = memServ.get(desc); else if (rule && rule.serv) res.serv = rule.serv;
      if (memDesc.has(desc)) res.descFinal = memDesc.get(desc); else if (rule && rule.desc) res.descFinal = rule.desc;
      if (!res.cat) res.origem = 'CLASSIFICAR';
      return res;
    };
  }

  /** quantos de cada lançamento já existem (data|descrição|valor), para não lançar em dobro */
  function keyOf(date, desc, value) { return date + '|' + norm(desc) + '|' + Number(value).toFixed(2); }
  function dedupe(newRows, existingRows) {
    const have = new Map();
    existingRows.forEach(r => { const k = keyOf(r.date, r.desc, r.value); have.set(k, (have.get(k) || 0) + 1); });
    const seen = new Map();
    return newRows.map(r => {
      const k = keyOf(r.date, r.desc, r.value);
      const n = (seen.get(k) || 0) + 1; seen.set(k, n);
      return Object.assign({}, r, { duplicate: n <= (have.get(k) || 0) });
    });
  }

  const api = { parseDate, parseMoney, parseCSV, decodeBytes, detectTable, rowsFromTable, parseOFX, pdfLinesToRows, pdfToLines,
    readStatement, buildClassifier, dedupe, excelSerial, serialToYMD, MESES, norm };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.RoboParser = api;
})(typeof self !== 'undefined' ? self : this);
