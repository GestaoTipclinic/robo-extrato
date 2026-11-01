/* Robô de Extrato — painel do Excel */
(function () {
  'use strict';
  const P = window.RoboParser;
  const SHEET_IMP = 'IMPORTAR EXTRATO INF';
  const SHEET_REG = 'REGRAS EXTRATO INF';
  const SHEET_CFG = 'CONFIGURAÇÃO INF';
  const MONTH_SHEETS = P.MESES.map(m => m + ' CONCILIAÇÃO INF');
  const FIRST = 10, LAST = 3009;           // linhas da aba IMPORTAR EXTRATO
  const RULE_FIRST = 8, RULE_LAST = 307;

  if (window.pdfjsLib) window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';

  const $ = id => document.getElementById(id);
  const fmtMoney = v => (v < 0 ? '-' : '') + 'R$ ' + Math.abs(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fmtDate = s => { const d = P.serialToYMD(s); return String(d.d).padStart(2, '0') + '/' + String(d.m).padStart(2, '0') + '/' + d.y; };
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  // ------------------------------------------------------------ acesso à planilha (Office.js) — trocável por um simulador nos testes
  const OfficeWB = {
    async load() {
      return Excel.run(async ctx => {
        const wb = ctx.workbook;
        const imp = wb.worksheets.getItemOrNullObject(SHEET_IMP);
        const reg = wb.worksheets.getItemOrNullObject(SHEET_REG);
        const cfg = wb.worksheets.getItemOrNullObject(SHEET_CFG);
        await ctx.sync();
        if (imp.isNullObject || reg.isNullObject || cfg.isNullObject) throw new Error('Esta planilha não tem as abas IMPORTAR EXTRATO, REGRAS e CONFIGURAÇÃO. Abra a planilha de conciliação correta (versão nova).');
        const inRange = imp.getRange(`A${FIRST}:N${LAST}`); inRange.load('values');
        const ano = imp.getRange('Q4'); ano.load('values');
        const rules = reg.getRange(`B${RULE_FIRST}:F${RULE_LAST}`); rules.load('values');
        const docs = cfg.getRange('B7:B10'); docs.load('values');
        const lists = cfg.getRange('X16:Z35'); lists.load('values');
        await ctx.sync();
        const col = i => lists.values.map(r => String(r[i] || '').trim()).filter(Boolean);
        return {
          imp: inRange.values, year: +ano.values[0][0] || new Date().getFullYear(), rules: rules.values,
          docs: docs.values.map(r => String(r[0] || '').trim()).filter(Boolean),
          cats: { DESPESA: col(0), RECEITA: col(1), SERV: col(2) }
        };
      });
    },
    async write(startRow, values) {
      return Excel.run(async ctx => {
        const ws = ctx.workbook.worksheets.getItem(SHEET_IMP);
        const r = ws.getRange(`A${startRow}:K${startRow + values.length - 1}`);
        r.values = values;
        await ctx.sync();
      });
    },
    async activate(name) {
      return Excel.run(async ctx => { ctx.workbook.worksheets.getItem(name).activate(); await ctx.sync(); });
    },
    getSetting(k) { try { return Office.context.document.settings.get(k); } catch (e) { return null; } },
    setSetting(k, v) { try { Office.context.document.settings.set(k, v); Office.context.document.settings.saveAsync(); } catch (e) { /* ignore */ } }
  };
  let WB = window.__MOCK_WB__ || OfficeWB;

  // ------------------------------------------------------------ estado
  const S = { book: null, file: null, parsed: null, rows: [], month: null, filter: 'all', lastMonthSheet: null };

  function setStatus(msg, cls) { const el = $('status'); el.textContent = msg; el.className = 'status ' + (cls || ''); }

  async function connect() {
    try {
      S.book = await WB.load();
      const used = S.book.imp.filter(r => r[0] !== '' || r[1] !== '' || r[2] !== '').length;
      setStatus(`Planilha conectada · ano ${S.book.year} · ${S.book.docs.length > 1 ? S.book.docs.length + ' médicos · ' : ''}${used} linhas de extrato já importadas`, 'ok');
      return true;
    } catch (e) {
      setStatus(e.message || String(e), 'err');
      return false;
    }
  }

  function existingRows() {
    // usa as colunas calculadas (I = data, J = valor) e B = descrição original
    return S.book.imp
      .filter(r => typeof r[12] === 'number' && typeof r[13] === 'number')
      .map(r => ({ date: Math.floor(r[12]), desc: String(r[1]), value: r[13], cat: String(r[3] || ''), serv: String(r[4] || ''), descPadrao: String(r[5] || '') }));
  }
  function memoryRows() {
    return S.book.imp.map(r => ({ descExtrato: String(r[1] || ''), cat: String(r[3] || ''), serv: String(r[4] || ''), descPadrao: String(r[5] || '') }));
  }

  // ------------------------------------------------------------ passo 1: arquivo
  const drop = $('drop'), fileInput = $('file');
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('over'));
  drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('over'); if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); });
  fileInput.addEventListener('change', () => { if (fileInput.files[0]) handleFile(fileInput.files[0]); });

  function headerSig(t) { return t && t.header && t.header.length ? t.header.map(P.norm).join('|') : null; }

  async function handleFile(file, mapOverride) {
    $('filewarn').innerHTML = ''; $('stepDone').classList.add('hidden');
    if (!S.book && !(await connect())) return;
    S.file = file;
    $('fileinfo').classList.remove('hidden');
    $('fileinfo').textContent = 'Lendo ' + file.name + '…';
    try {
      const buf = await file.arrayBuffer();
      let parsed = await P.readStatement(file.name, buf, { XLSX: window.XLSX, pdfjsLib: window.pdfjsLib }, { year: S.book.year });
      // layout de banco já ajustado antes?
      const sig = headerSig(parsed.table);
      const maps = WB.getSetting('robo_maps') || {};
      const remembered = mapOverride || (sig && maps[sig]);
      if (remembered && parsed.table) {
        parsed = await P.readStatement(file.name, buf, { XLSX: window.XLSX, pdfjsLib: window.pdfjsLib }, { year: S.book.year, mapOverride: remembered });
      }
      S.buf = buf; S.parsed = parsed;
      if (!parsed.rows.length) {
        $('fileinfo').innerHTML = `<b>${esc(file.name)}</b>`;
        $('filewarn').innerHTML = '<div class="err">Não encontrei lançamentos neste arquivo. ' + (parsed.table ? 'Use “Ajustar colunas” para indicar onde estão a data, a descrição e o valor.' : 'Tente baixar o extrato em outro formato (CSV, Excel ou OFX).') + '</div>';
      } else {
        $('fileinfo').innerHTML = `<b>${esc(file.name)}</b> · ${parsed.kind} · <b>${parsed.rows.length}</b> lançamentos lidos`;
      }
      parsed.warnings.forEach(w => $('filewarn').insertAdjacentHTML('beforeend', '<div class="warn">' + esc(w) + '</div>'));
      buildMapUI(parsed);
      if (parsed.rows.length) buildMonths();
    } catch (e) {
      $('fileinfo').innerHTML = `<b>${esc(file.name)}</b>`;
      $('filewarn').innerHTML = '<div class="err">' + esc(e.message || String(e)) + '</div>';
      ['step2', 'step3', 'footer'].forEach(id => $(id).classList.add('hidden'));
    }
  }

  const MAP_FIELDS = [['date', 'Data'], ['desc', 'Descrição'], ['desc2', 'Complemento'], ['value', 'Valor'], ['credit', 'Crédito'], ['debit', 'Débito'], ['type', 'Coluna D/C']];
  function buildMapUI(parsed) {
    const box = $('mapbox'), grid = $('mapgrid');
    if (!parsed.table) { box.classList.add('hidden'); return; }
    box.classList.remove('hidden');
    if (!parsed.rows.length) box.open = true;
    const t = parsed.table;
    const names = [];
    for (let c = 0; c < t.ncol; c++) {
      const ex = (t.body.find(r => r[c] !== '' && r[c] != null) || [])[c];
      names.push((t.header[c] || ('Coluna ' + String.fromCharCode(65 + c))) + (ex != null && ex !== '' ? ' (ex.: ' + String(ex instanceof Date ? ex.toLocaleDateString('pt-BR') : ex).slice(0, 18) + ')' : ''));
    }
    grid.innerHTML = MAP_FIELDS.map(([k, lbl]) =>
      `<label>${lbl}</label><select data-k="${k}"><option value="-1">—</option>${names.map((n, i) => `<option value="${i}" ${parsed.map[k] === i ? 'selected' : ''}>${esc(n)}</option>`).join('')}</select>`).join('');
  }
  $('applymap').addEventListener('click', async () => {
    const m = {};
    $('mapgrid').querySelectorAll('select').forEach(s => m[s.dataset.k] = +s.value);
    const sig = headerSig(S.parsed && S.parsed.table);
    if (sig) { const maps = WB.getSetting('robo_maps') || {}; maps[sig] = m; WB.setSetting('robo_maps', maps); }
    await handleFile(S.file, m);
  });

  // ------------------------------------------------------------ passo 2: mês
  function buildMonths() {
    const counts = {};
    S.parsed.rows.forEach(r => { const d = P.serialToYMD(r.date); const k = d.y + '-' + String(d.m).padStart(2, '0'); counts[k] = (counts[k] || 0) + 1; });
    const keys = Object.keys(counts).sort();
    const best = keys.slice().sort((a, b) => counts[b] - counts[a])[0];
    const sel = $('month');
    sel.innerHTML = keys.map(k => { const [y, m] = k.split('-').map(Number);
      return `<option value="${k}" ${k === best ? 'selected' : ''}>${P.MESES[m - 1]} / ${y} — ${counts[k]} lançamento(s)</option>`; }).join('')
      + (keys.length > 1 ? `<option value="all">Todos os meses do arquivo — ${S.parsed.rows.length} lançamento(s)</option>` : '');
    $('step2').classList.remove('hidden');
    onMonth();
  }
  $('month').addEventListener('change', onMonth);
  function onMonth() {
    S.month = $('month').value;
    const yrs = new Set(S.parsed.rows.map(r => P.serialToYMD(r.date).y));
    const other = [...yrs].filter(y => y !== S.book.year);
    $('yearwarn').innerHTML = other.length ? `<div class="warn">A planilha está no ano <b>${S.book.year}</b>. Lançamentos de ${other.join(', ')} ficam na aba IMPORTAR EXTRATO marcados como “OUTRO ANO” e não entram nos meses. Para trocar o ano, altere a célula M4 da aba IMPORTAR EXTRATO.</div>` : '';
    prepareRows();
  }

  // ------------------------------------------------------------ passo 3: prévia e classificação
  function prepareRows() {
    let rows = S.parsed.rows;
    if (S.month !== 'all') {
      const [y, m] = S.month.split('-').map(Number);
      rows = rows.filter(r => { const d = P.serialToYMD(r.date); return d.y === y && d.m === m; });
    }
    const classify = P.buildClassifier(S.book.rules, memoryRows());
    rows = P.dedupe(rows, existingRows());
    S.rows = rows.map((r, i) => {
      const c = classify(r.desc, r.value);
      return Object.assign({}, r, { id: i, include: !r.duplicate, sug: c, cat: c.cat, serv: c.serv, changed: false, pct: S.book.docs.map(() => null) });
    });
    $('step3').classList.remove('hidden'); $('footer').classList.remove('hidden');
    render();
  }
  function reclassify(row) {
    const classify = P.buildClassifier(S.book.rules, memoryRows());
    const c = classify(row.desc, row.value);
    row.sug = c; if (!row.changed) { row.cat = c.cat; row.serv = c.serv; }
  }

  $('tabs').addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    S.filter = b.dataset.f;
    $('tabs').querySelectorAll('button').forEach(x => x.classList.toggle('on', x === b));
    render();
  });

  function render() {
    const list = $('list');
    const show = S.rows.filter(r => S.filter === 'all' || (S.filter === 'todo' ? !r.duplicate && (!r.cat || !hasSplit(r)) : r.duplicate));
    list.innerHTML = show.length ? show.map(r => {
      const tipo = r.value < 0 ? 'DESPESA' : 'RECEITA';
      const opts = S.book.cats[tipo];
      const origin = r.duplicate ? 'DUP' : (r.changed ? 'MANUAL' : (r.cat ? r.sug.origem : 'CLASSIFICAR'));
      const chipTxt = { DUP: 'já lançado', MANUAL: 'manual', REGRA: 'regra', APRENDIDO: 'aprendido', CLASSIFICAR: 'classificar' }[origin];
      return `<div class="item ${r.duplicate ? 'dup' : ''}" data-id="${r.id}">
        <input type="checkbox" data-a="inc" ${r.include ? 'checked' : ''} title="Incluir">
        <div><div class="d">${fmtDate(r.date)} · ${tipo === 'DESPESA' ? 'saída' : 'entrada'}</div><div class="t">${esc(r.sug.descFinal)}</div>${r.sug.descFinal !== r.desc ? `<div class="o">extrato: ${esc(r.desc)}</div>` : ''}</div>
        <div class="v ${r.value < 0 ? 'neg' : 'pos'}">${fmtMoney(r.value)} <button class="sign" data-a="sign" title="Inverter entrada/saída">±</button></div>
        <div class="cat"><select data-a="cat"><option value="">— escolher categoria —</option>${opts.map(o => `<option ${o === r.cat ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>
          <span class="chip ${origin}">${chipTxt}</span></div>
        ${splitHTML(r)}
      </div>`;
    }).join('') : '<div class="item"><span></span><span class="muted">Nada aqui.</span></div>';
    $('splitBar').classList.toggle('hidden', !multi());
    if (multi()) $('splitMode').innerHTML = '<option value="">Dividir os lançamentos…</option><option value="eq">Igual entre todos os médicos</option>' + S.book.docs.map((d, i) => `<option value="${i}">100% para ${esc(d)}</option>`).join('');
    updateSummary();
  }
  function updateSummary() {
    const inc = S.rows.filter(r => r.include);
    $('sNew').textContent = inc.length;
    $('sTodo').textContent = inc.filter(r => !r.cat).length + (multi() ? inc.filter(r => !hasSplit(r)).length : 0);
    $('sDup').textContent = S.rows.filter(r => r.duplicate).length;
    const bad = multi() ? inc.filter(r => hasSplit(r) && Math.abs(pctSum(r) - 100) >= 0.01).length : 0;
    const label = S.month === 'all' ? 'nos meses do arquivo' : 'em ' + P.MESES[+S.month.split('-')[1] - 1];
    $('launch').textContent = bad ? `${bad} divisão(ões) não somam 100% — corrija` : (inc.length ? `Lançar ${inc.length} lançamento(s) ${label}` : 'Nada novo para lançar');
    $('launch').disabled = !inc.length || bad > 0;
  }

  function multi() { return S.book && S.book.docs.length > 1; }
  function pctSum(r) { return r.pct.reduce((a, b) => a + (+b || 0), 0); }
  function hasSplit(r) { return !multi() || r.pct.some(v => v !== null && v !== '' && +v > 0); }
  function splitHTML(r) {
    if (!multi() || r.duplicate) return '';
    const sum = pctSum(r), ok = !hasSplit(r) || Math.abs(sum - 100) < 0.01;
    return `<div class="split">${S.book.docs.map((d, i) => `<label title="${esc(d)}">${esc(initials(d))}<input type="number" min="0" max="100" step="1" data-a="pct" data-d="${i}" value="${r.pct[i] == null ? '' : r.pct[i]}" placeholder="%"></label>`).join('')}
      <span class="sum ${ok ? '' : 'bad'}">${hasSplit(r) ? (ok ? '100%' : sum + '%') : 'dividir'}</span></div>`;
  }
  function initials(n) {
    const w = n.replace(/^(dra?\.?|dr\.?)\s+/i, '').split(/\s+/).filter(Boolean);
    return (w[0] || n).slice(0, 10);
  }
  $('splitAll').addEventListener('click', () => {
    const mode = $('splitMode').value; if (!mode) return;
    const n = S.book.docs.length;
    S.rows.forEach(r => {
      if (!r.include || r.duplicate) return;
      if ($('splitOnlyEmpty').checked && hasSplit(r)) return;
      if (mode === 'eq') { const base = Math.floor(10000 / n) / 100; r.pct = S.book.docs.map((_, i) => i === n - 1 ? Math.round((100 - base * (n - 1)) * 100) / 100 : base); }
      else { const k = +mode; r.pct = S.book.docs.map((_, i) => i === k ? 100 : 0); }
    });
    render();
  });
  $('list').addEventListener('input', e => {
    if (e.target.dataset.a !== 'pct') return;
    const row = S.rows[+e.target.closest('.item').dataset.id];
    row.pct[+e.target.dataset.d] = e.target.value === '' ? null : Math.max(0, Math.min(100, +e.target.value));
    const sp = e.target.closest('.split').querySelector('.sum');
    const sum = pctSum(row), ok = !hasSplit(row) || Math.abs(sum - 100) < 0.01;
    sp.textContent = hasSplit(row) ? (ok ? '100%' : sum + '%') : 'dividir'; sp.className = 'sum ' + (ok ? '' : 'bad');
    updateSummary();
  });
  $('list').addEventListener('change', e => {
    if (e.target.dataset.a === 'pct') return;
    const it = e.target.closest('.item'); if (!it) return;
    const row = S.rows[+it.dataset.id];
    if (e.target.dataset.a === 'inc') row.include = e.target.checked;
    if (e.target.dataset.a === 'cat') {
      row.cat = e.target.value; row.changed = row.cat !== row.sug.cat;
      // mesma descrição na lista recebe a mesma categoria (aprende na hora)
      S.rows.forEach(o => { if (o !== row && o.desc === row.desc && Math.sign(o.value) === Math.sign(row.value) && !o.changed) { o.cat = row.cat; o.changed = row.changed; } });
    }
    render();
  });
  $('list').addEventListener('click', e => {
    if (e.target.dataset.a !== 'sign') return;
    const row = S.rows[+e.target.closest('.item').dataset.id];
    row.value = -row.value; row.changed = false; reclassify(row); render();
  });

  // ------------------------------------------------------------ lançar
  $('launch').addEventListener('click', async () => {
    const inc = S.rows.filter(r => r.include);
    if (!inc.length) return;
    const btn = $('launch'); btn.disabled = true; btn.textContent = 'Lançando…';
    try {
      await connect();                                   // relê a planilha (outra pessoa pode ter lançado)
      const lastUsed = S.book.imp.reduce((acc, r, i) => (r[0] !== '' || r[1] !== '' || r[2] !== '') ? i : acc, -1);
      const start = FIRST + lastUsed + 1;
      if (start + inc.length - 1 > LAST) throw new Error(`A aba IMPORTAR EXTRATO comporta ${LAST - FIRST + 1} linhas e só restam ${LAST - start + 1}. Fale com a TipClinic para ampliar.`);
      const nd = S.book.docs.length;
      const values = inc.map(r => [
        r.date, r.desc, r.value,
        r.changed || (!r.sug.cat && r.cat) ? r.cat : '',  // categoria manual vira aprendizado
        '', '', '',
        ...[0, 1, 2, 3].map(i => (nd > 1 && i < nd && hasSplit(r)) ? Math.round((+r.pct[i] || 0) * 100) / 10000 : '')
      ]);
      await WB.write(start, values);
      const months = [...new Set(inc.map(r => P.serialToYMD(r.date)).filter(d => d.y === S.book.year).map(d => d.m))].sort((a, b) => a - b);
      S.lastMonthSheet = months.length ? MONTH_SHEETS[months[0] - 1] : null;
      const todo = inc.filter(r => !r.cat).length;
      $('doneTitle').textContent = `${inc.length} lançamento(s) feitos`;
      $('doneText').innerHTML = (months.length ? 'Já aparecem em: <b>' + months.map(m => P.MESES[m - 1]).join(', ') + '</b>.' : '')
        + (todo ? `<br>${todo} ficaram sem categoria: escolha na aba IMPORTAR EXTRATO (coluna D, linhas marcadas ⚠).` : '')
        + (multi() && inc.some(r => !hasSplit(r)) ? `<br>${inc.filter(r => !hasSplit(r)).length} ficaram sem divisão entre os médicos (aparecem como “Não dividido”). Dá para completar na aba IMPORTAR EXTRATO, colunas H a K.` : '');
      $('goMonth').classList.toggle('hidden', !S.lastMonthSheet);
      ['step2', 'step3', 'footer'].forEach(id => $(id).classList.add('hidden'));
      $('stepDone').classList.remove('hidden');
      $('fileinfo').classList.add('hidden'); $('mapbox').classList.add('hidden'); $('filewarn').innerHTML = '';
      fileInput.value = '';
      await connect();
    } catch (e) {
      btn.disabled = false; render();
      $('filewarn').innerHTML = '<div class="err">' + esc(e.message || String(e)) + '</div>';
    }
  });
  $('goMonth').addEventListener('click', () => S.lastMonthSheet && WB.activate(S.lastMonthSheet));
  $('again').addEventListener('click', () => { $('stepDone').classList.add('hidden'); fileInput.click(); });

  // ------------------------------------------------------------ início
  if (window.__MOCK_WB__) connect();
  else if (window.Office) Office.onReady(info => {
    if (info.host === Office.HostType.Excel) connect();
    else setStatus('Abra este painel dentro do Excel.', 'err');
  });
  else setStatus('Abra este painel dentro do Excel.', 'err');
})();
