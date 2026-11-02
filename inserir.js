/* Inserir extrato — página sem instalação: lê o extrato e copia linhas prontas para colar na aba IMPORTAR EXTRATO */
(function () {
  'use strict';
  const P = window.RoboParser;
  if (window.pdfjsLib) window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';
  const $ = id => document.getElementById(id);
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const money = v => (v < 0 ? '-' : '') + 'R$ ' + Math.abs(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const ymd = s => P.serialToYMD(s);
  const br = s => { const d = ymd(s); return String(d.d).padStart(2, '0') + '/' + String(d.m).padStart(2, '0') + '/' + d.y; };
  const iso = s => { const d = ymd(s); return d.y + '-' + String(d.m).padStart(2, '0') + '-' + String(d.d).padStart(2, '0'); };
  const store = { get(k) { try { return JSON.parse(localStorage.getItem(k) || 'null'); } catch (e) { return null; } },
                  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* sem armazenamento */ } } };

  const S = { file: null, buf: null, parsed: null, rows: [], month: null, hash: null };
  const qs = new URLSearchParams(location.search);
  const MES = +qs.get('m') || 0;                      // 1..12 quando aberto pelo botão de uma aba mensal
  if (MES) {
    document.querySelector('header p').textContent = 'Lançamentos de ' + P.MESES[MES - 1] + ': envie o arquivo do banco, clique em Copiar e cole na aba do mês.';
    document.title = 'Inserir extrato — ' + P.MESES[MES - 1];
  }
  const EN = () => !!store.get('excel_en');
  const year = new Date().getFullYear();

  const drop = $('drop'), input = $('file');
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('over'));
  drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('over'); if (e.dataTransfer.files[0]) load(e.dataTransfer.files[0]); });
  input.addEventListener('change', () => input.files[0] && load(input.files[0]));

  async function hashOf(buf) {
    try { const h = await crypto.subtle.digest('SHA-256', buf); return [...new Uint8Array(h)].slice(0, 8).map(b => b.toString(16).padStart(2, '0')).join(''); }
    catch (e) { return String(buf.byteLength); }
  }
  const sigOf = t => t && t.header && t.header.length ? t.header.map(P.norm).join('|') : null;

  async function load(file, mapOverride) {
    $('msgs').innerHTML = ''; $('done').classList.add('hidden');
    $('info').classList.remove('hidden'); $('info').textContent = 'Lendo ' + file.name + '…';
    try {
      const buf = mapOverride ? S.buf : await file.arrayBuffer();
      if (!mapOverride) S.password = null;
      S.file = file; S.buf = buf;
      const libs = { XLSX: window.XLSX, pdfjsLib: window.pdfjsLib };
      let parsed;
      for (let tries = 0; ; tries++) {
        try { parsed = await P.readStatement(file.name, buf, libs, { year, password: S.password }); break; }
        catch (e) {
          if (e && e.name === 'PasswordException' && tries < 3) {
            const pw = prompt(tries ? 'Senha incorreta. Digite a senha do PDF (muitos bancos usam o CPF ou CNPJ, só números):' : 'Este PDF tem senha. Digite a senha (muitos bancos usam o CPF ou CNPJ, só números):');
            if (pw === null) throw new Error('PDF protegido por senha: sem a senha não é possível ler.');
            S.password = pw; continue;
          }
          throw e;
        }
      }
      const maps = store.get('robo_maps') || {};
      const remembered = mapOverride || (parsed.table && maps[sigOf(parsed.table)]);
      if (remembered && parsed.table) parsed = await P.readStatement(file.name, buf, libs, { year, mapOverride: remembered, password: S.password });
      S.parsed = parsed; S.hash = await hashOf(buf);
      $('info').innerHTML = `<b>${esc(file.name)}</b> · ${parsed.kind} · <b>${parsed.rows.length}</b> lançamentos encontrados`;
      parsed.warnings.forEach(w => $('msgs').insertAdjacentHTML('beforeend', `<div class="warn">${esc(w)}</div>`));
      const seen = (store.get('robo_copiados') || {})[S.hash];
      if (seen) $('msgs').insertAdjacentHTML('beforeend', `<div class="warn">Este mesmo arquivo já foi copiado neste computador em ${esc(seen)}. Se já colou na planilha, não cole de novo${MES ? ', senão os lançamentos ficam em dobro' : ' (a planilha também avisa “⚠ DUPLICADO”)'}.</div>`);
      buildMap(parsed);
      if (!parsed.rows.length) {
        $('msgs').insertAdjacentHTML('beforeend', `<div class="err">Não encontrei lançamentos neste arquivo. ${parsed.table ? 'Abra “Ajustar aqui” e indique as colunas de data, descrição e valor.' : 'Tente baixar o extrato do banco em outro formato (Excel, CSV ou OFX).'}</div>`);
        ['c2', 'c3', 'bar'].forEach(id => $(id).classList.add('hidden'));
        return;
      }
      buildMonths();
    } catch (e) {
      $('info').innerHTML = `<b>${esc(file.name)}</b>`;
      $('msgs').innerHTML = `<div class="err">${esc(e.message || String(e))}</div>`;
      ['c2', 'c3', 'bar'].forEach(id => $(id).classList.add('hidden'));
    }
  }

  const FIELDS = [['date', 'Data'], ['desc', 'Descrição'], ['desc2', 'Complemento'], ['value', 'Valor'], ['credit', 'Crédito (entradas)'], ['debit', 'Débito (saídas)'], ['type', 'Coluna D/C']];
  function buildMap(parsed) {
    const box = $('mapbox');
    if (!parsed.table) { box.classList.add('hidden'); return; }
    box.classList.remove('hidden'); if (!parsed.rows.length) box.open = true;
    const t = parsed.table, names = [];
    for (let c = 0; c < t.ncol; c++) {
      const ex = (t.body.find(r => r[c] !== '' && r[c] != null) || [])[c];
      names.push((t.header[c] || 'Coluna ' + String.fromCharCode(65 + c)) + (ex != null && ex !== '' ? ' (ex.: ' + String(ex instanceof Date ? ex.toLocaleDateString('pt-BR') : ex).slice(0, 20) + ')' : ''));
    }
    $('mapgrid').innerHTML = FIELDS.map(([k, l]) => `<label>${l}</label><select data-k="${k}"><option value="-1">—</option>${names.map((n, i) => `<option value="${i}" ${parsed.map[k] === i ? 'selected' : ''}>${esc(n)}</option>`).join('')}</select>`).join('');
  }
  $('mapgrid').addEventListener('change', () => {
    const m = {}; $('mapgrid').querySelectorAll('select').forEach(s => m[s.dataset.k] = +s.value);
    const sig = sigOf(S.parsed.table);
    if (sig) { const maps = store.get('robo_maps') || {}; maps[sig] = m; store.set('robo_maps', maps); }
    load(S.file, m);
  });

  function buildMonths() {
    const cnt = {};
    S.parsed.rows.forEach(r => { const d = ymd(r.date); const k = d.y + '-' + String(d.m).padStart(2, '0'); cnt[k] = (cnt[k] || 0) + 1; });
    const keys = Object.keys(cnt).sort();
    let best = keys.slice().sort((a, b) => cnt[b] - cnt[a])[0];
    if (MES) {
      const ofMonth = keys.filter(k => +k.split('-')[1] === MES).sort((a, b) => cnt[b] - cnt[a]);
      if (!ofMonth.length) {
        $('msgs').insertAdjacentHTML('beforeend', `<div class="warn">Este extrato não tem lançamentos de <b>${P.MESES[MES - 1]}</b>. Escolha abaixo o mês que deseja copiar (ele deve ser colado na aba desse mês).</div>`);
      } else best = ofMonth[0];
    }
    $('month').innerHTML = keys.map(k => { const [y, m] = k.split('-').map(Number); return `<option value="${k}" ${k === best ? 'selected' : ''}>${P.MESES[m - 1]} / ${y} — ${cnt[k]} lançamento(s)</option>`; }).join('')
      + (keys.length > 1 ? `<option value="all">Todos os meses do arquivo — ${S.parsed.rows.length} lançamento(s)</option>` : '');
    $('c2').classList.remove('hidden');
    pickMonth();
  }
  $('month').addEventListener('change', pickMonth);
  function pickMonth() {
    S.month = $('month').value;
    let rows = S.parsed.rows;
    if (S.month !== 'all') { const [y, m] = S.month.split('-').map(Number); rows = rows.filter(r => { const d = ymd(r.date); return d.y === y && d.m === m; }); }
    rows = rows.map((r, k) => ({ r, k })).sort((a, b) => a.r.date - b.r.date || a.k - b.k).map(x => x.r);   // ordem cronológica
    S.rows = rows.map((r, i) => Object.assign({}, r, { id: i, on: true }));
    $('c3').classList.remove('hidden'); $('bar').classList.remove('hidden');
    render();
  }
  function render() {
    $('rows').innerHTML = S.rows.map(r => `<tr class="${r.on ? '' : 'off'}" data-id="${r.id}">
      <td><input type="checkbox" ${r.on ? 'checked' : ''}></td><td>${br(r.date)}</td><td>${esc(r.desc)}</td>
      <td class="v ${r.value < 0 ? 'neg' : 'pos'}">${money(r.value)}<button class="sign" title="Inverter entrada/saída">±</button></td></tr>`).join('');
    const on = S.rows.filter(r => r.on);
    const ent = on.filter(r => r.value > 0).reduce((a, r) => a + r.value, 0), sai = on.filter(r => r.value < 0).reduce((a, r) => a + r.value, 0);
    $('totals').textContent = `${on.length} lançamento(s) · entradas ${money(ent)} · saídas ${money(sai)}`;
    $('copy').disabled = !on.length;
    $('copy').textContent = on.length ? `Copiar ${on.length} lançamento(s)` : 'Nada selecionado';
    $('barinfo').textContent = MES ? 'Depois: na aba do mês, “▸ COLAR NA LINHA” e Ctrl+V' : 'Depois: na planilha, “② Ir p/ linha vazia” e Ctrl+V';
  }
  $('rows').addEventListener('click', e => {
    const tr = e.target.closest('tr'); if (!tr) return;
    const r = S.rows[+tr.dataset.id];
    if (e.target.classList.contains('sign')) r.value = -r.value;
    else if (e.target.type === 'checkbox') r.on = e.target.checked;
    else return;
    render();
  });

  // aba mensal: colunas C..J (DATA, DESCRIÇÃO, DESPESA, VALOR DESPESA, VALOR RECEITA, RECEITA, SERVIÇO, OBSERVAÇÃO=texto do extrato)
  // DESCRIÇÃO/DESPESA/RECEITA/SERVIÇO entram como fórmula que busca o que já foi definido antes para o mesmo texto do extrato
  function buildMonthTSV() {
    const f = col => EN() ? `=INDIRECT("${col}"&ROW())` : `=INDIRETO("${col}"&LIN())`;
    const clean = s => String(s).replace(/[\t\r\n]+/g, ' ').trim().replace(/^[=+\-@]/, "'$&");
    const num = v => v.toFixed(2).replace('.', EN() ? '.' : ',');
    return S.rows.filter(r => r.on).map(r => [iso(r.date), f('X'), r.value < 0 ? f('Y') : '', r.value < 0 ? num(r.value) : '', r.value > 0 ? num(r.value) : '',
      r.value > 0 ? f('Z') : '', f('AA'), clean(r.desc)].join('\t')).join('\r\n');
  }
  // linhas no formato das colunas A..L da aba IMPORTAR EXTRATO (D..K vazias; L = lote)
  function buildTSV() {
    if (MES) return buildMonthTSV();
    const now = new Date();
    const lote = 'R' + now.getFullYear() + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0')
      + String(now.getHours()).padStart(2, '0') + String(now.getMinutes()).padStart(2, '0') + String(now.getSeconds()).padStart(2, '0');
    const clean = s => String(s).replace(/[\t\r\n]+/g, ' ').trim();
    const num = v => v.toFixed(2).replace('.', ',');
    return S.rows.filter(r => r.on).map(r => [iso(r.date), clean(r.desc), num(r.value), '', '', '', '', '', '', '', '', lote].join('\t')).join('\r\n');
  }
  $('en').checked = EN();
  $('en').addEventListener('change', e => store.set('excel_en', e.target.checked));
  $('copy').addEventListener('click', async () => {
    const text = buildTSV();
    let ok = false;
    try { await navigator.clipboard.writeText(text); ok = true; } catch (e) { /* tenta o método antigo */ }
    if (!ok) { const ta = $('clip'); ta.value = text; ta.select(); try { ok = document.execCommand('copy'); } catch (e) { ok = false; } }
    if (!ok) { $('msgs').insertAdjacentHTML('beforeend', '<div class="err">O navegador bloqueou a cópia. Tente de novo ou use outro navegador (Chrome ou Edge).</div>'); return; }
    const n = S.rows.filter(r => r.on).length;
    const cop = store.get('robo_copiados') || {}; cop[S.hash] = new Date().toLocaleString('pt-BR'); store.set('robo_copiados', cop);
    $('doneTitle').textContent = `✅ ${n} lançamento(s) copiados`;
    if (MES) $('doneSteps').innerHTML = `<li>Volte para a planilha, na aba <b>${P.MESES[+S.month.split('-')[1] - 1] || P.MESES[MES - 1]}</b>.</li><li>Clique em <b>▸ COLAR NA LINHA…</b> (abaixo do botão verde).</li><li>Aperte <kbd>Ctrl</kbd> + <kbd>V</kbd> (no Mac, <kbd>⌘</kbd> + <kbd>V</kbd>).</li>`;
    $('done').classList.remove('hidden'); $('done').scrollIntoView({ behavior: 'smooth' });
    $('copy').textContent = 'Copiado ✓ — copiar de novo';
  });
})();
