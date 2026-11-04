(() => {
  'use strict';

  const SHEET_ID = '1fRqUo8vH4awjCwV12U0fhR2bdBSRGFUVMlU8PozUsoQ';
  const GID = '62175904';
  const POLL_MS = 20000;
  const TIMEOUT_MS = 15000;
  const VENDOR = 'assets/vendor/';

  const $ = id => document.getElementById(id);
  const nf = new Intl.NumberFormat('pt-BR');
  const fmt = n => (n === null || n === undefined ? '' : nf.format(n));
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const plural = (n, one, many) => `${fmt(n)} ${n === 1 ? one : many}`;

  const state = {
    rows: [],
    status: 'pendentes',
    date: '',
    q: '',
    sortKey: null,
    sortDir: 1,
    lastSig: null,
    busy: false,
    flash: new Set(),
    loaded: false,
  };

  /* ---------- Leitura da planilha ---------- */

  function isRed(color) {
    let m = String(color).trim().match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
    let rgb;
    if (m) {
      const h = m[1].length === 3 ? [...m[1]].map(c => c + c).join('') : m[1];
      rgb = [0, 2, 4].map(i => parseInt(h.slice(i, i + 2), 16));
    } else {
      m = String(color).match(/rgba?\((\d+)\D+(\d+)\D+(\d+)/i);
      if (!m) return false;
      rgb = m.slice(1, 4).map(Number);
    }
    const [r, g, b] = rgb.map(v => v / 255);
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    const light = (max + min) / 2;
    if (max !== r || max === min) return false;
    const sat = (max - min) / (1 - Math.abs(2 * light - 1));
    const hue = 60 * (((g - b) / (max - min) + 6) % 6);
    return sat > 0.4 && light >= 0.18 && light <= 0.75 && (hue <= 15 || hue >= 345);
  }

  const cellText = v => String(v ?? '').replace(/\s+/g, ' ').trim();

  // Acha cada coluna pelo nome; se o nome mudar, usa a posição original.
  function mapColumns(cols) {
    const find = (re, pos) => {
      const i = cols.findIndex(c => re.test(c.label || ''));
      return i >= 0 ? i : pos;
    };
    return {
      sku: find(/SKU/i, 0),
      desc: find(/DESC/i, 1),
      un: find(/UNID/i, 2),
      cx: find(/CAIXA/i, 3),
      date: find(/DATA/i, 4),
      obs: find(/OBS/i, 5),
      sent: find(/ENVIAD/i, 6),
    };
  }

  const pad = n => String(n).padStart(2, '0');
  const makeDate = (y, m, d) => ({ y, m, d, iso: `${y}-${pad(m)}-${pad(d)}`, label: `${pad(d)}/${pad(m)}/${y}` });

  function rawValue(cell) {
    const v = cell.value;
    if (v && typeof v === 'object' && !(v instanceof Date)) {
      if (Array.isArray(v.richText)) return v.richText.map(t => t.text).join('');
      if ('formula' in v || 'sharedFormula' in v) {
        const r = v.result;
        return r && typeof r === 'object' && !(r instanceof Date) ? '' : r ?? '';
      }
      if ('text' in v) return v.text;
      return ''; // erro de fórmula ou célula sem valor
    }
    return v;
  }

  function parseDate(v) {
    if (v instanceof Date) return makeDate(v.getUTCFullYear(), v.getUTCMonth() + 1, v.getUTCDate());
    const t = cellText(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (!t) return null;
    return makeDate(+t[3] < 100 ? 2000 + +t[3] : +t[3], +t[2], +t[1]);
  }

  const toNumber = v => {
    if (typeof v === 'number') return v;
    const t = cellText(v).replace(/\s/g, '');
    if (!t) return null;
    if (/^\d{1,3}([.,]\d{3})+$/.test(t)) return Number(t.replace(/[.,]/g, ''));
    const n = Number(t.replace(',', '.'));
    return Number.isFinite(n) ? n : null;
  };

  const fontIsRed = font => !!(font && font.color && font.color.argb && isRed(`#${font.color.argb.slice(2)}`));
  const cellIsRed = cell => {
    if (fontIsRed(cell.font)) return true;
    const v = cell.value;
    return !!(v && Array.isArray(v.richText) && v.richText.some(t => cellText(t.text) && fontIsRed(t.font)));
  };

  // A exportação da aba traz tudo: linhas agrupadas, cores e caixas de seleção.
  async function fetchRows() {
    await loadScript('exceljs.min.js');
    const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx&gid=${GID}&_=${Date.now()}`;
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), TIMEOUT_MS);
    let buffer;
    try {
      const res = await fetch(url, { cache: 'no-store', credentials: 'omit', signal: ctrl.signal });
      if (!res.ok) throw new Error(`planilha: ${res.status}`);
      buffer = await res.arrayBuffer();
    } finally {
      clearTimeout(timer);
    }
    const wb = new window.ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.worksheets[0];

    let c = null;
    const rows = [];
    for (let n = 1; n <= ws.rowCount; n++) {
      const row = ws.getRow(n);
      if (!c) {
        const labels = [];
        for (let k = 1; k <= ws.columnCount; k++) labels.push({ label: cellText(rawValue(row.getCell(k))) });
        if (labels.some(l => /^SKU$/i.test(l.label))) c = mapColumns(labels);
        continue;
      }
      const cell = k => row.getCell(c[k] + 1);
      const sent = rawValue(cell('sent'));
      const item = {
        idx: rows.length,
        sku: cellText(rawValue(cell('sku'))),
        desc: cellText(rawValue(cell('desc'))),
        un: toNumber(rawValue(cell('un'))),
        cx: toNumber(rawValue(cell('cx'))),
        date: parseDate(rawValue(cell('date'))),
        obs: cellText(rawValue(cell('obs'))),
        sent: sent === true || /^(TRUE|VERDADEIRO)$/i.test(cellText(sent)),
        grouped: !!(row.hidden || row.outlineLevel > 0),
        red: {},
      };
      ['sku', 'desc', 'un', 'cx', 'date', 'obs'].forEach(k => { item.red[k] = cellIsRed(cell(k)); });
      if (item.sku || item.desc || item.un !== null || item.cx !== null || item.obs) rows.push(item);
    }
    if (!c) throw new Error('cabeçalho SKU não encontrado');
    return rows;
  }

  async function refresh() {
    if (state.busy) return;
    state.busy = true;
    $('btnRefresh').setAttribute('aria-busy', 'true');
    try {
      const rows = await fetchRows();
      const first = !state.loaded;
      state.loaded = true;
      const sig = JSON.stringify(rows);
      if (sig !== state.lastSig || first) {
        if (state.lastSig !== null) markChanges(rows);
        state.rows = rows;
        state.lastSig = sig;
        render();
      }
      setSync(true);
    } catch (err) {
      console.warn(err);
      state.loaded = true;
      if (!state.rows.length) render();
      setSync(false);
    } finally {
      state.busy = false;
      $('btnRefresh').removeAttribute('aria-busy');
    }
  }

  function markChanges(next) {
    const key = r => `${r.sku}|${r.desc}`;
    const prev = new Map(state.rows.map(r => [key(r), JSON.stringify({ ...r, idx: 0 })]));
    state.flash = new Set(
      next.filter(r => prev.get(key(r)) !== JSON.stringify({ ...r, idx: 0 })).map(key)
    );
  }

  function setSync(ok) {
    const time = new Date().toLocaleTimeString('pt-BR');
    $('syncDot').className = ok ? 'on' : 'off';
    if (ok) {
      $('stSync').textContent = `Planilha lida às ${time}. Atualiza sozinho a cada ${POLL_MS / 1000} segundos.`;
      $('alert').hidden = true;
      return;
    }
    $('stSync').textContent = `Última tentativa às ${time} sem resposta.`;
    const box = $('alert');
    box.textContent = state.rows.length
      ? 'Não foi possível ler a planilha agora. A tela mostra a última versão lida; nova tentativa em instantes.'
      : 'Não foi possível ler a planilha. Confira a internet e se a planilha continua compartilhada para quem tem o link.';
    box.hidden = false;
  }

  /* ---------- Filtros ---------- */

  const today = () => {
    const n = new Date();
    return new Date(n.getFullYear(), n.getMonth(), n.getDate());
  };
  const daysUntil = d => Math.round((new Date(d.y, d.m - 1, d.d) - today()) / 86400000);

  const inDate = r => !state.date || (r.date ? r.date.iso : '') === state.date;
  const isHistory = () => state.status === 'historico';
  const inTab = r => (isHistory() ? r.grouped : !r.grouped);

  function visibleRows() {
    const q = state.q.trim().toLocaleLowerCase('pt-BR');
    let list = state.rows.filter(inTab).filter(inDate).filter(r => {
      if (state.status === 'pendentes' && r.sent) return false;
      if (state.status === 'enviados' && !r.sent) return false;
      return !q || `${r.sku} ${r.desc} ${r.obs}`.toLocaleLowerCase('pt-BR').includes(q);
    });
    if (state.sortKey) {
      const k = state.sortKey;
      const val = r => (k === 'date' ? (r.date ? r.date.iso : '') : r[k]);
      const empty = v => v === null || v === '' || v === undefined;
      list = [...list].sort((a, b) => {
        const x = val(a), y = val(b);
        if (empty(x) || empty(y)) return empty(x) - empty(y);
        if (typeof x === 'number' && typeof y === 'number') return (x - y) * state.sortDir;
        if (typeof x === 'boolean') return (x - y) * state.sortDir;
        return String(x).localeCompare(String(y), 'pt-BR', { numeric: true }) * state.sortDir;
      });
    } else if (isHistory()) {
      // Histórico: pedidos mais recentes primeiro.
      list = [...list].sort((a, b) => (b.date ? b.date.iso : '').localeCompare(a.date ? a.date.iso : ''));
    }
    return list;
  }

  const sum = (list, k) => list.reduce((t, r) => t + (r[k] || 0), 0);
  const hasRed = r => Object.values(r.red || {}).some(Boolean);
  const neon = (r, col) => (r.red && r.red[col] ? ' neon' : '');

  /* ---------- Tela ---------- */

  function render() {
    renderDateOptions();
    const active = state.rows.filter(r => !r.grouped);
    const scope = isHistory() ? active : active.filter(inDate);
    renderRoute(scope);

    const pending = scope.filter(r => !r.sent).length;
    $('cntH').textContent = fmt(state.rows.length - active.length);
    $('cntP').textContent = fmt(pending);
    $('cntE').textContent = fmt(scope.length - pending);
    $('cntT').textContent = fmt(scope.length);
    document.querySelectorAll('.tabs button').forEach(b => {
      b.setAttribute('aria-pressed', String(b.dataset.status === state.status));
    });
    document.querySelectorAll('.grid th').forEach(th => {
      if (th.dataset.key === state.sortKey) {
        th.setAttribute('aria-sort', state.sortDir > 0 ? 'ascending' : 'descending');
      } else {
        th.removeAttribute('aria-sort');
      }
    });

    const list = visibleRows();
    $('rows').innerHTML = list.map(rowHtml).join('');
    state.flash.clear();

    $('totals').innerHTML = list.length
      ? `<tr><td colspan="2">Total</td><td class="num">${fmt(sum(list, 'un'))}</td><td class="num">${fmt(sum(list, 'cx'))}</td><td colspan="3"></td></tr>`
      : '';
    renderEmpty(list, scope);
    $('stCount').textContent = `${plural(list.length, 'item', 'itens')} na tela de ${fmt(state.rows.length)} na planilha`;
    $('btnPdf').disabled = $('btnXlsx').disabled = !list.length;
  }

  function rowHtml(r) {
    const key = `${r.sku}|${r.desc}`;
    const cls = [r.sent && 'is-sent', hasRed(r) && 'has-red', state.flash.has(key) && 'flash']
      .filter(Boolean).join(' ');
    let dateHtml = r.date ? r.date.label : '<span class="missing">Sem data</span>';
    if (r.date && !r.sent && !r.grouped) {
      const days = daysUntil(r.date);
      if (days < 0) dateHtml += `<small>${plural(-days, 'dia', 'dias')} de atraso</small>`;
    }
    return `<tr class="${cls}">
      <td class="sku${neon(r, 'sku')}" data-label="SKU">${r.sku ? esc(r.sku) : '<span class="missing">Sem SKU</span>'}</td>
      <td class="desc${neon(r, 'desc')}" data-label="Descrição">${esc(r.desc)}</td>
      <td class="num${neon(r, 'un')}" data-label="Unidades">${fmt(r.un)}</td>
      <td class="num${neon(r, 'cx')}" data-label="Caixas">${fmt(r.cx)}</td>
      <td class="date${neon(r, 'date')}" data-label="Data prevista">${dateHtml}</td>
      <td class="${neon(r, 'obs').trim()}" data-label="Observação">${esc(r.obs)}</td>
      <td data-label="Situação"><span class="chip ${r.sent ? 'sent' : 'pending'}">${r.sent ? 'Enviado' : 'Pendente'}</span></td>
    </tr>`;
  }

  function renderEmpty(list, scope) {
    const box = $('empty');
    box.hidden = list.length > 0;
    document.querySelector('.grid').hidden = !list.length;
    if (list.length) return;
    if (!state.loaded) {
      box.innerHTML = '<strong>Lendo a planilha…</strong>Os itens aparecem aqui em instantes.';
    } else if (!state.q && isHistory() && !state.rows.some(r => r.grouped)) {
      box.innerHTML = '<strong>Nenhum pedido no histórico</strong>Agrupe as linhas de um pedido antigo na planilha e ele aparece aqui.';
    } else if (!state.rows.length) {
      box.innerHTML = '<strong>Nenhum item na lista</strong>Preencha a aba PEDIDO DF na planilha e os itens aparecem aqui.';
    } else if (state.q) {
      box.innerHTML = `<strong>Nada encontrado para “${esc(state.q)}”</strong>Busque por outro SKU ou limpe a busca.`;
    } else if (state.status === 'pendentes' && scope.length) {
      box.innerHTML = '<strong>Todos os itens foram enviados</strong>Abra a aba Enviados para ver a lista.';
    } else {
      box.innerHTML = '<strong>Nenhum item nesta situação</strong>Escolha outra aba ou outra data.';
    }
  }

  function renderDateOptions() {
    const dates = new Map();
    state.rows.filter(inTab).forEach(r => { if (r.date) dates.set(r.date.iso, r.date.label); });
    if (state.date && !dates.has(state.date)) state.date = '';
    const opts = [...dates.keys()].sort().reverse()
      .map(iso => `<option value="${iso}">${dates.get(iso)}</option>`).join('');
    $('fDate').innerHTML = `<option value="">Todas</option>${opts}`;
    $('fDate').value = state.date;
  }

  function renderRoute(scope) {
    const totalCx = sum(scope, 'cx');
    const sentCx = sum(scope.filter(r => r.sent), 'cx');
    const useBoxes = totalCx > 0;
    const total = useBoxes ? totalCx : scope.length;
    const sent = useBoxes ? sentCx : scope.filter(r => r.sent).length;
    const pct = total ? Math.round((sent / total) * 100) : 0;
    $('trackFill').style.width = `${pct}%`;
    $('track').setAttribute('aria-valuenow', pct);
    $('track').classList.toggle('done', total > 0 && sent === total);
    const noun = useBoxes ? 'caixas' : 'itens';
    $('routeCaption').textContent = total
      ? `${fmt(sent)} de ${fmt(total)} ${noun} enviadas`.replace('itens enviadas', 'itens enviados')
      : state.loaded ? 'Nenhum item na lista' : 'Lendo a planilha…';

    const pending = scope.filter(r => !r.sent);
    $('figItems').textContent = fmt(scope.length);
    $('figItemsNote').textContent = `${fmt(pending.length)} ${pending.length === 1 ? 'pendente' : 'pendentes'}`;
    $('figUnits').textContent = fmt(sum(scope, 'un'));
    $('figBoxes').textContent = fmt(totalCx);

    const next = pending.map(r => r.date).filter(Boolean).sort((a, b) => a.iso.localeCompare(b.iso))[0];
    const note = $('figDateNote');
    note.className = '';
    if (!next) {
      $('figDate').textContent = '–';
      note.textContent = pending.length ? 'sem data definida' : 'nada pendente';
      return;
    }
    $('figDate').textContent = next.label;
    const days = daysUntil(next);
    if (days < 0) { note.textContent = `${plural(-days, 'dia', 'dias')} de atraso`; note.className = 'is-late'; }
    else if (days === 0) { note.textContent = 'hoje'; note.className = 'is-soon'; }
    else if (days === 1) { note.textContent = 'amanhã'; note.className = 'is-soon'; }
    else note.textContent = `em ${fmt(days)} dias`;
  }

  /* ---------- Exportação ---------- */

  const loaded = {};
  const loadScript = file => (loaded[file] ||= new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = VENDOR + file;
    s.onload = resolve;
    s.onerror = () => { delete loaded[file]; reject(new Error(`Falha ao carregar ${file}`)); };
    document.head.append(s);
  }));

  function exportMeta(list) {
    const dates = [...new Set(list.map(r => (r.date ? r.date.label : null)).filter(Boolean))];
    const isos = [...new Set(list.map(r => (r.date ? r.date.iso : null)).filter(Boolean))].sort();
    const now = new Date();
    const statusLabel = { pendentes: 'Pendentes', enviados: 'Enviados', todos: 'Todos os itens', historico: 'Histórico' }[state.status];
    const stamp = isos.length === 1 ? isos[0] : now.toISOString().slice(0, 10);
    return {
      dates,
      dateLabel: dates.length === 1 ? dates[0] : dates.length ? 'Várias datas' : 'Sem data',
      statusLabel: state.q ? `${statusLabel} (busca: ${state.q})` : statusLabel,
      issued: `${now.toLocaleDateString('pt-BR')} ${now.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}`,
      fileBase: `LuzSollar-Lista-DF_${stamp}_${state.status}`,
    };
  }

  // A fonte padrão do PDF só cobre o alfabeto latino.
  const pdfText = s => String(s ?? '').normalize('NFC')
    .replace(/[–—]/g, '-').replace(/[“”]/g, '"').replace(/[‘’]/g, "'")
    .replace(/[^\x20-\xFF]/g, '?');

  async function buildPdf(list, meta) {
    await loadScript('jspdf.umd.min.js');
    await loadScript('jspdf.plugin.autotable.min.js');
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait' });
    const W = doc.internal.pageSize.getWidth();
    const H = doc.internal.pageSize.getHeight();
    const M = 14;
    const INK = [20, 39, 58], STEEL = [79, 101, 120], SOL = [245, 163, 0], LINE = [211, 220, 227];
    const showDate = meta.dates.length > 1 || list.some(r => !r.date);
    const showStatus = state.status === 'todos' || isHistory();

    const pageHeader = () => {
      doc.setFillColor(...INK); doc.rect(0, 0, W, 20, 'F');
      doc.setFillColor(...SOL); doc.rect(0, 20, W, 1.2, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(15); doc.setTextColor(255, 255, 255);
      doc.text('LuzSollar', M, 12.6);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(10);
      doc.text('Lista de pedido MKT-SP para Brasília', W - M, 12.4, { align: 'right' });
    };

    pageHeader();
    // Bloco de dados do pedido
    const info = [
      ['Origem', 'MKT-SP'], ['Destino', 'Brasília (DF)'],
      ['Data prevista', meta.dateLabel], ['Emitido em', meta.issued],
    ];
    const cellW = (W - 2 * M) / info.length;
    let y = 29;
    info.forEach(([label, value], i) => {
      const x = M + i * cellW;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(...STEEL);
      doc.text(label, x, y);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(11.5); doc.setTextColor(...INK);
      doc.text(pdfText(value), x, y + 5.2);
    });
    y += 11;
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...STEEL);
    doc.text(pdfText(`Situação: ${meta.statusLabel}`), M, y);
    y += 4;

    const columns = [
      { key: 'item', header: 'Item' },
      { key: 'sku', header: 'SKU' },
      { key: 'desc', header: 'Descrição' },
      { key: 'un', header: 'Unidades' },
      { key: 'cx', header: 'Caixas' },
    ];
    if (showDate) columns.push({ key: 'date', header: 'Data prev.' });
    columns.push({ key: 'obs', header: 'Observação' });
    if (showStatus) columns.push({ key: 'status', header: 'Situação' });
    columns.push({ key: 'conf', header: 'Conf.' });

    const body = list.map((r, i) => ({
      item: String(i + 1), sku: pdfText(r.sku), desc: pdfText(r.desc),
      un: fmt(r.un), cx: fmt(r.cx), date: r.date ? r.date.label : '-',
      obs: pdfText(r.obs), status: r.sent ? 'Enviado' : 'Pendente', conf: '',
    }));
    const pdfCol = { sku: 'sku', desc: 'desc', un: 'un', cx: 'cx', date: 'date', obs: 'obs' };

    const leading = columns.findIndex(c => c.key === 'un');
    const foot = columns.slice(leading).map(c => ({
      un: fmt(sum(list, 'un')), cx: fmt(sum(list, 'cx')),
    }[c.key] ?? ''));
    foot.unshift({ content: `Total: ${plural(list.length, 'item', 'itens')}`, colSpan: leading });

    doc.autoTable({
      startY: y,
      margin: { left: M, right: M, top: 26, bottom: 16 },
      columns: columns.map(c => ({ header: c.header, dataKey: c.key })),
      body,
      foot: [foot],
      showFoot: 'lastPage',
      theme: 'plain',
      styles: { font: 'helvetica', fontSize: 9, cellPadding: { top: 2.1, bottom: 2.1, left: 1.8, right: 1.8 }, textColor: INK, lineColor: LINE, valign: 'middle' },
      headStyles: { fillColor: [241, 244, 247], textColor: STEEL, fontStyle: 'bold', lineWidth: { bottom: 0.3 }, lineColor: INK },
      bodyStyles: { lineWidth: { bottom: 0.15 } },
      footStyles: { fontStyle: 'bold', fillColor: [241, 244, 247], textColor: INK, lineWidth: { top: 0.4 }, lineColor: INK },
      columnStyles: {
        item: { cellWidth: 10, textColor: STEEL },
        sku: { cellWidth: 16, fontStyle: 'bold' },
        un: { cellWidth: 18, halign: 'right' },
        cx: { cellWidth: 15, halign: 'right' },
        date: { cellWidth: 21 },
        obs: { cellWidth: 34 },
        status: { cellWidth: 18 },
        conf: { cellWidth: 14, halign: 'center' },
      },
      didParseCell(d) {
        if (d.section === 'head' && (d.column.dataKey === 'un' || d.column.dataKey === 'cx')) d.cell.styles.halign = 'right';
        if (d.section === 'foot' && (d.column.dataKey === 'un' || d.column.dataKey === 'cx')) d.cell.styles.halign = 'right';
        if (d.section === 'head' && d.column.dataKey === 'conf') d.cell.styles.halign = 'center';
        if (d.section === 'body' && list[d.row.index].red?.[pdfCol[d.column.dataKey]]) {
          d.cell.styles.fontStyle = 'bold';
          d.cell.styles.textColor = [216, 32, 26];
        }
      },
      didDrawCell(d) {
        if (d.section === 'body' && d.column.dataKey === 'conf') {
          doc.setDrawColor(...STEEL); doc.setLineWidth(0.25);
          const s = 3.6;
          doc.rect(d.cell.x + (d.cell.width - s) / 2, d.cell.y + (d.cell.height - s) / 2, s, s);
        }
      },
      didDrawPage(d) { if (d.pageNumber > 1) pageHeader(); },
    });

    // Assinaturas
    let sy = doc.lastAutoTable.finalY + 24;
    if (sy > H - 26) { doc.addPage(); pageHeader(); sy = 52; }
    const signs = ['Separado por', 'Conferido por', 'Recebido em DF por'];
    const gap = 8, sw = (W - 2 * M - gap * 2) / 3;
    signs.forEach((s, i) => {
      const x = M + i * (sw + gap);
      doc.setDrawColor(...STEEL); doc.setLineWidth(0.25); doc.line(x, sy, x + sw, sy);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...STEEL);
      doc.text(s, x, sy + 4.5);
      doc.text('Data: ____/____/______', x, sy + 9.5);
    });

    // Rodapé
    const pages = doc.getNumberOfPages();
    for (let p = 1; p <= pages; p++) {
      doc.setPage(p);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(...STEEL);
      doc.text(`LuzSollar Lista DF | Emitido em ${meta.issued}`, M, H - 8);
      doc.text(`Página ${p} de ${pages}`, W - M, H - 8, { align: 'right' });
    }
    return doc.output('blob');
  }

  async function buildXlsx(list, meta) {
    await loadScript('exceljs.min.js');
    const wb = new window.ExcelJS.Workbook();
    wb.creator = 'LuzSollar Lista DF';
    wb.created = new Date();
    const ws = wb.addWorksheet('Pedido DF', {
      views: [{ state: 'frozen', ySplit: 5, showGridLines: false }],
      pageSetup: {
        paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0,
        margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.6, header: 0.3, footer: 0.3 },
      },
      headerFooter: { oddFooter: '&LLuzSollar Lista DF&RPágina &P de &N' },
    });

    const INK = 'FF14273A', SOL = 'FFF5A300', STEEL = 'FF4F6578', SOFT = 'FFF1F4F7', LINE = 'FFD3DCE3';
    const cols = [
      { h: 'Item', w: 6 }, { h: 'SKU', w: 10 }, { h: 'Descrição', w: 44 },
      { h: 'Unidades', w: 11 }, { h: 'Caixas', w: 9 }, { h: 'Data prevista', w: 13 },
      { h: 'Observação', w: 24 }, { h: 'Situação', w: 11 },
    ];
    cols.forEach((c, i) => { ws.getColumn(i + 1).width = c.w; });
    const last = cols.length;
    const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
    const font = (o = {}) => ({ name: 'Calibri', size: 11, ...o });

    ws.mergeCells(1, 1, 1, last);
    const title = ws.getCell(1, 1);
    title.value = 'LuzSollar Lista DF';
    title.font = font({ size: 16, bold: true, color: { argb: 'FFFFFFFF' } });
    title.alignment = { vertical: 'middle', indent: 1 };
    ws.getRow(1).height = 30;
    for (let c = 1; c <= last; c++) {
      ws.getCell(1, c).fill = fill(INK);
      ws.getCell(1, c).border = { bottom: { style: 'thick', color: { argb: SOL } } };
    }

    ws.mergeCells(2, 1, 2, last);
    const sub = ws.getCell(2, 1);
    sub.value = 'Lista de pedido MKT-SP para Brasília (DF)';
    sub.font = font({ bold: true, color: { argb: INK } });
    sub.alignment = { indent: 1 };

    ws.mergeCells(3, 1, 3, last);
    const info = ws.getCell(3, 1);
    info.value = `Data prevista: ${meta.dateLabel}    Situação: ${meta.statusLabel}    Emitido em: ${meta.issued}`;
    info.font = font({ color: { argb: STEEL } });
    info.alignment = { indent: 1 };

    const headRow = ws.getRow(5);
    cols.forEach((c, i) => {
      const cell = headRow.getCell(i + 1);
      cell.value = c.h;
      cell.font = font({ bold: true, color: { argb: 'FFFFFFFF' } });
      cell.fill = fill(INK);
      cell.alignment = { vertical: 'middle', horizontal: ['Unidades', 'Caixas'].includes(c.h) ? 'right' : 'left' };
    });
    headRow.height = 22;

    const border = { bottom: { style: 'thin', color: { argb: LINE } } };
    const first = 6;
    list.forEach((r, i) => {
      const row = ws.getRow(first + i);
      const skuNum = /^\d+$/.test(r.sku) ? Number(r.sku) : r.sku;
      row.values = [
        i + 1, skuNum, r.desc, r.un, r.cx,
        r.date ? new Date(Date.UTC(r.date.y, r.date.m - 1, r.date.d)) : null,
        r.obs, r.sent ? 'Enviado' : 'Pendente',
      ];
      for (let c = 1; c <= last; c++) {
        row.getCell(c).font = font();
        row.getCell(c).border = border;
        row.getCell(c).alignment = { vertical: 'middle' };
      }
      row.getCell(1).font = font({ color: { argb: STEEL } });
      row.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
      row.getCell(2).alignment = { horizontal: 'left', vertical: 'middle' };
      row.getCell(2).font = font({ bold: true });
      row.getCell(4).numFmt = '#,##0';
      row.getCell(5).numFmt = '#,##0';
      row.getCell(6).numFmt = 'dd/mm/yyyy';
      row.getCell(6).alignment = { horizontal: 'left', vertical: 'middle' };
      Object.entries({ sku: 2, desc: 3, un: 4, cx: 5, date: 6, obs: 7 }).forEach(([k, n]) => {
        if (r.red && r.red[k]) row.getCell(n).font = font({ bold: true, color: { argb: 'FFD8201A' } });
      });
      row.getCell(8).font = font({ bold: true, color: { argb: r.sent ? 'FF16704B' : 'FF7A5000' } });
    });

    const end = first + list.length - 1;
    const total = ws.getRow(end + 1);
    total.getCell(1).value = `Total: ${plural(list.length, 'item', 'itens')}`;
    ws.mergeCells(end + 1, 1, end + 1, 3);
    total.getCell(4).value = { formula: `SUM(D${first}:D${end})`, result: sum(list, 'un') };
    total.getCell(5).value = { formula: `SUM(E${first}:E${end})`, result: sum(list, 'cx') };
    total.getCell(4).numFmt = '#,##0';
    total.getCell(5).numFmt = '#,##0';
    for (let c = 1; c <= last; c++) {
      total.getCell(c).font = font({ bold: true });
      total.getCell(c).fill = fill(SOFT);
      total.getCell(c).border = { top: { style: 'medium', color: { argb: INK } } };
    }
    total.height = 22;

    ws.autoFilter = { from: { row: 5, column: 1 }, to: { row: end, column: last } };
    ws.pageSetup.printTitlesRow = '5:5';

    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  /* ---------- Pré-visualização ---------- */

  const preview = { blob: null, name: '', url: null };

  // Mesmo conteúdo do arquivo; usado no Excel e onde o navegador não abre PDF na página.
  function previewSheet(list, meta) {
    const red = (r, k) => (r.red && r.red[k] ? ' class="is-red"' : '');
    const rows = list.map((r, i) => `<tr>
      <td class="muted">${i + 1}</td>
      <td${r.red && r.red.sku ? ' class="is-red"' : ' class="strong"'}>${esc(r.sku)}</td>
      <td${red(r, 'desc')}>${esc(r.desc)}</td>
      <td class="num${r.red && r.red.un ? ' is-red' : ''}">${fmt(r.un)}</td>
      <td class="num${r.red && r.red.cx ? ' is-red' : ''}">${fmt(r.cx)}</td>
      <td${red(r, 'date')}>${r.date ? r.date.label : ''}</td>
      <td${red(r, 'obs')}>${esc(r.obs)}</td>
      <td class="strong">${r.sent ? 'Enviado' : 'Pendente'}</td>
    </tr>`).join('');
    return `<div class="sheet">
      <div class="sheet-title">LuzSollar Lista DF</div>
      <div class="sheet-sub"><strong>Lista de pedido MKT-SP para Brasília (DF)</strong>
        <span>Data prevista: ${esc(meta.dateLabel)}&emsp;Situação: ${esc(meta.statusLabel)}&emsp;Emitido em: ${esc(meta.issued)}</span></div>
      <div class="sheet-scroll"><table>
        <thead><tr><th>Item</th><th>SKU</th><th>Descrição</th><th class="num">Unidades</th><th class="num">Caixas</th><th>Data prevista</th><th>Observação</th><th>Situação</th></tr></thead>
        <tbody>${rows}</tbody>
        <tfoot><tr><td colspan="3">Total: ${plural(list.length, 'item', 'itens')}</td><td class="num">${fmt(sum(list, 'un'))}</td><td class="num">${fmt(sum(list, 'cx'))}</td><td colspan="3"></td></tr></tfoot>
      </table></div>
    </div>`;
  }

  async function openPreview(kind) {
    const list = visibleRows();
    const meta = exportMeta(list);
    const isPdf = kind === 'pdf';
    preview.blob = isPdf ? await buildPdf(list, meta) : await buildXlsx(list, meta);
    preview.name = `${meta.fileBase}.${isPdf ? 'pdf' : 'xlsx'}`;

    $('pvTitle').textContent = isPdf ? 'Pré-visualização do PDF' : 'Pré-visualização do Excel';
    $('pvInfo').textContent = `${plural(list.length, 'item', 'itens')}, ${fmt(sum(list, 'un'))} unidades, ${fmt(sum(list, 'cx'))} caixas`;
    $('pvFile').textContent = preview.name;
    $('pvDownload').textContent = isPdf ? 'Baixar PDF' : 'Baixar Excel';

    const body = $('pvBody');
    if (isPdf && navigator.pdfViewerEnabled) {
      preview.url = URL.createObjectURL(preview.blob);
      body.innerHTML = '';
      const frame = document.createElement('iframe');
      frame.title = 'Pré-visualização do PDF';
      frame.src = `${preview.url}#toolbar=0&navpanes=0&view=FitH`;
      body.append(frame);
    } else {
      body.innerHTML = previewSheet(list, meta);
    }
    $('preview').showModal();
  }

  function clearPreview() {
    $('pvBody').innerHTML = '';
    if (preview.url) URL.revokeObjectURL(preview.url);
    preview.url = null;
    preview.blob = null;
  }

  function downloadPreview() {
    const url = URL.createObjectURL(preview.blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = preview.name;
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
    $('preview').close();
  }

  async function runExport(btn, label, fn) {
    btn.disabled = true;
    btn.textContent = 'Gerando…';
    try {
      await fn();
    } catch (err) {
      console.error(err);
      const box = $('alert');
      box.textContent = 'Não foi possível gerar o arquivo. Recarregue a página e gere de novo.';
      box.hidden = false;
    } finally {
      btn.textContent = label;
      btn.disabled = !visibleRows().length;
    }
  }

  /* ---------- Eventos ---------- */

  $('lnkSheet').href = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit?gid=${GID}#gid=${GID}`;
  $('btnRefresh').addEventListener('click', () => refresh());
  $('btnPdf').addEventListener('click', e => runExport(e.currentTarget, 'Gerar PDF', () => openPreview('pdf')));
  $('btnXlsx').addEventListener('click', e => runExport(e.currentTarget, 'Gerar Excel', () => openPreview('xlsx')));
  $('pvCancel').addEventListener('click', () => $('preview').close());
  $('pvDownload').addEventListener('click', downloadPreview);
  $('preview').addEventListener('close', clearPreview);
  $('fDate').addEventListener('change', e => { state.date = e.target.value; render(); });
  $('fSearch').addEventListener('input', e => { state.q = e.target.value; render(); });
  document.querySelectorAll('.tabs button').forEach(b => {
    b.addEventListener('click', () => { state.status = b.dataset.status; render(); });
  });
  document.querySelectorAll('.grid th').forEach(th => {
    th.querySelector('button').addEventListener('click', () => {
      const k = th.dataset.key;
      if (state.sortKey !== k) { state.sortKey = k; state.sortDir = 1; }
      else if (state.sortDir === 1) state.sortDir = -1;
      else { state.sortKey = null; state.sortDir = 1; }
      render();
    });
  });
  document.addEventListener('visibilitychange', () => { if (!document.hidden) refresh(); });
  setInterval(() => { if (!document.hidden) refresh(); }, POLL_MS);

  render();
  refresh();
})();
