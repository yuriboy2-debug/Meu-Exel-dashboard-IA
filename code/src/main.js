import './styles.css';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';

const SAMPLE_ROWS = [
  { id: '01', name: 'Ana Souza', product: 'Notebook', value: 3499.9, date: '2026-10-01' },
  { id: '02', name: 'Bruno Lima', product: 'Smartphone', value: 2199, date: '2026-10-02' },
  { id: '03', name: 'Carla Mendes', product: 'Headset', value: 249.9, date: '2026-10-03' },
  { id: '04', name: 'Diego Alves', product: 'Monitor', value: 1299, date: '2026-10-04' },
  { id: '05', name: 'Elisa Rocha', product: 'Teclado mecânico', value: 399.9, date: '2026-10-05' },
  { id: '06', name: 'Felipe Costa', product: 'Mouse sem fio', value: 149.9, date: '2026-10-06' },
  { id: '07', name: 'Gabriela Martins', product: 'Webcam', value: 299.9, date: '2026-10-07' },
  { id: '08', name: 'Henrique Barros', product: 'SSD 512 GB', value: 459, date: '2026-10-08' },
  { id: '09', name: 'Isabela Nunes', product: 'Impressora', value: 899, date: '2026-10-09' },
  { id: '10', name: 'João Ferreira', product: 'Cadeira de escritório', value: 1199.9, date: '2026-10-10' },
  { id: '11', name: 'Karen Oliveira', product: 'Tablet', value: 1599, date: '2026-10-11' },
  { id: '12', name: 'Lucas Santos', product: 'Roteador Wi-Fi', value: 279.9, date: '2026-10-12' },
  { id: '13', name: 'Mariana Castro', product: 'Smartwatch', value: 799, date: '2026-10-13' },
  { id: '14', name: 'Natália Ribeiro', product: 'Caixa de som', value: 329.9, date: '2026-10-14' },
  { id: '15', name: 'Otávio Martins', product: 'HD externo', value: 499, date: '2026-10-15' },
];

const state = {
  rows: cloneRows(SAMPLE_ROWS),
  filteredRows: [],
  query: '',
  minValue: '',
  maxValue: '',
  startDate: '',
  endDate: '',
  charts: {},
};

const brl = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });
const integer = new Intl.NumberFormat('pt-BR');
const dateFormatter = new Intl.DateTimeFormat('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' });

function cloneRows(rows) {
  return rows.map((row) => ({ ...row }));
}

function normalizeText(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function parseMoney(value) {
  if (typeof value === 'number' && Number.isFinite(value)) return Math.round(value * 100) / 100;
  if (value === null || value === undefined || String(value).trim() === '') return null;
  let text = String(value).replace(/R\$|BRL/gi, '').replace(/\s/g, '');
  const comma = text.lastIndexOf(',');
  const dot = text.lastIndexOf('.');
  if (comma > dot) text = text.replace(/\./g, '').replace(',', '.');
  else text = text.replace(/,/g, '');
  const parsed = Number(text.replace(/[^\d.-]/g, ''));
  return Number.isFinite(parsed) ? Math.round(parsed * 100) / 100 : null;
}

function dateToISO(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return null;
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function parseDate(value) {
  if (value instanceof Date) return dateToISO(value);
  if (typeof value === 'number' && Number.isFinite(value) && value > 20000 && value < 80000) {
    const excelDate = new Date(Date.UTC(1899, 11, 30) + value * 86400000);
    return `${excelDate.getUTCFullYear()}-${String(excelDate.getUTCMonth() + 1).padStart(2, '0')}-${String(excelDate.getUTCDate()).padStart(2, '0')}`;
  }
  const text = String(value ?? '').trim();
  if (!text) return null;
  const brDate = text.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
  if (brDate) return `${brDate[3]}-${brDate[2].padStart(2, '0')}-${brDate[1].padStart(2, '0')}`;
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) return `${iso[1]}-${iso[2].padStart(2, '0')}-${iso[3].padStart(2, '0')}`;
  const parsed = new Date(text);
  return dateToISO(parsed);
}

function formatDate(isoDate) {
  if (!isoDate) return '—';
  return dateFormatter.format(new Date(`${isoDate}T00:00:00`));
}

function formatMonth(isoDate) {
  return new Intl.DateTimeFormat('pt-BR', { day: '2-digit', month: 'short' }).format(new Date(`${isoDate}T00:00:00`)).replace('.', '');
}

function icon(name, size = 18) {
  const paths = {
    upload: '<path d="M12 16V4m0 0 4 4m-4-4L8 8"/><path d="M4 15v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3"/>',
    pdf: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 15h1.5a1.5 1.5 0 0 0 0-3H8v6m5-6h1a2 2 0 0 1 0 4h-1v2m-1-6v6"/>',
    csv: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 12v5m2.2-5v5m2.2-5v5m2.2-5v5"/>',
    word: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 13l1 4 1.2-3 1.2 3 1-4"/>',
    search: '<circle cx="11" cy="11" r="7"/><path d="m20 20-4-4"/>',
    calendar: '<rect x="3" y="4" width="18" height="17" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>',
    filter: '<path d="M4 5h16M7 12h10m-7 7h4"/>',
    refresh: '<path d="M20 11a8 8 0 0 0-14.9-4L3 10m0-5v5h5M4 13a8 8 0 0 0 14.9 4L21 14m0 5v-5h-5"/>',
    chart: '<path d="M4 19V5m0 14h17"/><path d="m7 15 3-4 3 2 5-7"/>',
    trend: '<path d="m4 16 5-5 4 3 7-8"/><path d="M15 6h5v5"/>',
    users: '<circle cx="9" cy="8" r="3"/><path d="M3 20a6 6 0 0 1 12 0M17 11a3 3 0 1 0-1-5.8M17 14a5 5 0 0 1 4 4"/>',
    money: '<rect x="3" y="6" width="18" height="12" rx="2"/><circle cx="12" cy="12" r="3"/><path d="M7 9h.01M17 15h.01"/>',
    arrow: '<path d="m9 18 6-6-6-6"/>',
    close: '<path d="m6 6 12 12M18 6 6 18"/>',
  };
  return `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${paths[name] || paths.chart}</svg>`;
}

function appTemplate() {
  return `
    <div class="shell">
      <aside class="sidebar">
        <div class="brand">
          <div class="brand-mark">C</div>
          <div><strong>client<span>dash</span></strong><small>insights studio</small></div>
        </div>
        <div class="sidebar-label">VISÃO GERAL</div>
        <nav class="nav-list">
          <a class="nav-item active" href="#overview">${icon('chart')}<span>Dashboard</span></a>
          <a class="nav-item" href="#data-table">${icon('users')}<span>Clientes</span></a>
          <a class="nav-item" href="#analysis">${icon('trend')}<span>Análises</span></a>
        </nav>
        <div class="sidebar-bottom">
          <div class="help-card">
            <div class="help-icon">${icon('upload', 17)}</div>
            <div><strong>Importe sua planilha</strong><span>Arraste um Excel para começar</span></div>
          </div>
          <button class="sidebar-reset" id="resetData">${icon('refresh', 16)} Restaurar amostra</button>
          <div class="sidebar-foot">Dados locais · privacidade em primeiro lugar</div>
        </div>
      </aside>

      <main class="main-content">
        <header class="topbar">
          <div class="breadcrumbs"><span>Workspace</span><span class="slash">/</span><strong>Dashboard de clientes</strong></div>
          <div class="top-actions"><span class="live-dot"></span><span class="live-label">Atualizado agora</span><div class="avatar">CD</div></div>
        </header>

        <section class="page-head" id="overview">
          <div>
            <p class="eyebrow">CENTRAL DE INSIGHTS <span>•</span> OUTUBRO 2026</p>
            <h1>Dashboard de clientes</h1>
            <p class="page-subtitle">Acompanhe preços, datas e produtos em uma visão simples e acionável.</p>
          </div>
          <div class="header-actions">
            <label class="upload-button" for="excelInput">${icon('upload', 18)} Importar Excel</label>
            <input id="excelInput" class="sr-only" type="file" accept=".xlsx,.xls,.csv" />
            <div class="export-group">
              <button class="export-trigger" id="exportMenuButton">Exportar ${icon('arrow', 15)}</button>
              <div class="export-menu" id="exportMenu">
                <button data-export="pdf">${icon('pdf', 16)} Relatório PDF</button>
                <button data-export="csv">${icon('csv', 16)} Dados CSV</button>
                <button data-export="word">${icon('word', 16)} Relatório Word</button>
              </div>
            </div>
          </div>
        </section>

        <section class="drop-zone" id="dropZone">
          <div class="drop-icon">${icon('upload', 22)}</div>
          <div class="drop-copy"><strong>Arraste seu arquivo Excel aqui</strong><span>ou use o botão “Importar Excel” · .xlsx, .xls ou .csv</span></div>
          <span class="drop-status" id="importStatus">Amostra carregada</span>
        </section>

        <section class="metric-grid" aria-label="Indicadores principais">
          <article class="metric-card accent-teal"><div class="metric-top"><span class="metric-label">CLIENTES ANALISADOS</span><span class="metric-icon">${icon('users', 19)}</span></div><strong class="metric-value" id="metricClients">—</strong><span class="metric-note" id="metricClientsNote">registros no recorte atual</span></article>
          <article class="metric-card accent-blue"><div class="metric-top"><span class="metric-label">VALOR TOTAL</span><span class="metric-icon">${icon('money', 19)}</span></div><strong class="metric-value" id="metricTotal">—</strong><span class="metric-note" id="metricTotalNote">soma dos pedidos</span></article>
          <article class="metric-card accent-purple"><div class="metric-top"><span class="metric-label">PREÇO MÉDIO</span><span class="metric-icon">${icon('trend', 19)}</span></div><strong class="metric-value" id="metricAverage">—</strong><span class="metric-note" id="metricAverageNote">média por registro</span></article>
          <article class="metric-card accent-orange"><div class="metric-top"><span class="metric-label">INTERVALO DE DATAS</span><span class="metric-icon">${icon('calendar', 19)}</span></div><strong class="metric-value compact" id="metricRange">—</strong><span class="metric-note" id="metricRangeNote">período coberto</span></article>
        </section>

        <section class="content-grid" id="analysis">
          <article class="panel chart-panel chart-wide">
            <div class="panel-heading"><div><p class="section-kicker">EVOLUÇÃO TEMPORAL</p><h2>Valor dos pedidos por data</h2></div><span class="panel-badge">${icon('trend', 14)} tendência</span></div>
            <div class="chart-wrap line-wrap"><canvas id="evolutionChart"></canvas></div>
          </article>
          <article class="panel chart-panel">
            <div class="panel-heading"><div><p class="section-kicker">MIX DE PRODUTOS</p><h2>Valor por produto</h2></div><span class="panel-badge muted">Top produtos</span></div>
            <div class="chart-wrap bar-wrap"><canvas id="productChart"></canvas></div>
          </article>
        </section>

        <section class="lower-grid">
          <article class="panel insight-panel">
            <div class="panel-heading"><div><p class="section-kicker">ESTATÍSTICA DE PREÇOS</p><h2>Faixa de valores</h2></div>${icon('money', 21)}</div>
            <div class="range-stat"><span>Mínimo</span><strong id="metricMin">—</strong></div>
            <div class="range-track"><span id="rangeTrackFill"></span><i id="rangeTrackMarker"></i></div>
            <div class="range-stat"><span>Máximo</span><strong id="metricMax">—</strong></div>
            <div class="insight-foot"><span>Mediana estimada</span><strong id="metricMedian">—</strong></div>
          </article>
          <article class="panel insight-panel date-panel">
            <div class="panel-heading"><div><p class="section-kicker">DISTRIBUIÇÃO TEMPORAL</p><h2>Registros por período</h2></div>${icon('calendar', 21)}</div>
            <div class="date-distribution" id="dateDistribution"></div>
            <div class="insight-foot"><span>Data mais recente</span><strong id="latestDate">—</strong></div>
          </article>
        </section>

        <section class="panel table-panel" id="data-table">
          <div class="table-head"><div><p class="section-kicker">BASE DE DADOS</p><h2>Clientes importados</h2><span class="table-count" id="tableCount">—</span></div><div class="table-tools"><label class="search-field">${icon('search', 17)}<input id="searchInput" type="search" placeholder="Buscar nome ou produto" /></label><button class="filter-toggle" id="filterToggle">${icon('filter', 16)} Filtros <span class="filter-count" id="filterCount">0</span></button></div></div>
          <div class="filter-drawer" id="filterDrawer"><div class="filter-field"><label>Valor mínimo</label><input id="minValue" type="number" min="0" step="0.01" placeholder="R$ 0,00" /></div><div class="filter-field"><label>Valor máximo</label><input id="maxValue" type="number" min="0" step="0.01" placeholder="Sem limite" /></div><div class="filter-field"><label>Data inicial</label><input id="startDate" type="date" /></div><div class="filter-field"><label>Data final</label><input id="endDate" type="date" /></div><button class="clear-filter" id="clearFilters">Limpar filtros</button></div>
          <div class="table-scroll"><table><thead><tr><th>Id</th><th>Nome</th><th>Produto</th><th>Valor</th><th>Data</th></tr></thead><tbody id="dataBody"></tbody></table></div>
          <div class="empty-state" id="emptyState" hidden>${icon('search', 30)}<strong>Nenhum registro encontrado</strong><span>Ajuste os filtros ou importe uma nova planilha.</span></div>
        </section>

        <footer class="footer"><span>Clientdash Insights Studio</span><span><span class="footer-dot"></span> Dados processados localmente no navegador</span></footer>
      </main>
      <div class="toast" id="toast" role="status"></div>
    </div>
  `;
}

document.querySelector('#app').innerHTML = appTemplate();

const refs = {
  body: document.querySelector('#dataBody'),
  empty: document.querySelector('#emptyState'),
  tableCount: document.querySelector('#tableCount'),
  importStatus: document.querySelector('#importStatus'),
  toast: document.querySelector('#toast'),
  exportMenu: document.querySelector('#exportMenu'),
  filterDrawer: document.querySelector('#filterDrawer'),
};

function showToast(message, kind = 'success') {
  refs.toast.textContent = message;
  refs.toast.className = `toast visible ${kind}`;
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => { refs.toast.className = 'toast'; }, 3600);
}

function getStats(rows) {
  if (!rows.length) return { count: 0, total: 0, average: 0, min: 0, max: 0, median: 0, first: null, last: null };
  const values = rows.map((row) => row.value).sort((a, b) => a - b);
  const dates = rows.map((row) => row.date).sort();
  const total = rows.reduce((sum, row) => sum + row.value, 0);
  return {
    count: rows.length,
    total,
    average: total / rows.length,
    min: values[0],
    max: values[values.length - 1],
    median: values.length % 2 ? values[(values.length - 1) / 2] : (values[values.length / 2 - 1] + values[values.length / 2]) / 2,
    first: dates[0],
    last: dates[dates.length - 1],
  };
}

function applyFilters() {
  const query = normalizeText(state.query);
  const min = state.minValue === '' ? null : Number(state.minValue);
  const max = state.maxValue === '' ? null : Number(state.maxValue);
  state.filteredRows = state.rows.filter((row) => {
    const matchesQuery = !query || normalizeText(`${row.name} ${row.product} ${row.id}`).includes(query);
    const matchesMin = min === null || row.value >= min;
    const matchesMax = max === null || row.value <= max;
    const matchesStart = !state.startDate || row.date >= state.startDate;
    const matchesEnd = !state.endDate || row.date <= state.endDate;
    return matchesQuery && matchesMin && matchesMax && matchesStart && matchesEnd;
  });
  renderTable();
  updateMetrics();
  renderCharts();
  renderSecondaryInsights();
  updateFilterCount();
}

function renderTable() {
  refs.tableCount.textContent = `${state.filteredRows.length} de ${state.rows.length} registros`;
  refs.empty.hidden = state.filteredRows.length !== 0;
  refs.body.innerHTML = state.filteredRows.map((row, index) => `
    <tr style="--row-delay:${Math.min(index, 10) * 22}ms">
      <td><span class="id-pill">${escapeHtml(row.id)}</span></td>
      <td><strong>${escapeHtml(row.name)}</strong></td>
      <td><span class="product-cell"><span class="product-dot"></span>${escapeHtml(row.product)}</span></td>
      <td class="value-cell">${brl.format(row.value)}</td>
      <td><span class="date-cell">${formatDate(row.date)}</span></td>
    </tr>
  `).join('');
}

function updateMetrics() {
  const stats = getStats(state.filteredRows);
  document.querySelector('#metricClients').textContent = integer.format(stats.count);
  document.querySelector('#metricClientsNote').textContent = stats.count === state.rows.length ? 'todos os registros' : `de ${integer.format(state.rows.length)} registros`;
  document.querySelector('#metricTotal').textContent = brl.format(stats.total);
  document.querySelector('#metricTotalNote').textContent = 'soma no recorte atual';
  document.querySelector('#metricAverage').textContent = brl.format(stats.average);
  document.querySelector('#metricAverageNote').textContent = stats.count ? `faixa de ${brl.format(stats.min)} a ${brl.format(stats.max)}` : 'sem dados no recorte';
  document.querySelector('#metricRange').textContent = stats.first ? `${formatDate(stats.first)} — ${formatDate(stats.last)}` : '—';
  document.querySelector('#metricRangeNote').textContent = stats.first ? `${dateDiff(stats.first, stats.last)} dias de cobertura` : 'sem datas no recorte';
  document.querySelector('#metricMin').textContent = stats.count ? brl.format(stats.min) : '—';
  document.querySelector('#metricMax').textContent = stats.count ? brl.format(stats.max) : '—';
  document.querySelector('#metricMedian').textContent = stats.count ? brl.format(stats.median) : '—';
  const track = document.querySelector('#rangeTrackFill');
  const marker = document.querySelector('#rangeTrackMarker');
  const pct = stats.max ? Math.max(8, Math.min(92, (stats.average / stats.max) * 100)) : 8;
  track.style.width = `${pct}%`;
  marker.style.left = `${pct}%`;
  document.querySelector('#latestDate').textContent = stats.last ? formatDate(stats.last) : '—';
}

function dateDiff(start, end) {
  const a = new Date(`${start}T00:00:00`);
  const b = new Date(`${end}T00:00:00`);
  return Math.round((b - a) / 86400000) + 1;
}

function renderSecondaryInsights() {
  const counts = {};
  state.filteredRows.forEach((row) => { counts[row.date] = (counts[row.date] || 0) + 1; });
  const entries = Object.entries(counts).sort((a, b) => a[0].localeCompare(b[0]));
  const maxCount = Math.max(1, ...entries.map(([, count]) => count));
  document.querySelector('#dateDistribution').innerHTML = entries.slice(-7).map(([date, count]) => `
    <div class="date-bar-row"><span>${formatMonth(date)}</span><div class="date-bar-track"><i style="width:${Math.max(16, (count / maxCount) * 100)}%"></i></div><strong>${count}</strong></div>
  `).join('') || '<span class="muted-copy">Sem registros para exibir.</span>';
}

function renderCharts() {
  Object.values(state.charts).forEach((chart) => chart.destroy());
  const styles = getComputedStyle(document.documentElement);
  const teal = styles.getPropertyValue('--teal').trim() || '#0c7c66';
  const blue = styles.getPropertyValue('--blue').trim() || '#3f7cff';
  const grid = 'rgba(125, 145, 165, .16)';
  const text = '#6d7b89';
  const ctxLine = document.querySelector('#evolutionChart');
  const ctxBar = document.querySelector('#productChart');

  const byDate = {};
  state.filteredRows.forEach((row) => { byDate[row.date] = (byDate[row.date] || 0) + row.value; });
  const dates = Object.keys(byDate).sort();
  state.charts.line = new Chart(ctxLine, {
    type: 'line',
    data: {
      labels: dates.map(formatMonth),
      datasets: [{ data: dates.map((date) => byDate[date]), borderColor: teal, backgroundColor: 'rgba(12,124,102,.12)', fill: true, tension: .38, pointRadius: 4, pointHoverRadius: 6, pointBackgroundColor: '#fff', pointBorderWidth: 2, pointBorderColor: teal }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: (item) => ` ${brl.format(item.raw)}` }, displayColors: false, padding: 12 } },
      scales: { x: { grid: { display: false }, ticks: { color: text, font: { family: 'Inter, sans-serif', size: 11 }, maxRotation: 0, autoSkip: true, maxTicksLimit: 8 } }, y: { grid: { color: grid }, border: { display: false }, ticks: { color: text, font: { size: 11 }, callback: (value) => brl.format(value).replace(',00', '') } } },
    },
  });

  const byProduct = {};
  state.filteredRows.forEach((row) => { byProduct[row.product] = (byProduct[row.product] || 0) + row.value; });
  const products = Object.entries(byProduct).sort((a, b) => b[1] - a[1]).slice(0, 7);
  state.charts.bar = new Chart(ctxBar, {
    type: 'bar',
    data: { labels: products.map(([product]) => product), datasets: [{ data: products.map(([, value]) => value), backgroundColor: [blue, '#6e8fff', '#89a4ff', '#a6b9ff', '#bdd0ff', '#ccd9ff', '#dce6ff'], borderRadius: 7, borderSkipped: false, barThickness: 18 }] },
    options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { callbacks: { label: (item) => ` ${brl.format(item.raw)}` }, displayColors: false, padding: 12 } }, scales: { x: { grid: { color: grid }, border: { display: false }, ticks: { color: text, font: { size: 10 }, callback: (value) => brl.format(value).replace(',00', '') } }, y: { grid: { display: false }, border: { display: false }, ticks: { color: '#465466', font: { size: 11, weight: '600' } } } } },
  });
}

function updateFilterCount() {
  const count = [state.query, state.minValue, state.maxValue, state.startDate, state.endDate].filter(Boolean).length;
  document.querySelector('#filterCount').textContent = count;
  document.querySelector('#filterCount').classList.toggle('has-filter', count > 0);
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>'"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#039;', '"': '&quot;' }[char]));
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function exportCSV() {
  const header = ['Id', 'nome', 'produto', 'valor', 'data'];
  const lines = [header, ...state.filteredRows.map((row) => [row.id, row.name, row.product, row.value.toFixed(2).replace('.', ','), formatDate(row.date)])]
    .map((line) => line.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(';'));
  downloadBlob(new Blob([`\uFEFF${lines.join('\n')}`], { type: 'text/csv;charset=utf-8' }), 'dashboard-clientes.csv');
  showToast('CSV exportado com sucesso.');
}

function reportTitle() {
  return `Relatório do dashboard de clientes — ${formatDate(new Date().toISOString().slice(0, 10))}`;
}

function exportPDF() {
  const stats = getStats(state.filteredRows);
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  doc.setFillColor(12, 124, 102);
  doc.rect(0, 0, 297, 28, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(18);
  doc.text('Clientdash · Dashboard de clientes', 16, 17);
  doc.setTextColor(47, 61, 76);
  doc.setFontSize(10);
  doc.text(`${stats.count} clientes  |  Total ${brl.format(stats.total)}  |  Média ${brl.format(stats.average)}`, 16, 37);
  autoTable(doc, { startY: 45, head: [['Id', 'Nome', 'Produto', 'Valor', 'Data']], body: state.filteredRows.map((row) => [row.id, row.name, row.product, brl.format(row.value), formatDate(row.date)]), headStyles: { fillColor: [12, 124, 102], textColor: 255 }, alternateRowStyles: { fillColor: [245, 249, 248] }, styles: { fontSize: 9, cellPadding: 3 }, columnStyles: { 3: { halign: 'right' }, 4: { halign: 'center' } } });
  doc.save('dashboard-clientes.pdf');
  showToast('PDF gerado com sucesso.');
}

async function exportWord() {
  const stats = getStats(state.filteredRows);
  const headerCells = ['Id', 'Nome', 'Produto', 'Valor', 'Data'].map((value) => new TableCell({ shading: { fill: '0C7C66' }, children: [new Paragraph({ children: [new TextRun({ text: value, bold: true, color: 'FFFFFF' })] })] }));
  const bodyRows = state.filteredRows.map((row) => new TableRow({ children: [row.id, row.name, row.product, brl.format(row.value), formatDate(row.date)].map((value) => new TableCell({ children: [new Paragraph(String(value))] })) }));
  const doc = new Document({ sections: [{ children: [new Paragraph({ text: 'Clientdash · Dashboard de clientes', heading: HeadingLevel.TITLE }), new Paragraph({ text: `Relatório gerado em ${formatDate(new Date().toISOString().slice(0, 10))}` }), new Paragraph({ text: `Clientes: ${stats.count}   |   Total: ${brl.format(stats.total)}   |   Média: ${brl.format(stats.average)}` }), new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [new TableRow({ children: headerCells }), ...bodyRows] })] }] });
  downloadBlob(await Packer.toBlob(doc), 'dashboard-clientes.docx');
  showToast('Word gerado com sucesso.');
}

function headerIndex(headers, aliases) {
  return headers.findIndex((header) => aliases.includes(normalizeText(header).replace(/\s/g, '')));
}

async function importSpreadsheet(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
  if (matrix.length < 2) throw new Error('A planilha não possui linhas de dados.');
  const headers = matrix[0].map((header) => normalizeText(header).replace(/\s/g, ''));
  const indices = {
    id: headerIndex(headers, ['id', 'codigo', 'cod', 'identificador']),
    name: headerIndex(headers, ['nome', 'name', 'cliente']),
    product: headerIndex(headers, ['produto', 'product', 'item']),
    value: headerIndex(headers, ['valor', 'preco', 'preço', 'price', 'total']),
    date: headerIndex(headers, ['data', 'date', 'datavenda', 'datacompra']),
  };
  if (indices.name < 0 || indices.product < 0 || indices.value < 0 || indices.date < 0) throw new Error('Colunas necessárias não encontradas. Use Id, nome, produto, valor e data.');
  let invalid = 0;
  const imported = matrix.slice(1).map((cells, rowIndex) => {
    const value = parseMoney(cells[indices.value]);
    const date = parseDate(cells[indices.date]);
    const name = String(cells[indices.name] ?? '').trim();
    const product = String(cells[indices.product] ?? '').trim();
    if (!name || !product || value === null || !date) { invalid += 1; return null; }
    return { id: String(indices.id >= 0 ? cells[indices.id] || rowIndex + 1 : rowIndex + 1).padStart(2, '0'), name, product, value, date };
  }).filter(Boolean);
  if (!imported.length) throw new Error('Nenhuma linha válida foi encontrada.');
  state.rows = imported;
  state.query = '';
  state.minValue = '';
  state.maxValue = '';
  state.startDate = '';
  state.endDate = '';
  document.querySelector('#searchInput').value = '';
  ['minValue', 'maxValue', 'startDate', 'endDate'].forEach((id) => { document.querySelector(`#${id}`).value = ''; });
  refs.importStatus.textContent = `${imported.length} linhas importadas`;
  applyFilters();
  showToast(`${imported.length} registros importados${invalid ? ` · ${invalid} linha(s) ignorada(s)` : ''}.`);
}

function bindEvents() {
  document.querySelector('#exportMenuButton').addEventListener('click', (event) => { event.stopPropagation(); refs.exportMenu.classList.toggle('open'); });
  document.addEventListener('click', (event) => { if (!event.target.closest('.export-group')) refs.exportMenu.classList.remove('open'); });
  document.querySelectorAll('[data-export]').forEach((button) => button.addEventListener('click', async () => {
    refs.exportMenu.classList.remove('open');
    if (button.dataset.export === 'pdf') exportPDF();
    if (button.dataset.export === 'csv') exportCSV();
    if (button.dataset.export === 'word') await exportWord();
  }));
  document.querySelector('#excelInput').addEventListener('change', async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try { await importSpreadsheet(file); } catch (error) { showToast(error.message || 'Não foi possível importar o arquivo.', 'error'); }
    event.target.value = '';
  });
  const dropZone = document.querySelector('#dropZone');
  ['dragenter', 'dragover'].forEach((eventName) => dropZone.addEventListener(eventName, (event) => { event.preventDefault(); dropZone.classList.add('dragging'); }));
  ['dragleave', 'drop'].forEach((eventName) => dropZone.addEventListener(eventName, (event) => { event.preventDefault(); dropZone.classList.remove('dragging'); }));
  dropZone.addEventListener('drop', async (event) => {
    const file = event.dataTransfer.files?.[0];
    if (!file) return;
    try { await importSpreadsheet(file); } catch (error) { showToast(error.message || 'Não foi possível importar o arquivo.', 'error'); }
  });
  document.querySelector('#searchInput').addEventListener('input', (event) => { state.query = event.target.value; applyFilters(); });
  ['minValue', 'maxValue', 'startDate', 'endDate'].forEach((id) => document.querySelector(`#${id}`).addEventListener('input', (event) => { state[{ minValue: 'minValue', maxValue: 'maxValue', startDate: 'startDate', endDate: 'endDate' }[id]] = event.target.value; applyFilters(); }));
  document.querySelector('#filterToggle').addEventListener('click', () => refs.filterDrawer.classList.toggle('open'));
  document.querySelector('#clearFilters').addEventListener('click', () => {
    state.query = ''; state.minValue = ''; state.maxValue = ''; state.startDate = ''; state.endDate = '';
    document.querySelector('#searchInput').value = '';
    ['minValue', 'maxValue', 'startDate', 'endDate'].forEach((id) => { document.querySelector(`#${id}`).value = ''; });
    applyFilters();
  });
  document.querySelector('#resetData').addEventListener('click', () => { state.rows = cloneRows(SAMPLE_ROWS); refs.importStatus.textContent = 'Amostra restaurada'; applyFilters(); showToast('Amostra restaurada.'); });
}

bindEvents();
applyFilters();
