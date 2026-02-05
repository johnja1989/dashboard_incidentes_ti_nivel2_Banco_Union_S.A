/** 
 * Dashboard Incidentes TI — Banco Unión
 * - CSV -> Visualizaciones (Chart.js + DataLabels)
 * - PDF Ejecutivo (jsPDF + AutoTable) con encabezado corporativo y logo
 * - Excel (SheetJS)
 * - Persistencia localStorage del estado
 * - Camino A: Narrativa ejecutiva sin LLM (reglas / umbrales)
 * - Camino B: Narrativa con LLM local (Ollama o LM Studio), fallback automático
 */

/* ===================== Constantes de PDF (Landscape forzado) ===================== */
const PAGE_ORIENTATION = 'landscape';  // ← Siempre horizontal
const PAGE_UNIT = 'pt';
const PAGE_FORMAT = 'a4';

/* Helper para añadir páginas en horizontal de forma consistente */
function addLandscapePage(doc) {
  doc.addPage(PAGE_ORIENTATION, PAGE_UNIT);
}

/* ===================== Configuración del LLM local (Camino B) ===================== */
/**
 * Cambia estas constantes según tu entorno:
 * - LLM_PROVIDER: 'ollama' | 'lmstudio'
 * - Para Ollama: usa http://localhost:11434/api/generate (no-stream)
 * - Para LM Studio: usa http://localhost:1234/v1/chat/completions (OpenAI compatible)
 */
const LLM_ENABLED = true;                 // Activa/desactiva la narrativa LLM
const LLM_PROVIDER = 'ollama';            // 'ollama' o 'lmstudio'
const LLM_OLLAMA_MODEL = 'llama3.2';      // Modelo local (ajústalo a tu preferido)
const LLM_OLLAMA_URL = 'http://localhost:11434/api/generate';
const LLM_LMSTUDIO_URL = 'http://localhost:1234/v1/chat/completions';
const LLM_TIMEOUT_MS = 9000;            // Timeout razonable para respuesta local

/* ===================== Utils ===================== */
function norm(s) {
  s = (s === undefined || s === null) ? '' : String(s);
  return s.trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9_ ]/g, '')
    .replace(/\s+/g, ' ');
}
function parseNumber(val) {
  if (val === undefined || val === null) return NaN;
  const s = String(val).replace(',', '.');
  const m = s.match(/-?\d+(\.\d+)?/);
  return m ? parseFloat(m[0]) : NaN;
}
function cleanText(value) {
  if (value === undefined || value === null) return '';
  let text = String(value).trim();
  text = text.replace(/\s+/g, ' ').replace(/\u00A0/g, ' ').replace(/\uFFFD/g, '').replace(/\ufffd/g, '');
  text = text.replace(/[\u{1F300}-\u{1FAFF}]/gu, ''); // quitar emojis 
  return text.trim();
}
function isDateVal(v) {
  const s = String(v === undefined || v === null ? '' : v).trim();
  if (!s) return false;
  const iso = /^\d{4}\-\d{1,2}\-\d{1,2}/.test(s);
  const latam = /^\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}$/.test(s);
  const dt = new Date(s);
  return iso || latam || !isNaN(dt);
}
function isNumVal(v) { return !isNaN(parseNumber(v)); }
function detectType(values) {
  const sample = values.slice(0, 200);
  const dateCount = sample.filter(isDateVal).length;
  const numCount = sample.filter(isNumVal).length;
  const distinct = new Set(sample.map(function (x) { return String(x); })).size;
  if (dateCount > sample.length * 0.5) return 'date';
  if (numCount > sample.length * 0.6) return 'number';
  if (distinct <= sample.length * 0.7) return 'category';
  return 'text';
}

/* ===================== Semántica y schema ===================== */
const semanticHints = {
  estado: ['estado', 'status', 'estado final', 'estado incidente', 'estado linea', 'estado proveedor'],
  responsable: ['responsable', 'asignado', 'ingeniero asignado', 'owner', 'persona', 'responsable escalamiento'],
  servicio: ['servicio', 'service', 'tipificacion', 'categoria', 'tipo', 'producto', 'componente'],
  proveedor: ['proveedor', 'vendor', 'proveedor a escalar'],
  tiempo: ['tiempo', 'duracion', 'duration', 'dias', 'edad incidente', 'edad'],
  fecha: ['fecha', 'date', 'fch', 'radicado', 'cierre', 'actualizacion', 'produccion'],
  rangoEdad: ['rango edad', 'rango_edad', 'age range', 'rango']
};
function inferSchema(headers, rows) {
  const schema = { roles: {}, types: {} };
  headers.forEach(function (h) {
    const vals = rows.map(function (r) { return r[h]; }).filter(function (v) { return v !== undefined; });
    schema.types[h] = detectType(vals);
    const nh = norm(h);
    Object.keys(semanticHints).forEach(function (role) {
      const list = semanticHints[role];
      for (var i = 0; i < list.length; i++) {
        if (nh.indexOf(norm(list[i])) >= 0) { schema.roles[role] = h; break; }
      }
    });
  });
  // Tiempo: fallback al campo numérico con más hits 
  if (!schema.roles.tiempo) {
    var best = null, bestHits = -1;
    headers.forEach(function (h) {
      if (schema.types[h] !== 'number' && schema.types[h] !== 'text') return;
      const hits = rows.map(function (r) { return parseNumber(r[h]); })
        .filter(function (n) { return !isNaN(n); })
        .length;
      if (hits > bestHits) { bestHits = hits; best = h; }
    });
    if (best) schema.roles.tiempo = best;
  }
  // Estado: categoría con 2..15 valores 
  if (!schema.roles.estado) {
    const cats = headers.filter(function (h) { return schema.types[h] === 'category' || schema.types[h] === 'text'; });
    var best = null, uniqBest = 9999;
    cats.forEach(function (h) {
      const uniq = new Set(rows.map(function (r) { return r[h]; })).size;
      if (uniq >= 2 && uniq <= 15 && uniq < uniqBest) { uniqBest = uniq; best = h; }
    });
    if (best) schema.roles.estado = best;
  }
  // Responsable 
  if (!schema.roles.responsable) {
    const nonNums = headers.filter(function (h) { return schema.types[h] !== 'number'; });
    var best = null, score = 0;
    nonNums.forEach(function (h) {
      const vals = rows.map(function (r) { return String(r[h] === undefined ? '' : r[h]); }).slice(0, 200);
      const hit = vals.filter(function (v) { return /[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+\s+[A-Za-zÁÉÍÓÚÑáéíóúñ]/.test(v); }).length;
      if (hit > score) { score = hit; best = h; }
    });
    if (best) schema.roles.responsable = best;
  }
  // Servicio 
  if (!schema.roles.servicio) {
    const cats = headers.filter(function (h) { return schema.types[h] === 'category' || schema.types[h] === 'text'; });
    var best = null, uniqBest = 0;
    cats.forEach(function (h) {
      const uniq = new Set(rows.map(function (r) { return r[h]; })).size;
      if (uniq > uniqBest) { uniqBest = uniq; best = h; }
    });
    if (best) schema.roles.servicio = best;
  }
  // Proveedor (regex) 
  if (!schema.roles.proveedor) {
    const prov = headers.find(function (h) { return /(proveedor|vendor)/i.test(h); });
    if (prov) schema.roles.proveedor = prov;
  }
  // Fecha 
  if (!schema.roles.fecha) {
    const d = headers.find(function (h) { return schema.types[h] === 'date'; });
    if (d) schema.roles.fecha = d;
  }
  return schema;
}

/* ===================== Agregaciones ===================== */
function countBy(arr, keyCol) {
  const m = new Map();
  arr.forEach(function (r) {
    const k = r[keyCol];
    m.set(k, (m.get(k) || 0) + 1);
  });
  return Array.from(m.entries()).map(function (pair) {
    return { label: pair[0], value: pair[1] };
  }).sort(function (a, b) {
    return String(a.label).localeCompare(String(b.label));
  });
}
function average(arr) { return !arr.length ? 0 : arr.reduce(function (a, b) { return a + b; }, 0) / arr.length; }

/* ===================== Orden lógico de rangos ===================== */
function getRangeWeight(labelRaw) {
  const label = norm(labelRaw);
  var m = label.match(/(>\s*\+)\s*(\d+)/);
  if (m) return Number(m[2]) + 0.1;
  m = label.match(/(\d+)\s*-\s*(\d+)/);
  if (m) return Number(m[2]);
  m = label.match(/<\s*(\d+)/);
  if (m) return Number(m[1]) - 0.1;
  m = label.match(/(\d+)/);
  if (m) return Number(m[1]);
  return 9999;
}

/* ===================== Estado global y persistencia ===================== */
var rawData = [];
var charts = {};
var schema = null;
const STORAGE_KEY = 'dashboard_incidentes_state_v1';
function saveState() {
  try {
    const payload = {
      rawData: rawData,
      schema: schema,
      meta: { fileName: (window.__lastCsvName || 'CSV'), savedAt: new Date().toISOString() }
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    updateBanner(payload.meta.fileName, new Date(payload.meta.savedAt));
  } catch (err) { console.error('Error guardando estado:', err); }
}
function loadState() {
  try {
    const txt = localStorage.getItem(STORAGE_KEY);
    if (!txt) return false;
    const payload = JSON.parse(txt);
    if (!payload || !Array.isArray(payload.rawData) || !payload.rawData.length) return false;
    rawData = payload.rawData;
    schema = payload.schema || null;
    renderAll();
    const dt = payload.meta && payload.meta.savedAt ? new Date(payload.meta.savedAt) : null;
    if (dt) {
      const lastUpdate = document.getElementById('lastUpdate');
      if (lastUpdate) lastUpdate.textContent = 'Última actualización: ' + dt.toLocaleString('es-CO');
      updateBanner(payload.meta.fileName || 'CSV', dt);
    }
    return true;
  } catch (err) { console.error('Error cargando estado:', err); return false; }
}
function clearState() {
  try {
    localStorage.removeItem(STORAGE_KEY);
    updateBanner('—', null);
    alert('Se limpió la caché del dashboard. Recarga la página y vuelve a cargar el CSV.');
  } catch (err) { console.error('Error limpiando estado:', err); }
}
function updateBanner(fileName, savedDate) {
  const f = document.getElementById('bannerFile');
  const s = document.getElementById('bannerSaved');
  if (f) f.textContent = cleanText(fileName || '—');
  if (s) s.textContent = savedDate ? savedDate.toLocaleString('es-CO') : '—';
}

/* ===================== Inicialización ===================== */
document.addEventListener('DOMContentLoaded', function () {
  const fileInput = document.getElementById('csvFile');
  const btnPdfReport = document.getElementById('btnPdfReport');
  const btnExcelReport = document.getElementById('btnExcelReport');
  const btnClear = document.getElementById('btnClear');
  loadState();
  if (fileInput) {
    fileInput.addEventListener('change', async function (e) {
      const file = e.target.files && e.target.files[0];
      if (!file) return;
      window.__lastCsvName = file.name;
      const buffer = await file.arrayBuffer();
      var csvText;
      try { csvText = new TextDecoder('windows-1252').decode(buffer); }
      catch (err) { csvText = new TextDecoder('utf-8').decode(buffer); }
      Papa.parse(csvText, {
        header: true,
        skipEmptyLines: true,
        complete: function (results) {
          rawData = (results.data || []).filter(function (row) {
            return Object.values(row).some(function (val) {
              return val !== null && val !== undefined && String(val).trim() !== '';
            });
          });
          if (!rawData.length) { alert('El archivo CSV no contiene datos válidos.'); return; }
          const headers = results.meta && results.meta.fields ? results.meta.fields : Object.keys(rawData[0] || {});
          schema = inferSchema(headers, rawData);
          renderAll();
          const lastUpdate = document.getElementById('lastUpdate');
          if (lastUpdate) lastUpdate.textContent = 'Última actualización: ' + new Date().toLocaleString('es-CO');
          saveState();
        },
        error: function (err) { console.error('Error parseando CSV:', err); alert('Error al procesar el archivo CSV.'); }
      });
    });
  }
  if (btnPdfReport) {
    // generatePDFReport es async por la llamada LLM
    btnPdfReport.addEventListener('click', async function (e) {
      e.preventDefault();
      try { await generatePDFReport(); }
      catch (err) { console.error('Error generando PDF:', err); alert('Ocurrió un error al generar el PDF.'); }
    });
  }
  if (btnExcelReport) {
    btnExcelReport.addEventListener('click', generateExcelReport);
  }
  if (btnClear) {
    btnClear.addEventListener('click', clearState);
  }
});
function renderAll() { renderTable(); renderChartsAdaptive(); }

/* ===================== Tabla ===================== */
function renderTable() {
  const tbody = document.querySelector('#tablaDatos tbody');
  const theadRow = document.querySelector('#tablaDatos thead tr');
  if (!tbody || !theadRow) return;
  tbody.innerHTML = '';
  if (!rawData.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="muted">Sin datos para mostrar.</td></tr>';
    return;
  }
  const cols = Object.keys(rawData[0]);
  theadRow.innerHTML = cols.map(function (c) { return '<th>' + cleanText(c) + '</th>'; }).join('');
  rawData.forEach(function (r) {
    const tr = document.createElement('tr');
    tr.innerHTML = cols.map(function (c) { return '<td>' + cleanText(r[c]) + '</td>'; }).join('');
    tbody.appendChild(tr);
  });
}

/* ===================== Gráficos ===================== */
function destroyChart(id) { if (charts[id]) { charts[id].destroy(); charts[id] = null; } }
function showPlaceholder(id) {
  const ctx = document.getElementById(id); if (!ctx) return;
  destroyChart(id);
  charts[id] = new Chart(ctx, {
    type: 'bar',
    data: { labels: ['Sin datos'], datasets: [{ label: '—', data: [0], backgroundColor: '#273043' }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { enabled: false } }, scales: { y: { display: false } } }
  });
}

/* ---- KPI: texto centrado ---- */
const CenterTextPlugin = {
  id: 'centerText',
  afterDraw: function (chart) {
    const meta = chart.getDatasetMeta(0);
    if (!meta || !meta.data || !meta.data[0]) return;
    const ctx = chart.ctx;
    const cx = meta.data[0].x;
    const cy = meta.data[0].y;
    const opts = chart.options && chart.options.plugins && chart.options.plugins.centerText ? chart.options.plugins.centerText : {};
    const title = String(opts.title || '');
    const value = String(opts.value || '');
    const titleSize = Number(opts.titleSize || 14);
    const valueSize = Number(opts.valueSize || 28);
    const color = opts.color || '#ffffff';
    const fontFamily = 'system-ui,Segoe UI,Roboto,Arial';
    const gap = Number(opts.gap || 8);
    ctx.save();
    ctx.fillStyle = color;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    if (title) {
      ctx.font = titleSize + 'px ' + fontFamily;
      const titleY = cy - (value ? (valueSize / 2 + gap) : 0);
      ctx.fillText(title, cx, titleY);
    }
    if (value) {
      ctx.font = 'bold ' + valueSize + 'px ' + fontFamily;
      const valueY = title ? (cy + (titleSize / 2 + gap)) : cy;
      ctx.fillText(value, cx, valueY);
    }
    ctx.restore();
  }
};

/* ---- Barras simples ---- */
function buildBar(id, data, label) {
  const ctx = document.getElementById(id); if (!ctx) return; destroyChart(id);
  data.sort(function (a, b) { return b.value - a.value; });
  const SimpleValueLabels = {
    id: 'simpleValueLabels',
    afterDatasetsDraw: function (chart) {
      const ctx = chart.ctx;
      const dataset = chart.data.datasets[0];
      const meta = chart.getDatasetMeta(0);
      if (!meta || !meta.data) return;
      ctx.save();
      ctx.fillStyle = '#ffffff';
      ctx.textAlign = 'center';
      ctx.font = 'bold 12px system-ui,Segoe UI,Roboto,Arial';
      meta.data.forEach(function (bar, i) {
        const v = dataset.data[i];
        if (v === null || v === undefined || isNaN(v)) return;
        const x = bar.x; const y = bar.y - 6;
        ctx.fillText(String(Math.round(Number(v))), x, y);
      });
      ctx.restore();
    }
  };
  const useDatalabels = (typeof ChartDataLabels !== 'undefined');
  charts[id] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: data.map(function (d) { return cleanText(String(d.label)); }),
      datasets: [{ label: cleanText(label), data: data.map(function (d) { return d.value; }), backgroundColor: '#90cdf4', borderColor: '#3182ce', borderWidth: 1 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: true, labels: { color: '#ffffff', font: { size: 14, weight: 'bold' } } },
        tooltip: {
          enabled: true, titleColor: '#ffffff', bodyColor: '#ffffff',
          callbacks: { label: function (ctx) { return ctx.dataset.label + ': ' + Math.round(Number(ctx.raw)); } }
        },
        datalabels: useDatalabels ? {
          color: '#ffffff', anchor: 'end', align: 'end', offset: 2, clamp: true, clip: false,
          font: { weight: 'bold', size: 12 },
          formatter: function (value) { return String(Math.round(Number(value))); }
        } : undefined
      },
      scales: {
        y: { beginAtZero: true, ticks: { color: '#ffffff' }, grid: { color: 'rgba(255,255,255,0.1)' } },
        x: { ticks: { color: '#ffffff' }, grid: { color: 'rgba(255,255,255,0.05)' } }
      }
    },
    plugins: useDatalabels ? [ChartDataLabels] : [SimpleValueLabels]
  });
}

/* ======= Plugins % por barra ======= */
const PerBarPercentLabelsPlugin = { id: 'perBarPercentLabels', afterDatasetsDraw: function (ch, a, o) { try { if (!ch.canvas || ch.canvas.id !== 'chartResponsables') return; drawPercentsGeneric(ch, o); } catch (e) { console.warn(e); } } };
const PerBarPercentLabelsPluginCat = { id: 'perBarPercentLabelsCat', afterDatasetsDraw: function (ch, a, o) { try { if (!ch.canvas || ch.canvas.id !== 'chartCategoria') return; drawPercentsGeneric(ch, o); } catch (e) { console.warn(e); } } };
const PerBarPercentLabelsPluginProv = { id: 'perBarPercentLabelsProv', afterDatasetsDraw: function (ch, a, o) { try { if (!ch.canvas || ch.canvas.id !== 'chartProveedor') return; drawPercentsGeneric(ch, o); } catch (e) { console.warn(e); } } };
const PerBarPercentLabelsPluginTime = { id: 'perBarPercentLabelsTime', afterDatasetsDraw: function (ch, a, o) { try { if (!ch.canvas || ch.canvas.id !== 'chartTiempo') return; drawPercentsGeneric(ch, o); } catch (e) { console.warn(e); } } };
const PerBarPercentLabelsPluginServ = { id: 'perBarPercentLabelsServ', afterDatasetsDraw: function (ch, a, o) { try { if (!ch.canvas || ch.canvas.id !== 'chartServicio') return; drawPercentsGeneric(ch, o); } catch (e) { console.warn(e); } } };
function drawPercentsGeneric(chart, pluginOptions) {
  const opts = pluginOptions || {};
  const mapping = opts.mapping || null;
  const color = opts.color || '#ffffff';
  const inside = !!opts.inside;
  if (!mapping) return;
  const ctx = chart.ctx;
  ctx.save();
  ctx.fillStyle = color;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.font = 'bold 12px system-ui,Segoe UI,Roboto,Arial';
  function drawText(txt, x, y) {
    ctx.strokeStyle = 'rgba(0,0,0,0.45)';
    ctx.lineWidth = 3;
    ctx.strokeText(txt, x, y);
    ctx.fillText(txt, x, y);
  }
  Object.keys(mapping).forEach(function (dsIndexStr) {
    const dsIndex = Number(dsIndexStr);
    const meta = chart.getDatasetMeta(dsIndex);
    const percArr = mapping[dsIndex];
    if (!meta || !meta.data || !Array.isArray(percArr)) return;
    meta.data.forEach(function (bar, i) {
      const p = percArr[i];
      if (p === null || p === undefined) return;
      const txt = String(Math.round(p)) + '%';
      const x = bar.x;
      var baseScaleY = undefined;
      if (chart.scales && chart.scales.y && typeof chart.scales.y.getPixelForValue === 'function') {
        baseScaleY = chart.scales.y.getPixelForValue(0);
      }
      const base = (typeof bar.base === 'number') ? bar.base : baseScaleY;
      var y = (typeof base === 'number') ? (bar.y + base) / 2 : (bar.y - 8);
      if (!inside) y = bar.y - 8;
      drawText(txt, x, y);
    });
  });
  ctx.restore();
}

/* ======= Barras multi-dataset con % ======= */
function buildMultiBar(id, labels, datasets, options) {
  options = options || {};
  const ctx = document.getElementById(id); if (!ctx) return; destroyChart(id);
  const useDatalabels = (typeof ChartDataLabels !== 'undefined');
  const pluginsArr = [];
  if (useDatalabels) pluginsArr.push(ChartDataLabels);
  if (options.showEachBarPercentLabels && options.percentsByDataset) {
    if (id === 'chartCategoria') pluginsArr.push(PerBarPercentLabelsPluginCat);
    else if (id === 'chartProveedor') pluginsArr.push(PerBarPercentLabelsPluginProv);
    else if (id === 'chartTiempo') pluginsArr.push(PerBarPercentLabelsPluginTime);
    else if (id === 'chartServicio') pluginsArr.push(PerBarPercentLabelsPluginServ);
    else pluginsArr.push(PerBarPercentLabelsPlugin);
  }
  charts[id] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels.map(function (l) { return cleanText(String(l)); }),
      datasets: datasets.map(function (ds) {
        return {
          label: cleanText(ds.label),
          data: ds.data,
          backgroundColor: ds.backgroundColor,
          borderColor: ds.borderColor || '#0b1220',
          borderWidth: 1
        };
      })
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: true, labels: { color: '#ffffff', font: { size: 14, weight: 'bold' } } },
        tooltip: {
          enabled: true,
          titleColor: '#ffffff',
          bodyColor: '#ffffff',
          callbacks: {
            label: function (ctx) {
              const dsIndex = ctx.datasetIndex;
              const i = ctx.dataIndex;
              const val = ctx.raw;
              var pctTxt = '';
              try {
                const mapping = options.percentsByDataset || {};
                const arr = mapping[dsIndex] || null;
                if (arr && typeof arr[i] === 'number') pctTxt = ' (' + Math.round(arr[i]) + '%)';
              } catch (e) { }
              const shownVal = Math.round(Number(val));
              return ctx.dataset.label + ': ' + shownVal + pctTxt;
            }
          }
        },
        datalabels: useDatalabels ? {
          color: '#ffffff', anchor: 'end', align: 'end', offset: 2, clamp: true, clip: false,
          font: { weight: 'bold', size: 12 },
          formatter: function (value) { return String(Math.round(Number(value))); }
        } : undefined,
        perBarPercentLabels: (id !== 'chartCategoria' && id !== 'chartProveedor' && id !== 'chartTiempo' && id !== 'chartServicio' && options.showEachBarPercentLabels)
          ? { mapping: options.percentsByDataset, color: options.percentTextColor || '#ffffff', inside: !!options.percentInside } : undefined,
        perBarPercentLabelsCat: (id === 'chartCategoria' && options.showEachBarPercentLabels)
          ? { mapping: options.percentsByDataset, color: options.percentTextColor || '#ffffff', inside: !!options.percentInside } : undefined,
        perBarPercentLabelsProv: (id === 'chartProveedor' && options.showEachBarPercentLabels)
          ? { mapping: options.percentsByDataset, color: options.percentTextColor || '#ffffff', inside: !!options.percentInside } : undefined,
        perBarPercentLabelsTime: (id === 'chartTiempo' && options.showEachBarPercentLabels)
          ? { mapping: options.percentsByDataset, color: options.percentTextColor || '#ffffff', inside: !!options.percentInside } : undefined,
        perBarPercentLabelsServ: (id === 'chartServicio' && options.showEachBarPercentLabels)
          ? { mapping: options.percentsByDataset, color: options.percentTextColor || '#ffffff', inside: !!options.percentInside } : undefined
      },
      scales: {
        y: { beginAtZero: true, stacked: !!options.stacked, ticks: { color: '#ffffff' }, grid: { color: 'rgba(255,255,255,0.1)' } },
        x: { stacked: !!options.stacked, ticks: { color: '#ffffff' }, grid: { color: 'rgba(255,255,255,0.05)' } }
      }
    },
    plugins: pluginsArr
  });
}

/* ===================== KPI Incidentes ===================== */
function buildKPIIncidentes(id, total) {
  destroyChart(id);
  const ctx = document.getElementById(id); if (!ctx) return;
  charts[id] = new Chart(ctx, {
    type: 'doughnut',
    data: { labels: ['Incidentes Reportados'], datasets: [{ data: [total], backgroundColor: ['#4cc9f0'], borderColor: '#0b1220', borderWidth: 1, cutout: '70%' }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        title: { display: true, text: cleanText('Incidentes Reportados'), color: '#ffffff', font: { size: 16, weight: 'bold' } },
        centerText: { title: cleanText('Total'), value: String(Math.round(total)), color: '#ffffff', titleSize: 14, valueSize: 32, gap: 8 }
      }
    },
    plugins: [CenterTextPlugin]
  });
}

/* ===================== Estado (pastel con %) ===================== */
function buildPieWithTitle(id, data, label, titleText) {
  destroyChart(id);
  const ctx = document.getElementById(id); if (!ctx) return;
  const colors = ['#4cc9f0', '#e76f51', '#b5179e', '#4361ee', '#2a9d8f', '#ffbe0b'];
  const labels = data.map(function (d) { return cleanText(String(d.label)); });
  const values = data.map(function (d) { return d.value; });
  const total = values.reduce(function (a, b) { return a + (Number(b) || 0); }, 0);
  const useDatalabels = (typeof ChartDataLabels !== 'undefined');
  const PiePercentFallback = {
    id: 'piePercentFallback',
    afterDatasetsDraw: function (chart) {
      if (useDatalabels) return;
      const ctx = chart.ctx;
      const meta = chart.getDatasetMeta(0);
      if (!meta || !meta.data) return;
      ctx.save();
      ctx.fillStyle = '#ffffff';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.font = 'bold 12px system-ui,Segoe UI,Roboto,Arial';
      meta.data.forEach(function (arc, i) {
        const val = Number(values[i]) || 0;
        const pct = total ? Math.round((val / total) * 100) : 0;
        const pos = arc.tooltipPosition();
        const txt = String(pct) + '%';
        ctx.strokeStyle = 'rgba(0,0,0,0.45)'; ctx.lineWidth = 3;
        ctx.strokeText(txt, pos.x, pos.y);
        ctx.fillText(txt, pos.x, pos.y);
      });
      ctx.restore();
    }
  };
  charts[id] = new Chart(ctx, {
    type: 'pie',
    data: {
      labels: labels,
      datasets: [{
        label: cleanText(label),
        data: values,
        backgroundColor: values.map(function (_, i) { return colors[i % colors.length]; }),
        borderColor: '#0b1220',
        borderWidth: 1
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: true, labels: { color: '#ffffff', font: { size: 16, weight: 'bold' } } },
        title: { display: !!titleText, text: cleanText(titleText || ''), color: '#ffffff', font: { size: 16, weight: 'bold' } },
        tooltip: {
          enabled: true, titleColor: '#ffffff', bodyColor: '#ffffff',
          callbacks: {
            label: function (ctx) {
              const val = Number(ctx.raw) || 0;
              const pct = total ? Math.round((val / total) * 100) : 0;
              return ctx.label + ': ' + Math.round(val) + ' (' + pct + '%)';
            }
          }
        },
        datalabels: useDatalabels ? {
          color: '#ffffff', anchor: 'center', align: 'center', clamp: true, clip: false,
          font: { weight: 'bold', size: 12 },
          formatter: function (val) {
            const pct = total ? Math.round((Number(val) / total) * 100) : 0;
            return String(pct) + '%';
          }
        } : undefined
      }
    },
    plugins: useDatalabels ? [ChartDataLabels] : [PiePercentFallback]
  });
}

/* ===================== Cálculos ===================== */
function getOpenClosedTotals(rows, stateCol) {
  const OPEN_TERMS = ['abierto', 'open', 'pendiente', 'en progreso', 'en curso', 'asignado', 'reabierto'];
  const CLOSED_TERMS = ['cerrado', 'closed', 'resuelto', 'finalizado', 'completado'];
  var abiertos = 0, cerrados = 0;
  rows.forEach(function (r) {
    const s = norm(r[stateCol] || '');
    if (!s) return;
    if (OPEN_TERMS.some(function (t) { return s.indexOf(norm(t)) >= 0; })) abiertos++;
    else if (CLOSED_TERMS.some(function (t) { return s.indexOf(norm(t)) >= 0; })) cerrados++;
  });
  return { abiertos: abiertos, cerrados: cerrados, total: abiertos + cerrados };
}
function getOpenClosedReturnedEmptyTotals(rows, stateCol) {
  const OPEN_TERMS = ['abierto', 'open', 'pendiente', 'en progreso', 'en curso', 'asignado', 'reabierto'];
  const CLOSED_TERMS = ['cerrado', 'closed', 'resuelto', 'finalizado', 'completado'];
  const RETURNED_TERMS = ['devuelto', 'returned'];
  var abiertos = 0, cerrados = 0, devuelto = 0, vacios = 0;
  rows.forEach(function (r) {
    const raw = r[stateCol];
    const s = norm(raw || '');
    if (!raw || s === '') { vacios++; return; }
    if (OPEN_TERMS.some(function (t) { return s.indexOf(norm(t)) >= 0; })) abiertos++;
    else if (CLOSED_TERMS.some(function (t) { return s.indexOf(norm(t)) >= 0; })) cerrados++;
    else if (RETURNED_TERMS.some(function (t) { return s.indexOf(norm(t)) >= 0; })) devuelto++;
  });
  const total = abiertos + cerrados + devuelto + vacios;
  return { abiertos: abiertos, cerrados: cerrados, devuelto: devuelto, vacios: vacios, total: total };
}
function getOpenClosedByResponsible(rows, respCol, stateCol) {
  const OPEN_TERMS = ['abierto', 'open', 'pendiente', 'en progreso', 'en curso', 'asignado', 'reabierto'];
  const CLOSED_TERMS = ['cerrado', 'closed', 'resuelto', 'finalizado', 'completado'];
  const map = new Map();
  rows.forEach(function (r) {
    const resp = cleanText(r[respCol] || '').trim();
    const s = norm(r[stateCol] || '');
    if (!resp || !s) return;
    const isOpen = OPEN_TERMS.some(function (t) { return s.indexOf(norm(t)) >= 0; });
    const isClosed = CLOSED_TERMS.some(function (t) { return s.indexOf(norm(t)) >= 0; });
    if (!isOpen && !isClosed) return;
    if (!map.has(resp)) map.set(resp, { abiertos: 0, cerrados: 0 });
    const entry = map.get(resp);
    if (isOpen) entry.abiertos += 1; else if (isClosed) entry.cerrados += 1;
  });
  return Array.from(map.entries())
    .map(function (pair) { return { label: pair[0], abiertos: pair[1].abiertos, cerrados: pair[1].cerrados, total: pair[1].abiertos + pair[1].cerrados }; })
    .sort(function (a, b) { return b.total - a.total; });
}
function getOpenClosedByCategory(rows, categoryCol, stateCol) {
  const OPEN_TERMS = ['abierto', 'open', 'pendiente', 'en progreso', 'en curso', 'asignado', 'reabierto'];
  const CLOSED_TERMS = ['cerrado', 'closed', 'resuelto', 'finalizado', 'completado'];
  const map = new Map();
  rows.forEach(function (r) {
    const cat = cleanText(r[categoryCol] || '').trim();
    const s = norm(r[stateCol] || '');
    if (!cat || !s) return;
    const isOpen = OPEN_TERMS.some(function (t) { return s.indexOf(norm(t)) >= 0; });
    const isClosed = CLOSED_TERMS.some(function (t) { return s.indexOf(norm(t)) >= 0; });
    if (!isOpen && !isClosed) return;
    if (!map.has(cat)) map.set(cat, { abiertos: 0, cerrados: 0, total: 0 });
    const entry = map.get(cat);
    if (isOpen) entry.abiertos += 1; else if (isClosed) entry.cerrados += 1;
    entry.total = entry.abiertos + entry.cerrados;
  });
  return Array.from(map.entries())
    .map(function (pair) { return { label: pair[0], abiertos: pair[1].abiertos, cerrados: pair[1].cerrados, total: pair[1].total }; })
    .sort(function (a, b) { return b.total - a.total; });
}
function getOpenClosedByProvider(rows, providerCol, stateCol) {
  const OPEN_TERMS = ['abierto', 'open', 'pendiente', 'en progreso', 'en curso', 'asignado', 'reabierto'];
  const CLOSED_TERMS = ['cerrado', 'closed', 'resuelto', 'finalizado', 'completado'];
  const map = new Map();
  rows.forEach(function (r) {
    const prov = cleanText(r[providerCol] || '').trim();
    const s = norm(r[stateCol] || '');
    if (!prov || !s) return;
    const isOpen = OPEN_TERMS.some(function (t) { return s.indexOf(norm(t)) >= 0; });
    const isClosed = CLOSED_TERMS.some(function (t) { return s.indexOf(norm(t)) >= 0; });
    if (!isOpen && !isClosed) return;
    if (!map.has(prov)) map.set(prov, { abiertos: 0, cerrados: 0, total: 0 });
    const entry = map.get(prov);
    if (isOpen) entry.abiertos += 1; else if (isClosed) entry.cerrados += 1;
    entry.total = entry.abiertos + entry.cerrados;
  });
  return Array.from(map.entries())
    .map(function (pair) { return { label: pair[0], abiertos: pair[1].abiertos, cerrados: pair[1].cerrados, total: pair[1].total }; })
    .sort(function (a, b) { return b.total - a.total; });
}

/* ===================== Tiempo (promedio) ===================== */
function getAvgTimeByRangeAndState(rows, rangeCol, timeCol, stateCol) {
  const OPEN_TERMS = ['abierto', 'open', 'pendiente', 'en progreso', 'en curso', 'asignado', 'reabierto'];
  const CLOSED_TERMS = ['cerrado', 'closed', 'resuelto', 'finalizado', 'completado'];
  const buckets = new Map();
  rows.forEach(function (r) {
    const range = cleanText(r[rangeCol] || '').trim();
    const s = norm(r[stateCol] || '');
    const time = parseNumber(r[timeCol]);
    if (!range || !s || isNaN(time)) return;
    const isOpen = OPEN_TERMS.some(function (t) { return s.indexOf(norm(t)) >= 0; });
    const isClosed = CLOSED_TERMS.some(function (t) { return s.indexOf(norm(t)) >= 0; });
    if (!isOpen && !isClosed) return;
    if (!buckets.has(range)) buckets.set(range, { open: { sum: 0, count: 0 }, closed: { sum: 0, count: 0 } });
    const b = buckets.get(range);
    if (isOpen) { b.open.sum += time; b.open.count += 1; }
    if (isClosed) { b.closed.sum += time; b.closed.count += 1; }
  });
  const rowsOut = Array.from(buckets.entries()).map(function (pair) {
    const range = pair[0], b = pair[1];
    const avgOpen = b.open.count ? b.open.sum / b.open.count : 0;
    const avgClosed = b.closed.count ? b.closed.sum / b.closed.count : 0;
    return {
      label: range,
      avgOpen: Number(avgOpen.toFixed(2)),
      avgClosed: Number(avgClosed.toFixed(2)),
      countOpen: b.open.count,
      countClosed: b.closed.count,
      weight: getRangeWeight(range)
    };
  });
  rowsOut.sort(function (a, b) { return (a.weight === b.weight) ? String(a.label).localeCompare(String(b.label)) : (a.weight - b.weight); });
  return rowsOut;
}
function getAvgTimeByStateGlobal(rows, timeCol, stateCol) {
  const OPEN_TERMS = ['abierto', 'open', 'pendiente', 'en progreso', 'en curso', 'asignado', 'reabierto'];
  const CLOSED_TERMS = ['cerrado', 'closed', 'resuelto', 'finalizado', 'completado'];
  const acc = { open: { sum: 0, count: 0 }, closed: { sum: 0, count: 0 } };
  rows.forEach(function (r) {
    const s = norm(r[stateCol] || '');
    const time = parseNumber(r[timeCol]);
    if (!s || isNaN(time)) return;
    const isOpen = OPEN_TERMS.some(function (t) { return s.indexOf(norm(t)) >= 0; });
    const isClosed = CLOSED_TERMS.some(function (t) { return s.indexOf(norm(t)) >= 0; });
    if (!isOpen && !isClosed) return;
    if (isOpen) { acc.open.sum += time; acc.open.count += 1; }
    if (isClosed) { acc.closed.sum += time; acc.closed.count += 1; }
  });
  const avgOpen = acc.open.count ? acc.open.sum / acc.open.count : 0;
  const avgClosed = acc.closed.count ? acc.closed.sum / acc.closed.count : 0;
  return {
    label: 'Global',
    avgOpen: Number(avgOpen.toFixed(2)),
    avgClosed: Number(avgClosed.toFixed(2)),
    countOpen: acc.open.count,
    countClosed: acc.closed.count
  };
}

/* ===================== Render adaptativo: gráficas ===================== */
function renderChartsAdaptive() {
  const ids = ['chartIncidentes', 'chartEstado', 'chartTiempo', 'chartResponsables', 'chartServicio', 'chartProveedor', 'chartCategoria'];
  ids.forEach(destroyChart);
  if (!rawData.length) { ids.forEach(showPlaceholder); return; }
  const headers = Object.keys(rawData[0] || {});
  schema = inferSchema(headers, rawData);
  // KPI 
  buildKPIIncidentes('chartIncidentes', rawData.length);
  // Estado (pastel con %) 
  const FIXED_STATE_COL = 'Estado Final Incidente';
  const colEstado = headers.indexOf(FIXED_STATE_COL) >= 0 ? FIXED_STATE_COL : (schema.roles.estado || null);
  if (colEstado) {
    const totals = getOpenClosedReturnedEmptyTotals(rawData, colEstado);
    const pieData = [
      { label: 'Abiertos (' + totals.abiertos + ')', value: totals.abiertos },
      { label: 'Cerrados (' + totals.cerrados + ')', value: totals.cerrados },
      { label: 'Devuelto (' + totals.devuelto + ')', value: totals.devuelto },
      { label: 'Vacíos (' + totals.vacios + ')', value: totals.vacios }
    ];
    buildPieWithTitle('chartEstado', pieData, 'Abiertos vs Cerrados', 'Total: ' + totals.total);
  } else { showPlaceholder('chartEstado'); }

  // Tiempo (Promedio) — Ahora muestra totales Abiertos vs Cerrados por categoría (Rango edad)
  var colTiempo = schema.roles.tiempo;
  // mantiene tu preferencia para detectar la columna de tiempo si hubiera varias
  const preferidasTiempo = [/edad\s*incidente/i, /dias?\s*(ans|entrega|compromiso)?/i, /tiempo/i];
  const preferida = headers.find(h => preferidasTiempo.some(rx => rx.test(h)));
  if (preferida) colTiempo = preferida;

  // Forzar el uso de la columna exacta "Rango edad"
  const FIXED_RANGE_COL = 'Rango edad';
  const colRango = headers.indexOf(FIXED_RANGE_COL) >= 0 ? FIXED_RANGE_COL : null; // ← NO permitimos fallback automático

  if (colEstado && colTiempo && colRango) {
    // En lugar de promedios, calculamos totales (abiertos/cerrados) por rango
    const rowsCat = getOpenClosedByCategory(rawData, colRango, colEstado);
    if (rowsCat.length) {
      // Orden preferido solicitado por el usuario
      const desiredOrder = ['Menor a 1 año', 'Menor a 180 días', 'Menor a 2 años', 'Menor a 3 años', 'Menor a 30 días', 'Menor a 60 días', 'Menor a 90 días'];
      const mapCat = {};
      rowsCat.forEach(function (r) { mapCat[cleanText(r.label)] = r; });

      const ordered = [];
      desiredOrder.forEach(function (d) {
        const key = cleanText(d);
        if (mapCat[key]) ordered.push(mapCat[key]);
      });
      // Añadir categorías restantes en el orden encontrado
      rowsCat.forEach(function (r) {
        const k = cleanText(r.label);
        if (!ordered.find(function (o) { return cleanText(o.label) === k; })) ordered.push(r);
      });

      const labels = ordered.map(function (r) { return r.label; });
      const abiertos = ordered.map(function (r) { return r.abiertos; });
      const cerrados = ordered.map(function (r) { return r.cerrados; });
      const percOpen = ordered.map(function (r) { var tot = r.abiertos + r.cerrados; return tot ? (r.abiertos / tot) * 100 : 0; });
      const percClosed = ordered.map(function (r) { var tot = r.abiertos + r.cerrados; return tot ? (r.cerrados / tot) * 100 : 0; });

      buildMultiBar('chartTiempo', labels, [
        { label: 'Abiertos', data: abiertos, backgroundColor: '#4cc9f0', borderColor: '#3182ce' },
        { label: 'Cerrados', data: cerrados, backgroundColor: '#2a9d8f', borderColor: '#1f776b' }
      ], {
        stacked: false,
        showEachBarPercentLabels: true,
        percentsByDataset: { 0: percOpen, 1: percClosed },
        percentTextColor: '#ffffff',
        percentInside: true
      });
    } else {
      showPlaceholder('chartTiempo');
    }
  } else {
    showPlaceholder('chartTiempo');
    console.warn('La gráfica "Tiempo Tickets (Promedio)" requiere la columna exacta "Rango edad".');
    if (!headers.includes(FIXED_RANGE_COL)) {
      alert('La gráfica "Tiempo Tickets (Promedio)" requiere la columna exacta "Rango edad" en el CSV.');
    }
  }

  // Responsables 
  const FIXED_RESP_COL = 'Ingeniero Asignado';
  const colResp = headers.indexOf(FIXED_RESP_COL) >= 0 ? FIXED_RESP_COL : (schema.roles.responsable || null);
  if (colResp && colEstado) {
    const rowsResp = getOpenClosedByResponsible(rawData, colResp, colEstado);
    const labels = rowsResp.map(function (r) { return r.label; });
    const abiertos = rowsResp.map(function (r) { return r.abiertos; });
    const cerrados = rowsResp.map(function (r) { return r.cerrados; });
    const percAbiertos = rowsResp.map(function (r) { var tot = r.abiertos + r.cerrados; return tot ? (r.abiertos / tot) * 100 : 0; });
    const percCerrados = rowsResp.map(function (r) { var tot = r.abiertos + r.cerrados; return tot ? (r.cerrados / tot) * 100 : 0; });
    buildMultiBar('chartResponsables', labels, [
      { label: 'Abiertos', data: abiertos, backgroundColor: '#4cc9f0', borderColor: '#3182ce' },
      { label: 'Cerrados', data: cerrados, backgroundColor: '#2a9d8f', borderColor: '#1f776b' }
    ], {
      stacked: false, showEachBarPercentLabels: true,
      percentsByDataset: { 0: percAbiertos, 1: percCerrados },
      percentTextColor: '#ffffff', percentInside: true
    });
  } else { showPlaceholder('chartResponsables'); }
  // Servicio (Solo Abiertos)
  const colServ = schema.roles.servicio;
  if (colServ) {
    const OPEN_TERMS = ['abierto', 'open', 'pendiente', 'en progreso', 'en curso', 'asignado', 'reabierto'];
    const casosAbiertos = rawData.filter(function (r) {
      const estado = norm(String(r[colEstado] || ''));
      return OPEN_TERMS.some(function (t) { return estado.indexOf(norm(t)) >= 0; });
    });
    const counts = countBy(casosAbiertos, colServ).map(function (d) { return { label: cleanText(d.label), value: Math.round(d.value) }; });
    const labels = counts.map(function (d) { return d.label; });
    const valores = counts.map(function (d) { return d.value; });
    const totalServ = valores.reduce(function (a, b) { return a + (Number(b) || 0); }, 0);
    const percServicios = valores.map(function (v) { return totalServ ? (Number(v) / totalServ) * 100 : 0; });
    buildMultiBar('chartServicio', labels, [
      { label: 'Casos Abiertos', data: valores, backgroundColor: '#90cdf4', borderColor: '#3182ce' }
    ], {
      stacked: false, showEachBarPercentLabels: true,
      percentsByDataset: { 0: percServicios },
      percentTextColor: '#ffffff', percentInside: true
    });
  } else { showPlaceholder('chartServicio'); }
  // Proveedor 
  const FIXED_PROV_COL = 'Proveedor a escalar';
  const colProv = headers.indexOf(FIXED_PROV_COL) >= 0 ? FIXED_PROV_COL : (schema.roles.proveedor || null);
  if (colProv && colEstado) {
    const rowsProv = getOpenClosedByProvider(rawData, colProv, colEstado);
    const labels = rowsProv.map(function (r) { return r.label; });
    const abiertos = rowsProv.map(function (r) { return r.abiertos; });
    const cerrados = rowsProv.map(function (r) { return r.cerrados; });
    const percAbiertosProv = rowsProv.map(function (r) { var tot = r.abiertos + r.cerrados; return tot ? (r.abiertos / tot) * 100 : 0; });
    const percCerradosProv = rowsProv.map(function (r) { var tot = r.abiertos + r.cerrados; return tot ? (r.cerrados / tot) * 100 : 0; });
    buildMultiBar('chartProveedor', labels, [
      { label: 'Abiertos', data: abiertos, backgroundColor: '#4cc9f0', borderColor: '#3182ce' },
      { label: 'Cerrados', data: cerrados, backgroundColor: '#2a9d8f', borderColor: '#1f776b' }
    ], {
      stacked: false, showEachBarPercentLabels: true,
      percentsByDataset: { 0: percAbiertosProv, 1: percCerradosProv },
      percentTextColor: '#ffffff', percentInside: true
    });
  } else { showPlaceholder('chartProveedor'); }
  // Categoría 
  const FIXED_CAT_COL = 'Categoría';
  const colCat = headers.indexOf(FIXED_CAT_COL) >= 0 ? FIXED_CAT_COL : (headers.find(function (h) { return /categor[ií]a/i.test(h); }) || null);
  if (colCat && colEstado) {
    const rowsCat = getOpenClosedByCategory(rawData, colCat, colEstado);
    const labels = rowsCat.map(function (r) { return r.label; });
    const abiertos = rowsCat.map(function (r) { return r.abiertos; });
    const cerrados = rowsCat.map(function (r) { return r.cerrados; });
    const percAbiertosCat = rowsCat.map(function (r) { var tot = r.abiertos + r.cerrados; return tot ? (r.abiertos / tot) * 100 : 0; });
    const percCerradosCat = rowsCat.map(function (r) { var tot = r.abiertos + r.cerrados; return tot ? (r.cerrados / tot) * 100 : 0; });
    buildMultiBar('chartCategoria', labels, [
      { label: 'Abiertos', data: abiertos, backgroundColor: '#4cc9f0', borderColor: '#3182ce' },
      { label: 'Cerrados', data: cerrados, backgroundColor: '#2a9d8f', borderColor: '#1f776b' }
    ], {
      stacked: false, showEachBarPercentLabels: true,
      percentsByDataset: { 0: percAbiertosCat, 1: percCerradosCat },
      percentTextColor: '#ffffff', percentInside: true
    });
  } else { showPlaceholder('chartCategoria'); }
}

/* ===================== Exportación PDF ===================== */
function fmtInt(n) { return Math.round(Number(n || 0)); }
function fmtPctInt(part, total) { var t = Number(total || 0); if (!t) return 0; return Math.round((Number(part || 0) / t) * 100); }
function buildCountTableRows(pairs) {
  const total = pairs.reduce(function (a, b) { return a + Number(b.value || 0); }, 0);
  return pairs.map(function (p) { return [cleanText(String(p.label)), fmtInt(p.value), String(fmtPctInt(p.value, total)) + '%']; });
}
function buildResponsablesTableRows(rowsResp) {
  return rowsResp.map(function (r) {
    const tot = (r.abiertos || 0) + (r.cerrados || 0);
    return [cleanText(String(r.label)), fmtInt(r.abiertos), String(fmtPctInt(r.abiertos, tot)) + '%', fmtInt(r.cerrados), String(fmtPctInt(r.cerrados, tot)) + '%', fmtInt(tot)];
  });
}
function buildABvsCTableRows(rows) {
  return rows.map(function (r) {
    const tot = (r.abiertos || 0) + (r.cerrados || 0);
    return [cleanText(String(r.label)), fmtInt(r.abiertos), String(fmtPctInt(r.abiertos, tot)) + '%', fmtInt(r.cerrados), String(fmtPctInt(r.cerrados, tot)) + '%', fmtInt(tot)];
  });
}
function buildServicioTableRows(counts) {
  const total = counts.reduce(function (a, b) { return a + Number(b.value || 0); }, 0);
  return counts.map(function (d) { return [cleanText(String(d.label)), fmtInt(d.value), String(fmtPctInt(d.value, total)) + '%']; });
}
function buildTiempoTableRows(rowsTiempo, isGlobal) {
  isGlobal = !!isGlobal;
  if (isGlobal) {
    const r = rowsTiempo;
    const tot = (r.countOpen || 0) + (r.countClosed || 0);
    return [['Global', fmtInt(r.avgOpen), String(fmtPctInt(r.countOpen, tot)) + '%', fmtInt(r.avgClosed), String(fmtPctInt(r.countClosed, tot)) + '%', fmtInt(tot)]];
  } else {
    return rowsTiempo.map(function (r) {
      const tot = (r.countOpen || 0) + (r.countClosed || 0);
      return [cleanText(String(r.label)), fmtInt(r.avgOpen), String(fmtPctInt(r.countOpen, tot)) + '%', fmtInt(r.avgClosed), String(fmtPctInt(r.countClosed, tot)) + '%', fmtInt(tot)];
    });
  }
}
function addTable(doc, head, body, startY, columnStyles) {
  columnStyles = columnStyles || {};
  doc.autoTable({
    startY: startY,
    head: [head],
    body: body,
    theme: 'grid',
    headStyles: { fillColor: [67, 97, 238], textColor: 255, fontStyle: 'bold', halign: 'center' },
    styles: { font: 'helvetica', fontSize: 9, cellPadding: 4, overflow: 'linebreak', cellWidth: 'wrap' },
    bodyStyles: { minCellHeight: 14 },
    alternateRowStyles: { fillColor: [245, 247, 250] },
    columnStyles: columnStyles,
    margin: { left: 40, right: 40 },
    pageBreak: 'auto',
    rowPageBreak: 'auto'
  });
  return (doc.lastAutoTable && doc.lastAutoTable.finalY) ? doc.lastAutoTable.finalY : startY;
}
function getLogoDataUrl() {
  const img = document.getElementById('logoCorp');
  if (!img || !(img.complete && img.naturalWidth)) return null;
  const w = img.naturalWidth || 600;
  const h = img.naturalHeight || 200;
  const c = document.createElement('canvas');
  c.width = w; c.height = h;
  const ctx = c.getContext('2d');
  ctx.drawImage(img, 0, 0, w, h);
  return c.toDataURL('image/png', 1.0);
}
function drawHeader(doc, titleText, palette) {
  palette = palette || { bg: { r: 12, g: 18, b: 32 }, text: { r: 255, g: 255, b: 255 } };
  const pageWidth = doc.internal.pageSize.getWidth();
  const headerH = 60;
  doc.setFillColor(palette.bg.r, palette.bg.g, palette.bg.b);
  doc.rect(0, 0, pageWidth, headerH, 'F');
  try {
    const logo = getLogoDataUrl();
    if (logo) doc.addImage(logo, 'PNG', 40, 5, 160, 50);
  } catch (e) { }
  doc.setTextColor(palette.text.r, palette.text.g, palette.text.b);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(16);
  doc.text(cleanText(titleText), pageWidth / 2, 34, { align: 'center' });
  return headerH + 20;
}
function getCanvasForPdfHD(canvasId, targetWidth, targetHeight, scaleFactor) {
  scaleFactor = scaleFactor || 2;
  const source = document.getElementById(canvasId);
  if (!source) return null;
  const temp = document.createElement('canvas');
  temp.width = Math.max(1, Math.floor(targetWidth * scaleFactor));
  temp.height = Math.max(1, Math.floor(targetHeight * scaleFactor));
  const tctx = temp.getContext('2d');
  tctx.fillStyle = '#ffffff';
  tctx.fillRect(0, 0, temp.width, temp.height);
  const sw = source.width || source.clientWidth || 1000;
  const sh = source.height || source.clientHeight || 600;
  const scale = Math.min((targetWidth * scaleFactor) / sw, (targetHeight * scaleFactor) / sh);
  const dw = Math.floor(sw * scale);
  const dh = Math.floor(sh * scale);
  const dx = Math.floor((temp.width - dw) / 2);
  const dy = Math.floor((temp.height - dh) / 2);
  tctx.imageSmoothingEnabled = true;
  tctx.imageSmoothingQuality = 'high';
  tctx.drawImage(source, dx, dy, dw, dh);
  return temp;
}
function applyPrintTheme(enable) {
  enable = !!enable;
  Object.keys(charts).forEach(function (id) {
    const ch = charts[id];
    if (!ch) return;
    const opts = ch.options || {};
    if (opts.plugins && opts.plugins.legend && opts.plugins.legend.labels) {
      opts.plugins.legend.labels.color = enable ? '#000000' : '#ffffff';
      const curr = opts.plugins.legend.labels.font || {};
      opts.plugins.legend.labels.font = { size: (enable ? 14 : 12), weight: 'bold' };
      for (var k in curr) { if (!(k in opts.plugins.legend.labels.font)) opts.plugins.legend.labels.font[k] = curr[k]; }
    }
    if (opts.plugins && opts.plugins.title) {
      opts.plugins.title.color = enable ? '#000000' : '#ffffff';
      const currT = opts.plugins.title.font || {};
      opts.plugins.title.font = { size: (enable ? 16 : 14), weight: 'bold' };
      for (var k2 in currT) { if (!(k2 in opts.plugins.title.font)) opts.plugins.title.font[k2] = currT[k2]; }
    }
    if (opts.scales && opts.scales.x && opts.scales.x.ticks) {
      opts.scales.x.ticks.color = enable ? '#000000' : '#ffffff';
      const currX = opts.scales.x.ticks.font || {};
      opts.scales.x.ticks.font = { size: (enable ? 12 : 11), weight: 'bold' };
      for (var k3 in currX) { if (!(k3 in opts.scales.x.ticks.font)) opts.scales.x.ticks.font[k3] = currX[k3]; }
    }
    if (opts.scales && opts.scales.y && opts.scales.y.ticks) {
      opts.scales.y.ticks.color = enable ? '#000000' : '#ffffff';
      const currY = opts.scales.y.ticks.font || {};
      opts.scales.y.ticks.font = { size: (enable ? 12 : 11), weight: 'bold' };
      for (var k4 in currY) { if (!(k4 in opts.scales.y.ticks.font)) opts.scales.y.ticks.font[k4] = currY[k4]; }
    }
    ch.update('none');
  });
}

/* ====== Narrativa Ejecutiva (Camino A, sin LLM) ====== */
function buildExecutiveNarrative(rawData, schema) {
  const headers = Object.keys(rawData[0] || {});
  const colEstado = headers.includes('Estado Final Incidente') ? 'Estado Final Incidente' : (schema && schema.roles ? schema.roles.estado : null);
  const colTiempo = (schema && schema.roles ? schema.roles.tiempo : null);
  const colResp = headers.includes('Ingeniero Asignado') ? 'Ingeniero Asignado' : (schema && schema.roles ? schema.roles.responsable : null);
  const colServ = (schema && schema.roles ? schema.roles.servicio : null);
  const colProv = headers.includes('Proveedor a escalar') ? 'Proveedor a escalar' : (schema && schema.roles ? schema.roles.proveedor : null);

  const total = rawData.length;
  const oc = colEstado ? getOpenClosedTotals(rawData, colEstado) : { abiertos: 0, cerrados: 0, total: total };

  let promedioTiempo = 0;
  if (colTiempo) {
    const tiempos = rawData.map(r => parseNumber(r[colTiempo])).filter(n => !isNaN(n));
    promedioTiempo = tiempos.length ? average(tiempos) : 0;
  }

  const respRows = (colResp && colEstado) ? getOpenClosedByResponsible(rawData, colResp, colEstado) : [];
  const provRows = (colProv && colEstado) ? getOpenClosedByProvider(rawData, colProv, colEstado) : [];
  const servCounts = (colServ) ? countBy(rawData, colServ) : [];

  const topRespAbiertos = respRows.slice().sort((a, b) => b.abiertos - a.abiertos).slice(0, 5);
  const topProvAbiertos = provRows.slice().sort((a, b) => b.abiertos - a.abiertos).slice(0, 5);
  const topServicios = servCounts.slice().sort((a, b) => b.value - a.value).slice(0, 5);

  const tasaResol = total ? (oc.cerrados / total) * 100 : 0;
  const cicloLento = promedioTiempo > 30;
  const backlogAlto = oc.abiertos > oc.cerrados;

  const resumen = [
    `Se analizaron ${total} incidentes.`,
    `La tasa de resolución es ${Math.round(tasaResol)}%.`,
    `El tiempo promedio del ciclo es ${Math.round(promedioTiempo)} días.`,
    backlogAlto ? `El backlog actual requiere atención (abiertos > cerrados).` : `El backlog se mantiene bajo control.`,
    cicloLento ? `El ciclo operativo muestra lentitud (promedio > 30 días).` : `El ciclo operativo se mantiene dentro de parámetros aceptables.`
  ];

  const riesgos = [
    backlogAlto ? `Incremento de casos abiertos frente a los cerrados.` : null,
    cicloLento ? `Tiempo de ciclo elevado—riesgo de incumplimiento de ANS.` : null,
    (topRespAbiertos[0]?.abiertos || 0) > 0 ? `Concentración de backlog en ciertos responsables.` : null,
    (topProvAbiertos[0]?.abiertos || 0) > 0 ? `Dependencia de proveedores con cola de cierre alta.` : null
  ].filter(Boolean);

  const acciones = [
    `Priorizar cierre de los Top-5 responsables con más abiertos.`,
    `Escalar con los Top-5 proveedores críticos.`,
    `Revisar causas en los servicios con mayor volumen (Top-5) y definir planes de mitigación.`,
    cicloLento ? `Implementar SLA internos para reducir el tiempo promedio y reforzar monitoreo diario.` : `Mantener el ritmo de cierre y monitoreo semanal.`
  ];

  return { resumen, riesgos, acciones, topRespAbiertos, topProvAbiertos, topServicios };
}
function addExecutiveInsightsToPdf(doc, narrative, startY) {
  let y = startY;
  doc.setFont('helvetica', 'bold'); doc.setFontSize(12);
  doc.text('Resumen Ejecutivo', 40, y); y += 14;
  doc.setFont('helvetica', 'normal'); doc.setFontSize(10);

  narrative.resumen.forEach(line => { doc.text(cleanText(line), 40, y); y += 12; });

  y += 6; doc.setFont('helvetica', 'bold'); doc.text('Riesgos/Señales', 40, y); y += 14;
  doc.setFont('helvetica', 'normal');
  narrative.riesgos.forEach(line => { doc.text('• ' + cleanText(line), 40, y); y += 12; });

  y += 6; doc.setFont('helvetica', 'bold'); doc.text('Acciones recomendadas (próximos 30 días)', 40, y); y += 14;
  doc.setFont('helvetica', 'normal');
  narrative.acciones.forEach(line => { doc.text('• ' + cleanText(line), 40, y); y += 12; });

  const headResp = ['Responsable', 'Abiertos', 'Cerrados', 'Total'];
  const bodyResp = narrative.topRespAbiertos.map(r => [cleanText(r.label), fmtInt(r.abiertos), fmtInt(r.cerrados), fmtInt(r.total)]);
  y = addTable(doc, headResp, bodyResp, y + 10, { 0: { cellWidth: 220 }, 1: { cellWidth: 80 }, 2: { cellWidth: 80 }, 3: { cellWidth: 80 } });

  const headProv = ['Proveedor', 'Abiertos', 'Cerrados', 'Total'];
  const bodyProv = narrative.topProvAbiertos.map(r => [cleanText(r.label), fmtInt(r.abiertos), fmtInt(r.cerrados), fmtInt(r.total)]);
  y = addTable(doc, headProv, bodyProv, y + 10, { 0: { cellWidth: 220 }, 1: { cellWidth: 80 }, 2: { cellWidth: 80 }, 3: { cellWidth: 80 } });

  const headServ = ['Servicio', 'Casos', '% del total'];
  const bodyServ = buildCountTableRows(narrative.topServicios.map(s => ({ label: s.label, value: s.value })));
  addTable(doc, headServ, bodyServ, y + 10, { 0: { cellWidth: 220 }, 1: { cellWidth: 80 }, 2: { cellWidth: 100 } });
}

/* ====== *** NUEVO (Camino B) *** Generar narrativa con LLM local ====== */
/** Construye el payload con KPIs y Top-5 para el prompt del LLM */
function buildKpiPayload(schema, headers) {
  const FIXED_STATE_COL = 'Estado Final Incidente';
  const colEstado = headers.indexOf(FIXED_STATE_COL) >= 0 ? FIXED_STATE_COL : (schema.roles.estado || null);
  const totalIncidentes = rawData.length;
  const openClosed = colEstado ? getOpenClosedTotals(rawData, colEstado) : { abiertos: 0, cerrados: 0, total: totalIncidentes };

  const colTiempo = schema.roles.tiempo || null;
  let promedioTiempo = 0;
  if (colTiempo) {
    const tiempos = rawData.map(r => parseNumber(r[colTiempo])).filter(n => !isNaN(n));
    if (tiempos.length) promedioTiempo = average(tiempos);
  }

  const FIXED_RESP_COL = 'Ingeniero Asignado';
  const colResp = headers.indexOf(FIXED_RESP_COL) >= 0 ? FIXED_RESP_COL : (schema.roles.responsable || null);
  const FIXED_PROV_COL = 'Proveedor a escalar';
  const colProv = headers.indexOf(FIXED_PROV_COL) >= 0 ? FIXED_PROV_COL : (schema.roles.proveedor || null);
  const colServ = schema.roles.servicio || null;

  const topResp = (colResp && colEstado) ? getOpenClosedByResponsible(rawData, colResp, colEstado).slice(0, 5) : [];
  const topProv = (colProv && colEstado) ? getOpenClosedByProvider(rawData, colProv, colEstado).slice(0, 5) : [];
  const topServ = (colServ) ? countBy(rawData, colServ).slice(0, 5) : [];

  const tasaResolucion = Math.round((openClosed.cerrados / Math.max(1, totalIncidentes)) * 100);

  return {
    totalIncidentes,
    abiertos: openClosed.abiertos,
    cerrados: openClosed.cerrados,
    tiempoPromedioDias: Math.round(promedioTiempo),
    tasaResolucion,
    topResponsables: topResp,
    topProveedores: topProv,
    topServicios: topServ
  };
}

/** LLM via Ollama: /api/generate (no stream) — respuesta en .response */
async function llmViaOllama(payloadKPI) {
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), LLM_TIMEOUT_MS);

  const prompt = `
Actúa como analista de operaciones TI para un comité directivo.
Con los siguientes indicadores, redacta una narrativa ejecutiva (150–220 palabras),
incluye 3 riesgos y 4 acciones priorizadas (viñetas). Sé claro y específico.

Datos:
${JSON.stringify(payloadKPI, null, 2)}
`.trim();

  const res = await fetch(LLM_OLLAMA_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: LLM_OLLAMA_MODEL,
      prompt,
      stream: false,        // respuesta completa en un JSON
      options: { temperature: 0.7 }
    }),
    signal: controller.signal,
    mode: 'cors'           // si usas CORS con OLLAMA_ORIGINS
  });
  clearTimeout(t);
  if (!res.ok) throw new Error(`Ollama error HTTP ${res.status}`);
  const json = await res.json();
  // Ollama devuelve { response: "..." , done: true, ... }
  return String(json.response || '').trim();
}

/** LLM via LM Studio: /v1/chat/completions (OpenAI-compatible) */
async function llmViaLmStudio(payloadKPI) {
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), LLM_TIMEOUT_MS);

  const messages = [
    { role: 'system', content: 'Eres analista de operaciones TI. Respondes con claridad ejecutiva.' },
    {
      role: 'user', content: `
Redacta una narrativa ejecutiva (150–220 palabras) para comité directivo.
Incluye 3 riesgos y 4 acciones priorizadas (viñetas). Sé claro y específico.

Datos:
${JSON.stringify(payloadKPI, null, 2)}
`.trim()
    }
  ];

  const res = await fetch(LLM_LMSTUDIO_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: 'your-model-name',   // Ajusta al nombre exacto cargado en LM Studio
      messages,
      stream: false
    }),
    signal: controller.signal
  });
  clearTimeout(t);
  if (!res.ok) throw new Error(`LM Studio error HTTP ${res.status}`);
  const json = await res.json();
  const choice = json && json.choices && json.choices[0];
  const content = choice && (choice.message && choice.message.content);
  return String(content || '').trim();
}

/** Unified: intenta provider seleccionado; si falla, devuelve null */
async function getLLMExecutiveSummary(payloadKPI) {
  if (!LLM_ENABLED) return null;
  try {
    if (LLM_PROVIDER === 'ollama') {
      // Ollama /api/generate, stream=false. Campo "response".  [1](https://docs.ollama.com/api/generate)
      return await llmViaOllama(payloadKPI);
    } else if (LLM_PROVIDER === 'lmstudio') {
      // LM Studio OpenAI-compatible /v1/chat/completions.  [4](https://lmstudio.ai/docs/developer/core/server)
      return await llmViaLmStudio(payloadKPI);
    } else {
      return null;
    }
  } catch (err) {
    console.warn('LLM no disponible / error:', err);
    return null;
  }
}

/** Añade página LLM al PDF si hay texto */
function addLLMPageToPdf(doc, llmText) {
  if (!llmText) return;
  addLandscapePage(doc);
  const y = drawHeader(doc, 'Narrativa (LLM)', { bg: { r: 12, g: 18, b: 32 }, text: { r: 255, g: 255, b: 255 } });
  doc.setFont('helvetica', 'normal'); doc.setFontSize(10);
  const pageWidth = doc.internal.pageSize.getWidth();
  // Ajuste de ancho de texto (márgenes 40–40)
  const lines = doc.splitTextToSize(cleanText(llmText), pageWidth - 80);
  doc.text(lines, 40, y + 10);
}

/* ====== Reporte PDF Ejecutivo — Landscape + LLM ====== */
async function generatePDFReport() {
  if (!rawData.length) {
    alert('No hay datos cargados. Por favor carga un archivo CSV primero.');
    return;
  }
  var jsPDFlib = (window.jspdf && window.jspdf.jsPDF) ? window.jspdf.jsPDF : window.jsPDF;
  if (!jsPDFlib) {
    alert('No se encontró jsPDF. Verifica que el CDN esté cargado antes de app.js.');
    return;
  }
  const doc = new jsPDFlib(PAGE_ORIENTATION, PAGE_UNIT, PAGE_FORMAT);
  var hasAutoTable = (typeof doc.autoTable === 'function')
    || (jsPDFlib && jsPDFlib.API && typeof jsPDFlib.API.autoTable === 'function')
    || (window.jspdf && typeof window.jspdf.autoTable === 'function');
  if (!hasAutoTable) {
    alert('No se encontró AutoTable. Verifica que el CDN de jspdf-autotable esté cargado.');
    return;
  }
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const COLOR_BG = { r: 12, g: 18, b: 32 };
  const COLOR_TEXT = { r: 25, g: 25, b: 25 };

  // Portada + resumen (landscape)
  var y = drawHeader(doc, 'Reporte Ejecutivo — Dashboard Incidentes TI', { bg: COLOR_BG, text: { r: 255, g: 255, b: 255 } });
  doc.setTextColor(COLOR_TEXT.r, COLOR_TEXT.g, COLOR_TEXT.b);
  doc.setFont('helvetica', 'normal'); doc.setFontSize(10);
  doc.text(cleanText('Fecha de generación: ' + new Date().toLocaleString('es-CO')), 40, y);
  doc.text(cleanText('Archivo fuente: ' + (window.__lastCsvName || 'Datos cargados')), 40, y + 16);

  const headers = Object.keys(rawData[0] || {});
  schema = inferSchema(headers, rawData);
  const FIXED_STATE_COL = 'Estado Final Incidente';
  const colEstado = headers.indexOf(FIXED_STATE_COL) >= 0 ? FIXED_STATE_COL : (schema.roles.estado || null);
  const totalIncidentes = rawData.length;
  const openClosed = colEstado ? getOpenClosedTotals(rawData, colEstado) : { abiertos: 0, cerrados: 0 };
  const colTiempo = schema.roles.tiempo;
  var promedioTiempo = 0;
  if (colTiempo) {
    const tiempos = rawData.map(function (r) { return parseNumber(r[colTiempo]); }).filter(function (n) { return !isNaN(n); });
    if (tiempos.length) promedioTiempo = average(tiempos);
  }
  y = addTable(
    doc,
    ['Métrica', 'Valor', 'Interpretación'],
    [
      ['Total de Incidentes', fmtInt(totalIncidentes), (totalIncidentes > 0 ? 'Carga activa' : 'Sin carga')],
      ['Incidentes Abiertos', fmtInt(openClosed.abiertos), (openClosed.abiertos > openClosed.cerrados ? 'Requiere atención' : 'Bajo control')],
      ['Incidentes Cerrados', fmtInt(openClosed.cerrados), (openClosed.cerrados >= openClosed.abiertos ? 'Buena resolución' : 'Mejorar cierre')],
      ['Tiempo Promedio (días)', fmtInt(promedioTiempo), (promedioTiempo > 30 ? 'Ciclo lento' : 'Ciclo adecuado')],
      ['Tasa de Resolución', String(fmtInt((openClosed.cerrados / Math.max(1, totalIncidentes)) * 100)) + '%', 'Efectividad del equipo']
    ],
    y + 30,
    { 0: { cellWidth: 180 }, 1: { cellWidth: 80 }, 2: { cellWidth: 220 } }
  );

  // Narrativa ejecutiva (Camino A, sin LLM)
  const narrative = buildExecutiveNarrative(rawData, schema);
  addExecutiveInsightsToPdf(doc, narrative, y + 16);

  // *** Camino B: intentar añadir Narrativa (LLM) ***
  const payloadKPI = buildKpiPayload(schema, headers);
  const llmText = await getLLMExecutiveSummary(payloadKPI);
  addLLMPageToPdf(doc, llmText);

  // Resto de secciones (Estado / Responsables / Servicio / Proveedores / Categoría / Tiempo) como antes…
  if (colEstado) {
    const t = getOpenClosedReturnedEmptyTotals(rawData, colEstado);
    const estadoPairs = [
      { label: 'Abiertos (' + t.abiertos + ')', value: t.abiertos },
      { label: 'Cerrados (' + t.cerrados + ')', value: t.cerrados },
      { label: 'Devuelto (' + t.devuelto + ')', value: t.devuelto },
      { label: 'Vacíos (' + t.vacios + ')', value: t.vacios }
    ];
    addLandscapePage(doc);
    y = drawHeader(doc, 'Estado — Distribución (Tabla)', { bg: COLOR_BG, text: { r: 255, g: 255, b: 255 } });
    addTable(doc, ['Estado', 'Casos', '% del total'], buildCountTableRows(estadoPairs), y,
      { 0: { cellWidth: 200 }, 1: { cellWidth: 80 }, 2: { cellWidth: 100 } }
    );
    addLandscapePage(doc);
    y = drawHeader(doc, 'Estado — Distribución (Gráfico)', { bg: COLOR_BG, text: { r: 255, g: 255, b: 255 } });
    applyPrintTheme(true);
    const graphLeft = 40, graphTop = y;
    const maxWidth = Math.floor(pageWidth - 80);
    const maxHeight = Math.floor(pageHeight - (y + 40));
    const w16_9 = maxWidth;
    const h16_9 = Math.min(maxHeight, Math.floor(w16_9 * 9 / 16));
    const canvasHD = getCanvasForPdfHD('chartEstado', w16_9, h16_9, 3);
    if (canvasHD) doc.addImage(canvasHD.toDataURL('image/png', 1.0), 'PNG', graphLeft, graphTop, w16_9, h16_9);
    applyPrintTheme(false);
  }

  const FIXED_RESP_COL = 'Ingeniero Asignado';
  const colResp = headers.indexOf(FIXED_RESP_COL) >= 0 ? FIXED_RESP_COL : (schema.roles.responsable || null);
  if (colResp && colEstado) {
    const rowsResp = getOpenClosedByResponsible(rawData, colResp, colEstado);
    addLandscapePage(doc);
    y = drawHeader(doc, 'Responsables — Abiertos vs Cerrados (Tabla)', { bg: COLOR_BG, text: { r: 255, g: 255, b: 255 } });
    addTable(doc, ['Responsable', 'Abiertos', '% fila', 'Cerrados', '% fila', 'Total'], buildResponsablesTableRows(rowsResp), y,
      { 0: { cellWidth: 200 }, 1: { cellWidth: 80 }, 2: { cellWidth: 80 }, 3: { cellWidth: 80 }, 4: { cellWidth: 80 }, 5: { cellWidth: 80 } }
    );
    addLandscapePage(doc);
    y = drawHeader(doc, 'Responsables — Abiertos vs Cerrados (Gráfico)', { bg: COLOR_BG, text: { r: 255, g: 255, b: 255 } });
    applyPrintTheme(true);
    const graphLeft = 40, graphTop = y;
    const maxWidth = Math.floor(pageWidth - 80);
    const maxHeight = Math.floor(pageHeight - (y + 40));
    const w16_9 = maxWidth;
    const h16_9 = Math.min(maxHeight, Math.floor(w16_9 * 9 / 16));
    const canvasHD = getCanvasForPdfHD('chartResponsables', w16_9, h16_9, 3);
    if (canvasHD) doc.addImage(canvasHD.toDataURL('image/png', 1.0), 'PNG', graphLeft, graphTop, w16_9, h16_9);
    applyPrintTheme(false);
  }

  const colServ = schema.roles.servicio;
  if (colServ) {
    const countsServ = countBy(rawData, colServ).map(function (d) { return { label: cleanText(d.label), value: Math.round(d.value) }; });
    addLandscapePage(doc);
    y = drawHeader(doc, 'Tipificación por Servicio (Tabla)', { bg: COLOR_BG, text: { r: 255, g: 255, b: 255 } });
    addTable(doc, ['Servicio', 'Casos', '% del total'], buildServicioTableRows(countsServ), y,
      { 0: { cellWidth: 220 }, 1: { cellWidth: 80 }, 2: { cellWidth: 100 } }
    );
    addLandscapePage(doc);
    y = drawHeader(doc, 'Tipificación por Servicio (Gráfico)', { bg: COLOR_BG, text: { r: 255, g: 255, b: 255 } });
    applyPrintTheme(true);
    const graphLeft = 40, graphTop = y;
    const maxWidth = Math.floor(pageWidth - 80);
    const maxHeight = Math.floor(pageHeight - (y + 40));
    const w16_9 = maxWidth;
    const h16_9 = Math.min(maxHeight, Math.floor(w16_9 * 9 / 16));
    const canvasHD = getCanvasForPdfHD('chartServicio', w16_9, h16_9, 3);
    if (canvasHD) doc.addImage(canvasHD.toDataURL('image/png', 1.0), 'PNG', graphLeft, graphTop, w16_9, h16_9);
    applyPrintTheme(false);
  }

  const FIXED_PROV_COL = 'Proveedor a escalar';
  const colProv = headers.indexOf(FIXED_PROV_COL) >= 0 ? FIXED_PROV_COL : (schema.roles.proveedor || null);
  if (colProv && colEstado) {
    const rowsProv = getOpenClosedByProvider(rawData, colProv, colEstado);
    addLandscapePage(doc);
    y = drawHeader(doc, 'Proveedores — Abiertos vs Cerrados (Tabla)', { bg: COLOR_BG, text: { r: 255, g: 255, b: 255 } });
    addTable(doc, ['Proveedor', 'Abiertos', '% fila', 'Cerrados', '% fila', 'Total'], buildABvsCTableRows(rowsProv), y,
      { 0: { cellWidth: 220 }, 1: { cellWidth: 80 }, 2: { cellWidth: 80 }, 3: { cellWidth: 80 }, 4: { cellWidth: 80 }, 5: { cellWidth: 80 } }
    );
    addLandscapePage(doc);
    y = drawHeader(doc, 'Proveedores — Abiertos vs Cerrados (Gráfico)', { bg: COLOR_BG, text: { r: 255, g: 255, b: 255 } });
    applyPrintTheme(true);
    const graphLeft = 40, graphTop = y;
    const maxWidth = Math.floor(pageWidth - 80);
    const maxHeight = Math.floor(pageHeight - (y + 40));
    const w16_9 = maxWidth;
    const h16_9 = Math.min(maxHeight, Math.floor(w16_9 * 9 / 16));
    const canvasHD = getCanvasForPdfHD('chartProveedor', w16_9, h16_9, 3);
    if (canvasHD) doc.addImage(canvasHD.toDataURL('image/png', 1.0), 'PNG', graphLeft, graphTop, w16_9, h16_9);
    applyPrintTheme(false);
  }

  const FIXED_CAT_COL = 'Categoría';
  const colCat = headers.indexOf(FIXED_CAT_COL) >= 0 ? FIXED_CAT_COL : (headers.find(function (h) { return /categor[ií]a/i.test(h); }) || null);
  if (colCat && colEstado) {
    const rowsCat = getOpenClosedByCategory(rawData, colCat, colEstado);
    addLandscapePage(doc);
    y = drawHeader(doc, 'Categoría — Abiertos vs Cerrados (Tabla)', { bg: COLOR_BG, text: { r: 255, g: 255, b: 255 } });
    addTable(doc, ['Categoría', 'Abiertos', '% fila', 'Cerrados', '% fila', 'Total'], buildABvsCTableRows(rowsCat), y,
      { 0: { cellWidth: 220 }, 1: { cellWidth: 80 }, 2: { cellWidth: 80 }, 3: { cellWidth: 80 }, 4: { cellWidth: 80 }, 5: { cellWidth: 80 } }
    );
    addLandscapePage(doc);
    y = drawHeader(doc, 'Categoría — Abiertos vs Cerrados (Gráfico)', { bg: COLOR_BG, text: { r: 255, g: 255, b: 255 } });
    applyPrintTheme(true);
    const graphLeft = 40, graphTop = y;
    const maxWidth = Math.floor(pageWidth - 80);
    const maxHeight = Math.floor(pageHeight - (y + 40));
    const w16_9 = maxWidth;
    const h16_9 = Math.min(maxHeight, Math.floor(w16_9 * 9 / 16));
    const canvasHD = getCanvasForPdfHD('chartCategoria', w16_9, h16_9, 3);
    if (canvasHD) doc.addImage(canvasHD.toDataURL('image/png', 1.0), 'PNG', graphLeft, graphTop, w16_9, h16_9);
    applyPrintTheme(false);
  }

  const colTiempoPDF = schema.roles.tiempo;
  if (colTiempoPDF && colEstado) {
    const colRango = schema.roles.rangoEdad || headers.find(function (h) { return /rango\s*\_?\s*edad/i.test(h); });
    addLandscapePage(doc);
    y = drawHeader(doc, 'Tiempo Tickets (Promedio) — Abiertos vs Cerrados (Tabla)', { bg: COLOR_BG, text: { r: 255, g: 255, b: 255 } });
    if (colRango) {
      const rowsT = getAvgTimeByRangeAndState(rawData, colRango, colTiempoPDF, colEstado);
      addTable(doc, ['Rango', 'Prom. Abiertos (días)', '% fila', 'Prom. Cerrados (días)', '% fila', 'Casos fila'], buildTiempoTableRows(rowsT, false), y,
        { 0: { cellWidth: 180 }, 1: { cellWidth: 120 }, 2: { cellWidth: 80 }, 3: { cellWidth: 120 }, 4: { cellWidth: 80 }, 5: { cellWidth: 80 } }
      );
    } else {
      const g = getAvgTimeByStateGlobal(rawData, colTiempoPDF, colEstado);
      addTable(doc, ['Rango', 'Prom. Abiertos (días)', '% fila', 'Prom. Cerrados (días)', '% fila', 'Casos fila'], buildTiempoTableRows(g, true), y,
        { 0: { cellWidth: 180 }, 1: { cellWidth: 120 }, 2: { cellWidth: 80 }, 3: { cellWidth: 120 }, 4: { cellWidth: 80 }, 5: { cellWidth: 80 } }
      );
    }
    addLandscapePage(doc);
    y = drawHeader(doc, 'Tiempo Tickets (Promedio) — Abiertos vs Cerrados (Gráfico)', { bg: COLOR_BG, text: { r: 255, g: 255, b: 255 } });
    applyPrintTheme(true);
    const graphLeft = 40, graphTop = y;
    const maxWidth = Math.floor(pageWidth - 80);
    const maxHeight = Math.floor(pageHeight - (y + 40));
    const w16_9 = maxWidth;
    const h16_9 = Math.min(maxHeight, Math.floor(w16_9 * 9 / 16));
    const canvasHD = getCanvasForPdfHD('chartTiempo', w16_9, h16_9, 3);
    if (canvasHD) doc.addImage(canvasHD.toDataURL('image/png', 1.0), 'PNG', graphLeft, graphTop, w16_9, h16_9);
    applyPrintTheme(false);
  }

  const totalPages = doc.internal.getNumberOfPages();
  for (var i = 1; i <= totalPages; i++) {
    doc.setPage(i);
    const pw = doc.internal.pageSize.getWidth();
    const ph = doc.internal.pageSize.getHeight();
    doc.setFontSize(9);
    doc.setTextColor(120);
    doc.text(cleanText('Página ' + i + ' de ' + totalPages + ' — Dashboard Incidentes TI'), pw / 2, ph - 30, { align: 'center' });
    doc.text(cleanText('John Jairo Vargas González — Ingeniero de Soluciones TI — john.vargas@bancounion.com'), pw / 2, ph - 16, { align: 'center' });
  }
  const nombreArchivo = 'Reporte_Ejecutivo_Incidentes_' + new Date().toISOString().slice(0, 10) + '.pdf';
  doc.save(nombreArchivo);
}

/* ===================== Exportación Excel ===================== */
function generateExcelReport() {
  if (!rawData.length) {
    alert('No hay datos cargados. Por favor carga un archivo CSV primero.');
    return;
  }
  const wb = XLSX.utils.book_new();
  const totalIncidentes = rawData.length;
  const headers = Object.keys(rawData[0] || {});
  schema = inferSchema(headers, rawData);
  const colEstado = (schema && schema.roles && schema.roles.estado) ? schema.roles.estado : 'Estado Final Incidente';
  const openClosed = getOpenClosedTotals(rawData, colEstado);
  const colTiempo = (schema && schema.roles && schema.roles.tiempo) ? schema.roles.tiempo : null;
  var promedioTiempo = 0;
  if (colTiempo) {
    const tiempos = rawData.map(function (r) { return parseNumber(r[colTiempo]); }).filter(function (n) { return !isNaN(n); });
    if (tiempos.length) promedioTiempo = average(tiempos);
  }
  const resumenData = [
    ['REPORTE EJECUTIVO - BACKLOG INCIDENTES TI'],
    ['Fecha de Generación:', new Date().toLocaleString('es-CO')],
    ['Archivo Fuente:', cleanText(window.__lastCsvName || 'Datos cargados')],
    [''],
    ['MÉTRICAS PRINCIPALES'],
    ['Métrica', 'Valor', 'Notas'],
    ['Total de Incidentes', Math.round(totalIncidentes), totalIncidentes > 0 ? 'Carga activa' : 'Sin carga'],
    ['Incidentes Abiertos', Math.round(openClosed.abiertos), openClosed.abiertos > openClosed.cerrados ? 'Requiere atención' : 'Bajo control'],
    ['Incidentes Cerrados', Math.round(openClosed.cerrados), openClosed.cerrados >= openClosed.abiertos ? 'Buena resolución' : 'Mejorar cierre'],
    ['Tiempo Promedio (días)', Math.round(Number(promedioTiempo)), promedioTiempo > 30 ? 'Ciclo lento' : 'Ciclo adecuado'],
    ['Tasa de Resolución', String(Math.round((openClosed.cerrados / Math.max(1, totalIncidentes)) * 100)) + '%', 'Efectividad del equipo']
  ];
  const wsResumen = XLSX.utils.aoa_to_sheet(resumenData);
  wsResumen['!cols'] = [{ wch: 32 }, { wch: 24 }, { wch: 44 }];
  wsResumen['!freeze'] = { rows: 1, cols: 0 };
  XLSX.utils.book_append_sheet(wb, wsResumen, 'Resumen Ejecutivo');

  const wsData = XLSX.utils.json_to_sheet(rawData);
  const colsKeys = Object.keys(rawData[0] || {});
  wsData['!cols'] = colsKeys.map(function (k) { return { wch: Math.min(30, Math.max(10, k.length + 6)) }; });
  wsData['!freeze'] = { rows: 1, cols: 0 };
  XLSX.utils.book_append_sheet(wb, wsData, 'Datos Completos');

  // Hoja “Hallazgos” — narrativa sin LLM
  const narrative = buildExecutiveNarrative(rawData, schema);
  const sheetHallazgos = [
    ['HALLAZGOS Y ACCIONES'],
    ['Resumen Ejecutivo'],
    ...narrative.resumen.map(line => [cleanText(line)]),
    [''],
    ['Riesgos/Señales'],
    ...narrative.riesgos.map(line => ['• ' + cleanText(line)]),
    [''],
    ['Acciones recomendadas (30 días)'],
    ...narrative.acciones.map(line => ['• ' + cleanText(line)]),
    [''],
    ['Top-5 Responsables (por abiertos)'],
    ['Responsable', 'Abiertos', 'Cerrados', 'Total'],
    ...narrative.topRespAbiertos.map(r => [cleanText(r.label), fmtInt(r.abiertos), fmtInt(r.cerrados), fmtInt(r.total)]),
    [''],
    ['Top-5 Proveedores (por abiertos)'],
    ['Proveedor', 'Abiertos', 'Cerrados', 'Total'],
    ...narrative.topProvAbiertos.map(r => [cleanText(r.label), fmtInt(r.abiertos), fmtInt(r.cerrados), fmtInt(r.total)]),
    [''],
    ['Top-5 Servicios (por casos)'],
    ['Servicio', 'Casos', '% del total'],
    ...buildCountTableRows(narrative.topServicios.map(s => ({ label: s.label, value: s.value })))
  ];
  const wsHall = XLSX.utils.aoa_to_sheet(sheetHallazgos);
  wsHall['!cols'] = [{ wch: 40 }, { wch: 20 }, { wch: 20 }, { wch: 20 }];
  wsHall['!freeze'] = { rows: 1, cols: 0 };
  XLSX.utils.book_append_sheet(wb, wsHall, 'Hallazgos');

  // *** NUEVO: Hoja Narrativa (LLM) si disponible ***
  (async () => {
    if (LLM_ENABLED) {
      const payloadKPI = buildKpiPayload(schema, headers);
      const llmText = await getLLMExecutiveSummary(payloadKPI);
      if (llmText) {
        const lines = llmText.split('\n').map(s => [cleanText(s)]);
        const wsLLM = XLSX.utils.aoa_to_sheet([['NARRATIVA (LLM)'], ...lines]);
        wsLLM['!cols'] = [{ wch: 96 }];
        XLSX.utils.book_append_sheet(wb, wsLLM, 'Narrativa (LLM)');
      }
    }
    const nombreArchivo = 'Reporte_Detallado_Incidentes_' + new Date().toISOString().slice(0, 10) + '.xlsx';
    XLSX.writeFile(wb, nombreArchivo);
  })();
}

/* ===================== Mask plugins (ocultan % al ocultar datasets) ===================== */
const PercentVisibilityMaskPlugin = {
  id: 'percentVisibilityMask',
  beforeDatasetsDraw: function (chart) {
    try {
      if (!chart.canvas || chart.canvas.id !== 'chartResponsables') return;
      const cfg = chart.options && chart.options.plugins ? chart.options.plugins.perBarPercentLabels : null;
      if (!cfg || !cfg.mapping || typeof cfg.mapping !== 'object') return;
      if (!chart.$percentMappingOriginal) {
        const orig = {};
        Object.keys(cfg.mapping).forEach(function (k) {
          const arr = cfg.mapping[k];
          orig[k] = Array.isArray(arr) ? arr.slice() : arr;
        });
        chart.$percentMappingOriginal = orig;
      }
      const restored = {};
      Object.keys(chart.$percentMappingOriginal).forEach(function (k) {
        const arr = chart.$percentMappingOriginal[k];
        restored[k] = Array.isArray(arr) ? arr.slice() : arr;
      });
      const masked = {};
      Object.keys(restored).forEach(function (key) {
        const dsIndex = Number(key);
        const meta = chart.getDatasetMeta(dsIndex);
        const isVisible = (typeof chart.isDatasetVisible === 'function')
          ? chart.isDatasetVisible(dsIndex)
          : !(meta && meta.hidden === true);
        masked[dsIndex] = isVisible ? restored[dsIndex] : [];
      });
      chart.options.plugins.perBarPercentLabels.mapping = masked;
    } catch (e) { console.warn('PercentVisibilityMaskPlugin error:', e); }
  }
};
const PercentVisibilityMaskPluginCat = {
  id: 'percentVisibilityMaskCat',
  beforeDatasetsDraw: function (chart) {
    try {
      if (!chart.canvas || chart.canvas.id !== 'chartCategoria') return;
      const cfg = chart.options && chart.options.plugins ? chart.options.plugins.perBarPercentLabelsCat : null;
      if (!cfg || !cfg.mapping || typeof cfg.mapping !== 'object') return;
      if (!chart.$percentMappingOriginalCat) {
        const orig = {};
        Object.keys(cfg.mapping).forEach(function (k) {
          const arr = cfg.mapping[k];
          orig[k] = Array.isArray(arr) ? arr.slice() : arr;
        });
        chart.$percentMappingOriginalCat = orig;
      }
      const restored = {};
      Object.keys(chart.$percentMappingOriginalCat).forEach(function (k) {
        const arr = chart.$percentMappingOriginalCat[k];
        restored[k] = Array.isArray(arr) ? arr.slice() : arr;
      });
      const masked = {};
      Object.keys(restored).forEach(function (key) {
        const dsIndex = Number(key);
        const meta = chart.getDatasetMeta(dsIndex);
        const isVisible = (typeof chart.isDatasetVisible === 'function')
          ? chart.isDatasetVisible(dsIndex)
          : !(meta && meta.hidden === true);
        masked[dsIndex] = isVisible ? restored[dsIndex] : [];
      });
      chart.options.plugins.perBarPercentLabelsCat.mapping = masked;
    } catch (e) { console.warn('PercentVisibilityMaskPluginCat error:', e); }
  }
};
const PercentVisibilityMaskPluginProv = {
  id: 'percentVisibilityMaskProv',
  beforeDatasetsDraw: function (chart) {
    try {
      if (!chart.canvas || chart.canvas.id !== 'chartProveedor') return;
      const cfg = chart.options && chart.options.plugins ? chart.options.plugins.perBarPercentLabelsProv : null;
      if (!cfg || !cfg.mapping || typeof cfg.mapping !== 'object') return;
      if (!chart.$percentMappingOriginalProv) {
        const orig = {};
        Object.keys(cfg.mapping).forEach(function (k) {
          const arr = cfg.mapping[k];
          orig[k] = Array.isArray(arr) ? arr.slice() : arr;
        });
        chart.$percentMappingOriginalProv = orig;
      }
      const restored = {};
      Object.keys(chart.$percentMappingOriginalProv).forEach(function (k) {
        const arr = chart.$percentMappingOriginalProv[k];
        restored[k] = Array.isArray(arr) ? arr.slice() : arr;
      });
      const masked = {};
      Object.keys(restored).forEach(function (key) {
        const dsIndex = Number(key);
        const meta = chart.getDatasetMeta(dsIndex);
        const isVisible = (typeof chart.isDatasetVisible === 'function')
          ? chart.isDatasetVisible(dsIndex)
          : !(meta && meta.hidden === true);
        masked[dsIndex] = isVisible ? restored[dsIndex] : [];
      });
      chart.options.plugins.perBarPercentLabelsProv.mapping = masked;
    } catch (e) { console.warn('PercentVisibilityMaskPluginProv error:', e); }
  }
};
const PercentVisibilityMaskPluginTime = {
  id: 'percentVisibilityMaskTime',
  beforeDatasetsDraw: function (chart) {
    try {
      if (!chart.canvas || chart.canvas.id !== 'chartTiempo') return;
      const cfg = chart.options && chart.options.plugins ? chart.options.plugins.perBarPercentLabelsTime : null;
      if (!cfg || !cfg.mapping || typeof cfg.mapping !== 'object') return;
      if (!chart.$percentMappingOriginalTime) {
        const orig = {};
        Object.keys(cfg.mapping).forEach(function (k) {
          const arr = cfg.mapping[k];
          orig[k] = Array.isArray(arr) ? arr.slice() : arr;
        });
        chart.$percentMappingOriginalTime = orig;
      }
      const restored = {};
      Object.keys(chart.$percentMappingOriginalTime).forEach(function (k) {
        const arr = chart.$percentMappingOriginalTime[k];
        restored[k] = Array.isArray(arr) ? arr.slice() : arr;
      });
      const masked = {};
      Object.keys(restored).forEach(function (key) {
        const dsIndex = Number(key);
        const meta = chart.getDatasetMeta(dsIndex);
        const isVisible = (typeof chart.isDatasetVisible === 'function')
          ? chart.isDatasetVisible(dsIndex)
          : !(meta && meta.hidden === true);
        masked[dsIndex] = isVisible ? restored[dsIndex] : [];
      });
      chart.options.plugins.perBarPercentLabelsTime.mapping = masked;
    } catch (e) { console.warn('PercentVisibilityMaskPluginTime error:', e); }
  }
};
const PercentVisibilityMaskPluginServ = {
  id: 'percentVisibilityMaskServ',
  beforeDatasetsDraw: function (chart) {
    try {
      if (!chart.canvas || chart.canvas.id !== 'chartServicio') return;
      const cfg = chart.options && chart.options.plugins ? chart.options.plugins.perBarPercentLabelsServ : null;
      if (!cfg || !cfg.mapping || typeof cfg.mapping !== 'object') return;
      if (!chart.$percentMappingOriginalServ) {
        const orig = {};
        Object.keys(cfg.mapping).forEach(function (k) {
          const arr = cfg.mapping[k];
          orig[k] = Array.isArray(arr) ? arr.slice() : arr;
        });
        chart.$percentMappingOriginalServ = orig;
      }
      const restored = {};
      Object.keys(chart.$percentMappingOriginalServ).forEach(function (k) {
        const arr = chart.$percentMappingOriginalServ[k];
        restored[k] = Array.isArray(arr) ? arr.slice() : arr;
      });
      const masked = {};
      Object.keys(restored).forEach(function (key) {
        const dsIndex = Number(key);
        const meta = chart.getDatasetMeta(dsIndex);
        const isVisible = (typeof chart.isDatasetVisible === 'function')
          ? chart.isDatasetVisible(dsIndex)
          : !(meta && meta.hidden === true);
        masked[dsIndex] = isVisible ? restored[dsIndex] : [];
      });
      chart.options.plugins.perBarPercentLabelsServ.mapping = masked;
    } catch (e) { console.warn('PercentVisibilityMaskPluginServ error:', e); }
  }
};
Chart.register(PercentVisibilityMaskPlugin);
Chart.register(PercentVisibilityMaskPluginCat);
Chart.register(PercentVisibilityMaskPluginProv);
Chart.register(PercentVisibilityMaskPluginTime);
Chart.register(PercentVisibilityMaskPlugin);


/* ===================== OVERRIDE: SISTEMA DE REPORTES MEJORADO ===================== */
/*  Este bloque se apoya en las utilidades ya presentes en app.js:
    - rawData, schema, inferSchema, countBy, average, parseNumber, cleanText
    - getOpenClosedReturnedEmptyTotals, getOpenClosedByResponsible, getOpenClosedByProvider, getOpenClosedByCategory
    - getLogoDataUrl, getCanvasForPdfHD, applyPrintTheme
    No elimines el código original: este bloque solo sobreescribe los exports PDF/Excel.
*/

/* ==================== PDF EJECUTIVO MODERNO (Portrait A4) ==================== */
async function generateModernPDFReport() {
  if (!rawData.length) { alert('No hay datos cargados. Por favor carga un archivo CSV primero.'); return; }
  // Mostrar indicador de carga
  showLoadingIndicator('Generando PDF profesional...');

  try {
    const jsPDFlib = (window.jspdf && window.jspdf.jsPDF) ? window.jspdf.jsPDF : window.jsPDF;
    if (!jsPDFlib) { throw new Error('jsPDF no está disponible'); }

    const doc = new jsPDFlib('portrait', 'mm', 'a4');
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();

    // Paleta moderna
    const colors = {
      primary: [76, 201, 240],    // Cyan brillante
      secondary: [67, 97, 238],   // Azul profundo
      dark: [15, 23, 42],         // Azul oscuro
      accent: [231, 111, 81],     // Naranja
      success: [42, 157, 143],    // Verde agua
      warning: [255, 190, 11],    // Amarillo
      text: [30, 30, 30],         // Negro suave
      lightGray: [245, 247, 250]  // Gris claro
    };

    /* ============ PORTADA MODERNA ============ */
    drawModernCover(doc, colors, pageWidth, pageHeight);

    /* ============ ÍNDICE ============ */
    doc.addPage();
    drawTableOfContents(doc, colors, pageWidth);

    /* ============ RESUMEN EJECUTIVO ============ */
    doc.addPage();
    await drawExecutiveSummary(doc, colors, pageWidth);

    /* ============ MÉTRICAS CLAVE ============ */
    doc.addPage();
    await drawKeyMetrics(doc, colors, pageWidth);

    /* ============ ANÁLISIS POR ESTADO ============ */
    doc.addPage();
    await drawStateAnalysis(doc, colors, pageWidth, pageHeight);

    /* ============ ANÁLISIS POR RESPONSABLE ============ */
    doc.addPage();
    await drawResponsibleAnalysis(doc, colors, pageWidth, pageHeight);

    /* ============ ANÁLISIS POR SERVICIO ============ */
    doc.addPage();
    await drawServiceAnalysis(doc, colors, pageWidth, pageHeight);

    /* ============ TIEMPO Y RENDIMIENTO ============ */
    doc.addPage();
    await drawTimeAnalysis(doc, colors, pageWidth, pageHeight);

    /* ============ RECOMENDACIONES ============ */
    doc.addPage();
    drawRecommendations(doc, colors, pageWidth);

    /* ============ PIE DE PÁGINA EN TODAS LAS PÁGINAS ============ */
    addModernFooters(doc, colors, pageWidth, pageHeight);

    // Guardar
    const fileName = `Dashboard_Ejecutivo_${new Date().toISOString().slice(0, 10)}.pdf`;
    doc.save(fileName);

    hideLoadingIndicator();
    showSuccessNotification('PDF generado exitosamente');
  } catch (err) {
    console.error('Error generando PDF:', err);
    hideLoadingIndicator();
    showErrorNotification('Error al generar el PDF: ' + err.message);
  }
}

/* ==================== FUNCIONES DE DISEÑO PDF ==================== */
function drawModernCover(doc, colors, pageWidth, pageHeight) {
  // Fondo con gradiente simulado
  doc.setFillColor(...colors.dark);
  doc.rect(0, 0, pageWidth, pageHeight, 'F');

  // Banda superior
  doc.setFillColor(...colors.primary);
  doc.rect(0, 0, pageWidth, 60, 'F');

  // Banda diagonal
  doc.setFillColor(...colors.secondary);
  doc.triangle(0, 60, pageWidth, 60, pageWidth, 90, 'F');

  // Logo (si existe)
  try {
    const logo = getLogoDataUrl();
    if (logo) doc.addImage(logo, 'PNG', 20, 10, 60, 40);
  } catch (e) { console.warn('No se pudo cargar el logo'); }

  // Títulos
  doc.setTextColor(255, 255, 255);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(32);
  doc.text('Dashboard Ejecutivo', pageWidth / 2, 120, { align: 'center' });
  doc.setFontSize(24);
  doc.text('Backlog Incidentes TI', pageWidth / 2, 135, { align: 'center' });

  doc.setFont('helvetica', 'normal'); doc.setFontSize(14);
  doc.setTextColor(200, 200, 200);
  doc.text('Análisis y Métricas de Gestión', pageWidth / 2, 150, { align: 'center' });

  // Info reporte
  doc.setFontSize(11); doc.setTextColor(180, 180, 180);
  const fecha = new Date().toLocaleDateString('es-CO', { year: 'numeric', month: 'long', day: 'numeric' });
  doc.text(`Fecha de generación: ${fecha}`, pageWidth / 2, 180, { align: 'center' });
  doc.text(`Total de registros: ${rawData.length}`, pageWidth / 2, 190, { align: 'center' });

  // Caja info
  doc.setFillColor(...colors.secondary);
  doc.roundedRect(40, 210, pageWidth - 80, 40, 3, 3, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(10);
  doc.text('CONFIDENCIAL - USO INTERNO', pageWidth / 2, 225, { align: 'center' });
  doc.setFont('helvetica', 'normal'); doc.setFontSize(9);
  doc.text('Este documento contiene información sensible de la organización', pageWidth / 2, 235, { align: 'center' });

  // Pie portada
  doc.setFontSize(8); doc.setTextColor(150, 150, 150);
  doc.text('John Jairo Vargas González - Ingeniero de Soluciones TI', pageWidth / 2, pageHeight - 20, { align: 'center' });
  doc.text('john.vargas@bancounion.com', pageWidth / 2, pageHeight - 15, { align: 'center' });
}

function drawTableOfContents(doc, colors, pageWidth) {
  let y = 30;
  // Título
  doc.setFillColor(...colors.secondary);
  doc.rect(0, y - 10, pageWidth, 15, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(18);
  doc.text('Índice de Contenidos', 20, y); y += 25;

  const sections = [
    { title: '1. Resumen Ejecutivo', page: 3 },
    { title: '2. Métricas Clave', page: 4 },
    { title: '3. Análisis por Estado', page: 5 },
    { title: '4. Análisis por Responsable', page: 6 },
    { title: '5. Análisis por Servicio', page: 7 },
    { title: '6. Tiempo y Rendimiento', page: 8 },
    { title: '7. Recomendaciones', page: 9 }
  ];

  doc.setTextColor(...colors.text);
  doc.setFont('helvetica', 'normal'); doc.setFontSize(12);
  sections.forEach(section => {
    doc.setFillColor(...colors.primary);
    doc.circle(25, y - 2, 2, 'F');
    doc.text(section.title, 35, y);

    doc.setDrawColor(200, 200, 200);
    doc.setLineDash([1, 2]);
    doc.line(120, y, pageWidth - 40, y);
    doc.setLineDash([]);

    doc.setFont('helvetica', 'bold');
    doc.text(String(section.page), pageWidth - 30, y);
    doc.setFont('helvetica', 'normal');

    y += 12;
  });
}

async function drawExecutiveSummary(doc, colors, pageWidth) {
  let y = 30;
  drawSectionHeader(doc, colors, 'Resumen Ejecutivo', pageWidth, y); y += 20;

  const headers = Object.keys(rawData[0] || {});
  schema = inferSchema(headers, rawData);

  const colEstado = headers.includes('Estado Final Incidente') ? 'Estado Final Incidente' : schema.roles?.estado;
  const totals = getOpenClosedReturnedEmptyTotals(rawData, colEstado);
  const tasaResolucion = ((totals.cerrados / totals.total) * 100).toFixed(1);

  const colTiempo = schema.roles?.tiempo;
  let promedioTiempo = 0;
  if (colTiempo) {
    const tiempos = rawData.map(r => parseNumber(r[colTiempo])).filter(n => !isNaN(n));
    promedioTiempo = tiempos.length ? average(tiempos) : 0;
  }

  const metrics = [
    { label: 'Total Incidentes', value: totals.total, color: colors.primary, icon: '📊' },
    { label: 'Tasa Resolución', value: `${tasaResolucion}%`, color: colors.success, icon: '✓' },
    { label: 'Tiempo Promedio', value: `${Math.round(promedioTiempo)} días`, color: colors.warning, icon: '⏱' },
    { label: 'Backlog Actual', value: totals.abiertos, color: colors.accent, icon: '⚠' }
  ];

  const cardWidth = (pageWidth - 60) / 2;
  const cardHeight = 30;
  let x = 20;

  metrics.forEach((metric, i) => {
    if (i % 2 === 0 && i > 0) { y += cardHeight + 10; x = 20; }

    doc.setFillColor(...metric.color);
    doc.roundedRect(x, y, cardWidth, cardHeight, 3, 3, 'F');

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(10); doc.setFont('helvetica', 'normal');
    doc.text(metric.icon + ' ' + metric.label, x + 5, y + 10);

    doc.setFontSize(20); doc.setFont('helvetica', 'bold');
    doc.text(String(metric.value), x + 5, y + 24);

    x += cardWidth + 10;
  });

  y += cardHeight + 20;

  doc.setFillColor(...colors.lightGray);
  doc.roundedRect(20, y, pageWidth - 40, 80, 3, 3, 'F');

  doc.setTextColor(...colors.text);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(12);
  doc.text('Análisis de Situación', 25, y + 10);

  doc.setFont('helvetica', 'normal'); doc.setFontSize(10);
  const narrative = buildExecutiveNarrative(rawData, schema); // Camino A (sin LLM)
  let textY = y + 20;
  narrative.resumen.slice(0, 5).forEach(line => {
    const wrapped = doc.splitTextToSize(cleanText(line), pageWidth - 50);
    doc.text(wrapped, 25, textY);
    textY += wrapped.length * 5;
  });
}

async function drawKeyMetrics(doc, colors, pageWidth) {
  let y = 30;
  drawSectionHeader(doc, colors, 'Métricas Clave de Rendimiento', pageWidth, y); y += 20;

  const headers = Object.keys(rawData[0] || {});
  schema = inferSchema(headers, rawData);
  const colEstado = headers.includes('Estado Final Incidente') ? 'Estado Final Incidente' : schema.roles?.estado;
  const totals = getOpenClosedReturnedEmptyTotals(rawData, colEstado);

  const metricsData = [
    ['Métrica', 'Valor', 'Interpretación', 'Estado'],
    ['Incidentes Abiertos', String(totals.abiertos), `${((totals.abiertos / totals.total) * 100).toFixed(1)}% del total`, totals.abiertos > totals.cerrados ? '⚠ Alto' : '✓ Normal'],
    ['Incidentes Cerrados', String(totals.cerrados), `${((totals.cerrados / totals.total) * 100).toFixed(1)}% del total`, totals.cerrados >= totals.abiertos ? '✓ Bueno' : '⚠ Mejorar'],
    ['Casos Devueltos', String(totals.devuelto), `${((totals.devuelto / totals.total) * 100).toFixed(1)}% del total`, totals.devuelto > 0 ? '⚠ Revisar' : '✓ OK'],
    ['Registros Vacíos', String(totals.vacios), 'Requiere validación de datos', totals.vacios > 0 ? '⚠ Acción requerida' : '✓ OK']
  ];

  doc.autoTable({
    startY: y,
    head: [metricsData[0]],
    body: metricsData.slice(1),
    theme: 'grid',
    headStyles: { fillColor: colors.secondary, textColor: [255, 255, 255], fontStyle: 'bold', halign: 'center', fontSize: 10 },
    styles: { fontSize: 9, cellPadding: 5 },
    columnStyles: { 0: { cellWidth: 45, fontStyle: 'bold' }, 1: { cellWidth: 30, halign: 'center' }, 2: { cellWidth: 60 }, 3: { cellWidth: 35, halign: 'center', fontStyle: 'bold' } },
    alternateRowStyles: { fillColor: colors.lightGray }
  });

  y = doc.lastAutoTable.finalY + 15;
  await embedChartInPDF(doc, 'chartEstado', 20, y, pageWidth - 40, 80);
}

async function drawStateAnalysis(doc, colors, pageWidth, pageHeight) {
  let y = 30;
  drawSectionHeader(doc, colors, 'Análisis por Estado', pageWidth, y); y += 20;

  const headers = Object.keys(rawData[0] || {});
  const colEstado = headers.includes('Estado Final Incidente') ? 'Estado Final Incidente' : schema?.roles?.estado;

  if (colEstado) {
    const totals = getOpenClosedReturnedEmptyTotals(rawData, colEstado);
    const stateData = [
      ['Estado', 'Cantidad', '% del Total'],
      ['Abiertos', String(totals.abiertos), `${((totals.abiertos / totals.total) * 100).toFixed(1)}%`],
      ['Cerrados', String(totals.cerrados), `${((totals.cerrados / totals.total) * 100).toFixed(1)}%`],
      ['Devueltos', String(totals.devuelto), `${((totals.devuelto / totals.total) * 100).toFixed(1)}%`],
      ['Vacíos', String(totals.vacios), `${((totals.vacios / totals.total) * 100).toFixed(1)}%`],
      ['TOTAL', String(totals.total), '100%']
    ];

    doc.autoTable({
      startY: y,
      head: [stateData[0]],
      body: stateData.slice(1, -1),
      foot: [stateData[stateData.length - 1]],
      theme: 'grid',
      headStyles: { fillColor: colors.secondary, textColor: [255, 255, 255], fontStyle: 'bold', halign: 'center' },
      footStyles: { fillColor: colors.primary, textColor: [255, 255, 255], fontStyle: 'bold', halign: 'center' },
      columnStyles: { 0: { cellWidth: 60, fontStyle: 'bold' }, 1: { cellWidth: 50, halign: 'center' }, 2: { cellWidth: 50, halign: 'center' } }
    });

    y = doc.lastAutoTable.finalY + 15;
    await embedChartInPDF(doc, 'chartEstado', 20, y, pageWidth - 40, 100);
  }
}

async function drawResponsibleAnalysis(doc, colors, pageWidth, pageHeight) {
  let y = 30;
  drawSectionHeader(doc, colors, 'Análisis por Responsable', pageWidth, y); y += 20;

  const headers = Object.keys(rawData[0] || {});
  const colEstado = headers.includes('Estado Final Incidente') ? 'Estado Final Incidente' : schema?.roles?.estado;
  const colResp = headers.includes('Ingeniero Asignado') ? 'Ingeniero Asignado' : schema?.roles?.responsable;

  if (colResp && colEstado) {
    const rowsResp = getOpenClosedByResponsible(rawData, colResp, colEstado);
    const topResp = rowsResp.slice(0, 10);
    const tableData = [['Responsable', 'Abiertos', '% Fila', 'Cerrados', '% Fila', 'Total']];

    topResp.forEach(r => {
      const total = r.abiertos + r.cerrados;
      tableData.push([
        cleanText(r.label),
        String(r.abiertos),
        `${((r.abiertos / total) * 100).toFixed(1)}%`,
        String(r.cerrados),
        `${((r.cerrados / total) * 100).toFixed(1)}%`,
        String(total)
      ]);
    });

    doc.autoTable({
      startY: y, head: [tableData[0]], body: tableData.slice(1), theme: 'striped',
      headStyles: { fillColor: colors.secondary, textColor: [255, 255, 255], fontStyle: 'bold', halign: 'center', fontSize: 9 },
      styles: { fontSize: 8, cellPadding: 3 },
      columnStyles: { 0: { cellWidth: 70 }, 1: { cellWidth: 20, halign: 'center' }, 2: { cellWidth: 20, halign: 'center' }, 3: { cellWidth: 20, halign: 'center' }, 4: { cellWidth: 20, halign: 'center' }, 5: { cellWidth: 20, halign: 'center', fontStyle: 'bold' } }
    });

    y = doc.lastAutoTable.finalY + 10;
    await embedChartInPDF(doc, 'chartResponsables', 20, y, pageWidth - 40, 80);
  }
}

async function drawServiceAnalysis(doc, colors, pageWidth, pageHeight) {
  let y = 30;
  drawSectionHeader(doc, colors, 'Análisis por Servicio', pageWidth, y); y += 20;

  const colServ = schema?.roles?.servicio;
  if (colServ) {
    const counts = countBy(rawData, colServ);
    const total = counts.reduce((sum, c) => sum + c.value, 0);
    const topServices = counts.slice(0, 15);

    const tableData = [['Servicio', 'Cantidad', '% del Total']];
    topServices.forEach(s => {
      tableData.push([cleanText(s.label), String(Math.round(s.value)), `${((s.value / total) * 100).toFixed(1)}%`]);
    });

    doc.autoTable({
      startY: y,
      head: [tableData[0]],
      body: tableData.slice(1),
      theme: 'grid',
      headStyles: { fillColor: colors.secondary, textColor: [255, 255, 255], fontStyle: 'bold', halign: 'center' },
      columnStyles: { 0: { cellWidth: 100 }, 1: { cellWidth: 40, halign: 'center' }, 2: { cellWidth: 40, halign: 'center' } }
    });

    y = doc.lastAutoTable.finalY + 10;
    await embedChartInPDF(doc, 'chartServicio', 20, y, pageWidth - 40, 80);
  }
}

async function drawTimeAnalysis(doc, colors, pageWidth, pageHeight) {
  let y = 30;
  drawSectionHeader(doc, colors, 'Análisis de Tiempo y Rendimiento', pageWidth, y); y += 20;

  await embedChartInPDF(doc, 'chartTiempo', 20, y, pageWidth - 40, 100);
  y += 110;

  const colTiempo = schema?.roles?.tiempo;
  if (colTiempo) {
    const tiempos = rawData.map(r => parseNumber(r[colTiempo])).filter(n => !isNaN(n));
    if (tiempos.length) {
      const promedio = average(tiempos);
      const max = Math.max(...tiempos);
      const min = Math.min(...tiempos);

      const timeMetrics = [
        ['Métrica', 'Valor (días)'],
        ['Tiempo Promedio', String(Math.round(promedio))],
        ['Tiempo Máximo', String(Math.round(max))],
        ['Tiempo Mínimo', String(Math.round(min))]
      ];

      doc.autoTable({
        startY: y,
        head: [timeMetrics[0]],
        body: timeMetrics.slice(1),
        theme: 'grid',
        headStyles: { fillColor: colors.secondary, textColor: [255, 255, 255], fontStyle: 'bold' },
        columnStyles: { 0: { cellWidth: 80, fontStyle: 'bold' }, 1: { cellWidth: 60, halign: 'center', fontSize: 12 } }
      });
    }
  }
}

function drawRecommendations(doc, colors, pageWidth) {
  let y = 30;
  drawSectionHeader(doc, colors, 'Recomendaciones Estratégicas', pageWidth, y); y += 20;

  const narrative = buildExecutiveNarrative(rawData, schema);

  doc.setFont('helvetica', 'bold'); doc.setFontSize(12); doc.setTextColor(...colors.accent);
  doc.text('⚠ Riesgos Identificados', 20, y); y += 10;

  doc.setFont('helvetica', 'normal'); doc.setFontSize(10); doc.setTextColor(...colors.text);
  narrative.riesgos.forEach(riesgo => {
    const wrapped = doc.splitTextToSize('• ' + cleanText(riesgo), pageWidth - 45);
    doc.text(wrapped, 25, y);
    y += wrapped.length * 6 + 3;
  });

  y += 10;
  doc.setFont('helvetica', 'bold'); doc.setFontSize(12); doc.setTextColor(...colors.success);
  doc.text('✓ Acciones Recomendadas', 20, y); y += 10;

  doc.setFont('helvetica', 'normal'); doc.setFontSize(10); doc.setTextColor(...colors.text);
  narrative.acciones.forEach(accion => {
    const wrapped = doc.splitTextToSize('• ' + cleanText(accion), pageWidth - 45);
    doc.text(wrapped, 25, y);
    y += wrapped.length * 6 + 3;
  });
}

function drawSectionHeader(doc, colors, title, pageWidth, y) {
  doc.setFillColor(...colors.primary);
  doc.rect(0, y - 5, pageWidth, 12, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(14);
  doc.text(title, 20, y + 4);
}

async function embedChartInPDF(doc, chartId, x, y, width, height) {
  try {
    applyPrintTheme(true);
    await new Promise(resolve => setTimeout(resolve, 100));
    const canvas = getCanvasForPdfHD(chartId, width * 3, height * 3, 3);
    if (canvas) doc.addImage(canvas.toDataURL('image/png', 1.0), 'PNG', x, y, width, height);
    applyPrintTheme(false);
  } catch (err) {
    console.error('Error embebiendo gráfico:', err);
  }
}

function addModernFooters(doc, colors, pageWidth, pageHeight) {
  const totalPages = doc.internal.getNumberOfPages();
  for (let i = 1; i <= totalPages; i++) {
    doc.setPage(i);
    doc.setDrawColor(...colors.primary); doc.setLineWidth(0.5);
    doc.line(20, pageHeight - 20, pageWidth - 20, pageHeight - 20);

    doc.setFontSize(8); doc.setTextColor(100, 100, 100); doc.setFont('helvetica', 'normal');
    const leftText = 'Dashboard Incidentes TI - Banco Unión';
    const centerText = `Página ${i} de ${totalPages}`;
    const rightText = new Date().toLocaleDateString('es-CO');
    doc.text(leftText, 20, pageHeight - 10);
    doc.text(centerText, pageWidth / 2, pageHeight - 10, { align: 'center' });
    doc.text(rightText, pageWidth - 20, pageHeight - 10, { align: 'right' });
  }
}

/* ==================== EXCEL PROFESIONAL MEJORADO ==================== */
function generateModernExcelReport() {
  if (!rawData.length) { alert('No hay datos cargados. Por favor carga un archivo CSV primero.'); return; }

  showLoadingIndicator('Generando Excel profesional...');
  try {
    const wb = XLSX.utils.book_new();

    const headers = Object.keys(rawData[0] || {});
    schema = inferSchema(headers, rawData);
    const colEstado = headers.includes('Estado Final Incidente') ? 'Estado Final Incidente' : schema?.roles?.estado;
    const totals = getOpenClosedReturnedEmptyTotals(rawData, colEstado);

    /* ============ HOJA 1: PORTADA Y RESUMEN ============ */
    const coverData = buildExcelCoverSheet(totals);
    const wsCover = XLSX.utils.aoa_to_sheet(coverData);
    wsCover['!cols'] = [{ wch: 35 }, { wch: 20 }, { wch: 40 }];
    wsCover['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 2 } },
      { s: { r: 4, c: 0 }, e: { r: 4, c: 2 } }
    ];
    XLSX.utils.book_append_sheet(wb, wsCover, '📊 Resumen Ejecutivo');

    /* ============ HOJA 2: DATOS COMPLETOS ============ */
    const wsData = XLSX.utils.json_to_sheet(rawData);
    const colsKeys = Object.keys(rawData[0] || {});
    wsData['!cols'] = colsKeys.map(k => ({ wch: Math.min(35, Math.max(12, k.length + 4)) }));
    wsData['!freeze'] = { xSplit: 0, ySplit: 1, activePane: 'bottomLeft' };
    colsKeys.forEach((key, idx) => {
      const cellRef = XLSX.utils.encode_cell({ r: 0, c: idx });
      if (!wsData[cellRef]) return;
      wsData[cellRef].s = {
        fill: { fgColor: { rgb: "4361EE" } },
        font: { bold: true, color: { rgb: "FFFFFF" }, sz: 12 },
        alignment: { horizontal: "center", vertical: "center" }
      };
    });
    XLSX.utils.book_append_sheet(wb, wsData, '📋 Datos Completos');

    /* ============ HOJA 3..10 ============ */
    const wsEstado = buildStateAnalysisSheet(colEstado, totals);
    XLSX.utils.book_append_sheet(wb, wsEstado, '🔄 Análisis Estado');

    const wsResp = buildResponsibleAnalysisSheet(colEstado);
    XLSX.utils.book_append_sheet(wb, wsResp, '👥 Por Responsable');

    const wsServ = buildServiceAnalysisSheet();
    XLSX.utils.book_append_sheet(wb, wsServ, '🛠️ Por Servicio');

    const wsTime = buildTimeAnalysisSheet(colEstado);
    XLSX.utils.book_append_sheet(wb, wsTime, '⏱️ Tiempo');

    const wsProv = buildProviderAnalysisSheet(colEstado);
    XLSX.utils.book_append_sheet(wb, wsProv, '🏢 Proveedores');

    const wsCat = buildCategoryAnalysisSheet(colEstado);
    XLSX.utils.book_append_sheet(wb, wsCat, '📑 Categorías');

    const wsInsights = buildInsightsSheet();
    XLSX.utils.book_append_sheet(wb, wsInsights, '💡 Hallazgos');

    const wsDashboard = buildMetricsDashboardSheet(totals);
    XLSX.utils.book_append_sheet(wb, wsDashboard, '📈 Dashboard KPIs');

    const fileName = `Dashboard_Completo_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);

    hideLoadingIndicator();
    showSuccessNotification('Excel generado exitosamente con 10 hojas de análisis');
  } catch (err) {
    console.error('Error generando Excel:', err);
    hideLoadingIndicator();
    showErrorNotification('Error al generar Excel: ' + err.message);
  }
}

/* ============ BUILDERS HOJAS EXCEL ============ */
function buildExcelCoverSheet(totals) {
  const fecha = new Date().toLocaleDateString('es-CO', { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' });
  const tasaResolucion = ((totals.cerrados / totals.total) * 100).toFixed(1);
  return [
    ['DASHBOARD EJECUTIVO - BACKLOG INCIDENTES TI'],
    ['Banco Unión S.A - Gestión de Incidentes Nivel 2'],
    [''],
    ['Fecha de Generación:', fecha],
    ['Archivo Fuente:', cleanText(window.__lastCsvName || 'Datos cargados')],
    ['Total de Registros:', totals.total],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['MÉTRICAS PRINCIPALES', '', ''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Métrica', 'Valor', 'Interpretación'],
    ['Total de Incidentes', totals.total, totals.total > 0 ? '✓ Carga activa' : '⚠ Sin carga'],
    ['Incidentes Abiertos', totals.abiertos, totals.abiertos > totals.cerrados ? '⚠ Requiere atención urgente' : '✓ Bajo control'],
    ['Incidentes Cerrados', totals.cerrados, totals.cerrados >= totals.abiertos ? '✓ Buena resolución' : '⚠ Mejorar cierre'],
    ['Tasa de Resolución', `${tasaResolucion}%`, Number(tasaResolucion) >= 70 ? '✓ Excelente' : Number(tasaResolucion) >= 50 ? '⚠ Aceptable' : '❌ Crítico'],
    ['Casos Devueltos', totals.devuelto, totals.devuelto > 10 ? '⚠ Revisar proceso' : totals.devuelto > 0 ? '⚠ Monitorear' : '✓ OK'],
    ['Registros Vacíos', totals.vacios, totals.vacios > 0 ? '⚠ Validar calidad de datos' : '✓ Datos completos'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['INDICADORES DE SALUD DEL BACKLOG', '', ''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Indicador', 'Estado', 'Observación'],
    ['Ratio Abiertos/Cerrados', totals.abiertos <= totals.cerrados ? '✓ SALUDABLE' : '⚠ CRÍTICO',
      `${(totals.abiertos / Math.max(1, totals.cerrados)).toFixed(2)} - ${totals.abiertos <= totals.cerrados ? 'Mantener ritmo' : 'Aumentar cierre'}`],
    ['Calidad de Datos', totals.vacios === 0 ? '✓ EXCELENTE' : totals.vacios < 5 ? '⚠ BUENO' : '❌ DEFICIENTE',
      `${totals.vacios} registros vacíos de ${totals.total}`],
    ['Devoluciones', totals.devuelto < 5 ? '✓ BAJO' : totals.devuelto < 20 ? '⚠ MEDIO' : '❌ ALTO',
      `${((totals.devuelto / totals.total) * 100).toFixed(1)}% del total`],
    [''],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['Elaborado por:', 'John Jairo Vargas González', ''],
    ['Cargo:', 'Ingeniero de Soluciones TI', ''],
    ['Contacto:', 'john.vargas@bancounion.com', ''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━']
  ];
}

function buildStateAnalysisSheet(colEstado, totals) {
  const stateData = [
    ['ANÁLISIS DETALLADO POR ESTADO'],
    [''],
    ['Estado', 'Cantidad', '% del Total', 'Interpretación', 'Acción Recomendada'],
    ['Abiertos', totals.abiertos, `${((totals.abiertos / totals.total) * 100).toFixed(2)}%`,
      totals.abiertos > totals.cerrados ? 'Backlog creciente' : 'Bajo control',
      totals.abiertos > totals.cerrados ? 'Priorizar cierre inmediato' : 'Mantener seguimiento'],
    ['Cerrados', totals.cerrados, `${((totals.cerrados / totals.total) * 100).toFixed(2)}%`, 'Incidentes resueltos', 'Analizar tiempos de resolución'],
    ['Devueltos', totals.devuelto, `${((totals.devuelto / totals.total) * 100).toFixed(2)}%`,
      totals.devuelto > 10 ? 'Alto nivel de devoluciones' : 'Normal',
      totals.devuelto > 10 ? 'Revisar causas raíz' : 'Monitoreo regular'],
    ['Vacíos', totals.vacios, `${((totals.vacios / totals.total) * 100).toFixed(2)}%`, 'Registros sin estado',
      totals.vacios > 0 ? 'Validar y corregir datos' : 'N/A'],
    ['TOTAL', totals.total, '100%', '', ''],
    [''],
    ['TENDENCIAS Y OBSERVACIONES'],
    [''],
    ['Ratio Abiertos/Cerrados:', (totals.abiertos / Math.max(1, totals.cerrados)).toFixed(2)],
    ['Eficiencia de Cierre:', `${((totals.cerrados / totals.total) * 100).toFixed(1)}%`],
    ['Casos Problemáticos (Devueltos + Vacíos):', totals.devuelto + totals.vacios],
    [''],
    ['ALERTAS AUTOMÁTICAS'],
    [''],
    totals.abiertos > totals.cerrados ? ['⚠ ALERTA: Backlog en crecimiento - Abiertos superan cerrados'] : ['✓ Estado normal - Ritmo de cierre adecuado'],
    totals.devuelto > 15 ? ['⚠ ALERTA: Alto nivel de devoluciones - Revisar proceso'] : ['✓ Devoluciones dentro de parámetros normales'],
    totals.vacios > 10 ? ['⚠ ALERTA: Calidad de datos comprometida'] : ['✓ Calidad de datos aceptable']
  ];
  const ws = XLSX.utils.aoa_to_sheet(stateData);
  ws['!cols'] = [{ wch: 20 }, { wch: 15 }, { wch: 15 }, { wch: 30 }, { wch: 35 }];
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
    { s: { r: 9, c: 0 }, e: { r: 9, c: 4 } },
    { s: { r: 14, c: 0 }, e: { r: 14, c: 4 } }
  ];
  return ws;
}

function buildResponsibleAnalysisSheet(colEstado) {
  const headers = Object.keys(rawData[0] || {});
  const colResp = headers.includes('Ingeniero Asignado') ? 'Ingeniero Asignado' : schema?.roles?.responsable;
  if (!colResp || !colEstado) return XLSX.utils.aoa_to_sheet([['No hay datos de responsables disponibles']]);

  const rowsResp = getOpenClosedByResponsible(rawData, colResp, colEstado);
  const data = [['ANÁLISIS POR RESPONSABLE - TOP 20'], [''], ['Responsable', 'Abiertos', '% Abiertos', 'Cerrados', '% Cerrados', 'Total', 'Eficiencia', 'Estado', 'Acción']];

  rowsResp.slice(0, 20).forEach(r => {
    const total = r.abiertos + r.cerrados;
    const eficiencia = ((r.cerrados / total) * 100).toFixed(1);
    const estado = Number(eficiencia) >= 70 ? '✓ Bueno' : Number(eficiencia) >= 50 ? '⚠ Regular' : '❌ Crítico';
    const accion = Number(eficiencia) < 50 ? 'Requiere soporte urgente' : Number(eficiencia) < 70 ? 'Monitorear de cerca' : 'Mantener ritmo';
    data.push([cleanText(r.label), r.abiertos, `${((r.abiertos / total) * 100).toFixed(1)}%`, r.cerrados, `${((r.cerrados / total) * 100).toFixed(1)}%`, total, `${eficiencia}%`, estado, accion]);
  });

  const totalAbiertos = rowsResp.reduce((sum, r) => sum + r.abiertos, 0);
  const totalCerrados = rowsResp.reduce((sum, r) => sum + r.cerrados, 0);
  const totalGeneral = totalAbiertos + totalCerrados;

  data.push(['']);
  data.push(['TOTALES', totalAbiertos, '', totalCerrados, '', totalGeneral, '', '', '']);
  data.push(['']);
  data.push(['ESTADÍSTICAS']);
  data.push(['Responsables analizados:', rowsResp.length]);
  data.push(['Promedio casos por responsable:', Math.round(totalGeneral / rowsResp.length)]);
  data.push(['Responsable con más abiertos:', rowsResp[0]?.label || 'N/A', rowsResp[0]?.abiertos || 0]);
  data.push(['Responsable con más cerrados:',
    rowsResp.slice().sort((a, b) => b.cerrados - a.cerrados)[0]?.label || 'N/A',
    rowsResp.slice().sort((a, b) => b.cerrados - a.cerrados)[0]?.cerrados || 0
  ]);

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 35 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 15 }, { wch: 30 }];
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 8 } }];
  return ws;
}

function buildServiceAnalysisSheet() {
  const colServ = schema?.roles?.servicio;
  if (!colServ) return XLSX.utils.aoa_to_sheet([['No hay datos de servicios disponibles']]);

  const counts = countBy(rawData, colServ);
  const total = counts.reduce((sum, c) => sum + c.value, 0);
  const data = [['ANÁLISIS POR SERVICIO / TIPIFICACIÓN'], [''], ['Servicio', 'Cantidad', '% del Total', 'Impacto', 'Prioridad']];

  counts.forEach((s) => {
    const porcentaje = (s.value / total) * 100;
    const impacto = porcentaje > 20 ? 'ALTO' : porcentaje > 10 ? 'MEDIO' : 'BAJO';
    const prioridad = porcentaje > 20 ? 'P1 - Crítico' : porcentaje > 10 ? 'P2 - Alto' : porcentaje > 5 ? 'P3 - Medio' : 'P4 - Bajo';
    data.push([cleanText(s.label), Math.round(s.value), `${porcentaje.toFixed(2)}%`, impacto, prioridad]);
  });

  data.push(['']);
  data.push(['TOTAL', total, '100%', '', '']);
  data.push(['']);
  data.push(['CONCENTRACIÓN DE CASOS']);
  data.push(['Top 5 servicios representan:', `${(((counts.slice(0, 5).reduce((sum, s) => sum + s.value, 0)) / total) * 100).toFixed(1)}%`]);
  data.push(['Top 10 servicios representan:', `${(((counts.slice(0, 10).reduce((sum, s) => sum + s.value, 0)) / total) * 100).toFixed(1)}%`]);
  data.push(['']);
  data.push(['RECOMENDACIONES']);
  counts.slice(0, 3).forEach((s, idx) => {
    data.push([`${idx + 1}. ${cleanText(s.label)}`, `Focalizar recursos - ${Math.round((s.value / total) * 100)}% del volumen`]);
  });

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 50 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 20 }];
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }];
  return ws;
}

function buildTimeAnalysisSheet(colEstado) {
  const colTiempo = schema?.roles?.tiempo;
  if (!colTiempo) return XLSX.utils.aoa_to_sheet([['No hay datos de tiempo disponibles']]);

  const tiempos = rawData.map(r => parseNumber(r[colTiempo])).filter(n => !isNaN(n));
  if (!tiempos.length) return XLSX.utils.aoa_to_sheet([['No hay datos válidos de tiempo']]);

  const promedio = average(tiempos);
  const max = Math.max(...tiempos);
  const min = Math.min(...tiempos);
  const mediana = tiempos.sort((a, b) => a - b)[Math.floor(tiempos.length / 2)];

  const rangos = [
    { label: '0-7 días', min: 0, max: 7, count: 0 },
    { label: '8-15 días', min: 8, max: 15, count: 0 },
    { label: '16-30 días', min: 16, max: 30, count: 0 },
    { label: '31-60 días', min: 31, max: 60, count: 0 },
    { label: '61-90 días', min: 61, max: 90, count: 0 },
    { label: 'Más de 90 días', min: 91, max: Infinity, count: 0 }
  ];
  tiempos.forEach(t => { const r = rangos.find(rg => t >= rg.min && t <= rg.max); if (r) r.count++; });

  const data = [
    ['ANÁLISIS DE TIEMPO Y RENDIMIENTO'],
    [''],
    ['ESTADÍSTICAS GENERALES'],
    ['Métrica', 'Valor (días)', 'Interpretación'],
    ['Tiempo Promedio', Math.round(promedio), promedio > 30 ? '⚠ Ciclo lento' : '✓ Aceptable'],
    ['Tiempo Máximo', Math.round(max), max > 90 ? '❌ Caso crítico' : max > 60 ? '⚠ Alto' : '✓ Normal'],
    ['Tiempo Mínimo', Math.round(min), min < 1 ? '⚠ Revisar' : '✓ Rápido'],
    ['Mediana', Math.round(mediana), ''],
    ['Casos Analizados', tiempos.length, ''],
    [''],
    ['DISTRIBUCIÓN POR RANGOS DE TIEMPO'],
    ['Rango', 'Cantidad', '% del Total', 'Estado'],
    ...rangos.map(r => [r.label, r.count, `${((r.count / tiempos.length) * 100).toFixed(1)}%`,
      r.label.includes('Más de 90') && r.count > 0 ? '❌ Crítico' :
      r.label.includes('61-90') && r.count > tiempos.length * 0.2 ? '⚠ Alto' : '✓']),
    ['TOTAL', tiempos.length, '100%', ''],
    [''],
    ['ANÁLISIS DE ALERTAS'],
    [''],
    promedio > 30 ? ['⚠ Tiempo promedio superior a 30 días - Revisar procesos'] : ['✓ Tiempo promedio dentro de parámetros'],
    rangos[5].count > 0 ? [`⚠ ${rangos[5].count} casos con más de 90 días - Atención urgente`] : ['✓ No hay casos críticos por tiempo'],
    max > 180 ? [`❌ Caso más antiguo: ${Math.round(max)} días - Escalamiento necesario`] : ['✓ Casos antiguos bajo control']
  ];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 25 }, { wch: 18 }, { wch: 18 }, { wch: 35 }];
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 3 } },
    { s: { r: 10, c: 0 }, e: { r: 10, c: 3 } },
    { s: { r: 15, c: 0 }, e: { r: 15, c: 3 } }
  ];
  return ws;
}

function buildProviderAnalysisSheet(colEstado) {
  const headers = Object.keys(rawData[0] || {});
  const colProv = headers.includes('Proveedor a escalar') ? 'Proveedor a escalar' : schema?.roles?.proveedor;
  if (!colProv || !colEstado) return XLSX.utils.aoa_to_sheet([['No hay datos de proveedores disponibles']]);

  const rowsProv = getOpenClosedByProvider(rawData, colProv, colEstado);
  const data = [['ANÁLISIS POR PROVEEDOR'], [''], ['Proveedor', 'Abiertos', '% Abiertos', 'Cerrados', '% Cerrados', 'Total', 'SLA Cumpl.', 'Criticidad']];

  rowsProv.forEach(p => {
    const total = p.abiertos + p.cerrados;
    const sla = ((p.cerrados / total) * 100).toFixed(1);
    const criticidad = Number(sla) < 50 ? '❌ Crítico' : Number(sla) < 70 ? '⚠ Medio' : '✓ Bueno';
    data.push([cleanText(p.label), p.abiertos, `${((p.abiertos / total) * 100).toFixed(1)}%`, p.cerrados, `${((p.cerrados / total) * 100).toFixed(1)}%`, total, `${sla}%`, criticidad]);
  });

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 35 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 18 }];
  return ws;
}

function buildCategoryAnalysisSheet(colEstado) {
  const headers = Object.keys(rawData[0] || {});
  const colCat = headers.includes('Categoría') ? 'Categoría' : headers.find(h => /categor[ií]a/i.test(h));
  if (!colCat || !colEstado) return XLSX.utils.aoa_to_sheet([['No hay datos de categorías disponibles']]);

  const rowsCat = getOpenClosedByCategory(rawData, colCat, colEstado);
  const data = [['ANÁLISIS POR CATEGORÍA'], [''], ['Categoría', 'Abiertos', '% Abiertos', 'Cerrados', '% Cerrados', 'Total', 'Ratio A/C']];

  rowsCat.forEach(c => {
    const ratio = c.cerrados > 0 ? (c.abiertos / c.cerrados).toFixed(2) : 'N/A';
    data.push([cleanText(c.label), c.abiertos, `${((c.abiertos / c.total) * 100).toFixed(1)}%`, c.cerrados, `${((c.cerrados / c.total) * 100).toFixed(1)}%`, c.total, ratio]);
  });

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 40 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }];
  return ws;
}

function buildInsightsSheet() {
  const narrative = buildExecutiveNarrative(rawData, schema);
  const data = [
    ['HALLAZGOS Y RECOMENDACIONES ESTRATÉGICAS'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['RESUMEN EJECUTIVO'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ...narrative.resumen.map(line => [cleanText(line)]),
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['⚠ RIESGOS IDENTIFICADOS'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ...narrative.riesgos.map(line => ['• ' + cleanText(line)]),
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['✓ ACCIONES RECOMENDADAS (PRÓXIMOS 30 DÍAS)'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ...narrative.acciones.map(line => ['• ' + cleanText(line)]),
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['TOP 5 RESPONSABLES CON MÁS CASOS ABIERTOS'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Responsable', 'Abiertos', 'Cerrados', 'Total'],
    ...narrative.topRespAbiertos.map(r => [cleanText(r.label), r.abiertos, r.cerrados, r.total]),
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['TOP 5 PROVEEDORES CRÍTICOS'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Proveedor', 'Abiertos', 'Cerrados', 'Total'],
    ...narrative.topProvAbiertos.map(r => [cleanText(r.label), r.abiertos, r.cerrados, r.total]),
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['TOP 5 SERVICIOS POR VOLUMEN'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Servicio', 'Casos', '% del Total'],
    ...(() => {
      const totalTop = narrative.topServicios.reduce((sum, srv) => sum + srv.value, 0);
      return narrative.topServicios.map(s => [cleanText(s.label), Math.round(s.value), `${(((s.value) / totalTop) * 100).toFixed(1)}%`]);
    })()
  ];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 80 }, { wch: 15 }, { wch: 15 }, { wch: 15 }];
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];
  return ws;
}

function buildMetricsDashboardSheet(totals) {
  const colTiempo = schema?.roles?.tiempo;
  let promedioTiempo = 0;
  if (colTiempo) {
    const tiempos = rawData.map(r => parseNumber(r[colTiempo])).filter(n => !isNaN(n));
    if (tiempos.length) promedioTiempo = average(tiempos);
  }

  const tasaResolucion = ((totals.cerrados / totals.total) * 100).toFixed(1);
  const ratio = (totals.abiertos / Math.max(1, totals.cerrados)).toFixed(2);

  const data = [
    ['📈 DASHBOARD DE INDICADORES CLAVE (KPIs)'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['KPI PRINCIPAL', 'VALOR ACTUAL', 'META', 'CUMPLIMIENTO', 'TENDENCIA'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Total Incidentes', totals.total, 'N/A', '—', totals.total > 500 ? '📈 Alto volumen' : '📊 Normal'],
    ['Tasa de Resolución', `${tasaResolucion}%`, '≥ 70%', Number(tasaResolucion) >= 70 ? '✅ CUMPLE' : '❌ NO CUMPLE', Number(tasaResolucion) >= 70 ? '📈 Excelente' : '📉 Mejorar'],
    ['Tiempo Promedio Resolución', `${Math.round(promedioTiempo)} días`, '≤ 30 días', promedioTiempo <= 30 ? '✅ CUMPLE' : '❌ NO CUMPLE', promedioTiempo <= 30 ? '✓ Dentro de SLA' : '⚠ Fuera de SLA'],
    ['Ratio Abiertos/Cerrados', ratio, '≤ 1.0', Number(ratio) <= 1.0 ? '✅ CUMPLE' : '❌ NO CUMPLE', Number(ratio) <= 1.0 ? '📈 Saludable' : '📉 Crítico'],
    ['Backlog Actual (Abiertos)', totals.abiertos, `≤ ${Math.round(totals.total * 0.3)}`, totals.abiertos <= totals.total * 0.3 ? '✅ CUMPLE' : '❌ NO CUMPLE', totals.abiertos <= totals.cerrados ? '📉 Decreciendo' : '📈 Creciendo'],
    ['Calidad de Datos', `${totals.total - totals.vacios} registros`, '100%', totals.vacios === 0 ? '✅ CUMPLE' : '⚠ REVISAR', totals.vacios === 0 ? '✓ Excelente' : '⚠ Validar'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['SEMÁFORO DE SALUD DEL BACKLOG'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['Indicador', 'Estado Visual', 'Descripción'],
    ...calcularSemaforo(totals, tasaResolucion, promedioTiempo, ratio),
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['SCORE GENERAL DEL BACKLOG'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    [''],
    ...calcularScore(totals, tasaResolucion, promedioTiempo, ratio)
  ];

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 35 }, { wch: 20 }, { wch: 15 }, { wch: 18 }, { wch: 25 }];
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }];
  return ws;
}

function calcularSemaforo(totals, tasaResolucion, promedioTiempo, ratio) {
  const items = [];
  // Tasa de resolución
  if (Number(tasaResolucion) >= 70) items.push(['Tasa de Resolución', '🟢 VERDE', `Excelente: ${tasaResolucion}%`]);
  else if (Number(tasaResolucion) >= 50) items.push(['Tasa de Resolución', '🟡 AMARILLO', `Regular: ${tasaResolucion}%`]);
  else items.push(['Tasa de Resolución', '🔴 ROJO', `Crítico: ${tasaResolucion}%`]);

  // Tiempo promedio
  if (promedioTiempo <= 30) items.push(['Tiempo Promedio', '🟢 VERDE', `Óptimo: ${Math.round(promedioTiempo)} días`]);
  else if (promedioTiempo <= 60) items.push(['Tiempo Promedio', '🟡 AMARILLO', `Aceptable: ${Math.round(promedioTiempo)} días`]);
  else items.push(['Tiempo Promedio', '🔴 ROJO', `Crítico: ${Math.round(promedioTiempo)} días`]);

  // Ratio
  if (Number(ratio) <= 1.0) items.push(['Ratio A/C', '🟢 VERDE', `Saludable: ${ratio}`]);
  else if (Number(ratio) <= 1.5) items.push(['Ratio A/C', '🟡 AMARILLO', `Atención: ${ratio}`]);
  else items.push(['Ratio A/C', '🔴 ROJO', `Crítico: ${ratio}`]);

  // Calidad de datos
  if (totals.vacios === 0) items.push(['Calidad de Datos', '🟢 VERDE', '100% completo']);
  else if (totals.vacios <= 10) items.push(['Calidad de Datos', '🟡 AMARILLO', `${totals.vacios} registros vacíos`]);
  else items.push(['Calidad de Datos', '🔴 ROJO', `${totals.vacios} registros vacíos`]);

  return items;
}

function calcularScore(totals, tasaResolucion, promedioTiempo, ratio) {
  let score = 0; const maxScore = 100;
  if (Number(tasaResolucion) >= 70) score += 30; else if (Number(tasaResolucion) >= 50) score += 20; else score += 10;
  if (promedioTiempo <= 30) score += 25; else if (promedioTiempo <= 60) score += 15; else score += 5;
  if (Number(ratio) <= 1.0) score += 25; else if (Number(ratio) <= 1.5) score += 15; else score += 5;
  if (totals.vacios === 0) score += 20; else if (totals.vacios <= 10) score += 10;

  const porcentaje = ((score / maxScore) * 100).toFixed(1);
  let clasificacion = '', recomendacion = '';
  if (score >= 80) { clasificacion = '🏆 EXCELENTE'; recomendacion = 'Backlog en estado óptimo. Mantener prácticas actuales.'; }
  else if (score >= 60) { clasificacion = '✅ BUENO'; recomendacion = 'Backlog controlado. Monitorear áreas de mejora.'; }
  else if (score >= 40) { clasificacion = '⚠️ REGULAR'; recomendacion = 'Requiere atención. Implementar plan de acción.'; }
  else { clasificacion = '❌ CRÍTICO'; recomendacion = 'Situación crítica. Intervención urgente necesaria.'; }

  return [
    ['Score Total:', `${score} / ${maxScore}`, '', `${porcentaje}%`, ''],
    ['Clasificación:', clasificacion, '', '', ''],
    ['Recomendación:', recomendacion, '', '', ''],
    [''],
    ['Desglose de Puntuación:'],
    ['• Tasa de Resolución', '', '30 puntos máx.', '', ''],
    ['• Tiempo Promedio', '', '25 puntos máx.', '', ''],
    ['• Ratio Abiertos/Cerrados', '', '25 puntos máx.', '', ''],
    ['• Calidad de Datos', '', '20 puntos máx.', '', '']
  ];
}

/* ============ UI: INDICADORES Y NOTIFICACIONES ============ */
function showLoadingIndicator(message) {
  let overlay = document.getElementById('loadingOverlay');
  if (!overlay) {
    overlay = document.createElement('div');
    overlay.id = 'loadingOverlay';
    overlay.style.cssText = `
      position: fixed; top: 0; left: 0; width: 100%; height: 100%;
      background: rgba(11, 18, 32, 0.95); display: flex; align-items: center; justify-content: center;
      z-index: 9999; backdrop-filter: blur(8px);
    `;
    overlay.innerHTML = `
      <div style="text-align: center; color: white;">
        <div style="
          width: 60px; height: 60px; border: 4px solid rgba(76, 201, 240, 0.2);
          border-top-color: #4cc9f0; border-radius: 50%; margin: 0 auto 20px; animation: spin 1s linear infinite;">
        </div>
        <div style="font-size: 18px; font-weight: 600; margin-bottom: 8px;">Generando reporte...</div>
        <div id="loadingMessage" style="font-size: 14px; color: #9ca3af;">Por favor espere</div>
      </div>
    `;
    const style = document.createElement('style');
    style.textContent = `@keyframes spin { to { transform: rotate(360deg); } }`;
    document.head.appendChild(style);
    document.body.appendChild(overlay);
  }
  const msgElement = document.getElementById('loadingMessage');
  if (msgElement) msgElement.textContent = message;
  overlay.style.display = 'flex';
}

function hideLoadingIndicator() {
  const overlay = document.getElementById('loadingOverlay');
  if (overlay) overlay.style.display = 'none';
}

function showSuccessNotification(message) { showNotification(message, 'success'); }
function showErrorNotification(message) { showNotification(message, 'error'); }

function showNotification(message, type) {
  const notification = document.createElement('div');
  notification.style.cssText = `
    position: fixed; top: 80px; right: 20px; background: ${type === 'success' ? '#2a9d8f' : '#e76f51'};
    color: white; padding: 16px 24px; border-radius: 12px; box-shadow: 0 10px 40px rgba(0,0,0,0.3);
    z-index: 10000; font-weight: 600; font-size: 14px; display: flex; align-items: center; gap: 12px; min-width: 300px;
    animation: slideIn 0.3s ease-out;
  `;
  const icon = type === 'success' ? '✓' : '✕';
  notification.innerHTML = `<span style="font-size: 20px;">${icon}</span> <span>${message}</span>`;
  const style = document.createElement('style');
  style.textContent = `
    @keyframes slideIn {
      from { transform: translateX(400px); opacity: 0; }
      to   { transform: translateX(0); opacity: 1; }
    }
  `;
  document.head.appendChild(style);
  document.body.appendChild(notification);
  setTimeout(() => {
    notification.style.animation = 'slideIn 0.3s ease-out reverse';
    setTimeout(() => notification.remove(), 300);
  }, 3000);
}

/* ============ OVERRIDE DE EXPORTS ORIGINALES ============ */
/* Para asegurar que los botones existentes llamen las versiones modernas: */
function generatePDFReport() { return generateModernPDFReport(); }
function generateExcelReport() { return generateModernExcelReport(); }

/* Además, exponemos en window por compatibilidad con cualquier código externo: */
window.generatePDFReport = generateModernPDFReport;
window.generateExcelReport = generateModernExcelReport;
function buildSummarySheet(totals, tasaResolucion, fecha) {
  return [
    ['REPORTE DE ANÁLISIS DE BACKLOG DE INCIDENTES'],
    [''],
    ['Fecha de Generación:', fecha],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['RESUMEN DEL REPORTE', '', ''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Parámetro', 'Valor'],
    ['────────────────────────────────────────────────────────────────────────────────────────────────────────────'],
    ['Total de Incidentes', totals.total],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['INDICADORES CLAVE DE DESEMPEÑO (KPIs)', '', ''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Indicador', 'Valor', 'Interpretación'],
    ['────────────────────────────────────────────────────────────────────────────────────────────────────────────'],
    ['Incidentes Abiertos', totals.abiertos, totals.abiertos > totals.cerrados ? '⚠ Atención requerida' : '✓ Bajo control'],
    ['Incidentes Cerrados', totals.cerrados, totals.cerrados >= totals.abiertos ? '✓ Buen ritmo' : '⚠ Mejorar cierre'],
    ['Tasa de Resolución', `${tasaResolucion}%`, Number(tasaResolucion) >= 70 ? '✓ Cumple objetivo' : '❌ Por debajo del objetivo'],
    ['Incidentes Devueltos', totals.devuelto, totals.devuelto < 5 ? '✓ Nivel bajo' : totals.devuelto < 20 ? '⚠ Nivel medio' : '❌ Nivel alto'],
    ['Registros Vacíos', totals.vacios, totals.vacios === 0 ? '✓ Excelente' : totals.vacios < 5 ? '⚠ Bueno' : '❌ Deficiente'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['ANÁLISIS RÁPIDO Y RECOMENDACIONES', '', ''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Métrica', 'Valor', 'Recomendación'],
    ['────────────────────────────────────────────────────────────────────────────────────────────────────────────'],
    ['Ratio Abiertos/Cerrados', (A/C), (totals.abiertos / Math.max(1, totals.cerrados)).toFixed(2),
      (totals.abiertos > totals.cerrados) ? 'Priorizar cierre de casos abiertos' : 'Mantener ritmo de cierre'],
    ['Tiempo Promedio de Resolución', schema?.roles?.tiempo ? `${Math.round(average(rawData.map(r => parseNumber(r[schema.roles.tiempo])).filter(n => !isNaN(n))))} días` : 'N/A',
      schema?.roles?.tiempo ? (average(rawData.map(r => parseNumber(r[schema.roles.tiempo])).filter(n => !isNaN(n))) > 30 ? 'Revisar procesos para mejorar tiempos' : 'Continuar con buenas prácticas') : 'N/A'],
    ['Calidad de Datos', totals.total - totals.vacios,
      totals.vacios === 0 ? 'Mantener calidad de datos' : 'Validar y corregir registros incompletos'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['NOTAS ADICIONALES', '', ''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['• Este reporte proporciona una visión general del estado actual del backlog de incidentes.'],
    ['• Se recomienda revisar los análisis detallados en las hojas siguientes para acciones específicas.'],
    ['• Mantener un monitoreo constante de los KPIs para asegurar la salud del backlog.']
  ];
}

function buildStateAnalysisSheet(totals) {
  const stateData = [
    ['ANÁLISIS POR ESTADO DE INCIDENTES'],
    [''],
    ['Estado', 'Cantidad', '% del Total', 'Descripción', 'Recomendación'],
    ['────────────────────────────────────────────────────────────────────────────────────────────────────────────'],
    ['Abiertos', totals.abiertos, `${((totals.abiertos / totals.total) * 100).toFixed(2)}%`, 'Incidentes pendientes',
      totals.abiertos > totals.cerrados ? 'Priorizar cierre' : 'Ritmo adecuado'],
    ['Cerrados', totals.cerrados, `${((totals.cerrados / totals.total) * 100).toFixed(2)}%`, 'Incidentes resueltos',
      totals.cerrados >= totals.abiertos ? 'Mantener ritmo' : 'Mejorar cierre'],
    ['Devueltos', totals.devuelto, `${((totals.devuelto / totals.total) * 100).toFixed(2)}%`, 'Incidentes rechazados',
      totals.devuelto > 15 ? 'Revisar calidad' : 'Nivel aceptable'],
    ['Vacíos', totals.vacios, `${((totals.vacios / totals.total) * 100).toFixed(2)}%`, 'Registros incompletos',
      totals.vacios > 10 ? 'Corregir datos' : 'Calidad aceptable'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['RESUMEN Y ALERTAS'],
    [''],
    ['RESUMEN GENERAL'],
    [''],
    ['Total Incidentes Analizados:', totals.total],
    ['Incidentes Abiertos:', totals.abiertos],
    ['Incidentes Cerrados:', totals.cerrados],
    ['Incidentes Devueltos:', totals.devuelto],
    ['Registros Vacíos:', totals.vacios],
    [''],
    ['INDICADORES CLAVE'],
    [''],
    ['Tasa de Resolución:', `${((totals.cerrados / totals.total) * 100).toFixed(1)}%`],
    ['Ratio Abiertos/Cerrados:', (totals.abiertos / Math.max(1, totals.cerrados)).toFixed(2)],
    [''],
    ['ALERTAS'],
    [''],
    totals.abiertos > totals.cerrados ? ['⚠ Alto volumen de incidentes abiertos - Priorizar cierre'] : ['✓ Volumen de abiertos bajo control'],
    totals.devuelto > 15 ? [`⚠ Elevado número de incidentes devueltos (${totals.devuelto}) - Revisar calidad`] : ['✓ Nivel de devueltos aceptable'],
    totals.vacios > 10 ? [`⚠ Muchos registros vacíos (${totals.vacios}) - Corregir datos`] : ['✓ Calidad de datos adecuada']
  ];
  const ws = XLSX.utils.aoa_to_sheet(stateData);
  ws['!cols'] = [{ wch: 20 }, { wch: 12 }, { wch: 15 }, { wch: 40 }, { wch: 40 }];
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 4 } },
    { s: { r: 12, c: 0 }, e: { r: 12, c: 4 } }
  ];
  return ws;
}

function buildResponsibleAnalysisSheet(colEstado) {
  const colResp = schema?.roles?.responsable;
  if (!colResp || !colEstado) return XLSX.utils.aoa_to_sheet([['No hay datos de responsables disponibles']]);

  const rowsResp = getOpenClosedByResponsible(rawData, colResp, colEstado);
  const data = [['ANÁLISIS POR RESPONSABLE'], [''], ['Responsable', 'Abiertos', '% Abiertos', 'Cerrados', '% Cerrados', 'Total', 'Eficiencia', 'Estado', 'Acción Recomendada']];

  rowsResp.forEach(r => {
    const total = r.abiertos + r.cerrados;
    const eficiencia = ((r.cerrados / total) * 100).toFixed(1);
    const estado = Number(eficiencia) < 50 ? '❌ Crítico' : Number(eficiencia) < 70 ? '⚠ Medio' : '✓ Bueno';
    const accion = Number(eficiencia) < 70 ? 'Mejorar gestión' : 'Mantener desempeño';
    data.push([cleanText(r.label), r.abiertos, `${((r.abiertos / total) * 100).toFixed(1)}%`, r.cerrados, `${((r.cerrados / total) * 100).toFixed(1)}%`, total, `${eficiencia}%`, estado, accion]);
  });

  const totalAbiertos = rowsResp.reduce((sum, r) => sum + r.abiertos, 0);
  const totalCerrados = rowsResp.reduce((sum, r) => sum + r.cerrados, 0);
  const totalGeneral = totalAbiertos + totalCerrados;
  data.push(['']);
  data.push(['TOTAL', totalAbiertos, `${((totalAbiertos / totalGeneral) * 100).toFixed(1)}%`, totalCerrados, `${((totalCerrados / totalGeneral) * 100).toFixed(1)}%`, totalGeneral, '', '', '']);
  data.push(['']);
  data.push(['RESUMEN Y DESTACADOS']);
  data.push(['']);
  data.push(['Responsable con más casos:', rowsResp.slice().sort((a, b) => (b.abiertos + b.cerrados) - (a.abiertos + a.cerrados))[0]?.label || 'N/A',
    (rowsResp.slice().sort((a, b) => (b.abiertos + b.cerrados) - (a.abiertos + a.cerrados))[0]?.abiertos + rowsResp.slice().sort((a, b) => (b.abiertos + b.cerrados) - (a.abiertos + a.cerrados))[0]?.cerrados) || 0]);
  data.push(['Responsable con mejor eficiencia:', rowsResp.slice().sort((a, b) => ((b.cerrados / (b.abiertos + b.cerrados)) - (a.cerrados / (a.abiertos + a.cerrados))))[0]?.label || 'N/A',
    `${rowsResp.slice().sort((a, b) => ((b.cerrados / (b.abiertos + b.cerrados)) - (a.cerrados / (a.abiertos + a.cerrados))))[0] ? ((rowsResp.slice().sort((a, b) => ((b.cerrados / (b.abiertos + b.cerrados)) - (a.cerrados / (a.abiertos + a.cerrados))))[0].cerrados / (rowsResp.slice().sort((a, b) => ((b.cerrados / (b.abiertos + b.cerrados)) - (a.cerrados / (a.abiertos + a.cerrados))))[0].abiertos + rowsResp.slice().sort((a, b) => ((b.cerrados / (b.abiertos + b.cerrados)) - (a.cerrados / (a.abiertos + a.cerrados))))[0].cerrados)) * 100).toFixed(1) : '0'}%`]);
  data.push(['']);

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 35 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 25 }];
  return ws;
}

function buildServiceAnalysisSheet() {
  const colServicio = schema?.roles?.servicio;
  if (!colServicio) return XLSX.utils.aoa_to_sheet([['No hay datos de servicios disponibles']]);

  const serviceCounts = {};
  rawData.forEach(r => {
    const servicio = r[colServicio] ? String(r[colServicio]).trim() : 'Sin Servicio';
    serviceCounts[servicio] = (serviceCounts[servicio] || 0) + 1;
  });

  const counts = Object.entries(serviceCounts).map(([label, value]) => ({ label, value }));
  counts.sort((a, b) => b.value - a.value);
  const total = counts.reduce((sum, s) => sum + s.value, 0);

  const data = [['ANÁLISIS POR SERVICIO'], [''], ['Servicio', 'Recomendación Estratégica']];
  counts.forEach(s => {
    const porcentaje = (s.value / total) * 100;
    let recomendacion = '';
    if (porcentaje > 30) recomendacion = '❌ Alto volumen - Revisar recursos';
    else if (porcentaje > 15) recomendacion = '⚠ Medio volumen - Monitorear desempeño';
    else recomendacion = '✓ Bajo volumen - Mantener seguimiento';
    data.push([cleanText(s.label), recomendacion]);
  });

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 50 }, { wch: 40 }];
  return ws;
}

function buildTimeAnalysisSheet(colTiempo) {
  if (!colTiempo) return XLSX.utils.aoa_to_sheet([['No hay datos de tiempo disponibles']]);

  const tiempos = rawData.map(r => parseNumber(r[colTiempo])).filter(n => !isNaN(n));
  if (tiempos.length === 0) return XLSX.utils.aoa_to_sheet([['No hay datos de tiempo disponibles']]);

  const promedio = average(tiempos);
  const max = Math.max(...tiempos);
  const min = Math.min(...tiempos);
  const mediana = median(tiempos);

  const rangos = [
    { label: '0-1 días', min: 0, max: 1, count: 0 },
    { label: '2-5 días', min: 2, max: 5, count: 0 },
    { label: '6-15 días', min: 6, max: 15, count: 0 },
    { label: '16-30 días', min: 16, max: 30, count: 0 },
    { label: '31-60 días', min: 31, max: 60, count: 0 },
    { label: '61-90 días', min: 61, max: 90, count: 0 },
    { label: 'Más de 90 días', min: 91, max: Infinity, count: 0 }
  ];

  tiempos.forEach(t => {
    for (const rango of rangos) {
      if (t >= rango.min && t <= rango.max) {
        rango.count += 1;
        break;
      }
    }
  });

  const data = [
    ['ANÁLISIS DE TIEMPOS DE RESOLUCIÓN'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['MÉTRICAS PRINCIPALES'],
    [''],
    ['Métrica', 'Valor'],
    ['────────────────────────────────────────────────────────────────────────────────────────────────────────────'],
    ['Tiempo Promedio de Resolución', `${Math.round(promedio)} días`],
    ['Tiempo Máximo de Resolución', `${Math.round(max)} días`],
    ['Tiempo Mínimo de Resolución', `${Math.round(min)} días`],
    ['Mediana de Tiempo de Resolución', `${Math.round(mediana)} días`],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['DISTRIBUCIÓN DE TIEMPOS'],
    [''],
    ['Rango de Días', 'Cantidad de Casos', '% del Total', 'Nivel de Riesgo'],
    ['────────────────────────────────────────────────────────────────────────────────────────────────────────────'],
    ...rangos.map(r => [r.label, r.count, `${((r.count / tiempos.length) * 100).toFixed(2)}%`,
      (r.count / tiempos.length) * 100 > 30 ? '❌ Alto' : (r.count / tiempos.length) * 100 > 15 ? '⚠ Medio' : '✓ Bajo'])
  ];

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 25 }, { wch: 20 }, { wch: 15 }, { wch: 20 }];
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } },
    { s: { r: 3, c: 0 }, e: { r: 3, c: 3 } },
    { s: { r: 12, c: 0 }, e: { r: 12, c: 3 } },
    { s: { r: 14, c: 0 }, e: { r: 14, c: 3 } },
  ];
  return ws;
}

function buildProviderAnalysisSheet(colEstado) {
  const headers = Object.keys(rawData[0] || {});
  const colProv = headers.includes('Proveedor') ? 'Proveedor' : headers.find(h => /proveedor/i.test(h));
  if (!colProv || !colEstado) return XLSX.utils.aoa_to_sheet([['No hay datos de proveedores disponibles']]);

  const rowsProv = getOpenClosedByProvider(rawData, colProv, colEstado);
  const data = [['ANÁLISIS POR PROVEEDOR'], [''], ['Proveedor', 'Abiertos', '% Abiertos', 'Cerrados', '% Cerrados', 'Total', 'SLA Cumpl.', 'Criticidad']];

  rowsProv.forEach(p => {
    const total = p.abiertos + p.cerrados;
    const sla = ((p.cerrados / total) * 100).toFixed(1);
    const criticidad = Number(sla) < 50 ? '❌ Crítico' : Number(sla) < 70 ? '⚠ Medio' : '✓ Bueno';
    data.push([cleanText(p.label), p.abiertos, `${((p.abiertos / total) * 100).toFixed(1)}%`, p.cerrados, `${((p.cerrados / total) * 100).toFixed(1)}%`, total, `${sla}%`, criticidad]);
  });

  const totalAbiertos = rowsProv.reduce((sum, p) => sum + p.abiertos, 0);
  const totalCerrados = rowsProv.reduce((sum, p) => sum + p.cerrados, 0);
  const totalGeneral = totalAbiertos + totalCerrados;
  data.push(['']);
  data.push(['TOTAL', totalAbiertos, `${((totalAbiertos / totalGeneral) * 100).toFixed(1)}%`, totalCerrados, `${((totalCerrados / totalGeneral) * 100).toFixed(1)}%`, totalGeneral, '', '']);
  data.push(['']);
  data.push(['RESUMEN Y DESTACADOS']);
  data.push(['']);
  data.push(['Proveedor con más casos:', rowsProv.slice().sort((a, b) => (b.abiertos + b.cerrados) - (a.abiertos + a.cerrados))[0]?.label || 'N/A',
    (rowsProv.slice().sort((a, b) => (b.abiertos + b.cerrados) - (a.abiertos + a.cerrados))[0]?.abiertos + rowsProv.slice().sort((a, b) => (b.abiertos + b.cerrados) - (a.abiertos + a.cerrados))[0]?.cerrados) || 0]);
  data.push(['Proveedor con mejor SLA:', rowsProv.slice().sort((a, b) => ((b.cerrados / (b.abiertos + b.cerrados)) - (a.cerrados / (a.abiertos + a.cerrados))))[0]?.label || 'N/A',
    `${rowsProv.slice().sort((a, b) => ((b.cerrados / (b.abiertos + b.cerrados)) - (a.cerrados / (a.abiertos + a.cerrados))))[0] ? ((rowsProv.slice().sort((a, b) => ((b.cerrados / (b.abiertos + b.cerrados)) - (a.cerrados / (a.abiertos + a.cerrados))))[0].cerrados / (rowsProv.slice().sort((a, b) => ((b.cerrados / (b.abiertos + b.cerrados)) - (a.cerrados / (a.abiertos + a.cerrados))))[0].abiertos + rowsProv.slice().sort((a, b) => ((b.cerrados / (b.abiertos + b.cerrados)) - (a.cerrados / (a.abiertos + a.cerrados))))[0].cerrados)) * 100).toFixed(1) : '0'}%`]);
  data.push(['']);

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 35 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 15 }, { wch: 25 }];
  return ws;
}

function buildNarrativeSheet(narrative) {
  const data = [
    ['ANÁLISIS NARRATIVO DEL BACKLOG DE INCIDENTES'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['TOP 5 RESPONSABLES CON MÁS CASOS ABIERTOS'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['────────────────────────────────────────────────────────────────────────────────────────────────────────────'],
    ['Nota: Este análisis identifica a los responsables con mayor número de incidentes abiertos, lo que puede indicar áreas que requieren atención prioritaria para mejorar la gestión y resolución de casos.'],
    [''],
    ['────────────────────────────────────────────────────────────────────────────────────────────────────────────'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Responsables con Más Casos Abiertos'],
    ['────────────────────────────────────────────────────────────────────────────────────────────────────────────'],
    ['Este apartado destaca a los responsables que tienen la mayor cantidad de incidentes abiertos en el backlog. Identificar a estos responsables es crucial para enfocar esfuerzos de mejora y optimización en la gestión de incidentes.'],
    [''],
    ['Responsable', 'Casos Abiertos', '% del Total Abiertos'],
    ...(() => {
      const totalAbiertos = narrative.topResponsables.reduce((sum, resp) => sum + resp.value, 0);
      return narrative.topResponsables.map(resp => [
        cleanText(resp.label),
        resp.value,
        `${((resp.value / totalAbiertos) * 100).toFixed(2)}%`
      ]);
    })(),
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['INSIGHTS Y RECOMENDACIONES'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    [''],
    ...narrative.insights.map(insight => [ `• ${insight}` ]),
    ['']
  ];

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 40 }, { wch: 20 }, { wch: 20 }];
  return ws;
}

function buildSummaryKpiSheet(totals) {
  let promedioTiempo = 0;
  const colTiempo = schema?.roles?.tiempo;
  if (colTiempo) {
    const tiempos = rawData.map(r => parseNumber(r[colTiempo])).filter(n => !isNaN(n));
    promedioTiempo = average(tiempos);
  }

  const tasaResolucion = ((totals.cerrados / totals.total) * 100).toFixed(1);
  const ratio = (totals.abiertos / Math.max(1, totals.cerrados)).toFixed(2);

  const data = [
    ['RESUMEN DE INDICADORES CLAVE DE DESEMPEÑO (KPIs) DEL BACKLOG DE INCIDENTES'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['INDICADORES CLAVE DE DESEMPEÑO (KPIs)'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Indicador', 'Valor Actual', 'Meta Objetivo', 'Cumplimiento', 'Análisis'],
    ['────────────────────────────────────────────────────────────────────────────────────────────────────────────'],
    ['Tasa de Resolución', `${tasaResolucion}%`, '≥ 70%', Number(tasaResolucion) >= 70 ? '✅ CUMPLE' : '❌ NO CUMPLE', Number(tasaResolucion) >= 70 ? '✓ Buen ritmo' : '⚠ Mejorar'],
    ['Tiempo Promedio de Resolución', schema?.roles?.tiempo ? `${Math.round(promedioTiempo)} días` : 'N/A', '≤ 30 días', schema?.roles?.tiempo ? (promedioTiempo <= 30 ? '✅ CUMPLE' : '❌ NO CUMPLE') : 'N/A',
      schema?.roles?.tiempo ? (promedioTiempo <= 30 ? '✓ Óptimo' : promedioTiempo <= 60 ? '⚠ Aceptable' : '❌ Crítico') : 'N/A'],
    ['Ratio Abiertos/Cerrados', ratio, '≤ 1.0', Number(ratio) <= 1.0 ? '✅ CUMPLE' : '❌ NO CUMPLE', Number(ratio) <= 1.0 ? '✓ Saludable' : Number(ratio) <= 1.5 ? '⚠ Atención' : '❌ Crítico'],
    ['Calidad de Datos', `${totals.total - totals.vacios} registros completos`, '100% completo', totals.vacios === 0 ? '✅ CUMPLE' : '❌ NO CUMPLE',
      totals.vacios === 0 ? '✓ Excelente' : totals.vacios <= 10 ? '⚠ Bueno' : '❌ Deficiente'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['SEMAFORO DE INDICADORES'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Indicador', 'Semáforo', 'Interpretación'],
    ['────────────────────────────────────────────────────────────────────────────────────────────────────────────'],
    ...calcularSemaforo(totals, tasaResolucion, promedioTiempo, ratio),
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['SCORE TOTAL Y RECOMENDACIONES'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ...calcularScore(totals, tasaResolucion, promedioTiempo, ratio)
  ];

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 30 }, { wch: 15 }, { wch: 25 }, { wch: 15 }, { wch: 40 }];
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
    { s: { r: 3, c: 0 }, e: { r: 3, c: 4 } },
    { s: { r: 13, c: 0 }, e: { r: 13, c: 4 } },
    { s: { r: 17, c: 0 }, e: { r: 17, c: 4 } },
    { s: { r: 27, c: 0 }, e: { r: 27, c: 4 } }
  ];
  return ws;
}

function calcularSemaforo(totals, tasaResolucion, promedioTiempo, ratio) {
  const items = [];

  // Tasa de Resolución
  if (Number(tasaResolucion) >= 70) items.push(['Tasa de Resolución', '🟢 VERDE', `Cumple: ${tasaResolucion}%`]);
  else if (Number(tasaResolucion) >= 50) items.push(['Tasa de Resolución', '🟡 AMARILLO', `Atención: ${tasaResolucion}%`]);
  else items.push(['Tasa de Resolución', '🔴 ROJO', `Crítico: ${tasaResolucion}%`]);
  
  // Tiempo Promedio
  if (promedioTiempo <= 30) items.push(['Tiempo Promedio de Resolución', '🟢 VERDE', `Óptimo: ${Math.round(promedioTiempo)} días`]);
  else if (promedioTiempo <= 60) items.push(['Tiempo Promedio de Resolución', '🟡 AMARILLO', `Aceptable: ${Math.round(promedioTiempo)} días`]);
  else items.push(['Tiempo Promedio de Resolución', '🔴 ROJO', `Crítico: ${Math.round(promedioTiempo)} días`]);
  // Ratio Abiertos/Cerrados
  if (Number(ratio) <= 1.0) items.push(['Ratio Abiertos/Cerrados', '🟢 VERDE', `Saludable: ${ratio}`]);
  else if (Number(ratio) <= 1.5) items.push(['Ratio Abiertos/Cerrados', '🟡 AMARILLO', `Atención: ${ratio}`]);
  else items.push(['Ratio Abiertos/Cerrados', '🔴 ROJO', `Crítico: ${ratio}`]);
  // Calidad de Datos
  if (totals.vacios === 0) items.push(['Calidad de Datos', '🟢 VERDE', 'Excelente calidad']);
  else if (totals.vacios <= 10) items.push(['Calidad de Datos', '🟡 AMARILLO', 'Buena calidad']);
  else items.push(['Calidad de Datos', '🔴 ROJO', 'Deficiente calidad']);
  
  return items;
}

function calcularScore(totals, tasaResolucion, promedioTiempo, ratio) {
  let score = 0;
  const maxScore = 100;

  // Tasa de Resolución (30 puntos)
  if (Number(tasaResolucion) >= 70) score += 30;
  else if (Number(tasaResolucion) >= 50) score += 20;
  else score += 10;

  // Tiempo Promedio (25 puntos)
  if (promedioTiempo <= 30) score += 25;
  else if (promedioTiempo <= 60) score += 15;
  else score += 5;

  // Ratio Abiertos/Cerrados (25 puntos)
  if (Number(ratio) <= 1.0) score += 25;
  else if (Number(ratio) <= 1.5) score += 15;
  else score += 5;

  // Calidad de Datos (20 puntos)
  if (totals.vacios === 0) score += 20;
  else if (totals.vacios <= 10) score += 10;
  else score += 5;

  const recommendations = [];
  if (Number(tasaResolucion) < 70) recommendations.push('Mejorar la tasa de resolución de incidentes.');
  if (promedioTiempo > 30) recommendations.push('Reducir el tiempo promedio de resolución.');
  if (Number(ratio) > 1.0) recommendations.push('Optimizar el ratio de abiertos a cerrados.');
  if (totals.vacios > 0) recommendations.push('Incrementar la calidad de los datos.');

  return [
    ['Score Total', `${score} / ${maxScore} puntos`],
    [''],
    ['Recomendaciones para Mejorar el Score:'],
    ...recommendations.map(rec => [`• ${rec}`])
  ];
}

function showLoadingIndicator(message = 'Por favor espere') {
  let overlay = document.getElementById('loadingOverlay');
  if (!overlay) {
    overlay = document.createElement('div');
    overlay.id = 'loadingOverlay';
    overlay.style.cssText = `
      position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5);
      display: flex; justify-content: center; align-items: center; z-index: 9999; flex-direction: column;
      font-family: Arial, sans-serif; color: white;
    `;
    overlay.innerHTML = `
      <div style="background: #1f2937; padding: 24px 32px; border-radius: 16px; text-align: center; box-shadow: 0 10px 40px rgba(0,0,0,0.3);">
        <div style="margin-bottom: 16px;">
          <div style="width: 48px; height: 48px; border: 6px solid #3b82f6; border-top-color: transparent; border-radius: 50%; animation: spin 1s linear infinite;"></div>
        </div>
        <div style="font-size: 18px; font-weight: 600;">${message}</div>
      </div>
      <style>
        @keyframes spin {
          from { transform: rotate(0deg); }
          to   { transform: rotate(360deg); }
        }
      </style>
    `;
    document.body.appendChild(overlay);
  }
}

function hideLoadingIndicator() {
  const overlay = document.getElementById('loadingOverlay');
  if (overlay) {
    overlay.remove();
  }
}

function showNotification(message, type = 'success') {
  const notification = document.createElement('div');
  notification.style.cssText = `
    position: fixed; bottom: 24px; right: 24px; background: ${type === 'success' ? '#16a34a' : '#dc2626'}; color: white; padding: 16px 24px; border-radius: 8px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.3); font-family: Arial, sans-serif; display: flex; align-items: center; gap: 12px; animation: slideIn 0.3s ease-out forwards;
    z-index: 10000;
  `;
  const icon = type === 'success' ? '✅' : '❌';
  notification.innerHTML = `
    <div style="font-size: 20px;">${icon}</div>
    <div style="font-size: 16px;">${message}</div>
  `;
  const style = document.createElement('style');
  style.innerHTML = `
    @keyframes slideIn {
      from { transform: translateX(100%); opacity: 0; }
      to { transform: translateX(0); opacity: 1; }
    }
  `;
  document.head.appendChild(style);
  document.body.appendChild(notification);

  setTimeout(() => {
    notification.style.animation = 'slideOut 0.3s ease-in forwards';
    notification.addEventListener('animationend', () => {
      notification.remove();
      style.remove();
    });
  }, 4000);
}

function parseNumber(value) {
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const num = parseFloat(value.replace(/[^0-9.-]+/g, ''));
    return isNaN(num) ? NaN : num;
  }
  return NaN;
}

function cleanText(text) {
  if (typeof text !== 'string') return text;
  return text.replace(/[\n\r\t]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function average(arr) {
  if (arr.length === 0) return 0;
  return arr.reduce((sum, val) => sum + val, 0) / arr.length;
}

function median(arr) {
  if (arr.length === 0) return 0;
  const sorted = arr.slice().sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  if (sorted.length % 2 === 0) {
    return (sorted[mid - 1] + sorted[mid]) / 2;
  } else {
    return sorted[mid];
  }
}

function getOpenClosedByResponsible(data, colResp, colEstado) {
  const respMap = {};
  data.forEach(r => {
    const resp = r[colResp] ? String(r[colResp]).trim() : 'Sin Responsable';
    const estado = r[colEstado] ? String(r[colEstado]).trim().toLowerCase() : '';
    if (!respMap[resp]) {
      respMap[resp] = { abiertos: 0, cerrados: 0 };
    }
    if (estado === 'abierto' || estado === 'open') {
      respMap[resp].abiertos += 1;
    } else if (estado === 'cerrado' || estado === 'closed') {
      respMap[resp].cerrados += 1;
    }
  });
  return Object.entries(respMap).map(([label, counts]) => ({ label, ...counts }));
}
function getOpenClosedByProvider(data, colProv, colEstado) {
  const provMap = {};
  data.forEach(r => {
    const prov = r[colProv] ? String(r[colProv]).trim() : 'Sin Proveedor';
    const estado = r[colEstado] ? String(r[colEstado]).trim().toLowerCase() : '';
    if (!provMap[prov]) {
      provMap[prov] = { abiertos: 0, cerrados: 0 };
    }
    if (estado === 'abierto' || estado === 'open') {
      provMap[prov].abiertos += 1;
    } else if (estado === 'cerrado' || estado === 'closed') {
      provMap[prov].cerrados += 1;
    }
  });
  return Object.entries(provMap).map(([label, counts]) => ({ label, ...counts }));
}

function buildOverallSummarySheet(totals) {
  return [
    ['RESUMEN GENERAL DEL BACKLOG DE INCIDENTES'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['MÉTRICAS PRINCIPALES', '', ''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Métrica', 'Valor', 'Análisis y Recomendación'],
    ['────────────────────────────────────────────────────────────────────────────────────────────────────────────'],
    ['Total de Incidentes Analizados', totals.total, ''],
    ['Incidentes Abiertos', totals.abiertos,
      totals.abiertos > totals.cerrados ? 'Alto volumen de abiertos - Priorizar cierre' : 'Volumen de abiertos bajo control'],
    ['Incidentes Cerrados', totals.cerrados,
      totals.cerrados >= totals.abiertos ? 'Buen ritmo de cierre' : 'Mejorar ritmo de cierre'],
    ['Incidentes Devueltos', totals.devuelto,
      totals.devuelto > 15 ? 'Elevado número de devueltos - Revisar calidad' : 'Nivel de devueltos aceptable'],
    ['Registros Vacíos', totals.vacios,
      totals.vacios > 10 ? 'Muchos registros vacíos - Corregir datos' : 'Calidad de datos adecuada'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['ANÁLISIS DETALLADO Y ALERTAS', '', ''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Estado', 'Cantidad', '% del Total', 'Descripción', 'Acción Recomendada'],
    ['────────────────────────────────────────────────────────────────────────────────────────────────────────────'],
    ['Abiertos', totals.abiertos, `${((totals.abiertos / totals.total) * 100).toFixed(2)}%`, 'Incidentes pendientes de resolución',
      totals.abiertos > totals.cerrados ? 'Priorizar cierre' : 'Bajo control'],
    ['Cerrados', totals.cerrados, `${((totals.cerrados / totals.total) * 100).toFixed(2)}%`, 'Incidentes resueltos',
      totals.cerrados >= totals.abiertos ? 'Buen ritmo de cierre' : 'Mejorar ritmo de cierre'],
    ['Devueltos', totals.devuelto, `${((totals.devuelto / totals.total) * 100).toFixed(2)}%`, 'Incidentes retornados por calidad',
      totals.devuelto > 15 ? 'Revisar calidad' : 'Nivel aceptable'],
    ['Registros Vacíos', totals.vacios, `${((totals.vacios / totals.total) * 100).toFixed(2)}%`, 'Registros con datos incompletos',
      totals.vacios > 10 ? 'Corregir datos' : 'Calidad adecuada'],
    ['']
  ];
};

function buildStateSummarySheet(totals) {
  const stateData = [
    ['RESUMEN DEL ESTADO DEL BACKLOG DE INCIDENTES'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['MÉTRICAS PRINCIPALES'],
    [''],
    ['Total de Incidentes Analizados:', totals.total],
    ['Incidentes Abiertos:', totals.abiertos],
    ['Incidentes Cerrados:', totals.cerrados],
    ['Incidentes Devueltos:', totals.devuelto],
    ['Registros Vacíos:', totals.vacios],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['ANÁLISIS Y RECOMENDACIONES'],
    [''],
    totals.abiertos > totals.cerrados ? [`⚠ Alto volumen de incidentes abiertos (${totals.abiertos}) - Priorizar cierre`] : ['✓ Volumen de abiertos bajo control'],
    totals.cerrados >= totals.abiertos ? ['✓ Buen ritmo de cierre de incidentes'] : [`⚠ Mejorar ritmo de cierre - Cerrados (${totals.cerrados})`],
    totals.devuelto > 15 ? [`⚠ Elevado número de incidentes devueltos (${totals.devuelto}) - Revisar calidad`] : ['✓ Nivel de devueltos aceptable'],
    totals.vacios > 10 ? [`⚠ Muchos registros vacíos (${totals.vacios}) - Corregir datos`] : ['✓ Calidad de datos adecuada'],
    ['']
  ];

  const ws = XLSX.utils.aoa_to_sheet(stateData);
  ws['!cols'] = [{ wch: 50 }, { wch: 50 }];
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 1 } },
    { s: { r: 3, c: 0 }, e: { r: 3, c: 1 } },
    { s: { r: 12, c: 0 }, e: { r: 12, c: 1 } },
  ];
  return ws;
}

function buildResponsibleAnalysisSheet(colResp, colEstado) {
  const rowsResp = getOpenClosedByResponsible(rawData, colResp, colEstado);
  const data = [['ANÁLISIS POR RESPONSABLE'], [''], ['Responsable', 'Abiertos', '% Abiertos', 'Cerrados', '% Cerrados', 'Total', 'Eficiencia', 'Estado', 'Acción Recomendada']];

  rowsResp.forEach(r => {
    const total = r.abiertos + r.cerrados;
    const eficiencia = ((r.cerrados / total) * 100).toFixed(1);
    const estado = Number(eficiencia) < 50 ? '❌ Crítico' : Number(eficiencia) < 70 ? '⚠ Medio' : '✓ Bueno';
    const accion = Number(eficiencia) < 50 ? 'Mejorar gestión' : Number(eficiencia) < 70 ? 'Monitorear desempeño' : 'Mantener seguimiento';
    data.push([cleanText(r.label), r.abiertos, `${((r.abiertos / total) * 100).toFixed(1)}%`, r.cerrados, `${((r.cerrados / total) * 100).toFixed(1)}%`, total, `${eficiencia}%`, estado, accion]);
  });

  const totalAbiertos = rowsResp.reduce((sum, r) => sum + r.abiertos, 0);
  const totalCerrados = rowsResp.reduce((sum, r) => sum + r.cerrados, 0);
  const totalGeneral = totalAbiertos + totalCerrados;
  data.push(['']);
  data.push(['TOTAL', totalAbiertos, `${((totalAbiertos / totalGeneral) * 100).toFixed(1)}%`, totalCerrados, `${((totalCerrados / totalGeneral) * 100).toFixed(1)}%`, totalGeneral, '', '', '']);
  data.push(['']);
  data.push(['RESUMEN Y DESTACADOS']);
  data.push(['']);
  data.push(['Responsable con más casos:', rowsResp.slice().sort((a, b) => (b.abiertos + b.cerrados) - (a.abiertos + a.cerrados))[0]?.label || 'N/A',
    (rowsResp.slice().sort((a, b) => (b.abiertos + b.cerrados) - (a.abiertos + a.cerrados))[0]?.abiertos + rowsResp.slice().sort((a, b) => (b.abiertos + b.cerrados) - (a.abiertos + a.cerrados))[0]?.cerrados) || 0]);
  data.push(['Responsable con mejor eficiencia:', rowsResp.slice().sort((a, b) => ((b.cerrados / (b.abiertos + b.cerrados)) - (a.cerrados / (a.abiertos + a.cerrados))))[0]?.label || 'N/A',
    `${rowsResp.slice().sort((a, b) => ((b.cerrados / (b.abiertos + b.cerrados)) - (a.cerrados / (a.abiertos + a.cerrados))))[0] ? ((rowsResp.slice().sort((a, b) => ((b.cerrados / (b.abiertos + b.cerrados)) - (a.cerrados / (a.abiertos + a.cerrados))))[0].cerrados / (rowsResp.slice().sort((a, b) => ((b.cerrados / (b.abiertos + b.cerrados)) - (a.cerrados / (a.abiertos + a.cerrados))))[0].abiertos + rowsResp.slice().sort((a, b) => ((b.cerrados / (b.abiertos + b.cerrados)) - (a.cerrados / (a.abiertos + a.cerrados))))[0].cerrados)) * 100).toFixed(1) : '0'}%`]);
  data.push(['']);

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 35 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 15 }, { wch: 15 }, { wch: 25 }];
  return ws;
}

function buildServiceAnalysisSheet(colServicio) {
  if (!colServicio) return XLSX.utils.aoa_to_sheet([['No hay datos de servicio disponibles']]);

  const serviceCounts = {};
  rawData.forEach(r => {
    const servicio = r[colServicio] ? String(r[colServicio]).trim() : 'Sin Servicio';
    serviceCounts[servicio] = (serviceCounts[servicio] || 0) + 1;
  });

  const counts = Object.entries(serviceCounts).map(([label, value]) => ({ label, value }));
  const total = counts.reduce((sum, s) => sum + s.value, 0);
  const data = [['ANÁLISIS POR SERVICIO'], [''], ['Servicio', 'Recomendación']];
  
  counts.forEach(s => {
    const porcentaje = (s.value / total) * 100;
    let recomendacion = '';
    if (porcentaje > 30) recomendacion = '❌ Alto volumen - Priorizar revisión';
    else if (porcentaje > 15) recomendacion = '⚠ Medio volumen - Monitorear desempeño';
    else recomendacion = '✓ Bajo volumen - Mantener seguimiento';
    data.push([cleanText(s.label), recomendacion]);
  });
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 50 }, { wch: 40 }];
  return ws;
}

function buildTimeAnalysisSheet(colTiempo) {
  const tiempos = rawData.map(r => parseNumber(r[colTiempo])).filter(n => !isNaN(n));
  if (tiempos.length === 0) return XLSX.utils.aoa_to_sheet([['No hay datos de tiempo disponibles']]);

  const promedio = average(tiempos);
  const max = Math.max(...tiempos);
  const min = Math.min(...tiempos);
  const mediana = median(tiempos);

  const rangos = [
    { label: '0-1 día', min: 0, max: 1, count: 0 },
    { label: '2-5 días', min: 2, max: 5, count: 0 },
    { label: '6-15 días', min: 6, max: 15, count: 0 },
    { label: '16-30 días', min: 16, max: 30, count: 0 },
    { label: '31-60 días', min: 31, max: 60, count: 0 },
    { label: '61-90 días', min: 61, max: 90, count: 0 },
    { label: '91+ días', min: 91, max: Infinity, count: 0 },
  ];

  tiempos.forEach(t => {
    for (const rango of rangos) {
      if (t >= rango.min && t <= rango.max) {
        rango.count += 1;
        break;
      }
    }
  });

  tiempos.forEach(t => {
    for (const rango of rangos) {
      if (t >= rango.min && t <= rango.max) {
        rango.count += 1;
        break;
      }
    }
  });

  const data = [
    ['ANÁLISIS DE TIEMPOS DE RESOLUCIÓN DE INCIDENTES'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['MÉTRICAS DE TIEMPO DE RESOLUCIÓN'],
    [''],
    ['Tiempo Promedio de Resolución', `${Math.round(promedio)} días`],
    ['Tiempo Máximo de Resolución', `${max} días`],
    ['Tiempo Mínimo de Resolución', `${min} días`],
    ['Tiempo Mediano de Resolución', `${mediana} días`],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['DISTRIBUCIÓN DE TIEMPOS DE RESOLUCIÓN'],
    ['────────────────────────────────────────────────────────────────────────────────────────────────────────────'],
    ['Rango de Días', 'Cantidad de Incidentes', '% del Total', 'Nivel de Atención Recomendado'],
    ...rangos.map(r => [r.label, r.count, `${((r.count / tiempos.length) * 100).toFixed(2)}%`,
      r.count / tiempos.length > 0.3 ? '❌ Alto - Priorizar' : r.count / tiempos.length > 0.15 ? '⚠ Medio - Monitorear' : '✓ Bajo - Mantener']),
    ['']
  ];

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 25 }, { wch: 25 }, { wch: 15 }, { wch: 30 }];
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } },
    { s: { r: 3, c: 0 }, e: { r: 3, c: 3 } },
    { s: { r: 11, c: 0 }, e: { r: 11, c: 3 } },
    { s: { r: 13, c: 0 }, e: { r: 13, c: 3 } },
    { s: { r: 9, c: 0 }, e: { r: 9, c: 3 } },
  ];
  return ws;
}

function buildProviderAnalysisSheet(colProv, colEstado) {
  const rowsProv = getOpenClosedByProvider(rawData, colProv, colEstado);
  const data = [['ANÁLISIS POR PROVEEDOR'], [''], ['Proveedor', 'Abiertos', '% Abiertos', 'Cerrados', '% Cerrados', 'Total', 'SLA', 'Criticidad']];

  rowsProv.forEach(p => {
    const total = p.abiertos + p.cerrados;
    const sla = total === 0 ? 0 : ((p.cerrados / total) * 100).toFixed(1);
    let criticidad = '';
    if (sla < 50) criticidad = '❌ Crítico';
    else if (sla < 70) criticidad = '⚠ Medio';
    else criticidad = '✓ Bueno';
    data.push([cleanText(p.label), p.abiertos, `${((p.abiertos / total) * 100).toFixed(1)}%`, p.cerrados, `${((p.cerrados / total) * 100).toFixed(1)}%`, total, `${sla}%`, criticidad]);
  });

  const totalAbiertos = rowsProv.reduce((sum, p) => sum + p.abiertos, 0);
  const totalCerrados = rowsProv.reduce((sum, p) => sum + p.cerrados, 0);
  const totalGeneral = totalAbiertos + totalCerrados;
  data.push(['']);
  data.push(['TOTAL', totalAbiertos, `${((totalAbiertos / totalGeneral) * 100).toFixed(1)}%`, totalCerrados, `${((totalCerrados / totalGeneral) * 100).toFixed(1)}%`, totalGeneral, '', '']);
  data.push(['']);
  data.push(['RESUMEN Y DESTACADOS']);
  data.push(['']);
  data.push(['Proveedor con más casos:', rowsProv.slice().sort((a, b) => (b.abiertos + b.cerrados) - (a.abiertos + a.cerrados))[0]?.label || 'N/A',
    (rowsProv.slice().sort((a, b) => (b.abiertos + b.cerrados) - (a.abiertos + a.cerrados))[0]?.abiertos + rowsProv.slice().sort((a, b) => (b.abiertos + b.cerrados) - (a.abiertos + a.cerrados))[0]?.cerrados) || 0]);
  data.push(['Proveedor con más casos cerrados:', rowsProv.slice().sort((a, b) => b.cerrados - a.cerrados)[0]?.label || 'N/A',
    (rowsProv.slice().sort((a, b) => b.cerrados - a.cerrados)[0]?.cerrados) || 0]);
  data.push(['']);

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 35 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 15 }, { wch: 15 }];
  return ws;
}
