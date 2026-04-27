#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');
const {
  imageSizingContain,
  safeOuterShadow,
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
} = require('./pptx_helpers_local');

const inputJsonPath = process.argv[2];
const outputPptxPath = process.argv[3];
const modeArg = process.argv.find((arg) => arg.startsWith('--mode='));
const renderMode = (modeArg ? modeArg.split('=')[1] : 'body').toLowerCase();

if (!inputJsonPath || !outputPptxPath) {
  console.error('Usage: node scripts/pptx_renderer.js <report.json> <output.pptx> [--mode=body|full]');
  process.exit(1);
}

const report = JSON.parse(fs.readFileSync(inputJsonPath, 'utf8'));
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'BBVA_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'BBVA_WIDE';
pptx.author = 'OpenAI';
pptx.company = 'BBVA';
pptx.subject = 'Comunicaciones Internas';
pptx.lang = 'es-AR';
pptx.theme = { headFontFace: 'Source Serif 4', bodyFontFace: 'Lato', lang: 'es-AR' };
pptx.margin = 0;

const COLORS = {
  electricBlue: '001391',
  sereneBlue: '85C8FF',
  midnight: '060E46',
  sand: 'F7F8F8',
  white: 'FFFFFF',
  ice: '8BE1E9',
  lime: '88E783',
  canary: 'FFE761',
  coral: 'F7893B',
  purple: '6754B8',
  grey4: '46536D',
  grey3: '8A94A6',
  grey2: 'CAD1D8',
  grey1: 'E2E6EA',
  paleBlue: 'EEF7FF',
  paleGreen: 'F1FBF3',
  paleYellow: 'FFF9D6',
  palePurple: 'F4F1FF',
};

const CHART_COLORS = [
  COLORS.electricBlue,
  COLORS.sereneBlue,
  COLORS.ice,
  COLORS.lime,
  COLORS.canary,
  COLORS.purple,
  COLORS.coral,
];

const DEFAULT_EXECUTIVE_TAKEAWAY = 'Se consolidó el desempeño mensual con métricas verificables.';
const BRAND_ASSETS_DIR = path.resolve(__dirname, '..', 'assets', 'brand');
const BBVA_LOGO_BLUE = path.join(BRAND_ASSETS_DIR, 'bbva_logo_blue.png');
const BBVA_LOGO_WHITE = path.join(BRAND_ASSETS_DIR, 'bbva_logo_white.png');
const SHOULD_WARN_LAYOUT = (process.env.PPTX_LAYOUT_WARNINGS || '').toLowerCase() === 'true';

function resolveAsset(assetPath) {
  if (!assetPath) return null;
  const maybe = path.isAbsolute(assetPath) ? assetPath : path.resolve(process.cwd(), assetPath);
  return fs.existsSync(maybe) ? maybe : null;
}

function cleanText(value, fallback = '-') {
  const raw = String(value ?? '').replace(/_/g, ' ').replace(/\s+/g, ' ').trim();
  return raw || fallback;
}

function parseNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  if (value === null || value === undefined || value === '-') return 0;
  const text = String(value).trim();
  if (!text) return 0;
  const normalized = text.includes(',') && text.includes('.') && text.lastIndexOf(',') > text.lastIndexOf('.')
    ? text.replace(/\./g, '').replace(',', '.')
    : text.replace(/,/g, '.');
  const cleaned = normalized.replace(/%/g, '').replace(/[^0-9.-]/g, '');
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function fmtNum(value, digits = 0) {
  return new Intl.NumberFormat('es-AR', { maximumFractionDigits: digits }).format(parseNumber(value));
}

function fmtPct(value) {
  if (value === '-' || value === '' || value === null || value === undefined) return '-';
  const n = parseNumber(value);
  return `${new Intl.NumberFormat('es-AR', { minimumFractionDigits: n % 1 ? 1 : 0, maximumFractionDigits: 2 }).format(n)}%`;
}

function clip(value, max = 80) {
  const text = cleanText(value, '-');
  return text.length > max ? `${text.slice(0, max - 1).trim()}…` : text;
}

function firstNonEmpty(...values) {
  for (const value of values) {
    if (Array.isArray(value) && value.length) return value;
    if (typeof value === 'string' && value.trim()) return value;
    if (value !== null && value !== undefined && typeof value !== 'string' && !Array.isArray(value)) return value;
  }
  return null;
}

function weightedRows(source, limit = 5) {
  if (!Array.isArray(source)) return [];
  return source
    .map((row) => {
      if (!row || typeof row !== 'object') return null;
      const label = cleanText(row.label || row.theme || row.channel || row.name || row.title, 'Sin dato');
      const value = parseNumber(firstNonEmpty(row.value, row.weight, row.pct, row.count, row.total, 0));
      return { label, value, raw: row };
    })
    .filter(Boolean)
    .sort((a, b) => b.value - a.value)
    .slice(0, limit);
}

function sumValues(rows) {
  return rows.reduce((acc, row) => acc + parseNumber(row.value), 0);
}

function valueAsPct(row, rows) {
  const total = sumValues(rows);
  const value = parseNumber(row.value);
  if (value <= 100 && total > 95 && total < 105) return fmtPct(value);
  if (total > 0) return fmtPct((value / total) * 100);
  return fmtNum(value);
}


function valueLabel(row, rows, options = {}) {
  if (options.valueMode === 'number') return fmtNum(row.value, options.digits || 0);
  if (options.valueMode === 'percent') return fmtPct(row.value);
  return valueAsPct(row, rows);
}

function axisMetricLabel(row, rows, options = {}) {
  const label = clip(row.label, options.labelMax || 18);
  const value = valueLabel(row, rows, options);
  return `${label}\n${value}`;
}

function wrapText(value, maxLine = 34, maxLines = 2) {
  const text = cleanText(value, '');
  if (!text) return '-';
  const words = text.split(' ');
  const lines = [];
  let line = '';
  words.forEach((word) => {
    const candidate = line ? `${line} ${word}` : word;
    if (candidate.length <= maxLine) {
      line = candidate;
      return;
    }
    if (line) lines.push(line);
    line = word;
  });
  if (line) lines.push(line);
  if (lines.length > maxLines) {
    const kept = lines.slice(0, maxLines);
    kept[maxLines - 1] = `${kept[maxLines - 1].slice(0, Math.max(0, maxLine - 1)).trim()}…`;
    return kept.join('\n');
  }
  return lines.join('\n');
}

function hasTimeline(rows) {
  return Array.isArray(rows) && rows.length >= 2 && rows.some((row) => parseNumber(row.value) > 0);
}

function bulletize(items, fallback) {
  const rows = Array.isArray(items) ? items.filter(Boolean).map((item) => cleanText(item)).filter(Boolean) : [];
  return rows.length ? rows : [fallback || DEFAULT_EXECUTIVE_TAKEAWAY];
}

function isGenericNarrative(value) {
  const text = cleanText(value, '').toLowerCase();
  if (!text) return true;
  const genericMarkers = [
    'métricas verificables',
    'sin inferencias',
    'desempeño de canales se consolidó',
    'resumen ejecutivo del período',
    'gestión consolidada',
  ];
  return genericMarkers.some((marker) => text.includes(marker));
}

function periodLabel() {
  const rp = report?.render_plan?.period;
  if (rp?.label) return cleanText(rp.label);
  if (report?.period?.label) return cleanText(report.period.label);
  if (report?.slide_1_cover?.period) return cleanText(report.slide_1_cover.period);
  return '-';
}

function buildExecutiveMessage(p) {
  const supplied = p.historical_note || p.headline;
  if (supplied && !isGenericNarrative(supplied) && cleanText(supplied).length > 40) return cleanText(supplied);
  return (
    `En ${periodLabel()}, la comunicación interna registró ${fmtNum(p.plan_total)} comunicaciones planificadas, `
    + `${fmtNum(p.mail_total)} envíos por mail y una tasa de apertura de ${fmtPct(p.mail_open_rate)}. `
    + `En paralelo, SITE/Intranet concentró ${fmtNum(p.site_notes_total)} publicaciones y ${fmtNum(p.site_total_views)} páginas vistas.`
  );
}

function buildExecutiveInsights(p) {
  const provided = Array.isArray(p.takeaways) ? p.takeaways.filter(Boolean).map(cleanText).filter((t) => !isGenericNarrative(t)) : [];
  const insights = [...provided];

  if (parseNumber(p.mail_open_rate) > 0) {
    const level = parseNumber(p.mail_open_rate) >= 75 ? 'alto' : 'moderado';
    insights.push(`La tasa de apertura se ubicó en un nivel ${level} (${fmtPct(p.mail_open_rate)}), señal de buena llegada del canal push.`);
  }
  if (parseNumber(p.mail_interaction_rate) > 0) {
    insights.push(`La interacción promedio fue de ${fmtPct(p.mail_interaction_rate)}, útil para identificar piezas con mayor capacidad de conversión.`);
  }
  if (parseNumber(p.site_total_views) > 0 && parseNumber(p.site_notes_total) > 0) {
    insights.push(`El ecosistema pull promedió ${fmtNum(parseNumber(p.site_total_views) / parseNumber(p.site_notes_total), 0)} vistas por publicación.`);
  }
  return insights.slice(0, 3);
}

function buildChannelNarrative(p) {
  const supplied = p.message;
  if (supplied && !isGenericNarrative(supplied) && cleanText(supplied).length > 60) return cleanText(supplied);
  const rows = weightedRows(p.channel_mix, 5);
  if (!rows.length) return 'No hay datos suficientes para describir la mezcla de canales del período.';
  const top = rows.slice(0, 3);
  return `La distribución se concentró en ${top.map((row) => `${row.label} (${valueAsPct(row, rows)})`).join(', ')}. La lectura permite comparar peso relativo entre canales y detectar oportunidades en soportes de menor participación.`;
}

function buildAxesNarrative(p) {
  const supplied = p.message;
  if (supplied && !isGenericNarrative(supplied) && cleanText(supplied).length > 50) return cleanText(supplied);
  const axes = weightedRows(p.strategic_axes, 5);
  if (!axes.length) return 'No hay datos suficientes para describir la distribución temática.';
  const top = axes.slice(0, 3).map((row) => `${row.label} (${valueAsPct(row, axes)})`).join(', ');
  return `Los ejes con mayor presencia fueron ${top}. Esta priorización ayuda a leer qué temas concentraron la agenda interna del período.`;
}

function buildPushNarrative(p) {
  const supplied = p.message;
  if (supplied && !isGenericNarrative(supplied) && cleanText(supplied).length > 50) return cleanText(supplied);
  const best = Array.isArray(p.by_interaction) ? p.by_interaction[0] : null;
  if (!best) return 'No hay ranking push suficiente para construir una lectura ejecutiva.';
  return `La pieza con mejor desempeño por interacción fue “${clip(best.name || best.title, 72)}”, con ${fmtNum(best.clicks)} clics y ${fmtPct(best.interaction || best.ctr)} de interacción.`;
}

function buildPullNarrative(p) {
  const supplied = p.message;
  if (supplied && !isGenericNarrative(supplied) && cleanText(supplied).length > 50) return cleanText(supplied);
  const best = Array.isArray(p.top_pull_notes) ? p.top_pull_notes[0] : null;
  if (!best) return 'No hay ranking pull suficiente para construir una lectura ejecutiva.';
  return `La nota con mayor tracción fue “${clip(best.title || best.name, 72)}”, con ${fmtNum(best.unique_reads || best.users)} lecturas únicas y ${fmtNum(best.total_reads || best.views)} lecturas totales.`;
}

function addLogo(slide, variant = 'blue') {
  const logo = resolveAsset(variant === 'white' ? BBVA_LOGO_WHITE : BBVA_LOGO_BLUE);
  if (logo) {
    slide.addImage({ path: logo, ...imageSizingContain(logo, 11.92, 0.22, 0.78, 0.26) });
  } else {
    slide.addText('BBVA', { x: 11.82, y: 0.22, w: 0.9, h: 0.24, align: 'right', bold: true, color: variant === 'white' ? COLORS.white : COLORS.electricBlue, fontFace: 'Lato', fontSize: 16, margin: 0 });
  }
}

function markSafeZone(slide, subtitle) {
  if (subtitle) {
    slide.addText(cleanText(subtitle, ''), {
      x: 0.62, y: 0.24, w: 4.8, h: 0.18, fontFace: 'Lato', fontSize: 9.2, color: COLORS.grey4, margin: 0,
    });
  }
  slide.addShape(pptx.ShapeType.rect, { x: 0.62, y: 1.12, w: 0.52, h: 0.05, fill: { color: COLORS.sereneBlue }, line: { color: COLORS.sereneBlue } });
}

function baseSlide(title, subtitle = '') {
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.sand };
  addLogo(slide);
  markSafeZone(slide, subtitle);
  slide.addText(cleanText(title, 'Reporte ejecutivo'), {
    x: 0.62, y: 0.52, w: 8.7, h: 0.46, fontFace: 'Source Serif 4', bold: true, fontSize: 24, color: COLORS.electricBlue, margin: 0, breakLine: false,
  });
  return slide;
}

function panel(slide, x, y, w, h, header = '', options = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: options.fill || COLORS.white },
    line: { color: options.line || COLORS.grey1, width: 1 },
    shadow: options.shadow === false ? undefined : safeOuterShadow('000000', 0.08, 45, 0.7, 0.35),
  });
  if (header) {
    slide.addText(cleanText(header), {
      x: x + 0.16, y: y + 0.13, w: w - 0.32, h: 0.18,
      fontFace: 'Lato', fontSize: 10, bold: true, color: options.headerColor || COLORS.electricBlue, margin: 0,
    });
  }
}

function card(slide, x, y, w, h, label, value, accent = COLORS.sereneBlue, options = {}) {
  panel(slide, x, y, w, h, '', { fill: options.fill || COLORS.white });
  slide.addShape(pptx.ShapeType.rect, { x, y, w, h: 0.06, fill: { color: accent }, line: { color: accent } });
  slide.addText(cleanText(label), { x: x + 0.12, y: y + 0.14, w: w - 0.24, h: 0.14, fontFace: 'Lato', fontSize: options.labelSize || 8.5, color: COLORS.grey4, margin: 0, fit: 'shrink' });
  slide.addText(String(value ?? '-'), { x: x + 0.12, y: y + 0.33, w: w - 0.24, h: 0.25, fontFace: 'Source Serif 4', bold: true, fontSize: options.valueSize || 18, color: COLORS.electricBlue, margin: 0, fit: 'shrink' });
}

function emptyState(slide, x, y, w, h, message, options = {}) {
  panel(slide, x, y, w, h, options.title || 'Sin datos suficientes', { fill: options.fill || COLORS.paleBlue, line: COLORS.grey1, shadow: false });
  slide.addText(clip(message || 'No hay datos consolidados para este módulo.', options.max || 150), {
    x: x + 0.20, y: y + 0.54, w: w - 0.4, h: h - 0.72,
    fontFace: 'Lato', fontSize: options.fontSize || 11, color: COLORS.midnight, align: 'center', valign: 'mid', margin: 0, breakLine: false,
  });
}

function notePill(slide, x, y, w, text) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h: 0.42, rectRadius: 0.08,
    fill: { color: COLORS.paleBlue },
    line: { color: COLORS.grey1, width: 1 },
  });
  slide.addText(clip(text, 120), {
    x: x + 0.16, y: y + 0.12, w: w - 0.32, h: 0.16,
    fontFace: 'Lato', fontSize: 9.2, color: COLORS.grey4, margin: 0,
  });
}

function paragraphBlock(slide, x, y, w, h, body, options = {}) {
  slide.addText(clip(body, options.max || 260), {
    x, y, w, h,
    fontFace: options.fontFace || 'Lato',
    fontSize: options.fontSize || 11,
    bold: !!options.bold,
    color: options.color || COLORS.midnight,
    valign: options.valign || 'top',
    breakLine: false,
    margin: 0,
    fit: 'shrink',
  });
}

function bulletList(slide, x, y, w, items, options = {}) {
  const rows = bulletize(items, options.fallback).slice(0, options.limit || 3);
  rows.forEach((item, idx) => {
    slide.addText(`• ${clip(item, options.max || 118)}`, {
      x, y: y + idx * (options.step || 0.46), w, h: 0.27,
      fontFace: 'Lato', fontSize: options.fontSize || 10.8, color: COLORS.midnight, margin: 0, fit: 'shrink',
    });
  });
}

function tableRows(slide, x, y, w, headers, rows, options = {}) {
  const widths = options.widths || headers.map(() => 1 / headers.length);
  const total = widths.reduce((a, b) => a + b, 0);
  const normalized = widths.map((n) => (n / total) * w);
  const rowHeight = options.rowHeight || 0.46;
  const maxRows = options.maxRows || 5;
  let cursor = x;
  headers.forEach((header, idx) => {
    const colW = normalized[idx];
    slide.addText(cleanText(header), {
      x: cursor + 0.02, y, w: colW - 0.04, h: 0.22,
      fontFace: 'Lato', fontSize: options.headerSize || 8.8, bold: true, color: COLORS.grey4, margin: 0,
      align: idx === 0 ? 'left' : 'right', fit: 'shrink', breakLine: false,
    });
    cursor += colW;
  });
  rows.slice(0, maxRows).forEach((row, rowIndex) => {
    const rowY = y + 0.34 + rowIndex * rowHeight;
    if (rowIndex % 2 === 0) {
      slide.addShape(pptx.ShapeType.roundRect, {
        x, y: rowY - 0.07, w, h: rowHeight - 0.06, rectRadius: 0.03,
        fill: { color: 'F9FBFD' }, line: { color: 'F9FBFD' },
      });
    }
    let xCursor = x;
    row.forEach((value, colIndex) => {
      const colW = normalized[colIndex] || normalized[normalized.length - 1];
      const isFirst = colIndex === 0;
      const cellText = isFirst && options.wrapFirst !== false
        ? wrapText(value, options.wrapMax || 34, options.wrapLines || 2)
        : clip(value, options.clip?.[colIndex] || 20);
      slide.addText(cellText, {
        x: xCursor + 0.04, y: rowY - 0.01, w: colW - 0.08, h: rowHeight - 0.12,
        fontFace: isFirst ? 'Lato' : 'Source Serif 4',
        bold: !isFirst,
        fontSize: isFirst ? (options.firstColSize || 8.7) : (options.valueSize || 9.6),
        color: isFirst ? COLORS.midnight : COLORS.electricBlue,
        align: isFirst ? 'left' : 'right',
        margin: 0,
        fit: 'shrink',
        breakLine: false,
        valign: 'mid',
      });
      xCursor += colW;
    });
  });
}


function renderHorizontalBarChart(slide, x, y, w, h, rows, options = {}) {
  if (!rows.length) return;
  slide.addChart('bar', [{
    name: options.name || 'Participación',
    labels: rows.map((row) => axisMetricLabel(row, rows, { ...options, labelMax: options.labelMax || 22 })),
    values: rows.map((row) => parseNumber(row.value)),
  }], {
    x, y, w, h,
    showLegend: false,
    chartColors: [options.color || COLORS.electricBlue],
    catAxisLabelFontFace: 'Lato', catAxisLabelFontSize: options.catSize || 8,
    valAxisLabelFontFace: 'Lato', valAxisLabelFontSize: 7.5,
    valGridLine: { color: COLORS.grey1, width: 1 },
    showValue: true,
    dataLabelColor: COLORS.electricBlue,
    dataLabelFontFace: 'Lato',
    dataLabelFontSize: options.dataSize || 8,
    dataLabelPosition: 'outEnd',
    showTitle: false,
  });
}

function renderColumnChart(slide, x, y, w, h, rows, options = {}) {
  if (!rows.length) return;
  const maxVal = Math.max(...rows.map((row) => parseNumber(row.value)), 1);
  const chartTop = y + 0.12;
  const chartH = Math.max(0.42, h - 0.64);
  const gap = options.gap || 0.12;
  const slotW = w / rows.length;
  const barW = Math.max(0.16, Math.min(slotW - gap, options.barW || 0.48));
  const baseY = chartTop + chartH;
  rows.forEach((row, idx) => {
    const value = parseNumber(row.value);
    const barH = Math.max(0.05, (value / maxVal) * chartH);
    const cx = x + idx * slotW + (slotW - barW) / 2;
    const color = CHART_COLORS[idx % CHART_COLORS.length];
    slide.addShape(pptx.ShapeType.rect, {
      x: cx, y: baseY - barH, w: barW, h: barH,
      fill: { color },
      line: { color },
    });
    slide.addText(`${wrapText(row.label, options.labelMax || 14, 1)}\n${valueLabel(row, rows, options)}`, {
      x: x + idx * slotW + 0.01, y: baseY + 0.06, w: slotW - 0.02, h: 0.48,
      fontFace: 'Lato', fontSize: options.catSize || 6.8, color: COLORS.midnight,
      align: 'center', margin: 0, fit: 'shrink', breakLine: false, valign: 'mid',
    });
  });
  // Línea base discreta: refuerza que el valor se lee desde la etiqueta inferior y el número sobre la barra.
  slide.addShape(pptx.ShapeType.line, { x, y: baseY, w, h: 0, line: { color: COLORS.grey1, width: 1 } });
}


function renderMiniLineChart(slide, x, y, w, h, rows, options = {}) {
  if (!hasTimeline(rows)) return;
  slide.addChart('line', [{
    name: options.name || 'Evolución',
    labels: rows.map((row) => clip(row.label || row.month || row.period, 12)),
    values: rows.map((row) => parseNumber(row.value)),
  }], {
    x, y, w, h,
    showLegend: false,
    chartColors: [options.color || COLORS.electricBlue],
    catAxisLabelFontFace: 'Lato', catAxisLabelFontSize: 7,
    valAxisLabelFontFace: 'Lato', valAxisLabelFontSize: 7,
    valGridLine: { color: COLORS.grey1, width: 0.5 },
    showValue: false,
    lineSize: 2,
    showTitle: false,
  });
}


function renderDoughnut(slide, x, y, w, h, rows, options = {}) {
  if (!rows.length) return;
  slide.addChart('doughnut', [{
    name: options.name || 'Mix',
    labels: rows.map((row) => clip(row.label, options.labelMax || 16)),
    values: rows.map((row) => parseNumber(row.value)),
  }], {
    x, y, w, h,
    holeSize: options.holeSize || 66,
    chartColors: CHART_COLORS,
    showLegend: false,
    showValue: false,
  });
}

function renderLegend(slide, x, y, w, rows, options = {}) {
  rows.slice(0, options.limit || 5).forEach((row, idx) => {
    const yy = y + idx * (options.step || 0.34);
    const color = CHART_COLORS[idx % CHART_COLORS.length];
    slide.addShape(pptx.ShapeType.roundRect, { x, y: yy + 0.03, w: 0.12, h: 0.12, rectRadius: 0.02, fill: { color }, line: { color } });
    slide.addText(`${clip(row.label, options.labelMax || 22)} ${valueAsPct(row, rows)}`, {
      x: x + 0.18, y: yy, w: w - 0.18, h: 0.18,
      fontFace: 'Lato', fontSize: options.fontSize || 8.6, color: COLORS.midnight, margin: 0, fit: 'shrink',
    });
  });
}

function finalizeSlide(slide) {
  if (!SHOULD_WARN_LAYOUT) return;
  warnIfSlideHasOverlaps(slide);
  warnIfSlideElementsOutOfBounds(slide, pptx);
}

function renderExecutiveSummary(module) {
  const p = module.payload || {};
  const slide = baseSlide(module.title || 'Resumen ejecutivo del período', 'Resumen');
  const kpis = [
    ['Planificación', fmtNum(p.plan_total), COLORS.sereneBlue],
    ['Noticias site', fmtNum(p.site_notes_total), COLORS.ice],
    ['Vistas site', fmtNum(p.site_total_views), COLORS.lime],
    ['Mails enviados', fmtNum(p.mail_total), COLORS.canary],
    ['Apertura', fmtPct(p.mail_open_rate), COLORS.sereneBlue],
    ['Interacción', fmtPct(p.mail_interaction_rate), COLORS.ice],
  ];
  kpis.forEach((item, idx) => {
    card(slide, 0.62 + idx * 2.03, 1.35, idx >= 4 ? 1.75 : 1.86, 0.98, item[0], item[1], item[2], { valueSize: idx >= 4 ? 17 : 18 });
  });

  panel(slide, 0.62, 2.62, 7.35, 3.72, 'Mensaje clave');
  paragraphBlock(slide, 0.9, 3.05, 6.78, 1.18, buildExecutiveMessage(p), { fontFace: 'Source Serif 4', fontSize: 18.5, bold: true, max: 260, color: COLORS.midnight });

  notePill(slide, 0.9, 4.55, 6.76, `Período: ${periodLabel()} · Fuente: dashboard mensual consolidado`);
  paragraphBlock(slide, 0.9, 5.18, 6.76, 0.62, 'Lectura basada en KPIs extraídos y validados de forma determinística para reducir dependencia de narrativa generativa.', { fontSize: 10.2, color: COLORS.grey4, max: 150 });

  panel(slide, 8.28, 2.62, 4.38, 3.72, 'Lectura ejecutiva');
  bulletList(slide, 8.56, 3.1, 3.82, buildExecutiveInsights(p), { fallback: DEFAULT_EXECUTIVE_TAKEAWAY, max: 115, step: 0.62, fontSize: 10.6 });
  finalizeSlide(slide);
}

function renderChannelManagement(module) {
  const p = module.payload || {};
  const slide = baseSlide(module.title || 'Gestión de canales', 'Canales');
  const cards = [
    ['Mails enviados', fmtNum(p.mail_total), COLORS.sereneBlue],
    ['Apertura', fmtPct(p.mail_open_rate), COLORS.ice],
    ['Interacción', fmtPct(p.mail_interaction_rate), COLORS.lime],
    ['Noticias site', fmtNum(p.site_notes_total), COLORS.canary],
    ['Páginas vistas', fmtNum(p.site_total_views), COLORS.sereneBlue],
  ];
  cards.forEach((item, idx) => {
    const x = [0.62, 2.78, 4.94, 7.1, 9.26][idx];
    const w = idx === 4 ? 3.4 : 1.98;
    card(slide, x, 1.28, w, 0.88, item[0], item[1], item[2], { valueSize: idx >= 1 && idx <= 2 ? 16 : 17 });
  });

  const mix = weightedRows(p.channel_mix, 6);
  panel(slide, 0.62, 2.42, 7.2, 4.18, 'Mix de canales');
  if (!mix.length) {
    emptyState(slide, 0.9, 3.0, 6.56, 2.98, p.site_has_no_data_sections ? 'El sitio reporta secciones sin datos en el período.' : 'No hay mix de canales para mostrar.');
  } else {
    renderColumnChart(slide, 0.96, 2.98, 6.34, 2.92, mix.slice(0, 6), { labelMax: 18, catSize: 7.2, valueMode: 'percent', dataSize: 7.5 });
    notePill(slide, 0.96, 6.02, 6.32, `Cada barra indica participación sobre el mix. Principal canal: ${mix[0].label} (${valueAsPct(mix[0], mix)})`);
  }

  panel(slide, 8.08, 2.42, 4.58, 1.84, 'Lectura ejecutiva');
  paragraphBlock(slide, 8.32, 2.9, 4.08, 0.94, buildChannelNarrative(p), { max: 210, fontSize: 9.8 });

  panel(slide, 8.08, 4.48, 2.16, 2.12, 'Evolución mail');
  if (hasTimeline(p.timeline_mail)) {
    renderMiniLineChart(slide, 8.30, 4.94, 1.72, 1.18, p.timeline_mail, { name: 'Mails', color: COLORS.electricBlue });
  } else if (mix.length) {
    tableRows(
      slide,
      8.30,
      4.92,
      1.72,
      ['Canal', 'Peso'],
      mix.slice(0, 3).map((row) => [row.label, valueAsPct(row, mix)]),
      { widths: [0.62, 0.38], rowHeight: 0.34, maxRows: 3, wrapLines: 1, wrapMax: 14, firstColSize: 7.6, valueSize: 8.4, headerSize: 7.5 }
    );
  } else {
    emptyState(slide, 8.28, 4.9, 1.78, 1.2, 'Sin timeline', { title: 'Obs.', fontSize: 8, max: 24 });
  }

  panel(slide, 10.50, 4.48, 2.16, 2.12, 'Evolución site');
  if (hasTimeline(p.timeline_site)) {
    renderMiniLineChart(slide, 10.72, 4.94, 1.72, 1.18, p.timeline_site, { name: 'Notas', color: COLORS.sereneBlue });
  } else if (mix.length) {
    renderLegend(slide, 10.72, 4.92, 1.70, mix.slice(0, 4), { labelMax: 14, step: 0.30, fontSize: 7.4 });
  } else {
    emptyState(slide, 10.70, 4.9, 1.78, 1.2, 'Sin timeline', { title: 'Obs.', fontSize: 8, max: 24 });
  }
  finalizeSlide(slide);
}


function renderMix(module) {
  const p = module.payload || {};
  const strategic = weightedRows(p.strategic_axes, 6);
  const clients = weightedRows(p.internal_clients, 6);
  const formats = weightedRows(p.format_mix, 6);
  const slide = baseSlide(module.title || 'Mix temático y áreas solicitantes', 'Contenido');

  panel(slide, 0.62, 1.28, 6.72, 3.02, 'Ejes estratégicos');
  if (!strategic.length) {
    emptyState(slide, 0.9, 1.86, 6.12, 1.9, 'No hay distribución temática consolidada.');
  } else {
    renderColumnChart(slide, 0.96, 1.86, 5.94, 1.94, strategic.slice(0, 6), { labelMax: 14, catSize: 6.8, dataSize: 7.2 });
  }

  panel(slide, 7.62, 1.28, 5.04, 3.02, 'Áreas solicitantes');
  if (!clients.length) {
    emptyState(slide, 7.92, 1.9, 4.42, 1.82, 'No se detectó el bloque de áreas solicitantes en planificación. Revisar extracción de página 1 o nombre de sección.', { title: 'Dato pendiente', fontSize: 9.2, max: 135, fill: COLORS.paleYellow });
  } else {
    renderHorizontalBarChart(slide, 7.92, 1.86, 4.24, 1.94, clients.slice(0, 5), { labelMax: 18, catSize: 7.0, dataSize: 7.2 });
    notePill(slide, 7.92, 3.82, 4.2, `Mayor demanda: ${clients[0].label} (${valueAsPct(clients[0], clients)})`);
  }

  panel(slide, 0.62, 4.54, 3.4, 1.94, 'Formatos');
  if (!formats.length) {
    emptyState(slide, 0.9, 5.05, 2.82, 0.92, 'Sin mix de formatos.', { title: 'Observación', fontSize: 8.6, max: 60 });
  } else {
    renderDoughnut(slide, 0.86, 4.96, 1.10, 1.10, formats.slice(0, 4), { labelMax: 12, holeSize: 62 });
    renderLegend(slide, 2.08, 4.95, 1.72, formats.slice(0, 4), { labelMax: 16, step: 0.28, fontSize: 7.4 });
  }

  panel(slide, 4.28, 4.54, 3.06, 1.94, 'Lectura temática');
  paragraphBlock(slide, 4.52, 5.02, 2.58, 0.86, buildAxesNarrative(p), { max: 160, fontSize: 8.8 });

  panel(slide, 7.62, 4.54, 5.04, 1.94, 'Resumen de distribución');
  const summaryRows = [];
  if (strategic[0]) summaryRows.push(['Eje líder', strategic[0].label, valueAsPct(strategic[0], strategic)]);
  if (clients[0]) summaryRows.push(['Área líder', clients[0].label, valueAsPct(clients[0], clients)]);
  if (formats[0]) summaryRows.push(['Formato líder', formats[0].label, valueAsPct(formats[0], formats)]);
  if (summaryRows.length) {
    tableRows(slide, 7.9, 5.0, 4.48, ['Métrica', 'Nombre', 'Peso'], summaryRows, {
      widths: [0.26, 0.49, 0.25], rowHeight: 0.34, maxRows: 3, wrapLines: 1, wrapMax: 20, firstColSize: 7.8, valueSize: 8.4, headerSize: 7.6, clip: [16, 26, 10]
    });
  } else {
    emptyState(slide, 7.92, 5.04, 4.42, 0.92, 'Sin distribuciones para resumir.', { title: 'Observación', fontSize: 8.6, max: 60 });
  }
  finalizeSlide(slide);
}


function renderPushRanking(module) {
  const p = module.payload || {};
  const interaction = Array.isArray(p.by_interaction) ? p.by_interaction.slice(0, 4) : [];
  const openRate = Array.isArray(p.by_open_rate) ? p.by_open_rate.slice(0, 4) : [];
  const slide = baseSlide(module.title || 'Ranking push', 'Push');

  if (!p.available || (!interaction.length && !openRate.length)) {
    emptyState(slide, 1.0, 2.0, 11.3, 3.4, 'No hay ranking push suficiente para el período.');
    finalizeSlide(slide);
    return;
  }

  panel(slide, 0.62, 1.28, 6.08, 3.12, 'Top por interacción');
  if (interaction.length) {
    tableRows(
      slide,
      0.84,
      1.82,
      5.62,
      ['Comunicación', 'Clics', 'Interacción'],
      interaction.map((row) => [cleanText(row.name || row.title), fmtNum(row.clicks), fmtPct(row.interaction || row.ctr)]),
      { widths: [0.62, 0.17, 0.21], rowHeight: 0.56, maxRows: 4, wrapMax: 34, wrapLines: 2, firstColSize: 8.2, valueSize: 9.4, headerSize: 8.4 }
    );
  } else {
    emptyState(slide, 0.9, 1.94, 5.5, 1.92, 'No hay ranking por interacción para este período.', { title: 'Observación' });
  }

  panel(slide, 6.92, 1.28, 5.74, 3.12, 'Top por apertura');
  if (openRate.length) {
    tableRows(
      slide,
      7.14,
      1.82,
      5.3,
      ['Comunicación', 'Clics', 'Open rate'],
      openRate.map((row) => [cleanText(row.name || row.title), fmtNum(row.clicks), fmtPct(row.open_rate)]),
      { widths: [0.60, 0.18, 0.22], rowHeight: 0.56, maxRows: 4, wrapMax: 32, wrapLines: 2, firstColSize: 8.2, valueSize: 9.4, headerSize: 8.4 }
    );
  } else {
    emptyState(slide, 7.18, 1.94, 5.22, 1.92, 'La fuente no informó ranking por apertura consolidado.', { title: 'Observación' });
  }

  panel(slide, 0.62, 4.64, 7.2, 1.86, 'Interacción por pieza');
  if (interaction.length) {
    const chartRows = interaction.map((row) => ({ label: row.name || row.title, value: parseNumber(row.interaction || row.ctr) }));
    renderColumnChart(slide, 0.96, 5.06, 6.34, 0.94, chartRows, { labelMax: 18, catSize: 6.4, valueMode: 'percent', dataSize: 6.8 });
  }

  panel(slide, 8.08, 4.64, 4.58, 1.86, 'Lectura ejecutiva');
  paragraphBlock(slide, 8.32, 5.08, 4.08, 0.68, buildPushNarrative(p), { max: 160, fontSize: 9.0 });
  finalizeSlide(slide);
}


function renderPullRanking(module) {
  const p = module.payload || {};
  const rows = Array.isArray(p.top_pull_notes) ? p.top_pull_notes.slice(0, 5) : [];
  const slide = baseSlide(module.title || 'Ranking pull', 'Site / intranet');

  card(slide, 0.62, 1.28, 2.8, 0.88, 'Promedio lecturas por nota', fmtNum(p.average_reads_per_note), COLORS.sereneBlue, { labelSize: 8.2, valueSize: 17 });
  card(slide, 3.62, 1.28, 2.8, 0.88, 'Vistas totales site', fmtNum(p.site_total_views), COLORS.ice, { valueSize: 17 });

  panel(slide, 0.62, 2.42, 7.42, 4.08, 'Top notas pull');
  if (!p.available || !rows.length) {
    emptyState(slide, 0.9, 3.04, 6.84, 2.86, 'No hay ranking pull suficiente para este período.');
  } else {
    tableRows(
      slide,
      0.9,
      2.9,
      6.88,
      ['Nota', 'Únicas', 'Totales'],
      rows.map((row) => [cleanText(row.title || row.name), fmtNum(row.unique_reads || row.users), fmtNum(row.total_reads || row.views)]),
      { widths: [0.66, 0.17, 0.17], rowHeight: 0.50, maxRows: 5, wrapMax: 42, wrapLines: 2, firstColSize: 8.1, valueSize: 9.2, headerSize: 8.3 }
    );
  }

  panel(slide, 8.30, 1.28, 4.36, 2.42, 'Lecturas totales por nota');
  if (rows.length) {
    const chartRows = rows.slice(0, 5).map((row) => ({ label: row.title || row.name, value: parseNumber(row.total_reads || row.views) }));
    renderHorizontalBarChart(slide, 8.58, 1.82, 3.78, 1.34, chartRows, { labelMax: 20, catSize: 6.8, valueMode: 'number', dataSize: 6.8 });
  }

  panel(slide, 8.30, 4.00, 4.36, 2.50, 'Lectura ejecutiva');
  paragraphBlock(slide, 8.58, 4.48, 3.80, 0.96, buildPullNarrative(p), { max: 170, fontSize: 9.2 });
  if (rows.length) notePill(slide, 8.58, 5.72, 3.78, `Top 1: ${clip(rows[0].title || rows[0].name, 54)}`);
  finalizeSlide(slide);
}


function renderMilestones(module) {
  const p = module.payload || {};
  const items = Array.isArray(p.items) ? p.items.slice(0, 3) : [];
  const slide = baseSlide(module.title || 'Hitos del mes', 'Gestión');

  if (!items.length) {
    panel(slide, 0.82, 1.55, 11.68, 4.82, 'Cierre de gestión');
    slide.addText('Sin hitos consolidados para este período', {
      x: 1.18, y: 2.32, w: 6.6, h: 0.42,
      fontFace: 'Source Serif 4', bold: true, fontSize: 24, color: COLORS.electricBlue, margin: 0,
    });
    paragraphBlock(slide, 1.18, 3.1, 5.9, 1.2, cleanText(p.message || 'No se registraron hitos consolidados en la fuente mensual. La slide se conserva como control de calidad para distinguir ausencia de datos de error de generación.'), { max: 220, fontSize: 11 });
    emptyState(slide, 8.08, 2.18, 3.58, 2.74, 'No se detectaron hitos manuales ni automáticos. Puede completarse desde manual_context si el equipo quiere destacar acciones cualitativas.', { title: 'Observación', fill: COLORS.paleYellow, fontSize: 10.2, max: 150 });
    finalizeSlide(slide);
    return;
  }

  items.forEach((item, idx) => {
    const x = 0.62 + idx * 4.18;
    panel(slide, x, 1.35, 3.56, 5.06, `Hito ${idx + 1}`);
    slide.addText(clip(item.title || item.description || '-', 50), {
      x: x + 0.18, y: 1.9, w: 3.15, h: 0.58,
      fontFace: 'Source Serif 4', bold: true, fontSize: 15.5, color: COLORS.electricBlue, margin: 0, fit: 'shrink',
    });
    const bullets = Array.isArray(item.bullets) ? item.bullets.filter(Boolean).slice(0, 3) : [];
    const lines = bullets.length ? bullets : [item.description || 'Sin detalle adicional.'];
    bulletList(slide, x + 0.2, 2.72, 3.08, lines, { max: 68, limit: 3, fontSize: 9.8, step: 0.48 });
    slide.addText(cleanText(item.period || ''), {
      x: x + 0.2, y: 5.9, w: 3.1, h: 0.14,
      fontFace: 'Lato', fontSize: 8.5, color: COLORS.grey4, margin: 0,
    });
  });
  finalizeSlide(slide);
}

function renderEvents(module) {
  const p = module.payload || {};
  const events = Array.isArray(p.events) ? p.events.slice(0, 5) : [];
  const slide = baseSlide(module.title || 'Eventos del mes', 'Activaciones');

  card(slide, 0.62, 1.35, 2.6, 0.92, 'Eventos', fmtNum(p.total_events || events.length), COLORS.sereneBlue);
  card(slide, 3.42, 1.35, 2.8, 0.92, 'Participaciones', fmtNum(p.total_participants), COLORS.ice);

  if (!events.length) {
    emptyState(slide, 0.62, 2.54, 7.8, 3.92, 'No hay detalle de eventos suficiente, por lo que este módulo debería omitirse.');
    finalizeSlide(slide);
    return;
  }

  panel(slide, 0.62, 2.54, 7.8, 3.92, 'Detalle de eventos');
  tableRows(
    slide,
    0.86,
    3.0,
    7.2,
    ['Evento', 'Participantes', 'Fecha'],
    events.map((row) => [row.name || row.title || '-', fmtNum(row.participants), row.date || '-']),
    { widths: [0.58, 0.22, 0.20], rowHeight: 0.42, maxRows: 5, clip: [46, 14, 14] }
  );

  panel(slide, 8.62, 2.54, 4.04, 3.92, 'Mensaje');
  paragraphBlock(slide, 8.88, 3.1, 3.55, 2.9, cleanText(p.message || '-'), { max: 170, fontSize: 11 });
  finalizeSlide(slide);
}

function buildLegacyRenderPlan(payload) {
  // Compatibilidad para inputs históricos que todavía vienen con estructura slide_*.
  const s2 = payload.slide_2_overview || {};
  const s3 = payload.slide_3_plan || {};
  const s4 = payload.slide_4_strategy || payload.slide_3_strategy || {};
  const s5 = payload.slide_5_push_ranking || payload.slide_4_push_ranking || {};
  const s6 = payload.slide_6_pull_performance || payload.slide_5_pull_performance || {};
  const s7 = payload.slide_7_hitos || payload.slide_6_hitos || [];
  const s8 = payload.slide_8_events || payload.slide_7_events || {};
  const includeEvents = Array.isArray(s8.event_breakdown) && s8.event_breakdown.length > 0;
  return {
    period: { label: payload?.slide_1_cover?.period || payload?.period?.label || '-' },
    modules: [
      {
        key: 'executive_summary',
        title: 'Resumen ejecutivo del período',
        payload: {
          headline: s2.headline,
          plan_total: s2.volume_current,
          site_notes_total: s2.pull_notes_current,
          site_total_views: s6.total_views,
          mail_total: s3.mail_total,
          mail_open_rate: s2.push_open_rate,
          mail_interaction_rate: s2.push_interaction_rate,
          historical_note: s2.comparative_note,
          takeaways: s2.highlights,
        },
      },
      {
        key: 'channel_management',
        title: 'Gestión de canales',
        payload: {
          mail_total: s3.mail_total,
          mail_open_rate: s3.open_rate,
          mail_interaction_rate: s2.push_interaction_rate,
          site_notes_total: s3.pull_total,
          site_total_views: s6.total_views,
          channel_mix: s2.audience_segments || [],
          message: s3.footer || s3.mail_message,
          site_has_no_data_sections: false,
        },
      },
      {
        key: 'mix_thematic_clients',
        title: 'Mix temático y áreas solicitantes',
        payload: {
          strategic_axes: s4.content_distribution || [],
          internal_clients: s4.internal_clients || [],
          format_mix: s4.format_mix || [],
          message: s4.conclusion || s4.theme_message,
        },
      },
      {
        key: 'ranking_push',
        title: 'Ranking push',
        payload: {
          by_interaction: s5.top_communications || [],
          by_open_rate: s5.top_by_open_rate || [],
          available: Array.isArray(s5.top_communications) && s5.top_communications.length > 0,
          message: s5.key_learning,
        },
      },
      {
        key: 'ranking_pull',
        title: 'Ranking pull',
        payload: {
          top_pull_notes: s6.top_notes || [],
          available: Array.isArray(s6.top_notes) && s6.top_notes.length > 0,
          average_reads_per_note: s6.avg_reads,
          site_total_views: s6.total_views,
          message: s6.conclusion,
        },
      },
      {
        key: 'milestones',
        title: 'Hitos del mes',
        payload: {
          items: s7,
          message: 'Hitos destacados del período',
        },
      },
      ...(includeEvents ? [{
        key: 'events',
        title: 'Eventos del mes',
        payload: {
          events: s8.event_breakdown || [],
          total_events: s8.total_events,
          total_participants: s8.total_participants,
          message: s8.conclusion || s8.secondary_message,
        },
      }] : []),
    ],
  };
}

function renderFullCover() {
  const s = report?.period || report?.slide_1_cover || {};
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.electricBlue };
  const whiteLogo = resolveAsset(BBVA_LOGO_WHITE);
  if (whiteLogo) slide.addImage({ path: whiteLogo, ...imageSizingContain(whiteLogo, 10.9, 0.45, 1.25, 0.42) });
  slide.addShape(pptx.ShapeType.rect, { x: 0.0, y: 0, w: 0.16, h: 7.5, fill: { color: COLORS.sereneBlue }, line: { color: COLORS.sereneBlue } });
  slide.addText(cleanText(s.label || s.period || periodLabel()), { x: 0.8, y: 2.58, w: 6.4, h: 0.6, fontFace: 'Source Serif 4', bold: true, fontSize: 35, color: COLORS.white, margin: 0 });
  slide.addText('Comunicaciones Internas', { x: 0.8, y: 3.43, w: 8, h: 0.4, fontFace: 'Lato', fontSize: 16, color: COLORS.white, margin: 0 });
  slide.addText('Informe ejecutivo mensual', { x: 0.8, y: 3.9, w: 8, h: 0.32, fontFace: 'Lato', fontSize: 11, color: COLORS.sereneBlue, margin: 0 });
  finalizeSlide(slide);
}

function renderFullClosing() {
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.electricBlue };
  const whiteLogo = resolveAsset(BBVA_LOGO_WHITE);
  if (whiteLogo) slide.addImage({ path: whiteLogo, ...imageSizingContain(whiteLogo, 10.9, 0.45, 1.25, 0.42) });
  slide.addShape(pptx.ShapeType.rect, { x: 0.0, y: 0, w: 0.16, h: 7.5, fill: { color: COLORS.sereneBlue }, line: { color: COLORS.sereneBlue } });
  slide.addText('Fin del informe', { x: 0.8, y: 3.1, w: 6, h: 0.5, fontFace: 'Source Serif 4', bold: true, fontSize: 30, color: COLORS.white, margin: 0 });
  slide.addText(periodLabel(), { x: 0.8, y: 3.7, w: 4, h: 0.25, fontFace: 'Lato', fontSize: 11, color: COLORS.sereneBlue, margin: 0 });
  finalizeSlide(slide);
}

const renderPlan = report?.render_plan && Array.isArray(report.render_plan.modules)
  ? report.render_plan
  : buildLegacyRenderPlan(report);

const renderers = {
  executive_summary: renderExecutiveSummary,
  channel_management: renderChannelManagement,
  mix_thematic_clients: renderMix,
  ranking_push: renderPushRanking,
  ranking_pull: renderPullRanking,
  milestones: renderMilestones,
  events: renderEvents,
};

if (renderMode === 'full') renderFullCover();
for (const module of renderPlan.modules || []) {
  const fn = renderers[module.key];
  if (fn) fn(module);
}
if (renderMode === 'full') renderFullClosing();

pptx.writeFile({ fileName: outputPptxPath }).catch((err) => {
  console.error(err);
  process.exit(1);
});
