#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');

const inputJsonPath = process.argv[2];
const outputPptxPath = process.argv[3];

if (!inputJsonPath || !outputPptxPath) {
  console.error('Usage: node scripts/pptx_renderer.js <report.json> <output.pptx> [--mode=body|full]');
  process.exit(1);
}

const report = JSON.parse(fs.readFileSync(inputJsonPath, 'utf8'));
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'BBVA_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'BBVA_WIDE';
pptx.author = 'BBVA Comunicaciones Internas';
pptx.company = 'BBVA';
pptx.subject = 'Informe de gestión de Comunicaciones Internas';
pptx.lang = 'es-AR';
pptx.theme = { headFontFace: 'Georgia', bodyFontFace: 'Arial', lang: 'es-AR' };
pptx.margin = 0;

const COLORS = {
  blue: '001391',
  dark: '07124A',
  white: 'FFFFFF',
  paper: 'F7F8FA',
  ink: '17233F',
  muted: '56657F',
  grey1: 'E8ECF2',
  grey2: 'CAD1D8',
  grey3: '8A94A6',
  paleBlue: 'EEF7FF',
  paleLilac: 'F1EDF9',
  cyan: '00D7E8',
  sky: '85C8FF',
  aqua: '2DCCCD',
  lime: '88E783',
  yellow: 'F8D44C',
  orange: 'F7893B',
  purple: '6754B8',
  red: 'DA3851',
};
const CHART_COLORS = [COLORS.blue, COLORS.sky, COLORS.aqua, COLORS.yellow, COLORS.orange, COLORS.purple, COLORS.lime, COLORS.red, COLORS.grey2];
const BRAND_ASSETS_DIR = path.resolve(__dirname, '..', 'assets', 'brand');
const BBVA_LOGO_BLUE = path.join(BRAND_ASSETS_DIR, 'bbva_logo_blue.png');
const BBVA_LOGO_WHITE = path.join(BRAND_ASSETS_DIR, 'bbva_logo_white.png');

function exists(p) { return fs.existsSync(p); }
function addLogo(slide, variant = 'blue', x = 11.95, y = 0.22, w = 0.75, h = 0.25) {
  const logo = variant === 'white' ? BBVA_LOGO_WHITE : BBVA_LOGO_BLUE;
  if (exists(logo)) {
    slide.addImage({ path: logo, x, y, w, h });
  } else {
    slide.addText('BBVA', { x, y, w, h, fontFace: 'Arial', fontSize: 14, bold: true, color: variant === 'white' ? COLORS.white : COLORS.blue, margin: 0 });
  }
}
function clean(value, fallback = '-') {
  let text = String(value ?? '').replace(/_/g, ' ').replace(/\s+/g, ' ').trim();
  text = text.replace(/\.\.\./g, '…').replace(/\s*(?:…|\.\.\.)\s*$/g, '').trim();
  return text || fallback;
}
function num(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  if (value === null || value === undefined || value === '') return 0;
  let text = String(value).trim();
  if (text.includes(',') && text.includes('.') && text.lastIndexOf(',') > text.lastIndexOf('.')) text = text.replace(/\./g, '').replace(',', '.');
  else if (text.includes(',') && !text.includes('.')) text = text.replace(',', '.');
  else text = text.replace(/,/g, '');
  const n = Number(text.replace(/%/g, '').replace(/[^0-9.-]/g, ''));
  return Number.isFinite(n) ? n : 0;
}
function fmtNum(value, digits = 0) {
  return new Intl.NumberFormat('es-AR', { maximumFractionDigits: digits }).format(num(value));
}
function fmtPct(value) {
  if (value === null || value === undefined || value === '') return '-';
  const n = num(value);
  return `${new Intl.NumberFormat('es-AR', { minimumFractionDigits: n % 1 ? 1 : 0, maximumFractionDigits: 2 }).format(n)}%`;
}
function truncate(value, max = 62) {
  const text = clean(value, '');
  if (!text) return '-';
  if (text.length <= max) return text;
  const cut = text.slice(0, max - 1).split(' ').slice(0, -1).join(' ').trim();
  return `${cut || text.slice(0, max - 1)}…`;
}
function rows(source, limit = 7) {
  if (!Array.isArray(source)) return [];
  return source.map((item) => {
    if (!item || typeof item !== 'object') return null;
    const label = clean(item.label || item.theme || item.channel || item.name || item.title, '');
    const value = num(item.value ?? item.weight ?? item.pct ?? item.count ?? item.total ?? item.views ?? 0);
    if (!label) return null;
    return { ...item, label, value };
  }).filter(Boolean).sort((a, b) => b.value - a.value).slice(0, limit);
}
function periodLabel() {
  return clean(report?.period?.label || report?.render_plan?.period?.label || 'Q1 2026');
}
function scopeData(scope) {
  return report?.kpis?.scopes?.[scope] || {};
}
function cropPath(scope, moduleName, cropName) {
  const crops = report?.dashboard_crops || {};
  const scopeCrops = crops[scope] || {};
  const moduleCrops = scopeCrops[moduleName] || {};
  const candidate = moduleCrops[cropName];
  if (!candidate) return null;
  const absolute = path.isAbsolute(candidate) ? candidate : path.resolve(process.cwd(), candidate);
  return fs.existsSync(absolute) ? absolute : null;
}
function addCropOrPlaceholder(slide, imagePath, x, y, w, h, label = 'Gráfico no disponible') {
  if (imagePath && fs.existsSync(imagePath)) {
    slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.04, fill: { color: COLORS.white }, line: { color: COLORS.grey1, width: 0.7 } });
    slide.addImage({ path: imagePath, x: x + 0.05, y: y + 0.05, w: w - 0.10, h: h - 0.10, sizing: { type: 'contain', x: x + 0.05, y: y + 0.05, w: w - 0.10, h: h - 0.10 } });
    return;
  }
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.04, fill: { color: COLORS.white }, line: { color: COLORS.grey1, width: 0.7 } });
  slide.addText(label, { x: x + 0.16, y: y + h / 2 - 0.09, w: w - 0.32, h: 0.18, fontFace: 'Arial', fontSize: 8, color: COLORS.muted, align: 'center', margin: 0, fit: 'shrink' });
}
function ensureScopes() {
  const scopes = report?.kpis?.scopes || {};
  for (const s of ['argentina', 'holding', 'combined']) {
    if (!scopes[s]) throw new Error(`Falta scope requerido para renderizar informe: ${s}`);
  }
}
function slideBase(title, subtitle = '') {
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.paper };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: COLORS.paper }, line: { color: COLORS.paper } });
  slide.addText(title, { x: 0.48, y: 0.25, w: 7.8, h: 0.33, fontFace: 'Georgia', fontSize: 18, bold: true, color: COLORS.blue, margin: 0, fit: 'shrink' });
  if (subtitle) slide.addText(subtitle, { x: 0.50, y: 0.67, w: 7.6, h: 0.16, fontFace: 'Arial', fontSize: 8.5, bold: true, color: COLORS.muted, margin: 0, fit: 'shrink' });
  addLogo(slide, 'blue');
  slide.addShape(pptx.ShapeType.line, { x: 0.48, y: 0.98, w: 12.38, h: 0, line: { color: COLORS.grey2, width: 0.75 } });
  return slide;
}
function panel(slide, x, y, w, h, title = '', fill = COLORS.white) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.05, fill: { color: fill }, line: { color: COLORS.grey1, width: 0.7 } });
  if (title) slide.addText(title, { x: x + 0.12, y: y + 0.09, w: w - 0.24, h: 0.14, fontFace: 'Arial', fontSize: 7.6, bold: true, color: COLORS.blue, margin: 0, fit: 'shrink' });
}
function kpiCard(slide, x, y, w, h, label, value, opts = {}) {
  const fill = opts.fill || COLORS.blue;
  const color = opts.color || COLORS.white;
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.05, fill: { color: fill }, line: { color: fill } });
  slide.addText(clean(label), { x: x + 0.10, y: y + 0.08, w: w - 0.20, h: 0.16, fontFace: 'Arial', fontSize: opts.labelSize || 7.4, color, bold: true, align: 'center', margin: 0, fit: 'shrink' });
  slide.addText(String(value ?? '-'), { x: x + 0.10, y: y + 0.29, w: w - 0.20, h: h - 0.34, fontFace: 'Georgia', fontSize: opts.valueSize || 17, bold: true, color, align: 'center', valign: 'mid', margin: 0, fit: 'shrink' });
}
function observationBox(slide, x, y, w, h) {
  panel(slide, x, y, w, h, 'Observaciones del manager', 'FFFFFF');
  slide.addText('Agregar análisis del período…', { x: x + 0.16, y: y + 0.36, w: w - 0.32, h: h - 0.46, fontFace: 'Arial', fontSize: 9, italic: true, color: COLORS.grey3, margin: 0.04, fit: 'shrink', valign: 'top' });
}
function barChart(slide, x, y, w, h, data, opts = {}) {
  const r = rows(data, opts.limit || 6);
  panel(slide, x, y, w, h, opts.title || 'Distribución');
  if (!r.length) {
    slide.addText('Sin datos disponibles', { x: x + 0.18, y: y + h / 2 - 0.08, w: w - 0.36, h: 0.16, fontFace: 'Arial', fontSize: 8, color: COLORS.muted, align: 'center', margin: 0 });
    return;
  }
  const chartX = x + 0.25, chartY = y + 0.45, chartW = w - 0.50, chartH = h - 0.72;
  const max = Math.max(...r.map((a) => a.value), 1);
  const gap = 0.07;
  const bw = Math.max(0.14, (chartW - gap * (r.length - 1)) / r.length);
  r.forEach((d, i) => {
    const bh = Math.max(0.02, chartH * (d.value / max));
    const bx = chartX + i * (bw + gap);
    const by = chartY + chartH - bh;
    const c = CHART_COLORS[i % CHART_COLORS.length];
    slide.addShape(pptx.ShapeType.rect, { x: bx, y: by, w: bw, h: bh, fill: { color: c }, line: { color: c } });
    slide.addText(opts.percent ? fmtPct(d.value) : fmtNum(d.value), { x: bx - 0.02, y: by - 0.16, w: bw + 0.04, h: 0.10, fontFace: 'Arial', fontSize: 5.7, color: COLORS.ink, bold: true, align: 'center', margin: 0, fit: 'shrink' });
    slide.addText(truncate(d.label, opts.labelMax || 12), { x: bx - 0.04, y: chartY + chartH + 0.06, w: bw + 0.08, h: 0.23, fontFace: 'Arial', fontSize: 5.2, color: COLORS.ink, align: 'center', rotate: opts.rotate ? 25 : 0, margin: 0, fit: 'shrink' });
  });
}
function donutLegend(slide, x, y, w, h, data, opts = {}) {
  const r = rows(data, opts.limit || 6);
  panel(slide, x, y, w, h, opts.title || 'Distribución');
  if (!r.length) {
    slide.addText('Sin datos disponibles', { x: x + 0.18, y: y + h / 2 - 0.08, w: w - 0.36, h: 0.16, fontFace: 'Arial', fontSize: 8, color: COLORS.muted, align: 'center', margin: 0 });
    return;
  }
  const total = r.reduce((a, b) => a + b.value, 0) || 1;
  r.forEach((d, i) => {
    const yy = y + 0.40 + i * ((h - 0.55) / Math.max(r.length, 1));
    const c = CHART_COLORS[i % CHART_COLORS.length];
    slide.addShape(pptx.ShapeType.rect, { x: x + 0.18, y: yy + 0.02, w: 0.08, h: 0.08, fill: { color: c }, line: { color: c } });
    slide.addText(truncate(d.label, opts.labelMax || 22), { x: x + 0.34, y: yy - 0.01, w: w - 1.28, h: 0.14, fontFace: 'Arial', fontSize: 6.6, color: COLORS.ink, margin: 0, fit: 'shrink' });
    const value = opts.percent === false ? fmtNum(d.value) : fmtPct(d.value > 100 ? (d.value / total) * 100 : d.value);
    slide.addText(value, { x: x + w - 0.86, y: yy - 0.01, w: 0.66, h: 0.14, fontFace: 'Arial', fontSize: 6.6, bold: true, color: COLORS.blue, align: 'right', margin: 0, fit: 'shrink' });
  });
}
function topTable(slide, x, y, w, h, title, data, valueKeys, opts = {}) {
  panel(slide, x, y, w, h, title, opts.fill || COLORS.white);
  const r = Array.isArray(data) ? data.slice(0, opts.limit || 5) : [];
  const headY = y + 0.34;
  slide.addShape(pptx.ShapeType.rect, { x: x + 0.12, y: headY, w: w - 0.24, h: 0.18, fill: { color: opts.headerColor || COLORS.sky }, line: { color: opts.headerColor || COLORS.sky } });
  slide.addText('Título', { x: x + 0.20, y: headY + 0.04, w: w * 0.62, h: 0.08, fontFace: 'Arial', fontSize: 5.8, bold: true, color: COLORS.dark, margin: 0, fit: 'shrink' });
  slide.addText(opts.valueLabel || 'Valor', { x: x + w * 0.70, y: headY + 0.04, w: w * 0.20, h: 0.08, fontFace: 'Arial', fontSize: 5.8, bold: true, color: COLORS.dark, align: 'right', margin: 0, fit: 'shrink' });
  if (!r.length) {
    slide.addText('Sin datos disponibles', { x: x + 0.18, y: y + h / 2, w: w - 0.36, h: 0.16, fontFace: 'Arial', fontSize: 7.2, color: COLORS.muted, align: 'center', margin: 0 });
    return;
  }
  const rowH = (h - 0.66) / Math.max(r.length, 1);
  r.forEach((row, i) => {
    const yy = y + 0.58 + i * rowH;
    slide.addShape(pptx.ShapeType.rect, { x: x + 0.12, y: yy - 0.03, w: w - 0.24, h: rowH - 0.02, fill: { color: i % 2 ? 'FFFFFF' : 'F3F6FA' }, line: { color: i % 2 ? 'FFFFFF' : 'F3F6FA' } });
    const titleText = row.title || row.name || row.label || row.communication || '-';
    let value = '-';
    for (const key of valueKeys) {
      if (row[key] !== undefined && row[key] !== null && row[key] !== '') { value = row[key]; break; }
    }
    const isPct = opts.percent !== false;
    slide.addText(truncate(titleText, opts.titleMax || 44), { x: x + 0.20, y: yy, w: w * 0.64, h: rowH - 0.06, fontFace: 'Arial', fontSize: opts.fontSize || 5.8, color: COLORS.ink, margin: 0, fit: 'shrink', valign: 'mid' });
    slide.addText(isPct ? fmtPct(value) : fmtNum(value), { x: x + w * 0.70, y: yy, w: w * 0.20, h: rowH - 0.06, fontFace: 'Arial', fontSize: opts.fontSize || 5.8, bold: true, color: COLORS.blue, align: 'right', margin: 0, fit: 'shrink', valign: 'mid' });
  });
}
function contentTable(slide, x, y, w, h, title, data, opts = {}) {
  panel(slide, x, y, w, h, title, opts.fill || COLORS.white);
  const r = Array.isArray(data) ? data.slice(0, opts.limit || 5) : [];
  const headY = y + 0.34;
  slide.addShape(pptx.ShapeType.rect, { x: x + 0.12, y: headY, w: w - 0.24, h: 0.18, fill: { color: opts.headerColor || COLORS.sky }, line: { color: opts.headerColor || COLORS.sky } });
  slide.addText('Titular', { x: x + 0.20, y: headY + 0.04, w: w * 0.56, h: 0.08, fontFace: 'Arial', fontSize: 5.7, bold: true, color: COLORS.dark, margin: 0, fit: 'shrink' });
  slide.addText('Equipo', { x: x + w * 0.62, y: headY + 0.04, w: w * 0.13, h: 0.08, fontFace: 'Arial', fontSize: 5.7, bold: true, color: COLORS.dark, margin: 0, fit: 'shrink' });
  slide.addText('UU', { x: x + w * 0.77, y: headY + 0.04, w: w * 0.08, h: 0.08, fontFace: 'Arial', fontSize: 5.7, bold: true, color: COLORS.dark, align: 'right', margin: 0, fit: 'shrink' });
  slide.addText('Vistas', { x: x + w * 0.86, y: headY + 0.04, w: w * 0.10, h: 0.08, fontFace: 'Arial', fontSize: 5.7, bold: true, color: COLORS.dark, align: 'right', margin: 0, fit: 'shrink' });
  if (!r.length) {
    slide.addText('Sin datos disponibles', { x: x + 0.18, y: y + h / 2, w: w - 0.36, h: 0.16, fontFace: 'Arial', fontSize: 7.2, color: COLORS.muted, align: 'center', margin: 0 });
    return;
  }
  const rowH = (h - 0.66) / Math.max(r.length, 1);
  r.forEach((row, i) => {
    const yy = y + 0.58 + i * rowH;
    slide.addShape(pptx.ShapeType.rect, { x: x + 0.12, y: yy - 0.03, w: w - 0.24, h: rowH - 0.02, fill: { color: i % 2 ? 'FFFFFF' : 'F3F6FA' }, line: { color: i % 2 ? 'FFFFFF' : 'F3F6FA' } });
    slide.addText(truncate(row.title || row.name || row.label, opts.titleMax || 54), { x: x + 0.20, y: yy, w: w * 0.56, h: rowH - 0.06, fontFace: 'Arial', fontSize: opts.fontSize || 5.5, color: COLORS.ink, margin: 0, fit: 'shrink', valign: 'mid' });
    slide.addText(truncate(row.team || row.equipo || row.scope_label || row.scope || '-', 16), { x: x + w * 0.62, y: yy, w: w * 0.13, h: rowH - 0.06, fontFace: 'Arial', fontSize: opts.fontSize || 5.5, color: COLORS.ink, margin: 0, fit: 'shrink', valign: 'mid' });
    slide.addText(fmtNum(row.unique_reads ?? row.users ?? row.uu ?? row.reads ?? 0), { x: x + w * 0.77, y: yy, w: w * 0.08, h: rowH - 0.06, fontFace: 'Arial', fontSize: opts.fontSize || 5.5, bold: true, color: COLORS.blue, align: 'right', margin: 0, fit: 'shrink', valign: 'mid' });
    slide.addText(fmtNum(row.total_reads ?? row.views ?? row.page_views ?? 0), { x: x + w * 0.86, y: yy, w: w * 0.10, h: rowH - 0.06, fontFace: 'Arial', fontSize: opts.fontSize || 5.5, bold: true, color: COLORS.blue, align: 'right', margin: 0, fit: 'shrink', valign: 'mid' });
  });
}
function mailVolume(scope) {
  const unique = num(scope.mail_unique_total);
  const send = num(scope.mail_send_total || scope.mail_total);
  if (unique > 0 && send > 0 && unique !== send) return `${fmtNum(unique)} únicos / ${fmtNum(send)} envíos`;
  if (send > 0) return `${fmtNum(send)} envíos`;
  if (unique > 0) return `${fmtNum(unique)} únicos`;
  return '-';
}
function scopedKpiRow(slide, x, y, scope, label, accent) {
  slide.addText(label, { x, y: y + 0.15, w: 1.05, h: 0.16, fontFace: 'Georgia', fontSize: 9.5, color: COLORS.dark, margin: 0, fit: 'shrink' });
  kpiCard(slide, x + 1.12, y, 1.56, 0.64, 'Mails', mailVolume(scope), { fill: accent, valueSize: 8.8, labelSize: 5.6 });
  kpiCard(slide, x + 2.84, y, 1.52, 0.64, 'Apertura promedio', fmtPct(scope.mail_open_rate), { fill: COLORS.blue, valueSize: 12, labelSize: 5.6 });
  kpiCard(slide, x + 4.50, y, 1.78, 0.64, 'Interacción enviados', fmtPct(scope.mail_interaction_rate), { fill: COLORS.blue, valueSize: 12, labelSize: 5.6 });
  kpiCard(slide, x + 6.43, y, 1.78, 0.64, 'Interacción abiertos', fmtPct(scope.mail_interaction_rate_over_opened), { fill: COLORS.blue, valueSize: 12, labelSize: 5.6 });
}
function trendCrop(slide, x, y, w, h, scopeName, title = 'Tendencia mensual de envíos y aperturas') {
  panel(slide, x, y, w, h, title);
  addCropOrPlaceholder(slide, cropPath(scopeName, 'mailing', 'monthly_trend'), x + 0.10, y + 0.34, w - 0.20, h - 0.46, 'Tendencia mensual no disponible');
}
function renderCover() {
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.dark };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: COLORS.dark }, line: { color: COLORS.dark } });
  addLogo(slide, 'white', 0.50, 0.40, 0.75, 0.25);
  slide.addText('Comunicaciones\nInternas', { x: 0.62, y: 1.78, w: 6.6, h: 0.95, fontFace: 'Georgia', fontSize: 29, bold: true, color: COLORS.white, margin: 0, breakLine: false, fit: 'shrink' });
  slide.addShape(pptx.ShapeType.line, { x: 0.62, y: 3.10, w: 7.8, h: 0, line: { color: COLORS.sky, transparency: 28, width: 1.2 } });
  slide.addText(`Gestión ${periodLabel().replace(/\s*\([^)]*\)/, '')}`, { x: 0.62, y: 3.55, w: 5.8, h: 0.42, fontFace: 'Arial', fontSize: 24, color: COLORS.white, margin: 0, fit: 'shrink' });
  slide.addShape(pptx.ShapeType.roundRect, { x: 8.15, y: 1.55, w: 3.70, h: 2.25, rectRadius: 0.25, fill: { color: COLORS.cyan, transparency: 10 }, line: { color: COLORS.sky, transparency: 35, width: 1.4 } });
}
function renderPlanningComparison() {
  const arg = scopeData('argentina');
  const hol = scopeData('holding');
  const slide = slideBase(`Gestión CI - ${periodLabel()}`, 'Planificación | Argentina vs Holding');
  kpiCard(slide, 1.45, 1.12, 3.10, 0.68, 'ARGENTINA · Acciones de Comunicación', fmtNum(arg.plan_total), { fill: COLORS.blue, valueSize: 19 });
  kpiCard(slide, 8.55, 1.12, 3.10, 0.68, 'HOLDING · Acciones de Comunicación', fmtNum(hol.plan_total), { fill: COLORS.muted, valueSize: 19 });
  slide.addText('Distribución por Eje Estratégico', { x: 0.58, y: 1.92, w: 3.8, h: 0.16, fontFace: 'Arial', fontSize: 8, bold: true, color: COLORS.blue, margin: 0 });
  slide.addText('Distribución por Canales', { x: 4.45, y: 1.92, w: 2.4, h: 0.16, fontFace: 'Arial', fontSize: 8, bold: true, color: COLORS.blue, margin: 0 });
  slide.addText('Distribución por Eje Estratégico', { x: 7.08, y: 1.92, w: 3.8, h: 0.16, fontFace: 'Arial', fontSize: 8, bold: true, color: COLORS.blue, margin: 0 });
  slide.addText('Distribución por Canales', { x: 10.75, y: 1.92, w: 2.1, h: 0.16, fontFace: 'Arial', fontSize: 8, bold: true, color: COLORS.blue, margin: 0 });
  addCropOrPlaceholder(slide, cropPath('argentina', 'planning', 'strategic_axes'), 0.55, 2.12, 3.55, 1.48, 'Eje estratégico no disponible');
  addCropOrPlaceholder(slide, cropPath('argentina', 'planning', 'channel_mix'), 4.28, 2.12, 2.52, 1.48, 'Canales no disponibles');
  addCropOrPlaceholder(slide, cropPath('holding', 'planning', 'strategic_axes'), 7.02, 2.12, 3.55, 1.48, 'Eje estratégico no disponible');
  addCropOrPlaceholder(slide, cropPath('holding', 'planning', 'channel_mix'), 10.73, 2.12, 2.18, 1.48, 'Canales no disponibles');
  slide.addText('Área solicitante · Argentina', { x: 0.58, y: 3.78, w: 3.6, h: 0.16, fontFace: 'Arial', fontSize: 8, bold: true, color: COLORS.blue, margin: 0 });
  slide.addText('Área solicitante · Holding', { x: 7.08, y: 3.78, w: 3.6, h: 0.16, fontFace: 'Arial', fontSize: 8, bold: true, color: COLORS.blue, margin: 0 });
  addCropOrPlaceholder(slide, cropPath('argentina', 'planning', 'internal_clients'), 0.55, 3.98, 6.25, 1.20, 'Área solicitante no disponible');
  addCropOrPlaceholder(slide, cropPath('holding', 'planning', 'internal_clients'), 7.02, 3.98, 5.89, 1.20, 'Área solicitante no disponible');
  observationBox(slide, 0.55, 5.52, 12.36, 1.18);
}
function renderPlanningCombined() {
  const c = scopeData('combined');
  const slide = slideBase(`Gestión CI - ${periodLabel()}`, 'Planificación | Argentina + Holding');
  kpiCard(slide, 4.70, 1.10, 3.90, 0.72, 'Acciones de Comunicación', fmtNum(c.plan_total), { fill: COLORS.blue, valueSize: 22 });
  slide.addText('Distribución por Eje Estratégico', { x: 0.65, y: 2.04, w: 3.6, h: 0.16, fontFace: 'Arial', fontSize: 8.4, bold: true, color: COLORS.blue, margin: 0 });
  slide.addText('Distribución por Canales', { x: 4.74, y: 2.04, w: 3.2, h: 0.16, fontFace: 'Arial', fontSize: 8.4, bold: true, color: COLORS.blue, margin: 0 });
  slide.addText('Área solicitante', { x: 8.44, y: 2.04, w: 3.2, h: 0.16, fontFace: 'Arial', fontSize: 8.4, bold: true, color: COLORS.blue, margin: 0 });
  addCropOrPlaceholder(slide, cropPath('combined', 'planning', 'strategic_axes'), 0.62, 2.28, 3.75, 2.08, 'Eje estratégico no disponible');
  addCropOrPlaceholder(slide, cropPath('combined', 'planning', 'channel_mix'), 4.68, 2.28, 3.40, 2.08, 'Canales no disponibles');
  addCropOrPlaceholder(slide, cropPath('combined', 'planning', 'internal_clients'), 8.38, 2.28, 4.32, 2.08, 'Área solicitante no disponible');
  observationBox(slide, 0.62, 4.82, 12.08, 1.62);
}
function renderMailComparison() {
  const arg = scopeData('argentina');
  const hol = scopeData('holding');
  const slide = slideBase(`Gestión CI - ${periodLabel()}`, 'Canal Mail | Argentina vs Holding');
  scopedKpiRow(slide, 0.72, 1.12, arg, 'Argentina', COLORS.blue);
  scopedKpiRow(slide, 0.72, 1.82, hol, 'Holding', COLORS.muted);
  trendCrop(slide, 0.70, 2.62, 5.65, 1.15, 'argentina', 'Tendencia mensual · Argentina');
  trendCrop(slide, 6.65, 2.62, 5.65, 1.15, 'holding', 'Tendencia mensual · Holding');
  topTable(slide, 0.70, 4.02, 5.65, 0.95, 'Top five - Mayor Tasa de Apertura · Argentina', arg.top_push_by_open_rate, ['open_rate', 'rate', 'value'], { valueLabel: 'Apertura', headerColor: COLORS.lime, fontSize: 5.2, titleMax: 42, limit: 5 });
  topTable(slide, 6.65, 4.02, 5.65, 0.95, 'Top five - Mayor Tasa de Interacción · Argentina', arg.top_push_by_interaction, ['interaction', 'interaction_rate', 'ctr', 'value'], { valueLabel: 'Interacción', headerColor: COLORS.sky, fontSize: 5.2, titleMax: 42, limit: 5 });
  topTable(slide, 0.70, 5.10, 5.65, 0.95, 'Top five - Mayor Tasa de Apertura · Holding', hol.top_push_by_open_rate, ['open_rate', 'rate', 'value'], { valueLabel: 'Apertura', headerColor: COLORS.lime, fontSize: 5.2, titleMax: 42, limit: 5 });
  topTable(slide, 6.65, 5.10, 5.65, 0.95, 'Top five - Mayor Tasa de Interacción · Holding', hol.top_push_by_interaction, ['interaction', 'interaction_rate', 'ctr', 'value'], { valueLabel: 'Interacción', headerColor: COLORS.sky, fontSize: 5.2, titleMax: 42, limit: 5 });
  observationBox(slide, 0.70, 6.22, 11.60, 0.58);
}
function renderMailCombined() {
  const c = scopeData('combined');
  const slide = slideBase(`Gestión CI - ${periodLabel()}`, 'Canal Mail | Argentina + Holding');
  kpiCard(slide, 0.82, 1.12, 2.55, 0.68, 'Mails', mailVolume(c), { fill: COLORS.blue, valueSize: 12 });
  kpiCard(slide, 3.62, 1.12, 2.35, 0.68, 'Tasa de apertura promedio', fmtPct(c.mail_open_rate), { fill: COLORS.blue, valueSize: 16 });
  kpiCard(slide, 6.22, 1.12, 2.70, 0.68, 'Interacción sobre enviados', fmtPct(c.mail_interaction_rate), { fill: COLORS.blue, valueSize: 16 });
  kpiCard(slide, 9.17, 1.12, 2.70, 0.68, 'Interacción sobre abiertos', fmtPct(c.mail_interaction_rate_over_opened), { fill: COLORS.blue, valueSize: 16 });
  trendCrop(slide, 0.82, 2.14, 11.05, 1.42, 'combined');
  topTable(slide, 0.82, 3.86, 5.35, 1.40, 'Top five - Mayor Tasa de Apertura', c.top_push_by_open_rate, ['open_rate', 'rate', 'value'], { valueLabel: 'Apertura', headerColor: COLORS.lime, fontSize: 5.7, titleMax: 46 });
  topTable(slide, 6.52, 3.86, 5.35, 1.40, 'Top five - Mayor Tasa de Interacción', c.top_push_by_interaction, ['interaction', 'interaction_rate', 'ctr', 'value'], { valueLabel: 'Interacción', headerColor: COLORS.sky, fontSize: 5.7, titleMax: 46 });
  observationBox(slide, 0.82, 5.58, 11.05, 0.88);
}
function renderContentComparison() {
  const arg = scopeData('argentina');
  const hol = scopeData('holding');
  const slide = slideBase(`Gestión CI - ${periodLabel()}`, 'Canal Intranet / Contenidos | Argentina vs Holding');
  kpiCard(slide, 1.50, 1.08, 1.75, 0.58, 'Noticias Publicadas · ARG', fmtNum(arg.site_notes_total), { fill: COLORS.blue, valueSize: 15, labelSize: 5.6 });
  kpiCard(slide, 3.44, 1.08, 1.95, 0.58, 'Total Páginas Vistas · ARG', fmtNum(arg.site_total_views), { fill: COLORS.blue, valueSize: 13, labelSize: 5.6 });
  kpiCard(slide, 5.58, 1.08, 1.75, 0.58, 'Promedio Vistas · ARG', fmtNum(arg.site_average_views), { fill: COLORS.blue, valueSize: 13, labelSize: 5.6 });
  kpiCard(slide, 7.72, 1.08, 1.75, 0.58, 'Noticias Publicadas · HOL', fmtNum(hol.site_notes_total), { fill: COLORS.muted, valueSize: 15, labelSize: 5.6 });
  kpiCard(slide, 9.66, 1.08, 1.95, 0.58, 'Total Páginas Vistas · HOL', fmtNum(hol.site_total_views), { fill: COLORS.muted, valueSize: 13, labelSize: 5.6 });
  kpiCard(slide, 11.80, 1.08, 1.35, 0.58, 'Promedio · HOL', fmtNum(hol.site_average_views), { fill: COLORS.muted, valueSize: 13, labelSize: 5.6 });
  contentTable(slide, 0.52, 1.96, 6.10, 1.45, 'Top five - Notas más leídas (uu) · Argentina', arg.top_pull_notes, { headerColor: COLORS.sky, fontSize: 5.1, titleMax: 48 });
  contentTable(slide, 6.85, 1.96, 6.10, 1.45, 'Top five - Notas más leídas (TGM) · Argentina', arg.top_pull_notes_tgm || arg.top_pull_notes, { headerColor: COLORS.yellow, fontSize: 5.1, titleMax: 48 });
  contentTable(slide, 0.52, 3.66, 6.10, 1.45, 'Top five - Notas más leídas (uu) · Holding', hol.top_pull_notes, { headerColor: COLORS.sky, fontSize: 5.1, titleMax: 48 });
  contentTable(slide, 6.85, 3.66, 6.10, 1.45, 'Top five - Notas más leídas (TGM) · Holding', hol.top_pull_notes_tgm || hol.top_pull_notes, { headerColor: COLORS.yellow, fontSize: 5.1, titleMax: 48 });
  observationBox(slide, 0.52, 5.45, 12.43, 0.86);
}
function renderContentCombined() {
  const c = scopeData('combined');
  const slide = slideBase(`Gestión CI - ${periodLabel()}`, 'Canal Intranet / Contenidos | Argentina + Holding');
  kpiCard(slide, 2.20, 1.10, 2.20, 0.66, 'Noticias Publicadas', fmtNum(c.site_notes_total), { fill: COLORS.blue, valueSize: 17 });
  kpiCard(slide, 5.40, 1.10, 2.45, 0.66, 'Total Páginas Vistas', fmtNum(c.site_total_views), { fill: COLORS.blue, valueSize: 17 });
  kpiCard(slide, 8.85, 1.10, 2.20, 0.66, 'Promedio Vistas', fmtNum(c.site_average_views), { fill: COLORS.blue, valueSize: 17 });
  contentTable(slide, 0.80, 2.12, 5.60, 2.05, 'Top five - Notas más leídas (uu)', c.top_pull_notes, { headerColor: COLORS.sky, titleMax: 52 });
  contentTable(slide, 6.70, 2.12, 5.60, 2.05, 'Top five - Notas más leídas (TGM)', c.top_pull_notes_tgm || c.top_pull_notes, { headerColor: COLORS.yellow, titleMax: 52 });
  observationBox(slide, 0.80, 4.60, 11.50, 1.24);
}
function renderClosing() {
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.blue };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: COLORS.blue }, line: { color: COLORS.blue } });
  addLogo(slide, 'white', 5.72, 2.35, 1.90, 0.62);
  slide.addText('Comunicaciones Internas', { x: 2.4, y: 3.34, w: 8.55, h: 0.42, fontFace: 'Georgia', fontSize: 24, bold: true, color: COLORS.white, align: 'center', margin: 0, fit: 'shrink' });
  slide.addText(`Gestión ${periodLabel()}`, { x: 2.9, y: 3.93, w: 7.55, h: 0.24, fontFace: 'Arial', fontSize: 12.5, color: COLORS.sky, align: 'center', margin: 0, fit: 'shrink' });
}

try {
  ensureScopes();
  renderCover();
  renderPlanningComparison();
  renderPlanningCombined();
  renderMailComparison();
  renderMailCombined();
  renderContentComparison();
  renderContentCombined();
  renderClosing();
  fs.mkdirSync(path.dirname(outputPptxPath), { recursive: true });
  pptx.writeFile({ fileName: outputPptxPath }).catch((err) => {
    console.error(err);
    process.exit(1);
  });
} catch (err) {
  console.error(err);
  process.exit(1);
}
