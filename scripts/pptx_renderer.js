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
pptx.author = 'BBVA Comunicaciones Internas';
pptx.company = 'BBVA';
pptx.subject = 'Informe de gestión de Comunicaciones Internas';
pptx.lang = 'es-AR';
pptx.theme = { headFontFace: 'Georgia', bodyFontFace: 'Arial', lang: 'es-AR' };
pptx.margin = 0;

const COLORS = {
  electricBlue: '001391',
  midnight: '060E46',
  deep: '85C8FF',
  cyan: '00D7E8',
  aqua: '8BE1E9',
  sky: '85C8FF',
  lime: '88E783',
  yellow: 'FFE761',
  orange: 'FFB56B',
  purple: '9694FF',
  red: 'DA3851',
  white: 'FFFFFF',
  ink: '17233F',
  muted: '56657F',
  grey3: '8A94A6',
  grey2: 'CAD1D8',
  grey1: 'E8ECF2',
  paper: 'F7F8FA',
  paleBlue: 'EEF7FF',
  paleCyan: 'E7FBFC',
  paleLime: 'F0FBF4',
  paleYellow: 'FFF8DA',
  palePurple: 'F4F1FF',
};

const CHART_COLORS = [
  COLORS.electricBlue,
  COLORS.cyan,
  COLORS.sky,
  COLORS.lime,
  COLORS.yellow,
  COLORS.purple,
  COLORS.orange,
  COLORS.red,
];

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
  let raw = String(value ?? '').replace(/_/g, ' ').replace(/\s+/g, ' ').trim();
  raw = raw.replace(/\.\.\./g, '…');
  const replacements = [
    [/Los beneficios de febrero van a llenarte el co(?:…)?$/i, 'Los beneficios de febrero van a llenarte el corazón'],
    [/Queremos escucharte: ayudanos a mejorar la(?:…)?$/i, 'Queremos escucharte: ayudanos a mejorar la comunicación interna'],
    [/Seguimos acompa(?:ñ|n)ando tu desarrollo acad(?:é|e)mi(?:…)?$/i, 'Seguimos acompañando tu desarrollo académico'],
    [/Empez(?:á|a)? el 2026 con estos beneficios\s*AACC(?:…)?$/i, 'Empezá el 2026 con estos beneficios - AACC'],
    [/Empez(?:á|a)? el 2026 con estos beneficios\s*RESTO(?:…)?$/i, 'Empezá el 2026 con estos beneficios - RESTO'],
  ];
  for (const [pattern, replacement] of replacements) raw = raw.replace(pattern, replacement);
  raw = raw.replace(/\s*(?:…|\.\.\.)\s*$/g, '').trim();
  return raw || fallback;
}

function parseNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  if (value === null || value === undefined || value === '-') return 0;
  let text = String(value).trim();
  if (!text) return 0;
  if (text.includes(',') && text.includes('.') && text.lastIndexOf(',') > text.lastIndexOf('.')) {
    text = text.replace(/\./g, '').replace(',', '.');
  } else if (text.includes(',') && !text.includes('.')) {
    text = text.replace(',', '.');
  } else {
    text = text.replace(/,/g, '');
  }
  const cleaned = text.replace(/%/g, '').replace(/[^0-9.-]/g, '');
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
  const text = cleanText(value, '');
  if (!text) return '-';
  if (text.length <= max) return text;
  const cut = text.slice(0, max).rsplit ? text.slice(0, max) : text.slice(0, max);
  const byWord = cut.split(' ').slice(0, -1).join(' ').trim();
  return byWord || cut.trim();
}

function periodLabel() {
  const rp = report?.render_plan?.period;
  if (rp?.label) return cleanText(rp.label);
  if (report?.period?.label) return cleanText(report.period.label);
  if (report?.slide_1_cover?.period) return cleanText(report.slide_1_cover.period);
  return '-';
}

function weightedRows(source, limit = 6) {
  if (!Array.isArray(source)) return [];
  return source
    .map((row) => {
      if (!row || typeof row !== 'object') return null;
      const label = cleanText(row.label || row.theme || row.channel || row.name || row.title, 'Sin dato');
      const value = parseNumber(row.value ?? row.weight ?? row.pct ?? row.count ?? row.total ?? 0);
      return { label, value, raw: row };
    })
    .filter((row) => row && row.label !== '-' && row.value >= 0)
    .sort((a, b) => b.value - a.value)
    .slice(0, limit);
}

function sumValues(rows) {
  return rows.reduce((acc, row) => acc + parseNumber(row.value), 0);
}

function valueAsPct(row, rows) {
  const total = sumValues(rows);
  const value = parseNumber(row.value);
  if (value <= 100 && total >= 95 && total <= 105) return fmtPct(value);
  if (total > 0) return fmtPct((value / total) * 100);
  return fmtNum(value);
}

function valueLabel(row, rows, options = {}) {
  if (options.valueMode === 'number') return fmtNum(row.value, options.digits || 0);
  if (options.valueMode === 'percent') return fmtPct(row.value);
  return valueAsPct(row, rows);
}

function ensureSentence(text) {
  const t = cleanText(text, '').trim();
  if (!t) return '';
  if (/[.!?]$/.test(t)) return t;
  return `${t}.`;
}

function bestPushCampaign() {
  const rankings = report?.kpis?.consolidated_rankings?.top_push_by_interaction || report?.render_plan?.modules?.find((m) => m.key === 'ranking_push')?.payload?.by_interaction || [];
  return Array.isArray(rankings) && rankings.length ? rankings[0] : null;
}

function splitSentences(text, limit = 3) {
  const base = cleanText(text, '');
  if (!base) return [];
  return base.split(/(?<=[.!?])\s+/).map(ensureSentence).filter(Boolean).slice(0, limit);
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
    const last = kept[maxLines - 1].slice(0, Math.max(0, maxLine)).split(' ').slice(0, -1).join(' ').trim();
    kept[maxLines - 1] = last || kept[maxLines - 1].slice(0, Math.max(0, maxLine)).trim();
    return kept.join('\n');
  }
  return lines.join('\n');
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
    'comparación histórica habilitada',
  ];
  return genericMarkers.some((marker) => text.includes(marker));
}

function bulletize(items, fallback) {
  const rows = Array.isArray(items) ? items.filter(Boolean).map(cleanText).filter((text) => text && text !== '-') : [];
  return rows.length ? rows : [fallback || 'Priorizar acciones de mayor impacto y medir el efecto en el siguiente período.'];
}

function buildExecutiveMessage(p) {
  const supplied = p.historical_note || p.headline;
  if (supplied && !isGenericNarrative(supplied) && cleanText(supplied).length > 40) return cleanText(supplied);
  return `La gestión de ${periodLabel()} combinó ${fmtNum(p.plan_total)} comunicaciones planificadas, ${fmtNum(p.mail_total)} envíos de mail y ${fmtNum(p.site_notes_total)} publicaciones en SITE/Intranet. El mailing sostuvo ${fmtPct(p.mail_open_rate)} de apertura y ${fmtPct(p.mail_interaction_rate)} de interacción, mientras el ecosistema pull acumuló ${fmtNum(p.site_total_views)} vistas.`;
}

function buildExecutiveInsights(p) {
  const provided = Array.isArray(p.takeaways) ? p.takeaways.filter(Boolean).map(cleanText).filter((t) => !isGenericNarrative(t)) : [];
  const insights = [...provided];
  if (parseNumber(p.mail_open_rate) > 0) {
    insights.push(`Sostener el canal mail como vehículo de llegada: alcanzó ${fmtPct(p.mail_open_rate)} de apertura promedio.`);
  }
  if (parseNumber(p.mail_interaction_rate) > 0) {
    insights.push(`Replicar formatos con call to action claro para elevar la interacción sobre el ${fmtPct(p.mail_interaction_rate)} actual.`);
  }
  if (parseNumber(p.site_total_views) > 0 && parseNumber(p.site_notes_total) > 0) {
    insights.push(`Usar SITE/Intranet para profundización: promedió ${fmtNum(parseNumber(p.site_total_views) / parseNumber(p.site_notes_total), 0)} vistas por publicación.`);
  }
  return insights.slice(0, 3);
}

function buildChannelNarrative(p) {
  const supplied = p.message;
  if (supplied && !isGenericNarrative(supplied) && cleanText(supplied).length > 50) return cleanText(supplied);
  const rows = weightedRows(p.channel_mix, 5);
  if (!rows.length) return `Mail aportó ${fmtNum(p.mail_total)} envíos, con ${fmtPct(p.mail_open_rate)} de apertura y ${fmtPct(p.mail_interaction_rate)} de interacción. Falta mix de canales para identificar oportunidades de balance.`;
  const top = rows.slice(0, 3).map((row) => `${row.label} (${valueAsPct(row, rows)})`).join(', ');
  return `El mix se concentró en ${top}. Mail sostuvo el alcance directo (${fmtPct(p.mail_open_rate)} de apertura) y SITE/Intranet funcionó como soporte de profundización con ${fmtNum(p.site_total_views)} vistas.`;
}

function buildAxesNarrative(p) {
  const supplied = p.message;
  if (supplied && !isGenericNarrative(supplied) && cleanText(supplied).length > 50) return cleanText(supplied);
  const axes = weightedRows(p.strategic_axes, 5);
  const clients = weightedRows(p.internal_clients, 4);
  if (!axes.length) return 'Falta información suficiente para leer la distribución temática del período.';
  const lead = axes[0];
  const clientText = clients.length ? ` El área solicitante con mayor presencia fue ${clients[0].label} (${fmtPct(clients[0].value)}).` : '';
  return `El eje líder fue ${lead.label} (${valueAsPct(lead, axes)}), seguido por ${axes.slice(1, 3).map((row) => row.label).join(' y ') || 'otros ejes'}.${clientText}`;
}

function buildPushNarrative(p) {
  const supplied = p.message;
  if (supplied && !isGenericNarrative(supplied) && cleanText(supplied).length > 50) return cleanText(supplied);
  const best = Array.isArray(p.by_interaction) ? p.by_interaction[0] : null;
  if (!best) return 'No hay ranking push suficiente para construir una lectura ejecutiva.';
  return `La pieza con mejor desempeño fue “${clip(best.name || best.title, 62)}”: ${fmtNum(best.clicks)} clics y ${fmtPct(best.interaction || best.ctr)} de interacción. Conviene reutilizar su lógica de asunto, beneficio percibido y llamado a la acción.`;
}

function buildPullNarrative(p) {
  const supplied = p.message;
  if (supplied && !isGenericNarrative(supplied) && cleanText(supplied).length > 50) return cleanText(supplied);
  const best = Array.isArray(p.top_pull_notes) ? p.top_pull_notes[0] : null;
  if (!best) return 'No hay ranking pull suficiente para construir una lectura ejecutiva.';
  return `La nota con mayor tracción fue “${clip(best.title || best.name, 64)}”, con ${fmtNum(best.unique_reads || best.users)} lecturas únicas. La lectura pull marca intereses concretos para alimentar próximos envíos segmentados.`;
}

function addLogo(slide, variant = 'blue', x = 11.58, y = 0.22, w = 1.05, h = 0.34) {
  const logo = resolveAsset(variant === 'white' ? BBVA_LOGO_WHITE : BBVA_LOGO_BLUE);
  if (logo) {
    slide.addImage({ path: logo, ...imageSizingContain(logo, x, y, w, h) });
  } else {
    slide.addText('BBVA', { x, y, w, h, align: 'right', bold: true, color: variant === 'white' ? COLORS.white : COLORS.electricBlue, fontFace: 'Arial', fontSize: 16, margin: 0 });
  }
}

function finalizeSlide(slide) {
  if (SHOULD_WARN_LAYOUT) {
    warnIfSlideHasOverlaps(slide);
    warnIfSlideElementsOutOfBounds(slide, pptx);
  }
}

function addSubtleGrid(slide) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: COLORS.paper }, line: { color: COLORS.paper } });
  slide.addShape(pptx.ShapeType.arc, { x: 9.85, y: -1.3, w: 4.2, h: 4.2, adjustPoint: 0.18, line: { color: COLORS.sky, transparency: 78, width: 1.1 } });
  slide.addShape(pptx.ShapeType.arc, { x: -1.3, y: 5.75, w: 3.4, h: 3.4, adjustPoint: 0.22, line: { color: COLORS.cyan, transparency: 80, width: 1.1 } });
}

function baseSlide(title, subtitle = '') {
  const slide = pptx.addSlide();
  addSubtleGrid(slide);
  addLogo(slide, 'blue');
  if (subtitle) {
    slide.addText(`Comunicaciones Internas / ${cleanText(subtitle)}`, {
      x: 0.62, y: 0.22, w: 8.2, h: 0.18,
      fontFace: 'Arial', fontSize: 8.5, color: COLORS.muted, margin: 0,
    });
  }
  slide.addText(cleanText(title, 'Reporte ejecutivo'), {
    x: 0.62, y: 0.48, w: 9.9, h: 0.48,
    fontFace: 'Georgia', bold: true, fontSize: 25, color: COLORS.electricBlue, margin: 0,
    fit: 'shrink', breakLine: false,
  });
  slide.addShape(pptx.ShapeType.line, { x: 0.62, y: 1.11, w: 12.08, h: 0, line: { color: COLORS.grey1, width: 1 } });
  return slide;
}

function panel(slide, x, y, w, h, header = '', options = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: options.fill || COLORS.white, transparency: options.transparency || 0 },
    line: { color: options.line || COLORS.grey1, width: options.lineWidth || 0.75, transparency: options.lineTransparency || 0 },
    shadow: options.shadow === false ? undefined : safeOuterShadow('000000', 0.08, 45, 0.7, 0.25),
  });
  if (header) {
    slide.addText(cleanText(header), {
      x: x + 0.18, y: y + 0.15, w: w - 0.36, h: 0.18,
      fontFace: 'Arial', fontSize: 8.8, bold: true, color: options.headerColor || COLORS.electricBlue,
      margin: 0, fit: 'shrink', breakLine: false,
    });
  }
}

function metricTile(slide, x, y, w, h, label, value, accent = COLORS.sky, options = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: options.fill || COLORS.white },
    line: { color: options.line || COLORS.grey1, width: 0.75 },
  });
  slide.addShape(pptx.ShapeType.rect, { x, y, w: 0.08, h, fill: { color: accent }, line: { color: accent } });
  slide.addText(cleanText(label), {
    x: x + 0.18, y: y + 0.12, w: w - 0.3, h: 0.14,
    fontFace: 'Arial', fontSize: options.labelSize || 7.8, color: COLORS.muted, margin: 0, fit: 'shrink', breakLine: false,
  });
  slide.addText(String(value ?? '-'), {
    x: x + 0.18, y: y + 0.31, w: w - 0.3, h: h - 0.36,
    fontFace: 'Georgia', bold: true, fontSize: options.valueSize || 19,
    color: COLORS.electricBlue, margin: 0, fit: 'shrink', breakLine: false,
  });
}

function paragraph(slide, x, y, w, h, body, options = {}) {
  slide.addText(clip(body, options.max || 330), {
    x, y, w, h,
    fontFace: options.fontFace || 'Arial', fontSize: options.fontSize || 10.2,
    bold: !!options.bold, color: options.color || COLORS.ink,
    valign: options.valign || 'top', margin: 0,
    fit: 'shrink', breakLine: false,
  });
}

function bulletList(slide, x, y, w, items, options = {}) {
  const rows = bulletize(items, options.fallback).slice(0, options.limit || 3);
  rows.forEach((item, idx) => {
    const cy = y + idx * (options.step || 0.53);
    slide.addShape(pptx.ShapeType.ellipse, { x, y: cy + 0.035, w: 0.11, h: 0.11, fill: { color: options.bulletColor || COLORS.cyan }, line: { color: options.bulletColor || COLORS.cyan } });
    slide.addText(clip(ensureSentence(item), options.max || 170), {
      x: x + 0.22, y: cy, w: w - 0.22, h: options.itemH || 0.38,
      fontFace: 'Arial', fontSize: options.fontSize || 9.4,
      color: COLORS.ink, margin: 0, fit: 'shrink', breakLine: false,
    });
  });
}

function renderProgressBar(slide, x, y, w, label, valueText, pct, color = COLORS.cyan) {
  const p = Math.max(0, Math.min(1, pct));
  slide.addText(cleanText(label), { x, y, w: w * 0.62, h: 0.15, fontFace: 'Arial', fontSize: 7.7, color: COLORS.ink, margin: 0, fit: 'shrink' });
  slide.addText(cleanText(valueText), { x: x + w * 0.62, y, w: w * 0.38, h: 0.15, align: 'right', fontFace: 'Arial', bold: true, fontSize: 7.7, color: COLORS.electricBlue, margin: 0, fit: 'shrink' });
  slide.addShape(pptx.ShapeType.roundRect, { x, y: y + 0.2, w, h: 0.08, rectRadius: 0.03, fill: { color: COLORS.grey1 }, line: { color: COLORS.grey1 } });
  if (p > 0) slide.addShape(pptx.ShapeType.roundRect, { x, y: y + 0.2, w: w * p, h: 0.08, rectRadius: 0.03, fill: { color }, line: { color } });
}

function renderHorizontalBarChart(slide, x, y, w, h, rows, options = {}) {
  const data = rows.filter((row) => parseNumber(row.value) >= 0).slice(0, options.limit || 6);
  if (!data.length) {
    paragraph(slide, x, y + h / 2 - 0.1, w, 0.24, options.empty || 'Sin datos para graficar.', { fontSize: 9, color: COLORS.muted });
    return;
  }
  const max = Math.max(...data.map((row) => parseNumber(row.value)), 1);
  const rowH = h / data.length;
  data.forEach((row, idx) => {
    const barY = y + idx * rowH;
    const color = options.colors ? options.colors[idx % options.colors.length] : CHART_COLORS[idx % CHART_COLORS.length];
    const label = clip(row.label, options.labelMax || 24);
    const val = valueLabel(row, data, options);
    slide.addText(label, { x, y: barY, w: w * 0.38, h: 0.18, fontFace: 'Arial', fontSize: options.catSize || 7.3, color: COLORS.ink, margin: 0, fit: 'shrink' });
    slide.addShape(pptx.ShapeType.roundRect, { x: x + w * 0.40, y: barY + 0.025, w: w * 0.47, h: 0.13, rectRadius: 0.025, fill: { color: COLORS.grey1 }, line: { color: COLORS.grey1 } });
    const barW = (w * 0.47) * (parseNumber(row.value) / max);
    slide.addShape(pptx.ShapeType.roundRect, { x: x + w * 0.40, y: barY + 0.025, w: Math.max(0.03, barW), h: 0.13, rectRadius: 0.025, fill: { color }, line: { color } });
    slide.addText(val, { x: x + w * 0.89, y: barY - 0.01, w: w * 0.11, h: 0.17, align: 'right', fontFace: 'Arial', bold: true, fontSize: options.dataSize || 7.3, color: COLORS.electricBlue, margin: 0, fit: 'shrink' });
  });
}

function renderVerticalBarChart(slide, x, y, w, h, rows, options = {}) {
  const data = rows.filter((row) => parseNumber(row.value) >= 0).slice(0, options.limit || 5);
  if (!data.length) {
    paragraph(slide, x, y + h / 2 - 0.1, w, 0.24, options.empty || 'Sin datos para graficar.', { fontSize: 9, color: COLORS.muted });
    return;
  }
  const max = Math.max(...data.map((row) => parseNumber(row.value)), 1);
  const slot = w / data.length;
  const chartH = h - 0.54;
  data.forEach((row, idx) => {
    const val = parseNumber(row.value);
    const bh = Math.max(0.04, chartH * val / max);
    const bx = x + idx * slot + slot * 0.22;
    const bw = slot * 0.56;
    const by = y + chartH - bh;
    const color = CHART_COLORS[idx % CHART_COLORS.length];
    slide.addShape(pptx.ShapeType.roundRect, { x: bx, y: by, w: bw, h: bh, rectRadius: 0.035, fill: { color }, line: { color } });
    slide.addText(valueLabel(row, data, options), { x: bx - 0.08, y: by - 0.22, w: bw + 0.16, h: 0.15, fontFace: 'Arial', bold: true, fontSize: 7.4, color: COLORS.electricBlue, align: 'center', margin: 0, fit: 'shrink' });
    slide.addText(wrapText(row.label, options.labelMax || 12, 2), { x: x + idx * slot + 0.04, y: y + chartH + 0.1, w: slot - 0.08, h: 0.38, fontFace: 'Arial', fontSize: 6.8, color: COLORS.ink, align: 'center', margin: 0, fit: 'shrink' });
  });
}

function tableRows(slide, x, y, w, headers, rows, options = {}) {
  const widths = options.widths || headers.map(() => 1 / headers.length);
  const total = widths.reduce((a, b) => a + b, 0);
  const norm = widths.map((n) => (n / total) * w);
  let cx = x;
  headers.forEach((header, idx) => {
    slide.addText(cleanText(header), { x: cx, y, w: norm[idx] - 0.06, h: 0.18, fontFace: 'Arial', fontSize: options.headerSize || 7.6, bold: true, color: COLORS.muted, align: idx === 0 ? 'left' : 'right', margin: 0, fit: 'shrink' });
    cx += norm[idx];
  });
  const rowH = options.rowHeight || 0.43;
  rows.slice(0, options.maxRows || 5).forEach((row, rowIndex) => {
    const rowY = y + 0.3 + rowIndex * rowH;
    if (rowIndex % 2 === 0) slide.addShape(pptx.ShapeType.roundRect, { x, y: rowY - 0.055, w, h: rowH - 0.055, rectRadius: 0.03, fill: { color: 'F9FBFD' }, line: { color: 'F9FBFD' } });
    let xCursor = x;
    row.forEach((value, colIndex) => {
      const isFirst = colIndex === 0;
      const cellText = isFirst ? wrapText(value, options.wrapMax || 34, options.wrapLines || 2) : clip(value, options.clip?.[colIndex] || 18);
      slide.addText(cellText, { x: xCursor + 0.03, y: rowY - 0.01, w: norm[colIndex] - 0.08, h: rowH - 0.11, fontFace: 'Arial', bold: !isFirst, fontSize: isFirst ? (options.firstColSize || 7.4) : (options.valueSize || 8.1), color: isFirst ? COLORS.ink : COLORS.electricBlue, align: isFirst ? 'left' : 'right', margin: 0, fit: 'shrink' });
      xCursor += norm[colIndex];
    });
  });
}

function emptyState(slide, x, y, w, h, title, body) {
  panel(slide, x, y, w, h, title || 'Sin datos suficientes', { fill: COLORS.paleYellow, shadow: false });
  paragraph(slide, x + 0.24, y + 0.58, w - 0.48, h - 0.76, body || 'Completar con contexto manual si el equipo quiere destacar acciones cualitativas.', { max: 180, fontSize: 9.6, color: COLORS.ink });
}

function renderDonutChart(slide, x, y, w, h, rows, title = 'Mix') {
  const data = rows.filter((row) => parseNumber(row.value) > 0).slice(0, 6);
  if (!data.length) {
    paragraph(slide, x, y + h / 2 - 0.1, w, 0.24, 'Sin datos para graficar.', { fontSize: 9, color: COLORS.muted });
    return;
  }
  const colors = data.map((_, idx) => (idx === 0 ? COLORS.electricBlue : ['D7DDE6', 'B7C1CE', 'E8ECF2', 'CAD1D8', 'AAB6C5'][idx - 1] || COLORS.grey2));
  try {
    slide.addChart(pptx.ChartType.doughnut, [{ name: title, labels: data.map((r) => r.label), values: data.map((r) => parseNumber(r.value)) }], {
      x, y, w, h,
      holeSize: 65,
      showLegend: true,
      legendPos: 'r',
      showValue: false,
      showCategoryName: false,
      showPercent: true,
      dataLabelPosition: 'bestFit',
      firstSliceAng: 270,
      ser: [{ dataLabelPosition: 'bestFit' }],
      chartColors: colors,
      showTitle: false,
      showLeaderLines: true,
    });
  } catch (err) {
    renderHorizontalBarChart(slide, x, y + 0.1, w, h - 0.2, data, { valueMode: 'percent', limit: 5 });
  }
}

function insightBulletsFromText(text, fallbackItems = []) {
  const fromText = splitSentences(text, 3);
  return fromText.length ? fromText : fallbackItems.map(ensureSentence).filter(Boolean).slice(0, 3);
}

function renderFullCover() {
  const s = report?.period || report?.slide_1_cover || {};
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.electricBlue };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: COLORS.electricBlue }, line: { color: COLORS.electricBlue } });
  slide.addShape(pptx.ShapeType.arc, { x: 6.8, y: -0.4, w: 7.8, h: 8.1, adjustPoint: 0.18, line: { color: COLORS.cyan, transparency: 35, width: 3.2 } });
  slide.addShape(pptx.ShapeType.arc, { x: 7.6, y: 0.35, w: 5.8, h: 6.0, adjustPoint: 0.28, line: { color: COLORS.sky, transparency: 50, width: 1.8 } });
  addLogo(slide, 'white', 0.62, 0.36, 1.45, 0.52);
  slide.addText(cleanText(s.label || s.period || periodLabel()), { x: 0.68, y: 3.34, w: 5.8, h: 0.42, fontFace: 'Arial', fontSize: 15, color: COLORS.white, margin: 0 });
  slide.addShape(pptx.ShapeType.line, { x: 0.68, y: 3.98, w: 11.5, h: 0, line: { color: COLORS.white, width: 0.5, transparency: 50 } });
  slide.addText('Comunicaciones\nInternas', { x: 0.68, y: 4.22, w: 8.8, h: 1.55, fontFace: 'Georgia', bold: true, fontSize: 49, color: COLORS.white, margin: 0, breakLine: true, fit: 'shrink' });
  slide.addText('Informe de gestión mensual', { x: 0.72, y: 6.28, w: 5.3, h: 0.26, fontFace: 'Arial', fontSize: 10.5, color: COLORS.sky, margin: 0 });
  finalizeSlide(slide);
}

function renderExecutiveSummary(module) {
  const p = module.payload || {};
  const slide = baseSlide(module.title || 'Resumen ejecutivo', 'Resumen');
  const best = bestPushCampaign() || { name: p.top_campaign_title, interaction: p.top_campaign_interaction, open_rate: p.top_campaign_open_rate };
  const headline = cleanText(p.executive_headline || p.headline || `${periodLabel()}: desempeño mensual de Comunicaciones Internas`);
  const topInteraction = parseNumber(p.top_campaign_interaction || best?.interaction || best?.ctr || 0);

  slide.addText(clip(headline, 118), {
    x: 0.72, y: 1.28, w: 11.0, h: 0.48,
    fontFace: 'Georgia', bold: true, fontSize: 22, color: COLORS.midnight, margin: 0, fit: 'shrink', breakLine: false,
  });

  metricTile(slide, 0.72, 2.02, 3.25, 1.02, 'Apertura promedio', fmtPct(p.mail_open_rate), COLORS.lime, { valueSize: 27, labelSize: 8.8 });
  metricTile(slide, 4.20, 2.02, 3.25, 1.02, 'Interacción promedio', fmtPct(p.mail_interaction_rate), COLORS.cyan, { valueSize: 27, labelSize: 8.8 });
  metricTile(slide, 7.68, 2.02, 3.25, 1.02, 'Top campaña', fmtPct(topInteraction), COLORS.yellow, { valueSize: 27, labelSize: 8.8 });

  slide.addText(clip(best?.name || p.top_campaign_title || 'Campaña con mejor desempeño', 64), {
    x: 7.88, y: 3.14, w: 2.85, h: 0.28,
    fontFace: 'Arial', fontSize: 8.2, bold: true, color: COLORS.ink, margin: 0, fit: 'shrink', align: 'center',
  });

  panel(slide, 0.72, 3.74, 7.15, 2.25, 'Insights del período', { fill: COLORS.white, shadow: false });
  bulletList(slide, 1.02, 4.24, 6.55, buildExecutiveInsights(p), { max: 180, step: 0.58, fontSize: 9.6, itemH: 0.42, bulletColor: COLORS.electricBlue });

  panel(slide, 8.22, 3.74, 4.12, 2.25, 'Volumen gestionado', { fill: COLORS.paleBlue, shadow: false });
  metricTile(slide, 8.52, 4.20, 1.55, 0.70, 'Planificación', fmtNum(p.plan_total), COLORS.sky, { valueSize: 17, labelSize: 7.3, fill: COLORS.white });
  metricTile(slide, 10.24, 4.20, 1.55, 0.70, 'Mails', fmtNum(p.mail_total), COLORS.cyan, { valueSize: 17, labelSize: 7.3, fill: COLORS.white });
  metricTile(slide, 8.52, 5.08, 1.55, 0.70, 'Notas SITE', fmtNum(p.site_notes_total), COLORS.purple, { valueSize: 17, labelSize: 7.3, fill: COLORS.white });
  metricTile(slide, 10.24, 5.08, 1.55, 0.70, 'Vistas SITE', fmtNum(p.site_total_views), COLORS.orange, { valueSize: 17, labelSize: 7.3, fill: COLORS.white });

  slide.addText(`Período: ${periodLabel()} · fuente: dashboard mensual consolidado`, { x: 0.72, y: 6.72, w: 7.8, h: 0.18, fontFace: 'Arial', fontSize: 7.8, color: COLORS.muted, margin: 0 });
  finalizeSlide(slide);
}


function renderChannelManagement(module) {
  const p = module.payload || {};
  const slide = baseSlide(module.title || 'Gestión de canales', 'Canales');
  const mixRows = weightedRows(p.channel_mix, 5);
  const lead = mixRows[0];
  const headline = lead ? `${lead.label} lidera el mix con ${valueAsPct(lead, mixRows)} y sostiene el alcance directo.` : 'El mix de canales consolida el alcance del período.';

  slide.addText(clip(headline, 112), { x: 0.72, y: 1.28, w: 10.8, h: 0.36, fontFace: 'Georgia', bold: true, fontSize: 19, color: COLORS.midnight, margin: 0, fit: 'shrink' });

  panel(slide, 0.72, 1.90, 6.95, 4.15, 'Mix de canales');
  renderDonutChart(slide, 1.04, 2.34, 6.15, 3.15, mixRows, 'Mix de canales');
  if (lead) slide.addText(`Canal principal: ${lead.label} (${valueAsPct(lead, mixRows)})`, { x: 1.12, y: 5.60, w: 5.7, h: 0.18, align: 'center', fontFace: 'Arial', bold: true, fontSize: 8.2, color: COLORS.electricBlue, margin: 0 });

  metricTile(slide, 8.05, 1.90, 1.90, 0.82, 'Mails enviados', fmtNum(p.mail_total), COLORS.cyan, { valueSize: 19 });
  metricTile(slide, 10.20, 1.90, 1.90, 0.82, 'Apertura', fmtPct(p.mail_open_rate), COLORS.lime, { valueSize: 18 });
  metricTile(slide, 8.05, 2.94, 1.90, 0.82, 'Interacción', fmtPct(p.mail_interaction_rate), COLORS.yellow, { valueSize: 18 });
  metricTile(slide, 10.20, 2.94, 1.90, 0.82, 'Vistas SITE', fmtNum(p.site_total_views), COLORS.sky, { valueSize: 18 });

  panel(slide, 8.05, 4.12, 4.05, 1.93, 'Lectura ejecutiva', { fill: COLORS.paleCyan, shadow: false });
  const bullets = [
    lead ? `${lead.label} concentra la mayor participación del mix de canales.` : 'El mix de canales requiere lectura mensual para ajustar presión comunicacional.',
    `Mail combina alcance con ${fmtPct(p.mail_open_rate)} de apertura promedio.`,
    `SITE/Intranet aporta profundidad con ${fmtNum(p.site_total_views)} vistas acumuladas.`,
  ];
  bulletList(slide, 8.34, 4.62, 3.45, bullets, { max: 108, step: 0.44, fontSize: 8.3, itemH: 0.34, bulletColor: COLORS.cyan });
  finalizeSlide(slide);
}


function renderMix(module) {
  const p = module.payload || {};
  const slide = baseSlide(module.title || 'Mix temático y áreas solicitantes', 'Contenido');
  const axes = weightedRows(p.strategic_axes, 6);
  const clients = weightedRows(p.internal_clients, 7);
  const formats = weightedRows(p.format_mix, 4);
  const leadAxis = axes[0];
  const leadClient = clients[0];
  const headline = leadAxis ? `${leadAxis.label} lidera la agenda; ${leadClient ? `${leadClient.label} concentra la demanda.` : 'faltan áreas solicitantes.'}` : 'Distribución temática del período.';

  slide.addText(clip(headline, 112), { x: 0.72, y: 1.28, w: 10.8, h: 0.36, fontFace: 'Georgia', bold: true, fontSize: 19, color: COLORS.midnight, margin: 0, fit: 'shrink' });

  panel(slide, 0.72, 1.90, 6.15, 3.82, 'Ejes estratégicos');
  renderHorizontalBarChart(slide, 1.05, 2.42, 5.52, 2.45, axes, { labelMax: 23, catSize: 8.0, dataSize: 8.0, valueMode: 'percent', colors: [COLORS.electricBlue, COLORS.sky, COLORS.grey2, COLORS.grey2, COLORS.grey2, COLORS.grey2] });

  panel(slide, 7.18, 1.90, 5.15, 3.82, 'Áreas solicitantes');
  if (clients.length) {
    renderHorizontalBarChart(slide, 7.50, 2.42, 4.52, 2.45, clients, { labelMax: 21, catSize: 7.6, dataSize: 7.6, valueMode: 'percent', colors: [COLORS.electricBlue, COLORS.sky, COLORS.grey2, COLORS.grey2, COLORS.grey2, COLORS.grey2, COLORS.grey2] });
  } else {
    emptyState(slide, 7.48, 2.30, 4.54, 2.60, 'Dato pendiente', 'No se detectó el bloque de áreas solicitantes en la página de planificación.');
  }

  panel(slide, 0.72, 6.00, 11.61, 0.74, 'Síntesis', { fill: COLORS.midnight, line: COLORS.midnight, shadow: false });
  const formatText = formats[0] ? `Formato dominante: ${formats[0].label} (${valueAsPct(formats[0], formats)}).` : 'Formato dominante pendiente.';
  const summary = `${leadAxis ? `Eje líder: ${leadAxis.label} (${valueAsPct(leadAxis, axes)}).` : ''} ${leadClient ? `Mayor solicitante: ${leadClient.label} (${fmtPct(leadClient.value)}).` : ''} ${formatText}`;
  slide.addText(summary.trim(), { x: 1.00, y: 6.25, w: 11.05, h: 0.18, fontFace: 'Arial', bold: true, fontSize: 8.6, color: COLORS.white, margin: 0, fit: 'shrink' });
  finalizeSlide(slide);
}


function renderPushRanking(module) {
  const p = module.payload || {};
  const slide = baseSlide(module.title || 'Ranking push', 'Mail');
  const byInteraction = Array.isArray(p.by_interaction) ? p.by_interaction.slice(0, 5) : [];
  if (p.available === false && !byInteraction.length) {
    emptyState(slide, 0.86, 1.55, 11.6, 4.8, 'Ranking no disponible', cleanText(p.message || 'No se detectó ranking de mails en la fuente.'));
    finalizeSlide(slide);
    return;
  }
  const best = byInteraction[0] || {};
  slide.addText(clip(`La campaña líder alcanzó ${fmtPct(best.interaction || best.ctr)} de interacción.`, 112), { x: 0.72, y: 1.28, w: 10.8, h: 0.36, fontFace: 'Georgia', bold: true, fontSize: 19, color: COLORS.midnight, margin: 0, fit: 'shrink' });

  byInteraction.slice(0, 3).forEach((row, idx) => {
    const x = 0.72 + idx * 4.02;
    const accent = idx === 0 ? COLORS.electricBlue : (idx === 1 ? COLORS.sky : COLORS.grey2);
    panel(slide, x, 1.92, 3.68, 1.64, `Top ${idx + 1}`, { fill: idx === 0 ? COLORS.paleBlue : COLORS.white, shadow: false });
    slide.addText(clip(row.name || row.title || '-', 64), { x: x + 0.25, y: 2.28, w: 3.18, h: 0.36, fontFace: 'Arial', bold: true, fontSize: 8.5, color: COLORS.ink, margin: 0, fit: 'shrink' });
    slide.addText(fmtPct(row.interaction || row.ctr), { x: x + 0.25, y: 2.78, w: 1.45, h: 0.34, fontFace: 'Georgia', bold: true, fontSize: 21, color: COLORS.electricBlue, margin: 0, fit: 'shrink' });
    slide.addText(`${fmtNum(row.clicks)} clics`, { x: x + 1.95, y: 2.88, w: 1.35, h: 0.18, fontFace: 'Arial', bold: true, fontSize: 8.5, color: COLORS.muted, margin: 0, fit: 'shrink', align: 'right' });
    slide.addShape(pptx.ShapeType.rect, { x, y: 1.92, w: 3.68, h: 0.07, fill: { color: accent }, line: { color: accent } });
  });

  panel(slide, 0.72, 3.92, 7.05, 2.25, 'Interacción por pieza');
  renderHorizontalBarChart(slide, 1.02, 4.43, 6.40, 1.32, byInteraction.map((row) => ({ label: row.name || row.title, value: parseNumber(row.interaction || row.ctr) })), { valueMode: 'percent', limit: 5, labelMax: 28, catSize: 7.2, dataSize: 7.2, colors: [COLORS.electricBlue, COLORS.sky, COLORS.grey2, COLORS.grey2, COLORS.grey2] });

  panel(slide, 8.05, 3.92, 4.28, 2.25, 'So what', { fill: COLORS.paleCyan, shadow: false });
  const bullets = [
    `${clip(best.name || best.title || 'La pieza líder', 55)} combinó beneficio claro y alta respuesta.`,
    `Usar ${fmtPct(p.by_open_rate?.[0]?.open_rate || best.open_rate)} de apertura como referencia aspiracional.`,
    'Replicar llamados a la acción concretos en beneficios y servicios.',
  ];
  bulletList(slide, 8.34, 4.42, 3.58, bullets, { max: 110, step: 0.48, fontSize: 8.4, itemH: 0.36, bulletColor: COLORS.electricBlue });
  finalizeSlide(slide);
}


function renderPullRanking(module) {
  const p = module.payload || {};
  const slide = baseSlide(module.title || 'Ranking pull', 'SITE / Intranet');
  const rows = Array.isArray(p.top_pull_notes) ? p.top_pull_notes.slice(0, 5) : [];
  const best = rows[0] || {};

  slide.addText(clip(best.title ? `Bienestar y servicios impulsan las lecturas del ecosistema pull.` : 'Ranking pull del período.', 112), { x: 0.72, y: 1.28, w: 10.8, h: 0.36, fontFace: 'Georgia', bold: true, fontSize: 19, color: COLORS.midnight, margin: 0, fit: 'shrink' });

  metricTile(slide, 0.72, 1.92, 2.45, 0.88, 'Promedio lecturas/nota', fmtNum(p.average_reads_per_note), COLORS.cyan, { valueSize: 21 });
  metricTile(slide, 3.42, 1.92, 2.45, 0.88, 'Vistas totales SITE', fmtNum(p.site_total_views), COLORS.sky, { valueSize: 21 });

  if (p.available === false && !rows.length) {
    emptyState(slide, 0.72, 3.08, 11.62, 3.0, 'Ranking no disponible', cleanText(p.message || 'No se detectó ranking de notas en la fuente.'));
    finalizeSlide(slide);
    return;
  }

  panel(slide, 0.72, 3.14, 7.25, 3.05, 'Top 5 notas por vistas');
  renderHorizontalBarChart(slide, 1.04, 3.66, 6.58, 1.94, rows.map((row) => ({ label: row.title || row.name, value: parseNumber(row.total_reads || row.views) })), { valueMode: 'number', labelMax: 34, catSize: 7.1, dataSize: 7.1, colors: [COLORS.electricBlue, COLORS.sky, COLORS.grey2, COLORS.grey2, COLORS.grey2] });

  panel(slide, 8.28, 1.92, 4.05, 4.27, 'Lectura ejecutiva', { fill: COLORS.paleBlue, shadow: false });
  slide.addText(clip(best.title || 'Nota líder', 72), { x: 8.58, y: 2.42, w: 3.45, h: 0.56, fontFace: 'Georgia', bold: true, fontSize: 14, color: COLORS.electricBlue, margin: 0, fit: 'shrink' });
  slide.addText(`${fmtNum(best.unique_reads || best.users)} usuarios únicos · ${fmtNum(best.total_reads || best.views)} vistas`, { x: 8.58, y: 3.14, w: 3.45, h: 0.20, fontFace: 'Arial', bold: true, fontSize: 8.5, color: COLORS.ink, margin: 0, fit: 'shrink' });
  const bullets = [
    'Los temas de bienestar funcionan como puerta de entrada al SITE.',
    'El top pull debe alimentar próximos envíos segmentados.',
    `La media de ${fmtNum(p.average_reads_per_note)} vistas permite detectar contenidos sobreperformers.`,
  ];
  bulletList(slide, 8.58, 3.74, 3.35, bullets, { max: 104, step: 0.50, fontSize: 8.4, itemH: 0.36, bulletColor: COLORS.sky });
  finalizeSlide(slide);
}


function renderMilestones(module) {
  const p = module.payload || {};
  const items = Array.isArray(p.items) ? p.items.slice(0, 3) : [];
  const slide = baseSlide(module.title || 'Hitos del mes', 'Gestión');

  if (!items.length) {
    panel(slide, 0.82, 1.55, 11.68, 4.82, 'Cierre de gestión');
    slide.addText('Sin hitos consolidados para este período', { x: 1.18, y: 2.24, w: 6.8, h: 0.42, fontFace: 'Georgia', bold: true, fontSize: 24, color: COLORS.electricBlue, margin: 0, fit: 'shrink' });
    paragraph(slide, 1.18, 3.03, 6.1, 1.1, cleanText(p.message || 'No se registraron hitos destacados en la fuente. Puede completarse desde manual_context si el equipo quiere incorporar acciones cualitativas del período.'), { max: 230, fontSize: 10.2 });
    emptyState(slide, 8.08, 2.16, 3.58, 2.55, 'Observación', 'Esta slide queda como recordatorio editorial: sumar campañas, coberturas, eventos o piezas clave del mes.');
    finalizeSlide(slide);
    return;
  }

  items.forEach((item, idx) => {
    const x = 0.72 + idx * 4.02;
    const accent = CHART_COLORS[idx % CHART_COLORS.length];
    panel(slide, x, 1.45, 3.56, 4.78, `Hito ${idx + 1}`);
    slide.addShape(pptx.ShapeType.rect, { x, y: 1.45, w: 3.56, h: 0.08, fill: { color: accent }, line: { color: accent } });
    slide.addText(clip(item.title || item.description || '-', 54), { x: x + 0.22, y: 1.96, w: 3.12, h: 0.58, fontFace: 'Georgia', bold: true, fontSize: 15, color: COLORS.electricBlue, margin: 0, fit: 'shrink' });
    const bullets = Array.isArray(item.bullets) ? item.bullets.filter(Boolean).slice(0, 3) : [];
    bulletList(slide, x + 0.24, 2.82, 3.05, bullets.length ? bullets : [item.description || 'Sin detalle adicional.'], { max: 72, limit: 3, fontSize: 8.8, step: 0.48, bulletColor: accent });
    if (item.period) slide.addText(cleanText(item.period), { x: x + 0.24, y: 5.74, w: 2.9, h: 0.14, fontFace: 'Arial', fontSize: 7.6, color: COLORS.muted, margin: 0 });
  });
  finalizeSlide(slide);
}

function renderRecommendations(module) {
  const p = module.payload || {};
  const recs = bulletize(p.recommendations || p.items, 'Sostener formatos de alto rendimiento y testear mejoras de segmentación.').slice(0, 4);
  const experiments = bulletize(p.experiments, 'Probar asunto A/B en mails de beneficios y medir lift de apertura.').slice(0, 3);
  const plan = bulletize(p.action_plan || p.plan, 'Definir foco editorial, ajustar segmentación y revisar resultados en el próximo cierre.').slice(0, 4);
  const slide = baseSlide(module.title || 'Conclusiones y próximos pasos', 'Plan de mejora');

  slide.addText('Qué haría el equipo el mes que viene', { x: 0.72, y: 1.34, w: 8.5, h: 0.34, fontFace: 'Georgia', bold: true, fontSize: 21, color: COLORS.midnight, margin: 0, fit: 'shrink' });
  paragraph(slide, 0.74, 1.82, 7.45, 0.52, cleanText(p.summary || p.message || 'Plan accionable construido a partir de performance de canales, rankings y distribución temática.'), { max: 220, fontSize: 9.8 });

  panel(slide, 0.72, 2.58, 3.8, 3.38, 'Recomendaciones');
  bulletList(slide, 1.0, 3.06, 3.18, recs, { limit: 4, step: 0.55, max: 82, fontSize: 8.6, bulletColor: COLORS.cyan });

  panel(slide, 4.82, 2.58, 3.8, 3.38, 'Experimentos');
  bulletList(slide, 5.1, 3.06, 3.18, experiments, { limit: 3, step: 0.64, max: 82, fontSize: 8.6, bulletColor: COLORS.lime });

  panel(slide, 8.92, 2.58, 3.42, 3.38, 'Plan 30 días', { fill: COLORS.paleBlue });
  bulletList(slide, 9.2, 3.06, 2.84, plan, { limit: 4, step: 0.55, max: 70, fontSize: 8.4, bulletColor: COLORS.electricBlue });

  slide.addShape(pptx.ShapeType.roundRect, { x: 0.72, y: 6.32, w: 11.62, h: 0.38, rectRadius: 0.06, fill: { color: COLORS.midnight }, line: { color: COLORS.midnight } });
  slide.addText(cleanText(p.owner_note || 'Seguimiento sugerido: comparar apertura, interacción y vistas por nota contra el período anterior, evitando cambios de criterio de fuente.'), { x: 0.98, y: 6.44, w: 11.06, h: 0.12, fontFace: 'Arial', fontSize: 7.8, bold: true, color: COLORS.white, margin: 0, fit: 'shrink' });
  finalizeSlide(slide);
}

function renderEvents(module) {
  const p = module.payload || {};
  const events = Array.isArray(p.events) ? p.events.slice(0, 6) : [];
  const slide = baseSlide(module.title || 'Eventos del mes', 'Activaciones');

  metricTile(slide, 0.72, 1.36, 2.25, 0.78, 'Eventos', fmtNum(p.total_events || events.length), COLORS.cyan, { valueSize: 18 });
  metricTile(slide, 3.18, 1.36, 2.45, 0.78, 'Participaciones', fmtNum(p.total_participants), COLORS.sky, { valueSize: 18 });

  if (!events.length) {
    emptyState(slide, 0.72, 2.44, 11.62, 3.9, 'Eventos omitidos', 'No hay detalle suficiente de eventos para graficar este módulo.');
    finalizeSlide(slide);
    return;
  }

  panel(slide, 0.72, 2.42, 7.08, 3.82, 'Detalle de eventos');
  tableRows(slide, 1.0, 2.92, 6.48, ['Evento', 'Participantes', 'Fecha'], events.map((row) => [row.name || row.title || '-', fmtNum(row.participants), row.date || '-']), { widths: [0.58, 0.22, 0.20], rowHeight: 0.44, maxRows: 6, wrapMax: 42, firstColSize: 7.4, valueSize: 8.0 });

  panel(slide, 8.12, 2.42, 4.22, 3.82, 'Participación por evento');
  renderHorizontalBarChart(slide, 8.42, 2.94, 3.62, 1.8, events.map((row) => ({ label: row.name || row.title, value: parseNumber(row.participants) })), { valueMode: 'number', labelMax: 19, catSize: 6.9, dataSize: 6.9 });
  paragraph(slide, 8.42, 5.12, 3.62, 0.54, cleanText(p.message || 'Los eventos complementan el alcance digital y permiten reforzar mensajes estratégicos.'), { max: 150, fontSize: 8.7 });
  finalizeSlide(slide);
}

function renderFullClosing() {
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.electricBlue };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: COLORS.electricBlue }, line: { color: COLORS.electricBlue } });
  slide.addShape(pptx.ShapeType.arc, { x: -1.4, y: -1.1, w: 4.5, h: 4.5, adjustPoint: 0.24, line: { color: COLORS.cyan, transparency: 45, width: 2.4 } });
  addLogo(slide, 'white', 5.62, 3.08, 2.08, 0.68);
  slide.addText(`Comunicaciones Internas · ${periodLabel()}`, { x: 2.55, y: 3.94, w: 8.3, h: 0.35, fontFace: 'Georgia', bold: true, fontSize: 18, color: COLORS.white, align: 'center', margin: 0 });
  finalizeSlide(slide);
}

function buildLegacyRenderPlan(payload) {
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
      { key: 'executive_summary', title: 'Resumen ejecutivo del período', payload: { headline: s2.headline, plan_total: s2.volume_current, site_notes_total: s2.pull_notes_current, site_total_views: s6.total_views, mail_total: s3.mail_total, mail_open_rate: s2.push_open_rate, mail_interaction_rate: s2.push_interaction_rate, historical_note: s2.comparative_note, takeaways: s2.highlights } },
      { key: 'channel_management', title: 'Gestión de canales', payload: { mail_total: s3.mail_total, mail_open_rate: s3.open_rate, mail_interaction_rate: s2.push_interaction_rate, site_notes_total: s3.pull_total, site_total_views: s6.total_views, channel_mix: s2.audience_segments || [], timeline_mail: s3.mail_timeline || [], timeline_site: s3.pull_timeline || [], message: s3.footer || s3.mail_message } },
      { key: 'mix_thematic_clients', title: 'Mix temático y áreas solicitantes', payload: { strategic_axes: s4.content_distribution || [], internal_clients: s4.internal_clients || [], format_mix: s4.format_mix || [], message: s4.conclusion || s4.theme_message } },
      { key: 'ranking_push', title: 'Ranking push', payload: { by_interaction: s5.top_communications || [], by_open_rate: s5.top_by_open_rate || [], available: Array.isArray(s5.top_communications) && s5.top_communications.length > 0, message: s5.key_learning } },
      { key: 'ranking_pull', title: 'Ranking pull', payload: { top_pull_notes: s6.top_notes || [], available: Array.isArray(s6.top_notes) && s6.top_notes.length > 0, average_reads_per_note: s6.avg_reads, site_total_views: s6.total_views, message: s6.conclusion } },
      { key: 'milestones', title: 'Hitos del mes', payload: { items: s7, message: 'Hitos destacados del período' } },
      { key: 'recommendations', title: 'Conclusiones y próximos pasos', payload: { message: s2.conclusion_message || s4.conclusion, recommendations: s2.highlights || [], experiments: ['Testear asuntos y horarios en piezas de beneficios.', 'Cruzar top notas SITE con próximos envíos segmentados.', 'Revisar canales subutilizados para campañas de bajo alcance.'], action_plan: ['Definir foco editorial del próximo mes.', 'Priorizar 2 campañas con call to action claro.', 'Medir variación de apertura, interacción y vistas.'] } },
      ...(includeEvents ? [{ key: 'events', title: 'Eventos del mes', payload: { events: s8.event_breakdown || [], total_events: s8.total_events, total_participants: s8.total_participants, message: s8.conclusion || s8.secondary_message } }] : []),
    ],
  };
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
  recommendations: renderRecommendations,
  events: renderEvents,
};

if (renderMode === 'full') renderFullCover();
for (const module of renderPlan.modules || []) {
  if (module.key === 'milestones') {
    const items = module?.payload?.items;
    if (!Array.isArray(items) || !items.length) continue;
  }
  const fn = renderers[module.key];
  if (fn) fn(module);
}
if (renderMode === 'full') renderFullClosing();

pptx.writeFile({ fileName: outputPptxPath }).catch((err) => {
  console.error(err);
  process.exit(1);
});
