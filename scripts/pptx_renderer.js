#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');
const {
  imageSizingCrop,
  imageSizingContain,
  safeOuterShadow,
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
} = require('./pptx_helpers_local');

const inputJsonPath = process.argv[2];
const outputPptxPath = process.argv[3];
if (!inputJsonPath || !outputPptxPath) {
  console.error('Usage: node scripts/pptx_renderer.js <report.json> <output.pptx>');
  process.exit(1);
}

const report = JSON.parse(fs.readFileSync(inputJsonPath, 'utf8'));
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'OpenAI';
pptx.company = 'BBVA';
pptx.subject = 'Informe de gestión';
pptx.title = `${report?.slide_1_cover?.area || 'Comunicaciones Internas'} - ${report?.slide_1_cover?.period || ''}`;
pptx.lang = 'es-AR';
pptx.theme = {
  headFontFace: 'Source Serif 4',
  bodyFontFace: 'Lato',
  lang: 'es-AR',
};
pptx.defineLayout({ name: 'BBVA_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'BBVA_WIDE';

const COLORS = {
  electricBlue: '001391',
  sereneBlue: '85C8FF',
  midnight: '060E46',
  sand: 'F7F8F8',
  white: 'FFFFFF',
  ice: '8BE1E9',
  lime: '88E783',
  canary: 'FFE761',
  purple: '9694FF',
  mandarin: 'FFB56B',
  grey5: '000519',
  grey4: '46536D',
  grey3: 'ADB8C2',
  grey2: 'CAD1D8',
  grey1: 'E2E6EA',
};
const ACCENTS = [COLORS.electricBlue, COLORS.sereneBlue, COLORS.ice, COLORS.lime, COLORS.canary, COLORS.purple, COLORS.mandarin, '6DB8FF'];
const BRAND_ASSETS_DIR = path.resolve(__dirname, '..', 'assets', 'brand');
const BBVA_LOGO_BLUE = path.join(BRAND_ASSETS_DIR, 'bbva_logo_blue.png');
const BBVA_LOGO_WHITE = path.join(BRAND_ASSETS_DIR, 'bbva_logo_white.png');

function coalesce(...values) {
  for (const value of values) {
    if (value !== undefined && value !== null && String(value).trim() !== '') return value;
  }
  return '-';
}

function parseNumber(value) {
  if (typeof value === 'number') return value;
  if (value === null || value === undefined || value === '-') return 0;
  const cleaned = String(value)
    .replace(/%/g, '')
    .replace(/\./g, '')
    .replace(/,/g, '.')
    .replace(/[^0-9.-]/g, '');
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function formatNumber(value) {
  const n = parseNumber(value);
  return new Intl.NumberFormat('es-AR', { maximumFractionDigits: 1 }).format(n);
}

function formatPercent(value) {
  if (value === '-' || value === '' || value === null || value === undefined) return '-';
  const raw = String(value);
  if (raw.includes('%')) return raw.replace('.', ',');
  const n = parseNumber(value);
  return `${new Intl.NumberFormat('es-AR', {
    minimumFractionDigits: n % 1 ? 1 : 0,
    maximumFractionDigits: 1,
  }).format(n)}%`;
}

function truncate(text, max = 62) {
  const clean = String(coalesce(text, '-')).replace(/\s+/g, ' ').trim();
  return clean.length > max ? `${clean.slice(0, max - 1)}…` : clean;
}

function estimateCharsPerLine(w, fontSize) {
  const safeW = Math.max(0.6, Number(w) || 1);
  const safeFs = Math.max(7, Number(fontSize) || 12);
  return Math.max(10, Math.floor((safeW * 10.2) * (12 / safeFs)));
}

function estimateLineCount(text, w, fontSize) {
  const parts = String(coalesce(text, '')).split(/\n+/);
  const cpl = estimateCharsPerLine(w, fontSize);
  return Math.max(1, parts.reduce((acc, part) => acc + Math.max(1, Math.ceil(part.length / cpl)), 0));
}

function fitFontSize(text, w, h, baseFontSize, minFontSize = 8, maxLines = null) {
  const boxH = Math.max(0.18, Number(h) || 0.3);
  let fs = Number(baseFontSize) || 12;
  const minFs = Math.min(fs, Number(minFontSize) || 8);
  while (fs > minFs) {
    const lines = estimateLineCount(text, w, fs);
    const requiredH = (lines * fs * 1.22) / 72 + 0.02;
    if (requiredH <= boxH && (!maxLines || lines <= maxLines)) return fs;
    fs -= 0.5;
  }
  return minFs;
}

function addSafeText(slide, text, opts = {}, fitOpts = {}) {
  const content = String(coalesce(text, '-'));
  const fontSize = fitFontSize(
    content,
    opts.w,
    opts.h,
    opts.fontSize || 12,
    fitOpts.minFontSize || Math.max(7, (opts.fontSize || 12) - 4),
    fitOpts.maxLines || null,
  );
  slide.addText(content, { ...opts, fontSize });
}

function resolveAsset(assetPath) {
  if (!assetPath) return null;
  const maybe = path.isAbsolute(assetPath) ? assetPath : path.resolve(process.cwd(), assetPath);
  return fs.existsSync(maybe) ? maybe : null;
}

function addLogo(slide, dark = false, opts = {}) {
  const logoPath = dark ? BBVA_LOGO_WHITE : BBVA_LOGO_BLUE;
  const resolved = resolveAsset(logoPath);
  const x = opts.x ?? 11.95;
  const y = opts.y ?? 0.20;
  const w = opts.w ?? 0.68;
  const h = opts.h ?? 0.22;

  if (resolved) {
    slide.addImage({ path: resolved, ...imageSizingContain(resolved, x, y, w, h) });
    return;
  }

  slide.addText('BBVA', {
    x, y, w, h,
    fontFace: 'Lato', bold: true, fontSize: 18,
    color: dark ? COLORS.white : COLORS.electricBlue,
    align: 'right', margin: 0,
  });
}

function baseSlide(slide, title, section) {
  slide.background = { color: COLORS.sand };
  addLogo(slide, false);
  if (section) {
    slide.addText(section.toUpperCase(), {
      x: 0.62, y: 0.28, w: 3.2, h: 0.18,
      fontFace: 'Lato', fontSize: 9.5, color: COLORS.grey4, margin: 0,
    });
  }
  slide.addText(title, {
    x: 0.62, y: 0.56, w: 5.8, h: 0.6,
    fontFace: 'Source Serif 4', bold: true, fontSize: 24,
    color: COLORS.electricBlue, margin: 0,
  });
}

function addFooter(slide, text) {
  slide.addText(text, {
    x: 0.72, y: 7.02, w: 11.86, h: 0.18,
    fontFace: 'Lato', fontSize: 9.5, color: COLORS.grey4, margin: 0,
    align: 'center',
  });
}

function addPanel(slide, x, y, w, h, opts = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: 0.06,
    fill: { color: opts.fill || COLORS.white },
    line: { color: opts.line || COLORS.grey1, width: opts.lineWidth || 1 },
    shadow: opts.shadow ? safeOuterShadow('000000', 0.14, 45, 1.2, 0.6) : undefined,
  });
  if (opts.header) {
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y, w, h: opts.headerH || 0.58,
      rectRadius: 0.06,
      fill: { color: opts.headerFill || COLORS.electricBlue },
      line: { color: opts.headerFill || COLORS.electricBlue, width: 1 },
    });
    slide.addText(opts.header, {
      x: x + 0.18, y: y + 0.11, w: w - 0.36, h: 0.17,
      fontFace: 'Lato', fontSize: 10.5, bold: true,
      color: opts.headerColor || COLORS.white, margin: 0,
    });
    if (opts.subheader) {
      slide.addText(opts.subheader, {
        x: x + 0.18, y: y + 0.30, w: w - 0.36, h: 0.12,
        fontFace: 'Lato', fontSize: 7.5, color: opts.headerColor || COLORS.white, margin: 0,
      });
    }
  }
}

function addMetricCard(slide, x, y, w, h, eyebrow, value, caption, accent = COLORS.sereneBlue) {
  addPanel(slide, x, y, w, h, { fill: COLORS.white, line: COLORS.grey1 });
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h: 0.07,
    fill: { color: accent }, line: { color: accent },
  });
  slide.addText(eyebrow, {
    x: x + 0.14, y: y + 0.12, w: w - 0.28, h: 0.14,
    fontFace: 'Lato', fontSize: 8.5, color: COLORS.grey4, margin: 0,
  });
  slide.addText(String(coalesce(value, '-')), {
    x: x + 0.14, y: y + 0.30, w: w - 0.28, h: 0.24,
    fontFace: 'Source Serif 4', bold: true, fontSize: 18,
    color: COLORS.electricBlue, margin: 0,
  });
  slide.addText(caption, {
    x: x + 0.14, y: y + h - 0.20, w: w - 0.28, h: 0.12,
    fontFace: 'Lato', fontSize: 8.5, color: COLORS.grey4, margin: 0,
  });
}

function addBulletList(slide, x, y, w, bullets, opts = {}) {
  const items = (bullets || []).slice(0, opts.maxItems || 4);
  const boxH = opts.h || ((items.length || 1) * (opts.gap || 0.36));
  const maxByHeight = Math.max(1, Math.floor(boxH / (opts.gap || 0.36)));
  let cursorY = y;
  items.slice(0, maxByHeight).forEach((bullet, idx) => {
    const gap = opts.gap || 0.36;
    const lineH = opts.lineH || 0.28;
    if (cursorY + lineH > y + boxH) return;
    slide.addShape(pptx.ShapeType.ellipse, {
      x, y: cursorY + 0.05, w: 0.10, h: 0.10,
      fill: { color: opts.bulletColors?.[idx] || ACCENTS[idx % ACCENTS.length] },
      line: { color: opts.bulletColors?.[idx] || ACCENTS[idx % ACCENTS.length] },
    });
    addSafeText(slide, bullet, {
      x: x + 0.16, y: cursorY, w: w - 0.16, h: lineH,
      fontFace: 'Lato', fontSize: opts.fontSize || 11, color: COLORS.midnight, margin: 0,
      breakLine: false,
    }, { minFontSize: opts.minFontSize || 7.8, maxLines: opts.maxLinesPerBullet || 2 });
    cursorY += gap;
  });
}

function addImagePanel(slide, imgPath, x, y, w, h) {
  const resolved = resolveAsset(imgPath);
  if (resolved) {
    slide.addImage({ path: resolved, ...imageSizingCrop(resolved, x, y, w, h) });
  } else {
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y, w, h, rectRadius: 0.04,
      fill: { color: 'EEF3F7' }, line: { color: COLORS.grey1 },
    });
    slide.addText('Asset visual pendiente', {
      x: x + 0.2, y: y + h / 2 - 0.12, w: w - 0.4, h: 0.24,
      fontFace: 'Lato', fontSize: 10, color: COLORS.grey4, align: 'center', margin: 0,
    });
  }
}

function slideCover() {
  const s1 = report.slide_1_cover || {};
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.electricBlue };
  addLogo(slide, true);

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 8.25, y: 1.15, w: 2.8, h: 2.1,
    rectRadius: 0.18,
    fill: { color: COLORS.sereneBlue, transparency: 8 },
    line: { color: COLORS.white, transparency: 55, width: 1.2 },
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 10.1, y: 1.95, w: 0.28, h: 0.28,
    fill: { color: COLORS.sereneBlue, transparency: 0 },
    line: { color: COLORS.sereneBlue },
  });
  slide.addText(String((s1.period || '').split(' ').slice(-1)[0] || ''), {
    x: 0.48, y: 2.05, w: 1.5, h: 0.36,
    fontFace: 'Source Serif 4', bold: true, fontSize: 28,
    color: COLORS.white, margin: 0,
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 0.48, y: 2.62, w: 8.0, h: 0,
    line: { color: COLORS.white, transparency: 35, width: 1 },
  });
  slide.addText(`${coalesce(s1.area, 'Comunicaciones Internas')}\n${coalesce(s1.subtitle, 'Informe de gestión')}`, {
    x: 0.48, y: 2.98, w: 5.9, h: 1.75,
    fontFace: 'Source Serif 4', bold: true, fontSize: 30,
    color: COLORS.white, margin: 0,
  });
  slide.addText(String(coalesce(s1.period, '-')), {
    x: 0.5, y: 5.95, w: 3.8, h: 0.2,
    fontFace: 'Lato', fontSize: 11.5, color: COLORS.white, margin: 0,
  });
}

function slideOverview() {
  const s2 = report.slide_2_overview || {};
  const timeline = (s2.comparison_timeline || []).slice(0, 6);
  const slide = pptx.addSlide();
  baseSlide(slide, s2.headline || '¿Cómo nos fue? CI', 'Resumen ejecutivo');

  addMetricCard(slide, 0.62, 1.36, 1.58, 0.92, 'Comunicaciones push', formatNumber(s2.volume_current), 'Período actual', COLORS.sereneBlue);
  addMetricCard(slide, 2.33, 1.36, 1.58, 0.92, 'Apertura promedio', formatPercent(s2.push_open_rate), 'Tasa media', COLORS.ice);
  addMetricCard(slide, 4.04, 1.36, 1.58, 0.92, 'Interacción', formatPercent(s2.push_interaction_rate), 'Tasa media', COLORS.lime);
  addMetricCard(slide, 5.75, 1.36, 1.58, 0.92, 'Lecturas promedio', formatNumber(s2.average_reads), 'Por nota pull', COLORS.canary);

  addPanel(slide, 0.62, 2.52, 3.55, 3.82, { fill: COLORS.white, line: COLORS.grey1 });
  slide.addText('Lectura ejecutiva', {
    x: 0.84, y: 2.78, w: 2.3, h: 0.18,
    fontFace: 'Lato', fontSize: 10.5, bold: true, color: COLORS.electricBlue, margin: 0,
  });
  addSafeText(slide, truncate(s2.conclusion_message || '-', 185), {
    x: 0.84, y: 3.06, w: 3.0, h: 0.92,
    fontFace: 'Source Serif 4', bold: true, fontSize: 17,
    color: COLORS.midnight, margin: 0,
  }, { minFontSize: 13, maxLines: 4 });
  addBulletList(slide, 0.86, 4.28, 3.0, (s2.highlights || []).length ? s2.highlights : [
    `${formatNumber(s2.volume_current)} comunicaciones push en el período.`,
    `${formatNumber(s2.pull_notes_current)} publicaciones pull relevadas.`,
    `${formatNumber(s2.average_reads)} lecturas promedio por nota.`,
  ], { maxItems: 4, gap: 0.42, fontSize: 10.5, lineH: 0.30, h: 1.2, minFontSize: 8.2, maxLinesPerBullet: 2 });

  addPanel(slide, 4.38, 2.52, 2.85, 3.82, { header: 'SEGMENTACIÓN DE AUDIENCIA', subheader: '(promedio del período)' });
  const segs = (s2.audience_segments || []).slice(0, 5);
  if (segs.length) {
    slide.addChart('doughnut', [{
      name: 'Segmentación',
      labels: segs.map((item) => truncate(item.label, 14)),
      values: segs.map((item) => parseNumber(item.value)),
    }], {
      x: 4.58, y: 3.16, w: 1.65, h: 1.68,
      holeSize: 68,
      chartColors: [COLORS.electricBlue, COLORS.sereneBlue, COLORS.ice, COLORS.lime, COLORS.canary],
      showLegend: false,
      showValue: false,
      showTitle: false,
    });
    addBulletList(slide, 4.62, 4.92, 2.2, segs.map((item) => `${truncate(item.label, 18)}  ${formatPercent(item.value)}`), {
      maxItems: 5, gap: 0.28, fontSize: 8.7, lineH: 0.22, h: 1.2, minFontSize: 7.2, maxLinesPerBullet: 2,
    });
  } else {
    slide.addText('Sin datos disponibles.', {
      x: 4.7, y: 4.0, w: 2.0, h: 0.2,
      fontFace: 'Lato', fontSize: 10, color: COLORS.grey4, margin: 0,
    });
  }

  addPanel(slide, 7.46, 2.52, 5.2, 3.82, {
    header: 'VOLUMEN DEL PERÍODO Y COMPARATIVO',
    subheader: '(si no hay histórico, se mantiene como referencia simple)'
  });
  if (timeline.length) {
    slide.addChart('column', [{
      name: 'Volumen',
      labels: timeline.map((item) => truncate(item.label, 10)),
      values: timeline.map((item) => parseNumber(item.value)),
    }], {
      x: 7.72, y: 3.18, w: 3.2, h: 2.0,
      showLegend: false,
      chartColors: [COLORS.sereneBlue],
      catAxisLabelFontFace: 'Lato',
      catAxisLabelFontSize: 8,
      valAxisLabelFontFace: 'Lato',
      valAxisLabelFontSize: 8,
      valGridLine: { color: COLORS.grey1, width: 1 },
      showValue: true,
      dataLabelColor: COLORS.electricBlue,
      dataLabelPosition: 'outEnd',
      showTitle: false,
    });
  }
  slide.addText(String(coalesce(s2.volume_previous, '-')), {
    x: 11.15, y: 3.24, w: 0.95, h: 0.28,
    fontFace: 'Source Serif 4', bold: true, fontSize: 18, color: COLORS.grey4, margin: 0,
  });
  slide.addText('Período anterior', {
    x: 11.15, y: 3.54, w: 1.15, h: 0.14,
    fontFace: 'Lato', fontSize: 8.5, color: COLORS.grey4, margin: 0,
  });
  slide.addText(String(coalesce(s2.volume_change, '-')), {
    x: 11.15, y: 4.02, w: 1.1, h: 0.28,
    fontFace: 'Source Serif 4', bold: true, fontSize: 20, color: COLORS.electricBlue, margin: 0,
  });
  slide.addText('Variación', {
    x: 11.15, y: 4.34, w: 0.9, h: 0.14,
    fontFace: 'Lato', fontSize: 8.5, color: COLORS.grey4, margin: 0,
  });
  addSafeText(slide, truncate(s2.comparative_note || 'Síntesis del desempeño general del período.', 85), {
    x: 11.0, y: 4.92, w: 1.35, h: 0.78,
    fontFace: 'Lato', fontSize: 8.7, color: COLORS.midnight, margin: 0,
  }, { minFontSize: 7, maxLines: 6 });
  addFooter(slide, 'Síntesis del desempeño general del período.');
}

function slidePlan() {
  const sPlan = report.slide_3_plan || {};
  const mailTimeline = (sPlan.mail_timeline || []).slice(0, 6);
  const pullTimeline = (sPlan.pull_timeline || []).slice(0, 6);
  const slide = pptx.addSlide();
  baseSlide(slide, sPlan.title || 'Gestión del plan CI', 'Canales');

  addPanel(slide, 0.62, 1.4, 7.2, 4.95, { header: 'MAIL', subheader: '(evolución del volumen y saturación)' });
  if (mailTimeline.length) {
    slide.addChart('column', [{
      name: 'Mail',
      labels: mailTimeline.map((item) => truncate(item.label, 10)),
      values: mailTimeline.map((item) => parseNumber(item.value)),
    }], {
      x: 0.88, y: 2.0, w: 4.3, h: 2.45,
      showLegend: false,
      chartColors: [COLORS.electricBlue],
      catAxisLabelFontFace: 'Lato', catAxisLabelFontSize: 8,
      valAxisLabelFontFace: 'Lato', valAxisLabelFontSize: 8,
      valGridLine: { color: COLORS.grey1, width: 1 },
      showValue: true,
      dataLabelColor: COLORS.electricBlue,
      dataLabelPosition: 'outEnd',
      showTitle: false,
    });
  }
  addMetricCard(slide, 5.45, 2.0, 1.9, 0.96, 'Envíos del período', formatNumber(sPlan.mail_total || '-'), 'Total', COLORS.sereneBlue);
  addMetricCard(slide, 5.45, 3.12, 1.9, 0.96, 'A todo BBVA', formatPercent(sPlan.segmented_share || '-'), 'Share de alcance', COLORS.ice);
  addSafeText(slide, truncate(sPlan.mail_message || '-', 150), {
    x: 0.88, y: 4.8, w: 6.2, h: 0.78,
    fontFace: 'Lato', fontSize: 11.2, color: COLORS.midnight, margin: 0,
  }, { minFontSize: 9, maxLines: 4 });

  addPanel(slide, 8.02, 1.4, 4.64, 4.95, { header: 'INTRANET / SITE', subheader: '(volumen y lectura)' });
  if (pullTimeline.length) {
    slide.addChart('bar', [{
      name: 'Notas',
      labels: pullTimeline.map((item) => truncate(item.label, 10)),
      values: pullTimeline.map((item) => parseNumber(item.value)),
    }], {
      x: 8.28, y: 2.0, w: 2.95, h: 2.1,
      showLegend: false,
      chartColors: [COLORS.sereneBlue],
      catAxisLabelFontFace: 'Lato', catAxisLabelFontSize: 8,
      valAxisLabelFontFace: 'Lato', valAxisLabelFontSize: 8,
      valGridLine: { color: COLORS.grey1, width: 1 },
      showValue: true,
      dataLabelColor: COLORS.electricBlue,
      dataLabelPosition: 'outEnd',
      showTitle: false,
    });
  }
  addMetricCard(slide, 11.45, 2.0, 0.95, 0.96, 'Notas', formatNumber(sPlan.pull_total || '-'), 'Total', COLORS.lime);
  addMetricCard(slide, 11.45, 3.12, 0.95, 0.96, 'Open', formatPercent(sPlan.open_rate || '-'), 'Prom.', COLORS.canary);
  addSafeText(slide, truncate(sPlan.pull_message || '-', 150), {
    x: 8.28, y: 4.55, w: 3.8, h: 0.88,
    fontFace: 'Lato', fontSize: 11.2, color: COLORS.midnight, margin: 0,
  }, { minFontSize: 9, maxLines: 4 });
  addFooter(slide, sPlan.footer || 'Esta slide resume la gestión operacional de los principales canales.');
}

function slideStrategy() {
  const s3 = report.slide_4_strategy || report.slide_3_strategy || {};
  const slide = pptx.addSlide();
  baseSlide(slide, 'Mix de contenidos y ejes estratégicos', 'Contenido');
  const contentDistribution = (s3.content_distribution || []).slice(0, 6);
  const clients = (s3.internal_clients || []).slice(0, 5);
  const balance = s3.canal_balance || {};

  addPanel(slide, 0.62, 1.4, 3.88, 4.95, { header: 'EJES ESTRATÉGICOS', subheader: '(participación promedio)' });
  if (contentDistribution.length) {
    slide.addChart('doughnut', [{
      name: 'Ejes',
      labels: contentDistribution.map((item) => truncate(item.theme, 16)),
      values: contentDistribution.map((item) => parseNumber(item.weight)),
    }], {
      x: 0.86, y: 2.0, w: 1.95, h: 2.05,
      holeSize: 68,
      chartColors: ACCENTS.slice(0, Math.max(3, contentDistribution.length)),
      showLegend: false,
      showValue: false,
    });
    addBulletList(slide, 2.65, 2.06, 1.45, contentDistribution.map((item) => `${truncate(item.theme, 16)} ${formatPercent(item.weight)}`), {
      maxItems: 6, gap: 0.28, fontSize: 8.7, lineH: 0.22, h: 1.9, minFontSize: 7, maxLinesPerBullet: 2,
    });
  }
  addSafeText(slide, truncate(s3.theme_message || s3.conclusion || '-', 110), {
    x: 0.88, y: 4.8, w: 3.0, h: 0.64,
    fontFace: 'Lato', fontSize: 10.8, color: COLORS.midnight, margin: 0,
  }, { minFontSize: 8.5, maxLines: 4 });

  addPanel(slide, 4.72, 1.4, 3.18, 4.95, { header: 'BALANCE INSTITUCIONAL / TRANSACCIONAL', subheader: '(share del período)' });
  slide.addChart('doughnut', [{
    name: 'Balance',
    labels: ['Institucional', 'Transaccional / talento'],
    values: [parseNumber(balance.institutional), parseNumber(balance.transactional_talent)],
  }], {
    x: 5.05, y: 2.08, w: 1.7, h: 1.8,
    holeSize: 72,
    chartColors: [COLORS.electricBlue, COLORS.sereneBlue],
    showLegend: false,
    showValue: true,
    dataLabelColor: COLORS.electricBlue,
    dataLabelFormatCode: '0%'
  });
  addBulletList(slide, 5.0, 4.08, 2.0, [
    `Institucional ${formatPercent(balance.institutional)}`,
    `Transaccional / talento ${formatPercent(balance.transactional_talent)}`,
    truncate(s3.balance_message || 'El objetivo es sostener un mix balanceado.', 50),
  ], { maxItems: 3, gap: 0.40, fontSize: 10.2, lineH: 0.30, h: 1.1, minFontSize: 8, maxLinesPerBullet: 2 });

  addPanel(slide, 8.12, 1.4, 4.54, 4.95, { header: 'PEDIDOS POR DIRECCIÓN / ÁREA', subheader: '(participación promedio)' });
  if (clients.length) {
    slide.addChart('bar', [{
      name: 'Áreas',
      labels: clients.map((item) => truncate(item.label, 16)),
      values: clients.map((item) => parseNumber(item.value)),
    }], {
      x: 8.42, y: 2.0, w: 3.68, h: 2.5,
      showLegend: false,
      chartColors: [COLORS.sereneBlue],
      catAxisLabelFontFace: 'Lato', catAxisLabelFontSize: 8.5,
      valAxisLabelFontFace: 'Lato', valAxisLabelFontSize: 8,
      valGridLine: { color: COLORS.grey1, width: 1 },
      showValue: true,
      dataLabelColor: COLORS.electricBlue,
      dataLabelPosition: 'outEnd',
      showTitle: false,
    });
  }
  addSafeText(slide, truncate(s3.conclusion || '-', 110), {
    x: 8.44, y: 4.95, w: 3.4, h: 0.66,
    fontFace: 'Lato', fontSize: 10.8, color: COLORS.midnight, margin: 0,
  }, { minFontSize: 8.5, maxLines: 4 });
  addFooter(slide, 'La lectura de contenido cruza ejes temáticos, balance editorial y origen de los pedidos.');
}

function slidePushRanking() {
  const s4 = report.slide_5_push_ranking || report.slide_4_push_ranking || {};
  const rows = (s4.top_communications || []).slice(0, 3);
  const slide = pptx.addSlide();
  baseSlide(slide, 'Comunicaciones push de mayor impacto', 'Push');

  addPanel(slide, 0.62, 1.4, 7.08, 5.04, { header: 'TOP COMUNICACIONES', subheader: '(clics e interacción)' });
  slide.addText('Comunicación', { x: 0.88, y: 2.0, w: 4.2, h: 0.16, fontFace: 'Lato', fontSize: 10.5, bold: true, color: COLORS.grey4, margin: 0 });
  slide.addText('Clics', { x: 5.5, y: 2.0, w: 0.65, h: 0.16, fontFace: 'Lato', fontSize: 10.5, bold: true, color: COLORS.grey4, align: 'right', margin: 0 });
  slide.addText('Interacción', { x: 6.35, y: 2.0, w: 0.95, h: 0.16, fontFace: 'Lato', fontSize: 10.5, bold: true, color: COLORS.grey4, align: 'right', margin: 0 });
  rows.forEach((row, idx) => {
    const y = 2.34 + idx * 1.08;
    if (idx % 2 === 0) {
      slide.addShape(pptx.ShapeType.rect, { x: 0.82, y, w: 6.66, h: 0.85, fill: { color: 'F9FBFD' }, line: { color: 'F9FBFD' } });
    }
    addSafeText(slide, `${idx + 1}. ${truncate(row.name, 72)}`, {
      x: 0.92, y: y + 0.13, w: 4.45, h: 0.40,
      fontFace: 'Lato', fontSize: 12.2, color: COLORS.midnight, margin: 0,
    }, { minFontSize: 9.5, maxLines: 2 });
    slide.addText(formatNumber(row.clicks), {
      x: 5.42, y: y + 0.12, w: 0.78, h: 0.28,
      fontFace: 'Source Serif 4', bold: true, fontSize: 16, color: COLORS.electricBlue, align: 'right', margin: 0,
    });
    slide.addText(formatPercent(row.interaction), {
      x: 6.24, y: y + 0.12, w: 0.98, h: 0.28,
      fontFace: 'Source Serif 4', bold: true, fontSize: 16, color: COLORS.electricBlue, align: 'right', margin: 0,
    });
  });
  if (rows.length) {
    slide.addChart('bar', [{
      name: 'Clics',
      labels: rows.map((item) => `${rows.indexOf(item)+1}`),
      values: rows.map((item) => parseNumber(item.clicks)),
    }], {
      x: 0.92, y: 5.38, w: 2.05, h: 0.56,
      showLegend: false,
      chartColors: [COLORS.electricBlue],
      catAxisLabelFontFace: 'Lato', catAxisLabelFontSize: 8,
      valAxisLabelFontFace: 'Lato', valAxisLabelFontSize: 7,
      showValue: false,
      valGridLine: { color: COLORS.grey1, width: 1 },
      showTitle: false,
    });
    slide.addText('Ranking visual por clics', {
      x: 3.1, y: 5.42, w: 1.4, h: 0.14,
      fontFace: 'Lato', fontSize: 8.8, color: COLORS.grey4, margin: 0,
    });
  }

  addPanel(slide, 7.92, 1.4, 4.74, 3.58, { header: 'PIEZAS / MINIATURAS', subheader: '(reales si el pipeline recibe assets)' });
  rows.forEach((row, idx) => {
    const thumbW = 1.34;
    const gap = 0.14;
    const x = 8.18 + idx * (thumbW + gap);
    addImagePanel(slide, row.thumbnail_path, x, 2.05, thumbW, 2.44);
    addSafeText(slide, truncate(row.name, 22), {
      x, y: 4.58, w: thumbW, h: 0.28,
      fontFace: 'Lato', fontSize: 7.8, color: COLORS.midnight, margin: 0, align: 'center',
    }, { minFontSize: 7, maxLines: 2 });
  });

  addPanel(slide, 7.92, 5.16, 4.74, 1.28, { fill: COLORS.white, line: COLORS.grey1 });
  slide.addText('Aprendizaje', {
    x: 8.16, y: 5.38, w: 1.1, h: 0.15,
    fontFace: 'Lato', fontSize: 10.5, bold: true, color: COLORS.electricBlue, margin: 0,
  });
  addSafeText(slide, truncate(s4.key_learning || '-', 135), {
    x: 8.16, y: 5.60, w: 4.05, h: 0.48,
    fontFace: 'Lato', fontSize: 10.5, color: COLORS.midnight, margin: 0,
  }, { minFontSize: 8.2, maxLines: 3 });
  addFooter(slide, 'Las miniaturas se muestran cuando el pipeline recibe assets de cada comunicación top.');
}

function slidePullPerformance() {
  const s5 = report.slide_6_pull_performance || report.slide_5_pull_performance || {};
  const notes = (s5.top_notes || []).slice(0, 5);
  const slide = pptx.addSlide();
  baseSlide(slide, 'Desempeño pull', 'Intranet / site');

  addPanel(slide, 0.62, 1.4, 7.18, 4.96, { header: 'TOP NOTAS Y PROFUNDIDAD DE LECTURA', subheader: '(lecturas únicas)' });
  if (notes.length) {
    slide.addChart('bar', [{
      name: 'Lecturas únicas',
      labels: notes.map((item) => truncate(item.title, 28)),
      values: notes.map((item) => parseNumber(item.unique_reads)),
    }], {
      x: 0.92, y: 2.02, w: 4.7, h: 3.25,
      showLegend: false,
      chartColors: [COLORS.sereneBlue],
      catAxisLabelFontFace: 'Lato', catAxisLabelFontSize: 8.5,
      valAxisLabelFontFace: 'Lato', valAxisLabelFontSize: 8,
      valGridLine: { color: COLORS.grey1, width: 1 },
      showValue: true,
      dataLabelColor: COLORS.electricBlue,
      dataLabelPosition: 'outEnd',
      showTitle: false,
    });
    notes.slice(0, 2).forEach((note, idx) => {
      const y = 2.12 + idx * 1.0;
      slide.addText(`${idx + 1}. ${truncate(note.title, 28)}`, {
        x: 5.95, y, w: 1.35, h: 0.24,
        fontFace: 'Lato', fontSize: 8.8, color: COLORS.midnight, margin: 0,
      });
      slide.addText(`${formatNumber(note.unique_reads)} únicas`, {
        x: 5.95, y: y + 0.24, w: 1.15, h: 0.16,
        fontFace: 'Source Serif 4', bold: true, fontSize: 11, color: COLORS.electricBlue, margin: 0,
      });
      slide.addText(`${formatNumber(note.total_reads)} totales`, {
        x: 5.95, y: y + 0.44, w: 1.15, h: 0.16,
        fontFace: 'Lato', fontSize: 8.3, color: COLORS.grey4, margin: 0,
      });
    });
  }
  addSafeText(slide, truncate(s5.conclusion || '-', 120), {
    x: 0.92, y: 5.52, w: 6.1, h: 0.46,
    fontFace: 'Lato', fontSize: 10.8, color: COLORS.midnight, margin: 0,
  }, { minFontSize: 9, maxLines: 2 });

  addPanel(slide, 8.0, 1.4, 4.66, 4.96, { fill: COLORS.white, line: COLORS.grey1 });
  slide.addText('Lectura ejecutiva', {
    x: 8.24, y: 1.72, w: 1.8, h: 0.18,
    fontFace: 'Lato', fontSize: 10.5, bold: true, color: COLORS.electricBlue, margin: 0,
  });
  addSafeText(slide, truncate(s5.conclusion || '-', 175), {
    x: 8.24, y: 2.02, w: 3.95, h: 0.96,
    fontFace: 'Source Serif 4', bold: true, fontSize: 17,
    color: COLORS.midnight, margin: 0,
  }, { minFontSize: 13, maxLines: 4 });
  addMetricCard(slide, 8.24, 3.36, 1.18, 0.96, 'Notas', formatNumber(s5.pub_current), 'Período', COLORS.sereneBlue);
  addMetricCard(slide, 9.56, 3.36, 1.18, 0.96, 'Promedio', formatNumber(s5.avg_reads), 'Lecturas', COLORS.ice);
  addMetricCard(slide, 10.88, 3.36, 1.52, 0.96, 'Vistas', formatNumber(s5.total_views), 'Totales', COLORS.lime);
  addBulletList(slide, 8.28, 4.72, 3.9, [
    `Publicaciones del período: ${coalesce(s5.pub_current, '-')}`,
    `Período anterior: ${coalesce(s5.pub_previous, '-')}`,
    truncate(s5.secondary_message || 'El canal pull sirve para detectar qué notas sostienen lectura en profundidad.', 70),
  ], { maxItems: 3, gap: 0.42, fontSize: 10.5, lineH: 0.30, h: 1.35, minFontSize: 8.5, maxLinesPerBullet: 2 });
  addFooter(slide, 'El canal pull ayuda a identificar profundidad de lectura, no solo volumen de impactos.');
}

function slideHitos() {
  const hitos = (report.slide_7_hitos || report.slide_6_hitos || []).slice(0, 3);
  const slide = pptx.addSlide();
  baseSlide(slide, 'Hitos destacados', 'Agenda y campañas');
  if (!hitos.length) {
    addPanel(slide, 1.0, 2.0, 11.3, 3.2, { fill: COLORS.white, line: COLORS.grey1 });
    slide.addText('Sin hitos destacados para este período.', {
      x: 1.4, y: 3.0, w: 10.5, h: 0.4,
      fontFace: 'Source Serif 4', bold: true, fontSize: 26,
      color: COLORS.electricBlue, align: 'center', margin: 0,
    });
    addFooter(slide, 'El módulo admite hasta tres hitos con título, bullets e imagen principal.');
    return;
  }

  hitos.forEach((hito, idx) => {
    const x = 0.62 + idx * 4.12;
    addPanel(slide, x, 1.42, 3.5, 5.0, { fill: COLORS.white, line: COLORS.grey1 });
    addImagePanel(slide, hito.thumbnail_path, x + 0.16, 1.58, 3.18, 1.86);
    addSafeText(slide, String(coalesce(hito.title || hito.period, '-')), {
      x: x + 0.18, y: 3.64, w: 2.95, h: 0.40,
      fontFace: 'Source Serif 4', bold: true, fontSize: 17, color: COLORS.electricBlue, margin: 0,
    }, { minFontSize: 12.5, maxLines: 2 });
    const bullets = (hito.bullets && hito.bullets.length ? hito.bullets : [hito.description]).filter(Boolean);
    addBulletList(slide, x + 0.20, 4.24, 2.95, bullets, { maxItems: 3, gap: 0.40, fontSize: 10.2, lineH: 0.30, h: 1.45, minFontSize: 8.2, maxLinesPerBullet: 2 });
    slide.addText(truncate(hito.period || report.slide_1_cover?.period || '-', 24), {
      x: x + 0.18, y: 5.95, w: 2.9, h: 0.16,
      fontFace: 'Lato', fontSize: 8.5, color: COLORS.grey4, margin: 0,
    });
  });
  addFooter(slide, 'Los hitos suman contexto de gestión más allá de las métricas de volumen y performance.');
}

function slideEvents() {
  const s7 = report.slide_8_events || report.slide_7_events || {};
  const breakdown = (s7.event_breakdown || []).slice(0, 5);
  const slide = pptx.addSlide();
  baseSlide(slide, 'Eventos y participación', 'Activaciones');

  addMetricCard(slide, 0.62, 1.42, 1.75, 0.92, 'Eventos', formatNumber(s7.total_events), 'Período', COLORS.sereneBlue);
  addMetricCard(slide, 2.52, 1.42, 1.95, 0.92, 'Participaciones', formatNumber(s7.total_participants), 'Total', COLORS.ice);

  if (breakdown.length) {
    addPanel(slide, 0.62, 2.6, 7.28, 3.86, { header: 'DETALLE DE EVENTOS', subheader: '(participaciones por acción)' });
    slide.addChart('bar', [{
      name: 'Participantes',
      labels: breakdown.map((item) => truncate(item.name, 24)),
      values: breakdown.map((item) => parseNumber(item.participants)),
    }], {
      x: 0.92, y: 3.12, w: 4.75, h: 2.7,
      showLegend: false,
      chartColors: [COLORS.electricBlue],
      catAxisLabelFontFace: 'Lato', catAxisLabelFontSize: 8.5,
      valAxisLabelFontFace: 'Lato', valAxisLabelFontSize: 8,
      valGridLine: { color: COLORS.grey1, width: 1 },
      showValue: true,
      dataLabelColor: COLORS.electricBlue,
      dataLabelPosition: 'outEnd',
      showTitle: false,
    });
    breakdown.slice(0, 4).forEach((item, idx) => {
      addSafeText(slide, `${idx + 1}. ${truncate(item.name, 24)} — ${formatNumber(item.participants)}`, {
        x: 5.95, y: 3.20 + idx * 0.46, w: 1.55, h: 0.20,
        fontFace: 'Lato', fontSize: 8.4, color: COLORS.midnight, margin: 0,
      }, { minFontSize: 7, maxLines: 2 });
    });
  } else {
    addPanel(slide, 0.62, 2.6, 7.28, 3.86, { fill: COLORS.white, line: COLORS.grey1 });
    addSafeText(slide, 'No se registró un desglose consolidado de eventos para este período.', {
      x: 1.0, y: 3.85, w: 6.4, h: 0.9,
      fontFace: 'Source Serif 4', bold: true, fontSize: 21,
      color: COLORS.electricBlue, align: 'center', margin: 0,
    }, { minFontSize: 15, maxLines: 3 });
  }

  addPanel(slide, 8.08, 1.42, 4.58, 5.04, { fill: COLORS.white, line: COLORS.grey1 });
  slide.addText('Conclusión', {
    x: 8.34, y: 1.74, w: 1.3, h: 0.16,
    fontFace: 'Lato', fontSize: 10.5, bold: true, color: COLORS.electricBlue, margin: 0,
  });
  addSafeText(slide, truncate(s7.conclusion || '-', 180), {
    x: 8.34, y: 2.05, w: 3.85, h: 0.95,
    fontFace: 'Source Serif 4', bold: true, fontSize: 17,
    color: COLORS.midnight, margin: 0,
  }, { minFontSize: 13, maxLines: 4 });
  addBulletList(slide, 8.34, 3.42, 3.7, [
    `Eventos del período: ${coalesce(s7.total_events, '-')}`,
    `Participaciones: ${coalesce(s7.total_participants, '-')}`,
    truncate(s7.secondary_message || 'Con dato consolidado, este módulo permite leer alcance y performance por activación.', 76),
  ], { maxItems: 3, gap: 0.48, fontSize: 10.5, lineH: 0.30, h: 1.55, minFontSize: 8.5, maxLinesPerBullet: 2 });
  addFooter(slide, 'El módulo de eventos funciona tanto con datos completos como con una versión elegante sin desglose.');
}

function slideClosure() {
  const s8 = report.slide_9_closure || report.slide_8_closure || {};
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.electricBlue };
  addLogo(slide, true);
  addSafeText(slide, String(coalesce(s8.title, 'Claves del período')), {
    x: 0.68, y: 0.82, w: 4.8, h: 0.5,
    fontFace: 'Source Serif 4', bold: true, fontSize: 26,
    color: COLORS.white, margin: 0,
  }, { minFontSize: 20, maxLines: 2 });
  addPanel(slide, 0.72, 1.75, 11.1, 4.5, { fill: COLORS.white, line: COLORS.white });
  addBulletList(slide, 1.08, 2.18, 10.2, (s8.bullets || []).slice(0, 4), {
    maxItems: 4, gap: 0.72, fontSize: 13, lineH: 0.42, h: 3.1, minFontSize: 10.5, maxLinesPerBullet: 2,
    bulletColors: [COLORS.electricBlue, COLORS.sereneBlue, COLORS.ice, COLORS.lime],
  });
  slide.addText('Reporte automatizado con lineamientos visuales BBVA.', {
    x: 0.76, y: 6.88, w: 4.8, h: 0.16,
    fontFace: 'Lato', fontSize: 8.5, color: COLORS.white, margin: 0,
  });
}

slideCover();
slideOverview();
slidePlan();
slideStrategy();
slidePushRanking();
slidePullPerformance();
slideHitos();
slideEvents();
slideClosure();

for (const slide of pptx._slides) {
  warnIfSlideHasOverlaps(slide, pptx);
  warnIfSlideElementsOutOfBounds(slide, pptx);
}

pptx.writeFile({ fileName: outputPptxPath });
