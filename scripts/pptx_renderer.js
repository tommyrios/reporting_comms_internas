#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');
const { imageSizingContain, safeOuterShadow } = require('./pptx_helpers_local');

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

const COLORS = {
  electricBlue: '001391',
  sereneBlue: '85C8FF',
  midnight: '060E46',
  sand: 'F7F8F8',
  white: 'FFFFFF',
  ice: '8BE1E9',
  lime: '88E783',
  canary: 'FFE761',
  grey4: '46536D',
  grey2: 'CAD1D8',
  grey1: 'E2E6EA',
};
const DEFAULT_EXECUTIVE_TAKEAWAY = 'Se consolidó el desempeño mensual con métricas verificables.';

const BRAND_ASSETS_DIR = path.resolve(__dirname, '..', 'assets', 'brand');
const BBVA_LOGO_BLUE = path.join(BRAND_ASSETS_DIR, 'bbva_logo_blue.png');

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
  if (typeof value === 'number') return value;
  if (value === null || value === undefined || value === '-') return 0;
  const cleaned = String(value).replace(/%/g, '').replace(/\./g, '').replace(/,/g, '.').replace(/[^0-9.-]/g, '');
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function fmtNum(value) {
  return new Intl.NumberFormat('es-AR', { maximumFractionDigits: 2 }).format(parseNumber(value));
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

function addLogo(slide) {
  const logo = resolveAsset(BBVA_LOGO_BLUE);
  if (logo) {
    slide.addImage({ path: logo, ...imageSizingContain(logo, 12.0, 0.2, 0.72, 0.24) });
  } else {
    slide.addText('BBVA', { x: 11.9, y: 0.22, w: 0.9, h: 0.24, align: 'right', bold: true, color: COLORS.electricBlue, fontFace: 'Lato', fontSize: 16, margin: 0 });
  }
}

function baseSlide(title, subtitle = '') {
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.sand };
  addLogo(slide);
  if (subtitle) {
    slide.addText(cleanText(subtitle, ''), {
      x: 0.62, y: 0.24, w: 4.8, h: 0.18, fontFace: 'Lato', fontSize: 9.2, color: COLORS.grey4, margin: 0,
    });
  }
  slide.addText(cleanText(title, 'Reporte ejecutivo'), {
    x: 0.62, y: 0.54, w: 8.3, h: 0.52, fontFace: 'Source Serif 4', bold: true, fontSize: 24, color: COLORS.electricBlue, margin: 0,
  });
  return slide;
}

function panel(slide, x, y, w, h, header = '') {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: COLORS.white },
    line: { color: COLORS.grey1, width: 1 },
    shadow: safeOuterShadow('000000', 0.1, 45, 0.8, 0.4),
  });
  if (header) {
    slide.addText(cleanText(header), {
      x: x + 0.16, y: y + 0.12, w: w - 0.32, h: 0.16,
      fontFace: 'Lato', fontSize: 10, bold: true, color: COLORS.electricBlue, margin: 0,
    });
  }
}

function card(slide, x, y, w, h, label, value, accent = COLORS.sereneBlue) {
  panel(slide, x, y, w, h);
  slide.addShape(pptx.ShapeType.rect, { x, y, w, h: 0.06, fill: { color: accent }, line: { color: accent } });
  slide.addText(cleanText(label), { x: x + 0.12, y: y + 0.14, w: w - 0.24, h: 0.14, fontFace: 'Lato', fontSize: 8.5, color: COLORS.grey4, margin: 0 });
  slide.addText(String(value ?? '-'), { x: x + 0.12, y: y + 0.30, w: w - 0.24, h: 0.24, fontFace: 'Source Serif 4', bold: true, fontSize: 18, color: COLORS.electricBlue, margin: 0 });
}

function emptyState(slide, x, y, w, h, message) {
  panel(slide, x, y, w, h, 'Sin datos suficientes');
  slide.addText(clip(message || 'No hay datos consolidados para este módulo.', 120), {
    x: x + 0.20, y: y + 0.52, w: w - 0.4, h: h - 0.72,
    fontFace: 'Source Serif 4', bold: true, fontSize: 18, color: COLORS.electricBlue, align: 'center', valign: 'mid', margin: 0,
  });
}

function tableRows(slide, x, y, w, headers, rows) {
  const colW = w / headers.length;
  headers.forEach((header, idx) => {
    slide.addText(cleanText(header), {
      x: x + idx * colW + 0.02, y, w: colW - 0.04, h: 0.2,
      fontFace: 'Lato', fontSize: 9.5, bold: true, color: COLORS.grey4, margin: 0,
      align: idx === 0 ? 'left' : 'right',
    });
  });
  rows.slice(0, 5).forEach((row, rowIndex) => {
    const rowY = y + 0.28 + rowIndex * 0.42;
    if (rowIndex % 2 === 0) {
      slide.addShape(pptx.ShapeType.rect, {
        x, y: rowY - 0.04, w, h: 0.34,
        fill: { color: 'F9FBFD' }, line: { color: 'F9FBFD' },
      });
    }
    row.forEach((value, colIndex) => {
      slide.addText(clip(value, colIndex === 0 ? 60 : 16), {
        x: x + colIndex * colW + 0.02, y: rowY, w: colW - 0.04, h: 0.18,
        fontFace: colIndex === 0 ? 'Lato' : 'Source Serif 4',
        bold: colIndex !== 0,
        fontSize: colIndex === 0 ? 10 : 11,
        color: colIndex === 0 ? COLORS.midnight : COLORS.electricBlue,
        align: colIndex === 0 ? 'left' : 'right',
        margin: 0,
      });
    });
  });
}

function renderExecutiveSummary(module) {
  const p = module.payload || {};
  const slide = baseSlide(module.title || 'Resumen ejecutivo del período', 'Resumen');
  card(slide, 0.62, 1.4, 1.95, 0.96, 'Planificación', fmtNum(p.plan_total), COLORS.sereneBlue);
  card(slide, 2.72, 1.4, 1.95, 0.96, 'Noticias site', fmtNum(p.site_notes_total), COLORS.ice);
  card(slide, 4.82, 1.4, 1.95, 0.96, 'Vistas site', fmtNum(p.site_total_views), COLORS.lime);
  card(slide, 6.92, 1.4, 1.95, 0.96, 'Mails enviados', fmtNum(p.mail_total), COLORS.canary);
  card(slide, 9.02, 1.4, 1.72, 0.96, 'Apertura', fmtPct(p.mail_open_rate), COLORS.sereneBlue);
  card(slide, 10.86, 1.4, 1.82, 0.96, 'Interacción', fmtPct(p.mail_interaction_rate), COLORS.ice);

  panel(slide, 0.62, 2.62, 8.2, 3.72, 'Lectura ejecutiva');
  slide.addText(clip(p.historical_note || p.headline || '-', 240), {
    x: 0.86, y: 3.04, w: 7.6, h: 0.84,
    fontFace: 'Source Serif 4', bold: true, fontSize: 20, color: COLORS.midnight, margin: 0,
  });
  const takeaways = Array.isArray(p.takeaways) ? p.takeaways.filter(Boolean).slice(0, 3) : [];
  (takeaways.length ? takeaways : [DEFAULT_EXECUTIVE_TAKEAWAY]).forEach((item, idx) => {
    slide.addText(`• ${clip(item, 120)}`, {
      x: 0.9, y: 4.1 + idx * 0.48, w: 7.5, h: 0.24,
      fontFace: 'Lato', fontSize: 11, color: COLORS.midnight, margin: 0,
    });
  });

  panel(slide, 8.98, 2.62, 3.7, 3.72, 'Mensaje clave');
  slide.addText(clip(p.historical_note || '-', 160), {
    x: 9.2, y: 3.1, w: 3.25, h: 2.8,
    fontFace: 'Lato', fontSize: 11, color: COLORS.midnight, valign: 'mid', margin: 0,
  });
}

function renderChannelManagement(module) {
  const p = module.payload || {};
  const slide = baseSlide(module.title || 'Gestión de canales', 'Canales');
  card(slide, 0.62, 1.4, 2.1, 0.96, 'Mails enviados', fmtNum(p.mail_total), COLORS.sereneBlue);
  card(slide, 2.88, 1.4, 2.1, 0.96, 'Apertura', fmtPct(p.mail_open_rate), COLORS.ice);
  card(slide, 5.14, 1.4, 2.1, 0.96, 'Interacción', fmtPct(p.mail_interaction_rate), COLORS.lime);
  card(slide, 7.4, 1.4, 2.1, 0.96, 'Noticias site', fmtNum(p.site_notes_total), COLORS.canary);
  card(slide, 9.66, 1.4, 3.0, 0.96, 'Páginas vistas', fmtNum(p.site_total_views), COLORS.sereneBlue);

  panel(slide, 0.62, 2.62, 6.2, 3.72, 'Mix de canales');
  const mix = Array.isArray(p.channel_mix) ? p.channel_mix.slice(0, 5) : [];
  if (!mix.length) {
    emptyState(slide, 0.82, 2.94, 5.8, 3.2, p.site_has_no_data_sections ? 'El sitio reporta secciones sin datos en el período.' : 'No hay mix de canales para mostrar.');
  } else {
    slide.addChart('bar', [{
      name: 'Canales',
      labels: mix.map((x) => clip(x.label || x.theme, 18)),
      values: mix.map((x) => parseNumber(x.value || x.weight)),
    }], {
      x: 0.9, y: 3.1, w: 5.45, h: 2.8,
      showLegend: false,
      chartColors: [COLORS.electricBlue],
      catAxisLabelFontFace: 'Lato', catAxisLabelFontSize: 8.5,
      valAxisLabelFontFace: 'Lato', valAxisLabelFontSize: 8,
      valGridLine: { color: COLORS.grey1, width: 1 },
      showValue: true, dataLabelColor: COLORS.electricBlue, dataLabelPosition: 'outEnd',
    });
  }

  panel(slide, 6.98, 2.62, 5.7, 3.72, 'Narrativa');
  slide.addText(clip(p.message || '-', 260), {
    x: 7.22, y: 3.1, w: 5.2, h: 2.9,
    fontFace: 'Lato', fontSize: 11, color: COLORS.midnight, valign: 'mid', margin: 0,
  });
}

function renderMix(module) {
  const p = module.payload || {};
  const strategic = Array.isArray(p.strategic_axes) ? p.strategic_axes.slice(0, 5) : [];
  const clients = Array.isArray(p.internal_clients) ? p.internal_clients.slice(0, 5) : [];
  const formats = Array.isArray(p.format_mix) ? p.format_mix.slice(0, 5) : [];
  const slide = baseSlide(module.title || 'Mix temático y áreas solicitantes', 'Contenido');

  panel(slide, 0.62, 1.4, 3.95, 4.94, 'Ejes estratégicos');
  if (!strategic.length) {
    emptyState(slide, 0.84, 2.0, 3.5, 4.1, 'No hay distribución temática consolidada.');
  } else {
    slide.addChart('doughnut', [{
      name: 'Ejes',
      labels: strategic.map((x) => clip(x.label || x.theme, 14)),
      values: strategic.map((x) => parseNumber(x.value || x.weight)),
    }], {
      x: 0.9, y: 2.0, w: 2.0, h: 2.1,
      holeSize: 68,
      chartColors: [COLORS.electricBlue, COLORS.sereneBlue, COLORS.ice, COLORS.lime, COLORS.canary],
      showLegend: false,
      showValue: false,
    });
    strategic.forEach((row, idx) => {
      slide.addText(`${idx + 1}. ${clip(row.label || row.theme, 16)} ${fmtPct(row.value || row.weight)}`, {
        x: 3.0, y: 2.06 + idx * 0.34, w: 1.4, h: 0.16,
        fontFace: 'Lato', fontSize: 8.5, color: COLORS.midnight, margin: 0,
      });
    });
  }

  panel(slide, 4.82, 1.4, 3.95, 4.94, 'Áreas solicitantes');
  if (!clients.length) {
    emptyState(slide, 5.04, 2.0, 3.5, 4.1, 'No hay datos de áreas solicitantes.');
  } else {
    slide.addChart('bar', [{
      name: 'Áreas',
      labels: clients.map((x) => clip(x.label, 14)),
      values: clients.map((x) => parseNumber(x.value)),
    }], {
      x: 5.1, y: 2.0, w: 3.4, h: 2.95,
      showLegend: false,
      chartColors: [COLORS.sereneBlue],
      catAxisLabelFontFace: 'Lato', catAxisLabelFontSize: 8,
      valAxisLabelFontFace: 'Lato', valAxisLabelFontSize: 8,
      showValue: true, dataLabelColor: COLORS.electricBlue, dataLabelPosition: 'outEnd',
    });
  }

  panel(slide, 9.02, 1.4, 3.66, 4.94, 'Mix de formatos');
  if (!formats.length) {
    emptyState(slide, 9.2, 2.0, 3.3, 2.5, 'No hay mix de formatos disponible.');
  } else {
    tableRows(slide, 9.24, 2.02, 3.2, ['Formato', 'Peso'], formats.map((x) => [x.label || x.theme || '-', fmtPct(x.value || x.weight)]));
  }
  slide.addText(clip(p.message || '-', 145), {
    x: 9.24, y: 4.95, w: 3.2, h: 1.1,
    fontFace: 'Lato', fontSize: 10.3, color: COLORS.midnight, margin: 0,
  });
}

function renderPushRanking(module) {
  const p = module.payload || {};
  const interaction = Array.isArray(p.by_interaction) ? p.by_interaction.slice(0, 5) : [];
  const openRate = Array.isArray(p.by_open_rate) ? p.by_open_rate.slice(0, 5) : [];
  const slide = baseSlide(module.title || 'Ranking push', 'Push');

  if (!p.available || (!interaction.length && !openRate.length)) {
    emptyState(slide, 1.0, 2.0, 11.3, 3.4, 'No hay ranking push suficiente para el período.');
    return;
  }

  panel(slide, 0.62, 1.4, 6.08, 4.94, 'Top por interacción');
  tableRows(
    slide,
    0.84,
    2.0,
    5.6,
    ['Comunicación', 'Clics', 'Interacción'],
    interaction.map((row) => [cleanText(row.name), fmtNum(row.clicks), fmtPct(row.interaction)])
  );

  panel(slide, 6.92, 1.4, 5.74, 4.94, 'Top por apertura');
  tableRows(
    slide,
    7.14,
    2.0,
    5.3,
    ['Comunicación', 'Clics', 'Open rate'],
    openRate.map((row) => [cleanText(row.name), fmtNum(row.clicks), fmtPct(row.open_rate)])
  );

  slide.addText(clip(p.message || '-', 180), {
    x: 0.84, y: 6.48, w: 11.5, h: 0.3,
    fontFace: 'Lato', fontSize: 10.5, color: COLORS.midnight, margin: 0,
  });
}

function renderPullRanking(module) {
  const p = module.payload || {};
  const rows = Array.isArray(p.top_pull_notes) ? p.top_pull_notes.slice(0, 5) : [];
  const slide = baseSlide(module.title || 'Ranking pull', 'Site / intranet');

  card(slide, 0.62, 1.4, 2.8, 0.96, 'Promedio lecturas por nota', fmtNum(p.average_reads_per_note), COLORS.sereneBlue);
  card(slide, 3.62, 1.4, 2.8, 0.96, 'Vistas totales site', fmtNum(p.site_total_views), COLORS.ice);

  panel(slide, 0.62, 2.62, 8.2, 3.72, 'Top notas pull');
  if (!p.available || !rows.length) {
    emptyState(slide, 0.84, 2.94, 7.8, 3.2, 'No hay ranking pull suficiente para este período.');
  } else {
    tableRows(
      slide,
      0.9,
      3.0,
      7.8,
      ['Nota', 'Lecturas únicas', 'Lecturas totales'],
      rows.map((row) => [cleanText(row.title), fmtNum(row.unique_reads), fmtNum(row.total_reads)])
    );
  }

  panel(slide, 8.98, 2.62, 3.7, 3.72, 'Lectura ejecutiva');
  slide.addText(clip(p.message || '-', 180), {
    x: 9.2, y: 3.2, w: 3.25, h: 2.9,
    fontFace: 'Lato', fontSize: 11, color: COLORS.midnight, valign: 'mid', margin: 0,
  });
}

function renderMilestones(module) {
  const p = module.payload || {};
  const items = Array.isArray(p.items) ? p.items.slice(0, 3) : [];
  const slide = baseSlide(module.title || 'Hitos del mes', 'Gestión');

  if (!items.length) {
    emptyState(slide, 1.0, 2.0, 11.3, 3.4, 'No se registraron hitos consolidados en este período.');
    return;
  }

  items.forEach((item, idx) => {
    const x = 0.62 + idx * 4.18;
    panel(slide, x, 1.4, 3.56, 4.94, `Hito ${idx + 1}`);
    slide.addText(clip(item.title || item.description || '-', 48), {
      x: x + 0.16, y: 1.92, w: 3.2, h: 0.48,
      fontFace: 'Source Serif 4', bold: true, fontSize: 16, color: COLORS.electricBlue, margin: 0,
    });
    const bullets = Array.isArray(item.bullets) ? item.bullets.filter(Boolean).slice(0, 3) : [];
    const lines = bullets.length ? bullets : [item.description || 'Sin detalle adicional.'];
    lines.forEach((line, lineIdx) => {
      slide.addText(`• ${clip(line, 70)}`, {
        x: x + 0.18, y: 2.64 + lineIdx * 0.46, w: 3.15, h: 0.24,
        fontFace: 'Lato', fontSize: 10.2, color: COLORS.midnight, margin: 0,
      });
    });
    slide.addText(cleanText(item.period || ''), {
      x: x + 0.18, y: 5.94, w: 3.1, h: 0.14,
      fontFace: 'Lato', fontSize: 8.5, color: COLORS.grey4, margin: 0,
    });
  });
}

function renderEvents(module) {
  const p = module.payload || {};
  const events = Array.isArray(p.events) ? p.events.slice(0, 5) : [];
  const slide = baseSlide(module.title || 'Eventos del mes', 'Activaciones');

  card(slide, 0.62, 1.4, 2.6, 0.96, 'Eventos', fmtNum(p.total_events || events.length), COLORS.sereneBlue);
  card(slide, 3.42, 1.4, 2.8, 0.96, 'Participaciones', fmtNum(p.total_participants), COLORS.ice);

  if (!events.length) {
    emptyState(slide, 0.62, 2.62, 7.8, 3.72, 'No hay detalle de eventos suficiente, por lo que este módulo debería omitirse.');
    return;
  }

  panel(slide, 0.62, 2.62, 7.8, 3.72, 'Detalle de eventos');
  tableRows(
    slide,
    0.86,
    3.0,
    7.2,
    ['Evento', 'Participantes', 'Fecha'],
    events.map((row) => [row.name || row.title || '-', fmtNum(row.participants), row.date || '-'])
  );

  panel(slide, 8.62, 2.62, 4.04, 3.72, 'Mensaje');
  slide.addText(clip(p.message || '-', 170), {
    x: 8.88, y: 3.2, w: 3.55, h: 2.9,
    fontFace: 'Lato', fontSize: 11, color: COLORS.midnight, valign: 'mid', margin: 0,
  });
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
          format_mix: [],
          message: s4.conclusion || s4.theme_message,
        },
      },
      {
        key: 'ranking_push',
        title: 'Ranking push',
        payload: {
          by_interaction: s5.top_communications || [],
          by_open_rate: [],
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
  slide.addText(cleanText(s.label || s.period || '-'), { x: 0.8, y: 2.8, w: 6, h: 0.6, fontFace: 'Source Serif 4', bold: true, fontSize: 34, color: COLORS.white, margin: 0 });
  slide.addText('Comunicaciones Internas BBVA', { x: 0.8, y: 3.6, w: 8, h: 0.4, fontFace: 'Lato', fontSize: 16, color: COLORS.white, margin: 0 });
}

function renderFullClosing() {
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.electricBlue };
  slide.addText('Fin del informe', { x: 0.8, y: 3.1, w: 6, h: 0.5, fontFace: 'Source Serif 4', bold: true, fontSize: 30, color: COLORS.white, margin: 0 });
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

pptx.writeFile({ fileName: outputPptxPath });
