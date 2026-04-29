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
  midnight: '07124A',
  deep: '030B31',
  cyan: '00D7E8',
  aqua: '2DCCCD',
  sky: '85C8FF',
  lime: '88E783',
  yellow: 'F8D44C',
  orange: 'F7893B',
  purple: '6754B8',
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
  COLORS.orange,
  COLORS.purple,
  COLORS.lime,
  COLORS.cyan,
  COLORS.yellow,
  COLORS.red,
];

const SECONDARY_CHART_COLORS = [
  COLORS.orange,
  COLORS.purple,
  COLORS.lime,
  COLORS.cyan,
  COLORS.yellow,
  COLORS.red,
];

const CHANNEL_CHART_COLORS = [
  COLORS.electricBlue,
  COLORS.orange,
  COLORS.lime,
  COLORS.purple,
  COLORS.cyan,
  COLORS.yellow,
];

const AXIS_CHART_COLORS = [
  COLORS.electricBlue,
  COLORS.purple,
  COLORS.orange,
  COLORS.lime,
  COLORS.cyan,
  COLORS.yellow,
];

const AREA_CHART_COLORS = [
  COLORS.electricBlue,
  COLORS.orange,
  COLORS.purple,
  COLORS.lime,
  COLORS.cyan,
  COLORS.yellow,
  COLORS.grey2,
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

function stripEmoji(value) {
  return String(value ?? '').replace(/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/gu, '').replace(/\s+/g, ' ').trim();
}

function repairSpanishText(value) {
  let text = String(value ?? '');
  const replacements = [
    [/\bIncentivaci\s+n\b/g, 'Incentivación'],
    [/\bincentivaci\s+n\b/g, 'incentivación'],
    [/\bEvaluaci\s+n\b/g, 'Evaluación'],
    [/\bevaluaci\s+n\b/g, 'evaluación'],
    [/\bComunicaci\s+n\b/g, 'Comunicación'],
    [/\bcomunicaci\s+n\b/g, 'comunicación'],
    [/\bPlanificaci\s+n\b/g, 'Planificación'],
    [/\bplanificaci\s+n\b/g, 'planificación'],
    [/\bInteracci\s+n\b/g, 'Interacción'],
    [/\binteracci\s+n\b/g, 'interacción'],
    [/\bValid\s+los\s+datos\b/g, 'Validá los datos'],
    [/\bvalid\s+los\s+datos\b/g, 'validá los datos'],
    [/\bAcced\s+a\s+tu\s+info\b/g, 'Accedé a tu info'],
    [/\bacced\s+a\s+tu\s+info\b/g, 'accedé a tu info'],
    [/\bConoc\s+los\s+resultados\b/g, 'Conocé los resultados'],
    [/\bconoc\s+los\s+resultados\b/g, 'conocé los resultados'],
    [/\bPlanific\s+tu\s+mes\b/g, 'Planificá tu mes'],
    [/\bplanific\s+tu\s+mes\b/g, 'planificá tu mes'],
    [/\bProteg\s+tu\b/g, 'Protegé tu'],
    [/\bproteg\s+tu\b/g, 'protegé tu'],
    [/\bCuid\s+nuestra\b/g, 'Cuidá nuestra'],
    [/\bcuid\s+nuestra\b/g, 'cuidá nuestra'],
    [/\bMa\s+ana\b/g, 'Mañana'],
    [/\bma\s+ana\b/g, 'mañana'],
    [/\bten\s+s\b/g, 'tenés'],
    [/\bTen\s+s\b/g, 'Tenés'],
    [/\bAs\s+vivimos\b/g, 'Así vivimos'],
    [/\bas\s+vivimos\b/g, 'así vivimos'],
    [/\bD\s+a\s+Interna/g, 'Día Interna'],
    [/\bd\s+a\s+interna/g, 'día interna'],
    [/\bm\s+sica\b/g, 'música'],
    [/\bM\s+sica\b/g, 'Música'],
    [/\bm\s+s\s+cerca\b/g, 'más cerca'],
    [/\bM\s+s\s+cerca\b/g, 'Más cerca'],
    [/\bltimos\s+d\s+as\b/g, 'últimos días'],
    [/\bLtimos\s+d\s+as\b/g, 'Últimos días'],
    [/\bacad\s+mico\b/g, 'académico'],
    [/\bAcad\s+mico\b/g, 'Académico'],
    [/\bacad\s+mica\b/g, 'académica'],
    [/\bAcad\s+mica\b/g, 'Académica'],
    [/\bacompa\s+ando\b/g, 'acompañando'],
    [/\bAcompa\s+ando\b/g, 'Acompañando'],
    [/\blegi\s+n\b/g, 'legión'],
    [/\bLegi\s+n\b/g, 'Legión'],
    [/\bM\s+s\s+formaciones\b/g, 'Más formaciones'],
    [/\bm\s+s\s+formaciones\b/g, 'más formaciones'],
  ];
  for (const [pattern, replacement] of replacements) text = text.replace(pattern, replacement);
  return text;
}

function cleanText(value, fallback = '-') {
  let raw = String(value ?? '').replace(/_/g, ' ').replace(/\s+/g, ' ').trim();
  raw = repairSpanishText(raw);
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

function splitPeriodDisplay(label) {
  const raw = cleanText(label || periodLabel(), '-');
  const match = raw.match(/^(.+?)\s*(\([^)]*\))$/);
  if (match) {
    return { primary: match[1].trim(), secondary: match[2].trim() };
  }
  return { primary: raw, secondary: '' };
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

function normalizeAreaLabel(label) {
  const text = cleanText(label, '');
  const replacements = [
    [/^Relaciones$/i, 'Relaciones Institucionales'],
    [/^Relaciones Institu(?:…|\.{3})?$/i, 'Relaciones Institucionales'],
    [/^Country Manager$/i, 'Country Manager Office'],
    [/^Country Manager Office.*$/i, 'Country Manager Office'],
    [/^Internal Control.*$/i, 'Internal Control & Compliance'],
  ];
  for (const [pattern, replacement] of replacements) {
    if (pattern.test(text)) return replacement;
  }
  return text;
}

function valueLabel(row, rows, options = {}) {
  if (options.valueMode === 'number') return fmtNum(row.value, options.digits || 0);
  if (options.valueMode === 'percent') return fmtPct(row.value);
  return valueAsPct(row, rows);
}

function chartColor(idx, options = {}) {
  if (options.colors) return options.colors[idx % options.colors.length];
  if (idx === 0) return COLORS.electricBlue;
  return SECONDARY_CHART_COLORS[(idx - 1) % SECONDARY_CHART_COLORS.length];
}

function hasIncompleteEnding(value) {
  const text = cleanText(value, '').replace(/…/g, '').trim();
  if (!text) return true;
  return /(\b(?:a|al|ante|bajo|con|contra|de|del|desde|durante|e|en|entre|hacia|hasta|mediante|o|para|por|según|sin|sobre|tras|u|y)\.?$|\b(?:nuestro|nuestra|nuestros|nuestras|este|esta|estos|estas|ese|esa|esos|esas|su|sus)\.?$)/i.test(text);
}

function completeSentence(value, max = 120) {
  const text = ensureSentence(cleanText(value, '')).replace(/…/g, '').trim();
  if (!text || hasIncompleteEnding(text)) return '';
  if (text.length <= max) return text;
  const firstSentence = text.split(/(?<=[.!?])\s+/)[0];
  if (firstSentence && firstSentence.length <= max && !hasIncompleteEnding(firstSentence)) return ensureSentence(firstSentence);
  const clause = text.split(/[,;:]/)[0];
  if (clause && clause.length >= 28 && clause.length <= max && !hasIncompleteEnding(clause)) return ensureSentence(clause);
  return '';
}

function shortSentence(value, max = 120) {
  const complete = completeSentence(value, max);
  if (complete) return complete;
  const text = ensureSentence(cleanText(value, '')).replace(/…/g, '').trim();
  if (!text) return '';
  const cut = text.slice(0, max).split(' ').slice(0, -1).join(' ').trim();
  if (!cut || hasIncompleteEnding(cut)) return '';
  return ensureSentence(cut);
}

function hasMetric(value) {
  return value !== null && value !== undefined && value !== '' && value !== '-';
}

function isCompletePushRow(row, metricKey = 'interaction') {
  if (!row || typeof row !== 'object') return false;
  const metric = parseNumber(row[metricKey] || (metricKey === 'interaction' ? row.ctr : row.interaction));
  const clicks = parseNumber(row.clicks);
  if (metric <= 0) return false;
  if (metricKey === 'interaction' && clicks <= 0 && metric > 20) return false;
  return true;
}

function usablePushRows(rows, metricKey = 'interaction', limit = 5) {
  if (!Array.isArray(rows)) return [];
  const seen = new Set();
  return rows
    .filter((row) => isCompletePushRow(row, metricKey))
    .map((row) => ({ ...row, name: cleanText(row.name || row.title || 'Sin título') }))
    .filter((row) => {
      const key = `${_compact(row.name)}-${Math.round(parseNumber(row[metricKey] || row.ctr || row.interaction) * 100)}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .sort((a, b) => parseNumber(b[metricKey] || b.ctr || b.interaction) - parseNumber(a[metricKey] || a.ctr || a.interaction))
    .slice(0, limit);
}

function _compact(value) {
  return cleanText(value, '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '');
}

function clicksLabel(row) {
  const clicks = parseNumber(row?.clicks);
  if (clicks <= 0) return 'Datos incompletos';
  return `${fmtNum(clicks)} clics`;
}

function ensureSentence(text) {
  const t = cleanText(text, '').trim();
  if (!t) return '';
  if (/[.!?]$/.test(t)) return t;
  return `${t}.`;
}

function bestPushCampaign() {
  const rankings = report?.kpis?.consolidated_rankings?.top_push_by_interaction || report?.render_plan?.modules?.find((m) => m.key === 'ranking_push')?.payload?.by_interaction || [];
  const usable = usablePushRows(rankings, 'interaction', 1);
  return usable.length ? usable[0] : null;
}

function averageMailInteraction() {
  const totals = report?.kpis?.calculated_totals || {};
  const summary = report?.render_plan?.modules?.find((m) => m.key === 'executive_summary')?.payload || {};
  return parseNumber(totals.mail_interaction_rate || totals.average_interaction_rate || summary.mail_interaction_rate);
}

function periodFrequencyLabel() {
  const slug = cleanText(report?.period?.slug || report?.render_plan?.period?.slug || '', '').toLowerCase();
  const label = cleanText(report?.period?.label || report?.render_plan?.period?.label || periodLabel(), '').toLowerCase();
  if (/year|año|anual/.test(slug) || /año|anual/.test(label)) return 'anual';
  if (/quarter|trimestre|q[1-4]/.test(slug) || /trimestre|q[1-4]/.test(label)) return 'trimestral';
  return 'mensual';
}

function coverSubtitle() {
  const freq = periodFrequencyLabel();
  return `Informe ejecutivo ${freq} de desempeño de canales, agenda y contenidos.`;
}

function upliftLabel(row, benchmark = averageMailInteraction()) {
  const interaction = parseNumber(row?.interaction || row?.ctr);
  const base = parseNumber(benchmark);
  if (interaction <= 0 || base <= 0 || interaction < base * 1.2) return '';
  const ratio = interaction / base;
  return `${new Intl.NumberFormat('es-AR', { maximumFractionDigits: 1 }).format(ratio)}x vs prom.`;
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
    x: x + 0.18, y: y + 0.10, w: w - 0.36, h: 0.12,
    fontFace: 'Arial', fontSize: options.labelSize || 7.6, color: COLORS.muted, margin: 0, fit: 'shrink', breakLine: false,
  });
  slide.addText(String(value ?? '-'), {
    x: x + 0.18, y: y + 0.28, w: w - 0.36, h: Math.max(0.24, h - 0.38),
    fontFace: 'Georgia', bold: true, fontSize: options.valueSize || 18,
    color: COLORS.electricBlue, margin: 0, fit: 'shrink', breakLine: false, valign: 'mid',
  });
}

function heroMetric(slide, x, y, w, h, label, value, detail, accent = COLORS.electricBlue, options = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: options.fill || COLORS.white },
    line: { color: options.line || COLORS.grey1, width: 0.75 },
    shadow: options.shadow === false ? undefined : safeOuterShadow('000000', 0.10, 45, 0.8, 0.24),
  });
  slide.addShape(pptx.ShapeType.rect, { x, y, w: 0.10, h, fill: { color: accent }, line: { color: accent } });
  slide.addText(cleanText(label), {
    x: x + 0.28, y: y + 0.14, w: w - 0.58, h: 0.14,
    fontFace: 'Arial', fontSize: options.labelSize || 8.0, bold: true, color: COLORS.muted,
    margin: 0, fit: 'shrink', breakLine: false,
  });
  slide.addText(String(value ?? '-'), {
    x: x + 0.28, y: y + 0.34, w: w - 0.58, h: 0.38,
    fontFace: 'Georgia', bold: true, fontSize: options.valueSize || 24,
    color: accent, margin: 0, fit: 'shrink', breakLine: false, valign: 'mid',
  });
  if (detail) {
    slide.addShape(pptx.ShapeType.line, { x: x + 0.28, y: y + h - 0.26, w: w - 0.56, h: 0, line: { color: COLORS.grey2, width: 0.6 } });
    slide.addText(shortSentence(detail, options.detailMax || 44), {
      x: x + 0.30, y: y + h - 0.19, w: w - 0.60, h: 0.08,
      fontFace: 'Arial', fontSize: options.detailSize || 6.2, color: COLORS.ink,
      margin: 0, fit: 'shrink', breakLine: false,
    });
  }
}


function topCampaignCard(slide, x, y, w, h, row, options = {}) {
  const accent = options.accent || COLORS.electricBlue;
  const title = stripEmoji(row?.name || row?.title || 'Campaña líder');
  const valueSize = Math.min(options.valueSize || 29, 30);
  const labelSize = options.labelSize || 7.6;
  const titleSize = Math.min(options.titleSize || 9.8, 10.6);
  const metaSize = options.metaSize || 5.9;
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: options.fill || COLORS.paleBlue },
    line: { color: options.line || COLORS.paleBlue, width: 0.75 },
    shadow: safeOuterShadow('000000', 0.08, 45, 0.62, 0.18),
  });
  slide.addShape(pptx.ShapeType.rect, { x, y, w: 0.12, h, fill: { color: accent }, line: { color: accent } });

  const left = x + 0.32;
  const innerW = w - 0.66;
  slide.addText(cleanText(options.label || 'Top campaña'), {
    x: left, y: y + 0.15, w: innerW, h: 0.16,
    fontFace: 'Arial', fontSize: labelSize, bold: true, color: COLORS.muted, margin: 0,
    fit: 'shrink', breakLine: false,
  });
  slide.addText(fmtPct(row?.interaction || row?.ctr), {
    x: left, y: y + 0.38, w: innerW, h: 0.40,
    fontFace: 'Georgia', fontSize: valueSize, bold: true, color: accent,
    margin: 0, fit: 'shrink', breakLine: false,
  });
  slide.addText('interacción', {
    x: left + 0.02, y: y + 0.82, w: 1.32, h: 0.13,
    fontFace: 'Arial', fontSize: options.unitSize || 6.2, bold: true, color: COLORS.muted,
    margin: 0, fit: 'shrink', breakLine: false,
  });

  const meta = [];
  if (parseNumber(row?.clicks) > 0) meta.push(`${fmtNum(row.clicks)} clics`);
  if (parseNumber(row?.open_rate) > 0) meta.push(`${fmtPct(row.open_rate)} apertura`);
  const uplift = upliftLabel(row, options.benchmark);
  if (uplift) meta.push(uplift);
  const metaText = meta.join(' · ') || 'Datos completos no disponibles';

  const metaY = y + h - 0.24;
  const lineY = y + h - 0.33;
  const titleY = y + 1.05;
  const titleH = Math.max(0.30, lineY - titleY - 0.09);
  slide.addText(wrapText(title, options.titleLine || 25, options.titleLines || 2), {
    x: left, y: titleY, w: innerW, h: titleH,
    fontFace: 'Georgia', fontSize: titleSize, bold: true, color: COLORS.midnight,
    margin: 0, fit: 'shrink', valign: 'mid', breakLine: false,
  });
  slide.addShape(pptx.ShapeType.line, { x: left, y: lineY, w: innerW, h: 0, line: { color: COLORS.grey2, width: 0.6 } });
  slide.addText(metaText, {
    x: left, y: metaY, w: innerW, h: 0.13,
    fontFace: 'Arial', fontSize: metaSize, bold: true, color: COLORS.ink,
    margin: 0, fit: 'shrink', breakLine: false,
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
    slide.addText(shortSentence(item, options.max || 170), {
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
    const color = chartColor(idx, options);
    const label = wrapText(row.label, options.labelMax || 24, options.labelLines || 1);
    const val = valueLabel(row, data, options);
    const labelW = w * (options.labelWidth || 0.38);
    const gap = w * 0.03;
    const valueW = w * (options.valueWidth || 0.12);
    const barX = x + labelW + gap;
    const barWMax = w - labelW - gap - valueW - 0.04;
    const barH = options.barH || 0.15;
    slide.addText(label, { x, y: barY - 0.02, w: labelW - 0.05, h: Math.max(0.2, rowH - 0.05), fontFace: 'Arial', fontSize: options.catSize || 7.3, color: COLORS.ink, margin: 0, fit: 'shrink' });
    slide.addShape(pptx.ShapeType.roundRect, { x: barX, y: barY + 0.035, w: barWMax, h: barH, rectRadius: 0.025, fill: { color: COLORS.grey1 }, line: { color: COLORS.grey1 } });
    const barW = barWMax * (parseNumber(row.value) / max);
    slide.addShape(pptx.ShapeType.roundRect, { x: barX, y: barY + 0.035, w: Math.max(0.03, barW), h: barH, rectRadius: 0.025, fill: { color }, line: { color } });
    slide.addText(val, { x: barX + barWMax + 0.05, y: barY - 0.005, w: valueW, h: 0.18, align: 'right', fontFace: 'Arial', bold: true, fontSize: options.dataSize || 7.3, color: idx === 0 ? COLORS.electricBlue : COLORS.ink, margin: 0, fit: 'shrink' });
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

function renderLegend(slide, x, y, w, rows, options = {}) {
  const data = rows.filter((row) => parseNumber(row.value) > 0).slice(0, options.limit || 6);
  data.forEach((row, idx) => {
    const yy = y + idx * (options.step || 0.34);
    const color = chartColor(idx, options);
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y: yy + 0.02, w: 0.16, h: 0.16, rectRadius: 0.03,
      fill: { color }, line: { color },
    });
    slide.addText(clip(row.label, options.labelMax || 28), {
      x: x + 0.24, y: yy - 0.005, w: w - 0.84, h: 0.17,
      fontFace: 'Arial', fontSize: options.fontSize || 7.5, color: COLORS.ink,
      margin: 0, fit: 'shrink', breakLine: false,
    });
    slide.addText(valueLabel(row, data, { valueMode: options.valueMode || 'percent' }), {
      x: x + w - 0.58, y: yy - 0.005, w: 0.58, h: 0.17,
      fontFace: 'Arial', bold: true, fontSize: options.fontSize || 7.5,
      color: idx === 0 ? COLORS.electricBlue : COLORS.muted,
      align: 'right', margin: 0, fit: 'shrink', breakLine: false,
    });
  });
}

function renderDonutChart(slide, x, y, w, h, rows, title = 'Mix', options = {}) {
  const data = rows.filter((row) => parseNumber(row.value) > 0).slice(0, options.limit || 6);
  if (!data.length) {
    paragraph(slide, x, y + h / 2 - 0.1, w, 0.24, 'Sin datos para graficar.', { fontSize: 9, color: COLORS.muted });
    return;
  }
  const colors = data.map((_, idx) => chartColor(idx, options));
  try {
    slide.addChart(pptx.ChartType.doughnut, [{ name: title, labels: data.map((r) => r.label), values: data.map((r) => parseNumber(r.value)) }], {
      x, y, w, h,
      holeSize: 68,
      showLegend: false,
      showValue: false,
      showCategoryName: false,
      showPercent: false,
      firstSliceAng: 270,
      chartColors: colors,
      showTitle: false,
      showLeaderLines: false,
      showBorder: false,
      showCatName: false,
    });

    if (options.showSliceLabels) {
      const total = sumValues(data) || 1;
      const cx = x + w / 2;
      const cy = y + h / 2;
      const rx = w * 0.47;
      const ry = h * 0.47;
      let angle = -Math.PI / 2;
      data.forEach((row, idx) => {
        const frac = parseNumber(row.value) / total;
        const mid = angle + frac * Math.PI;
        if (frac >= (options.minSliceLabelPct || 0.06)) {
          const lx = cx + Math.cos(mid) * rx;
          const ly = cy + Math.sin(mid) * ry;
          const boxW = 0.62;
          const boxH = 0.18;
          const textX = Math.max(x - 0.05, Math.min(x + w - boxW + 0.05, lx - boxW / 2));
          const textY = Math.max(y - 0.02, Math.min(y + h - boxH + 0.02, ly - boxH / 2));
          slide.addText(fmtPct(row.value), {
            x: textX, y: textY, w: boxW, h: boxH,
            fontFace: 'Arial', bold: true, fontSize: options.percentSize || 8.0,
            color: idx === 0 ? COLORS.electricBlue : COLORS.ink,
            margin: 0, fit: 'shrink', align: 'center', valign: 'mid',
            breakLine: false,
          });
        }
        angle += frac * Math.PI * 2;
      });
    }
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
  const period = splitPeriodDisplay(s.label || s.period || periodLabel());
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.electricBlue };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: COLORS.electricBlue }, line: { color: COLORS.electricBlue } });
  slide.addShape(pptx.ShapeType.arc, { x: 8.95, y: -1.05, w: 5.8, h: 5.8, adjustPoint: 0.18, line: { color: COLORS.cyan, transparency: 50, width: 2.2 } });
  slide.addShape(pptx.ShapeType.arc, { x: 9.6, y: -0.15, w: 4.2, h: 4.2, adjustPoint: 0.28, line: { color: COLORS.sky, transparency: 55, width: 1.4 } });
  addLogo(slide, 'white', 0.38, 0.28, 1.28, 0.42);

  // decorative bars / icon
  const bars = [
    { x: 7.78, y: 1.30, w: 0.52, h: 1.18, c: COLORS.sky },
    { x: 8.47, y: 0.88, w: 0.52, h: 1.60, c: COLORS.cyan },
    { x: 9.16, y: 1.10, w: 0.52, h: 1.38, c: COLORS.sky },
  ];
  bars.forEach((b) => slide.addShape(pptx.ShapeType.roundRect, {
    x: b.x, y: b.y, w: b.w, h: b.h, rectRadius: 0.06,
    fill: { color: b.c, transparency: 10 }, line: { color: COLORS.white, transparency: 65, width: 1.1 },
  }));

  slide.addText(period.primary, { x: 0.38, y: 2.50, w: 2.4, h: 0.22, fontFace: 'Georgia', fontSize: 12.5, color: COLORS.white, margin: 0, fit: 'shrink' });
  if (period.secondary) {
    slide.addText(period.secondary, { x: 0.38, y: 2.73, w: 2.4, h: 0.18, fontFace: 'Arial', fontSize: 8.8, color: COLORS.sky, margin: 0, fit: 'shrink' });
  }
  slide.addText('Informe de gestión', { x: 0.38, y: 2.98, w: 3.1, h: 0.18, fontFace: 'Arial', fontSize: 8.9, color: COLORS.white, margin: 0, fit: 'shrink' });
  slide.addShape(pptx.ShapeType.line, { x: 0.38, y: 3.24, w: 9.4, h: 0, line: { color: COLORS.white, width: 0.6, transparency: 34 } });

  slide.addText('Comunicaciones\ninternas', {
    x: 0.38, y: 3.46, w: 7.4, h: 2.05,
    fontFace: 'Arial', bold: true, fontSize: 40, color: COLORS.white,
    margin: 0, breakLine: true, fit: 'shrink',
  });
  slide.addText(coverSubtitle(), {
    x: 0.42, y: 5.82, w: 5.95, h: 0.24,
    fontFace: 'Arial', fontSize: 8.6, color: COLORS.sky, margin: 0, fit: 'shrink',
  });
  finalizeSlide(slide);
}

function renderExecutiveSummary(module) {
  const p = module.payload || {};
  const slide = baseSlide(module.title || 'Resumen ejecutivo', 'Resumen');
  const best = bestPushCampaign() || usablePushRows([{
    name: p.top_campaign_title,
    interaction: p.top_campaign_interaction,
    open_rate: p.top_campaign_open_rate,
    clicks: p.top_campaign_clicks,
  }], 'interaction', 1)[0];
  const headline = cleanText(
    p.executive_headline ||
    (best ? `${periodLabel()}: alto engagement impulsado por campañas de beneficios` : `${periodLabel()}: lectura ejecutiva del período`)
  );

  slide.addText(clip(headline, 105), {
    x: 0.72, y: 1.26, w: 11.1, h: 0.40,
    fontFace: 'Georgia', bold: true, fontSize: 21, color: COLORS.midnight, margin: 0, fit: 'shrink', breakLine: false,
  });

  if (best) {
    topCampaignCard(slide, 0.72, 1.94, 4.55, 2.10, best, { label: 'Top campaña', titleLine: 25, titleLines: 2, titleSize: 10.0, valueSize: 29, metaSize: 5.9, benchmark: p.mail_interaction_rate });
  } else {
    emptyState(slide, 0.72, 1.94, 4.55, 2.10, 'Top campaña', 'Ranking push no disponible con datos completos.');
  }

  metricTile(slide, 5.55, 1.94, 2.28, 0.86, 'Apertura promedio', fmtPct(p.mail_open_rate), COLORS.electricBlue, { valueSize: 21, labelSize: 7.8, fill: COLORS.white });
  metricTile(slide, 5.55, 3.00, 2.28, 0.86, 'Interacción promedio', fmtPct(p.mail_interaction_rate), COLORS.cyan, { valueSize: 21, labelSize: 7.8, fill: COLORS.white });

  panel(slide, 8.26, 1.94, 4.02, 3.05, 'Lectura ejecutiva', { fill: COLORS.white, shadow: false });
  const defaultInsights = [
    best ? 'Beneficios y servicios generaron la respuesta más alta.' : 'Priorizar piezas con métricas completas para el seguimiento.',
    'SITE funciona como soporte para profundizar contenidos.',
    'Revisar el balance temático antes del próximo cierre.',
  ];
  const insights = buildExecutiveInsights(p).map((item) => shortSentence(item, 96)).filter(Boolean);
  bulletList(slide, 8.56, 2.46, 3.30, insights.length ? insights : defaultInsights, { max: 96, step: 0.58, fontSize: 8.4, itemH: 0.38, bulletColor: COLORS.electricBlue });
  slide.addShape(pptx.ShapeType.roundRect, { x: 8.56, y: 4.48, w: 3.28, h: 0.28, rectRadius: 0.05, fill: { color: COLORS.paleYellow }, line: { color: COLORS.paleYellow } });
  slide.addText('Implicancia clave: replicar beneficio claro + CTA visible.', { x: 8.72, y: 4.575, w: 2.96, h: 0.11, fontFace: 'Arial', bold: true, fontSize: 6.95, color: COLORS.midnight, margin: 0, fit: 'shrink' });

  const volumeRows = [
    { label: 'Planificación', value: parseNumber(p.plan_total) },
    { label: 'Mails', value: parseNumber(p.mail_total) },
    { label: 'Notas SITE', value: parseNumber(p.site_notes_total) },
  ];
  renderHorizontalBarChart(slide, 1.04, 4.88, 6.25, 0.62, volumeRows, { valueMode: 'number', labelMax: 16, catSize: 7.0, dataSize: 7.0, labelWidth: 0.34, barH: 0.11, colors: [COLORS.electricBlue, COLORS.lime, COLORS.purple] });

  slide.addText(`Período: ${periodLabel()} · fuente: dashboard mensual consolidado`, { x: 0.72, y: 6.72, w: 7.8, h: 0.18, fontFace: 'Arial', fontSize: 7.6, color: COLORS.muted, margin: 0 });
  finalizeSlide(slide);
}

function renderChannelManagement(module) {
  const p = module.payload || {};
  const slide = baseSlide(module.title || 'Gestión de canales', 'Canales');
  const mixRows = weightedRows(p.channel_mix, 5);
  const lead = mixRows[0];
  const headline = lead ? `${lead.label} lidera el mix con ${valueAsPct(lead, mixRows)} de participación.` : 'El mix de canales consolida el alcance del período.';

  slide.addText(clip(headline, 106), { x: 0.72, y: 1.28, w: 10.8, h: 0.34, fontFace: 'Georgia', bold: true, fontSize: 19, color: COLORS.midnight, margin: 0, fit: 'shrink' });

  panel(slide, 0.72, 1.88, 7.15, 4.38, 'Mix de canales', { fill: COLORS.white, shadow: false });
  renderDonutChart(slide, 0.96, 2.34, 3.00, 2.95, mixRows, 'Mix de canales', { limit: 5, colors: CHANNEL_CHART_COLORS });
  renderLegend(slide, 4.30, 2.42, 3.05, mixRows, { limit: 5, labelMax: 24, fontSize: 7.5, colors: CHANNEL_CHART_COLORS });
  if (lead) {
    slide.addShape(pptx.ShapeType.roundRect, { x: 1.18, y: 5.48, w: 6.20, h: 0.34, rectRadius: 0.06, fill: { color: COLORS.paleBlue }, line: { color: COLORS.paleBlue } });
    slide.addText(`${lead.label}: canal principal del período`, { x: 1.40, y: 5.60, w: 5.75, h: 0.10, align: 'center', fontFace: 'Arial', bold: true, fontSize: 7.6, color: COLORS.electricBlue, margin: 0, fit: 'shrink' });
  }

  heroMetric(slide, 8.20, 1.88, 3.95, 1.14, 'Apertura promedio', fmtPct(p.mail_open_rate), `${fmtNum(p.mail_total)} mails enviados`, COLORS.electricBlue, { valueSize: 28, detailMax: 80, fill: COLORS.paleBlue, shadow: false });
  metricTile(slide, 8.20, 3.25, 1.85, 0.76, 'Interacción', fmtPct(p.mail_interaction_rate), COLORS.cyan, { valueSize: 17 });
  metricTile(slide, 10.30, 3.25, 1.85, 0.76, 'Vistas SITE', fmtNum(p.site_total_views), COLORS.orange, { valueSize: 17 });

  panel(slide, 8.20, 4.28, 3.95, 1.98, 'Lectura ejecutiva', { fill: COLORS.paleCyan, shadow: false });
  const bullets = [
    lead ? `${lead.label} concentra la llegada directa.` : 'El mix requiere seguimiento mensual.',
    `Mail mantiene ${fmtPct(p.mail_open_rate)} de apertura promedio.`,
    `SITE e Intranet suman ${fmtNum(p.site_total_views)} vistas.`,
  ];
  bulletList(slide, 8.48, 4.78, 3.35, bullets, { max: 82, step: 0.42, fontSize: 8.1, itemH: 0.30, bulletColor: COLORS.cyan });
  finalizeSlide(slide);
}


function renderMix(module) {
  const p = module.payload || {};
  const slide = baseSlide(module.title || 'Mix temático y áreas solicitantes', 'Contenido');
  const axes = weightedRows(p.strategic_axes, 6);
  const clients = weightedRows(p.internal_clients, 8).map((row) => ({ ...row, label: normalizeAreaLabel(row.label) }));
  const formats = weightedRows(p.format_mix, 4);
  const leadAxis = axes[0];
  const leadClient = clients[0];
  const headline = leadAxis ? `${leadAxis.label} lidera la agenda; ${leadClient ? `${leadClient.label} concentra la demanda.` : 'faltan áreas solicitantes.'}` : 'Distribución temática del período.';

  slide.addText(clip(headline, 106), { x: 0.72, y: 1.28, w: 10.8, h: 0.34, fontFace: 'Georgia', bold: true, fontSize: 19, color: COLORS.midnight, margin: 0, fit: 'shrink' });

  panel(slide, 0.72, 1.86, 5.85, 4.12, 'Ejes estratégicos', { fill: COLORS.white, shadow: false });
  renderHorizontalBarChart(slide, 1.04, 2.42, 5.20, 3.02, axes, { labelMax: 20, labelLines: 1, catSize: 7.3, dataSize: 7.4, valueMode: 'percent', labelWidth: 0.45, valueWidth: 0.13, barH: 0.13, colors: AXIS_CHART_COLORS });

  panel(slide, 6.86, 1.86, 5.48, 4.12, 'Áreas solicitantes', { fill: COLORS.white, shadow: false });
  if (clients.length) {
    renderHorizontalBarChart(slide, 7.18, 2.42, 4.82, 3.02, clients, { labelMax: 30, labelLines: 2, catSize: 6.4, dataSize: 7.0, valueMode: 'percent', labelWidth: 0.60, valueWidth: 0.13, barH: 0.12, colors: AREA_CHART_COLORS });
  } else {
    emptyState(slide, 7.18, 2.34, 4.75, 2.62, 'Dato pendiente', 'No se detectó el bloque de áreas solicitantes en la página de planificación.');
  }

  panel(slide, 0.72, 6.16, 11.62, 0.58, 'Síntesis ejecutiva', { fill: COLORS.midnight, line: COLORS.midnight, shadow: false });
  const formatText = formats[0] ? `Formato líder: ${formats[0].label} (${valueAsPct(formats[0], formats)}).` : 'Formato líder pendiente.';
  const clientText = leadClient ? `Demanda principal: ${leadClient.label} (${fmtPct(leadClient.value)}).` : 'Demanda principal pendiente.';
  const summary = `${leadAxis ? `Agenda: ${leadAxis.label} (${valueAsPct(leadAxis, axes)}).` : ''} ${clientText} ${formatText}`;
  slide.addText(summary.trim(), { x: 1.00, y: 6.39, w: 11.05, h: 0.10, fontFace: 'Arial', bold: true, fontSize: 7.6, color: COLORS.white, margin: 0, fit: 'shrink' });
  finalizeSlide(slide);
}


function renderPushRanking(module) {
  const p = module.payload || {};
  const slide = baseSlide(module.title || 'Ranking push', 'Mail');
  const byInteraction = usablePushRows(p.by_interaction, 'interaction', 5);
  if ((p.available === false && !byInteraction.length) || !byInteraction.length) {
    emptyState(slide, 0.86, 1.55, 11.6, 4.8, 'Ranking no disponible', 'No se detectó un ranking de mails con métricas completas para mostrar.');
    finalizeSlide(slide);
    return;
  }
  const best = byInteraction[0] || {};
  slide.addText(clip(`La campaña líder alcanzó ${fmtPct(best.interaction || best.ctr)} de interacción.`, 104), { x: 0.72, y: 1.28, w: 10.8, h: 0.34, fontFace: 'Georgia', bold: true, fontSize: 19, color: COLORS.midnight, margin: 0, fit: 'shrink' });

  topCampaignCard(slide, 0.72, 1.88, 4.36, 2.10, best, { label: 'Top 1 por interacción', titleLine: 26, titleLines: 2, titleSize: 10.4, valueSize: 30, metaSize: 5.9, benchmark: p.average_interaction_rate });

  byInteraction.slice(1, 3).forEach((row, idx) => {
    const y = 1.88 + idx * 1.04;
    const accent = idx === 0 ? COLORS.orange : COLORS.purple;
    panel(slide, 5.36, y, 3.22, 0.86, '', { fill: COLORS.white, shadow: false });
    slide.addShape(pptx.ShapeType.rect, { x: 5.36, y, w: 0.07, h: 0.86, fill: { color: accent }, line: { color: accent } });
    slide.addText(`Top ${idx + 2}`, { x: 5.58, y: y + 0.10, w: 0.72, h: 0.12, fontFace: 'Arial', bold: true, fontSize: 6.8, color: accent, margin: 0, fit: 'shrink' });
    slide.addText(wrapText(row.name || row.title || '-', 28, 2), { x: 5.58, y: y + 0.30, w: 1.94, h: 0.28, fontFace: 'Arial', bold: true, fontSize: 6.9, color: COLORS.ink, margin: 0, fit: 'shrink' });
    slide.addText(fmtPct(row.interaction || row.ctr), { x: 7.58, y: y + 0.20, w: 0.74, h: 0.21, fontFace: 'Georgia', bold: true, fontSize: 14, color: COLORS.electricBlue, margin: 0, fit: 'shrink', align: 'right' });
    slide.addText(clicksLabel(row), { x: 7.20, y: y + 0.60, w: 1.12, h: 0.10, fontFace: 'Arial', fontSize: 6.4, color: COLORS.muted, margin: 0, fit: 'shrink', align: 'right' });
  });

  panel(slide, 8.92, 1.88, 3.42, 2.10, 'Implicancias clave', { fill: COLORS.paleCyan, shadow: false });
  const bullets = [
    'Beneficio concreto + urgencia elevó la respuesta.',
    `${fmtPct(best.open_rate)} de apertura es referencia aspiracional.`,
    'Replicar CTA único en piezas de servicio.',
  ];
  bulletList(slide, 9.20, 2.38, 2.82, bullets, { max: 76, step: 0.44, fontSize: 7.8, itemH: 0.31, bulletColor: COLORS.cyan });

  panel(slide, 0.72, 4.38, 11.62, 1.84, 'Ranking de interacción', { fill: COLORS.white, shadow: false });
  renderHorizontalBarChart(slide, 1.04, 4.88, 10.70, 1.02, byInteraction.map((row) => ({ label: row.name || row.title, value: parseNumber(row.interaction || row.ctr) })), { valueMode: 'percent', limit: 5, labelMax: 32, labelLines: 1, catSize: 7.1, dataSize: 7.2, labelWidth: 0.34, valueWidth: 0.11, barH: 0.12, colors: [COLORS.electricBlue, COLORS.lime, COLORS.purple, COLORS.orange, COLORS.yellow] });
  finalizeSlide(slide);
}


function renderPullRanking(module) {
  const p = module.payload || {};
  const slide = baseSlide(module.title || 'Ranking pull', 'SITE / Intranet');
  const rows = Array.isArray(p.top_pull_notes) ? p.top_pull_notes.slice(0, 5) : [];
  const best = rows[0] || {};

  slide.addText(clip(best.title ? 'Bienestar y servicios impulsan las lecturas del ecosistema pull.' : 'Ranking pull del período.', 106), { x: 0.72, y: 1.28, w: 10.8, h: 0.34, fontFace: 'Georgia', bold: true, fontSize: 19, color: COLORS.midnight, margin: 0, fit: 'shrink' });

  if (p.available === false && !rows.length) {
    emptyState(slide, 0.72, 1.95, 11.62, 4.1, 'Ranking no disponible', cleanText(p.message || 'No se detectó ranking de notas en la fuente.'));
    finalizeSlide(slide);
    return;
  }

  panel(slide, 0.72, 1.88, 7.45, 4.38, 'Top 5 notas por vistas', { fill: COLORS.white, shadow: false });
  renderHorizontalBarChart(slide, 1.08, 2.48, 6.72, 3.10, rows.map((row) => ({ label: row.title || row.name, value: parseNumber(row.total_reads || row.views) })), { valueMode: 'number', labelMax: 31, labelLines: 2, catSize: 6.8, dataSize: 7.2, labelWidth: 0.53, valueWidth: 0.11, barH: 0.12, colors: [COLORS.electricBlue, COLORS.lime, COLORS.purple, COLORS.orange, COLORS.yellow] });

  heroMetric(slide, 8.46, 1.88, 3.70, 1.14, 'Promedio lecturas/nota', fmtNum(p.average_reads_per_note), `${fmtNum(p.site_total_views)} vistas totales SITE`, COLORS.electricBlue, { valueSize: 28, detailMax: 88, fill: COLORS.paleBlue, shadow: false });

  panel(slide, 8.46, 3.28, 3.70, 2.98, 'Lectura ejecutiva', { fill: COLORS.paleCyan, shadow: false });
  slide.addText(wrapText(best.title || 'Nota líder', 34, 2), { x: 8.76, y: 3.76, w: 3.10, h: 0.42, fontFace: 'Georgia', bold: true, fontSize: 11.2, color: COLORS.electricBlue, margin: 0, fit: 'shrink' });
  slide.addText(`${fmtNum(best.unique_reads || best.users)} usuarios únicos · ${fmtNum(best.total_reads || best.views)} vistas`, { x: 8.76, y: 4.28, w: 3.10, h: 0.15, fontFace: 'Arial', bold: true, fontSize: 7.5, color: COLORS.ink, margin: 0, fit: 'shrink' });
  const bullets = [
    'Bienestar funciona como driver de tráfico.',
    'El ranking pull alimenta envíos segmentados.',
    `${fmtNum(p.average_reads_per_note)} vistas por nota marca la base.`,
  ];
  bulletList(slide, 8.76, 4.76, 3.05, bullets, { max: 76, step: 0.39, fontSize: 7.7, itemH: 0.30, bulletColor: COLORS.sky });
  finalizeSlide(slide);
}


function renderMilestones(module) {
  const p = module.payload || {};
  const items = Array.isArray(p.items) ? p.items.slice(0, 3) : [];
  const slide = baseSlide(module.title || 'Hitos destacados', 'Hitos');
  slide.addText('Momentos relevantes del período', { x: 0.72, y: 1.30, w: 10.8, h: 0.32, fontFace: 'Georgia', bold: true, fontSize: 20, color: COLORS.midnight, margin: 0, fit: 'shrink' });

  if (!items.length) {
    emptyState(slide, 0.72, 2.0, 11.62, 3.8, 'Sin hitos cargados', 'El módulo de hitos requiere contexto manual o detección de eventos cualitativos en la fuente.');
    finalizeSlide(slide);
    return;
  }

  items.forEach((item, idx) => {
    const x = 0.72 + idx * 4.02;
    const accent = chartColor(idx, { colors: [COLORS.electricBlue, COLORS.cyan, COLORS.lime] });
    panel(slide, x, 2.05, 3.58, 3.72, '', { fill: COLORS.white, shadow: true });
    slide.addShape(pptx.ShapeType.rect, { x, y: 2.05, w: 3.58, h: 0.10, fill: { color: accent }, line: { color: accent } });
    if (item.image_path && resolveAsset(item.image_path)) {
      const img = resolveAsset(item.image_path);
      slide.addImage({ path: img, ...imageSizingContain(img, x + 0.25, 2.42, 3.08, 1.40) });
    } else {
      slide.addShape(pptx.ShapeType.ellipse, { x: x + 1.42, y: 2.50, w: 0.72, h: 0.72, fill: { color: accent, transparency: 8 }, line: { color: accent } });
      slide.addText(String(idx + 1), { x: x + 1.42, y: 2.69, w: 0.72, h: 0.16, fontFace: 'Georgia', bold: true, fontSize: 14, color: COLORS.white, align: 'center', margin: 0 });
    }
    slide.addText(wrapText(item.title || item.name || `Hito ${idx + 1}`, 31, 2), { x: x + 0.26, y: 4.10, w: 3.06, h: 0.44, fontFace: 'Georgia', bold: true, fontSize: 12, color: COLORS.electricBlue, margin: 0, fit: 'shrink' });
    paragraph(slide, x + 0.26, 4.64, 3.06, 0.62, item.description || item.detail || 'Acción destacada del período.', { max: 150, fontSize: 8.2 });
  });

  const msg = cleanText(p.message || 'Hitos destacados del período.');
  slide.addShape(pptx.ShapeType.roundRect, { x: 0.72, y: 6.14, w: 11.62, h: 0.46, rectRadius: 0.06, fill: { color: COLORS.paleCyan }, line: { color: COLORS.paleCyan } });
  slide.addText(shortSentence(msg, 170), { x: 1.0, y: 6.31, w: 11.04, h: 0.14, fontFace: 'Arial', bold: true, fontSize: 8.1, color: COLORS.midnight, margin: 0, fit: 'shrink' });
  finalizeSlide(slide);
}

function actionItems(source, defaults, limit = 3) {
  const maxWords = 12;
  const items = Array.isArray(source) ? source.map((item) => completeSentence(item, 92)).filter(Boolean) : [];
  const usable = items
    .filter((item) => item.length <= 96)
    .filter((item) => item.split(/\s+/).length <= maxWords)
    .filter((item) => !hasIncompleteEnding(item))
    .filter((item) => !item.includes('…'));
  return (usable.length >= limit ? usable : defaults).slice(0, limit);
}

function renderRecommendations(module) {
  const p = module.payload || {};
  const recs = actionItems(p.recommendations || p.items, [
    'Replicar campañas de beneficios con CTA directo.',
    'Usar SITE para ampliar bienestar y servicio.',
    'Balancear Innovación con ejes subrepresentados.',
  ], 3);
  const experiments = actionItems(p.experiments, [
    'Probar dos asuntos en una campaña de beneficios.',
    'Testear horario de envío no urgente.',
    'Vincular mail y nota SITE.',
  ], 3);
  const plan = actionItems(p.action_plan || p.plan, [
    'Cerrar calendario editorial de febrero.',
    'Revisar segmentación con Talento y Cultura.',
    'Auditar mails de baja interacción.',
  ], 3);
  const slide = baseSlide(module.title || 'Conclusiones y próximos pasos', 'Plan de mejora');

  slide.addText('Plan 30 días: foco, prueba y seguimiento', { x: 0.72, y: 1.30, w: 9.0, h: 0.34, fontFace: 'Georgia', bold: true, fontSize: 20, color: COLORS.midnight, margin: 0, fit: 'shrink' });
  const summaryText = completeSentence(p.summary || p.message, 105) || 'Priorizar beneficios, bienestar y segmentación para sostener engagement.';
  slide.addText(summaryText, { x: 0.74, y: 1.78, w: 8.0, h: 0.20, fontFace: 'Arial', fontSize: 8.4, color: COLORS.muted, margin: 0, fit: 'shrink' });

  const blocks = [
    { title: 'Quick wins', items: recs, color: COLORS.electricBlue, fill: COLORS.paleBlue },
    { title: 'Experimentos', items: experiments, color: COLORS.lime, fill: COLORS.paleLime },
    { title: 'Plan 30 días', items: plan, color: COLORS.orange, fill: COLORS.paleYellow },
  ];
  blocks.forEach((block, idx) => {
    const x = 0.72 + idx * 4.02;
    panel(slide, x, 2.30, 3.58, 3.58, block.title, { fill: block.fill, shadow: false, headerColor: block.color });
    slide.addShape(pptx.ShapeType.rect, { x, y: 2.30, w: 3.58, h: 0.08, fill: { color: block.color }, line: { color: block.color } });
    bulletList(slide, x + 0.28, 2.90, 3.02, block.items, { limit: 3, step: 0.74, max: 96, fontSize: 8.1, itemH: 0.50, bulletColor: block.color });
  });

  slide.addShape(pptx.ShapeType.roundRect, { x: 0.72, y: 6.20, w: 11.62, h: 0.45, rectRadius: 0.06, fill: { color: COLORS.midnight }, line: { color: COLORS.midnight } });
  const ownerNote = completeSentence(p.owner_note, 130) || 'Seguimiento: apertura, interacción, vistas por nota, ranking de campañas y balance temático.';
  slide.addText(ownerNote, { x: 0.98, y: 6.36, w: 11.02, h: 0.12, fontFace: 'Arial', fontSize: 7.5, bold: true, color: COLORS.white, margin: 0, fit: 'shrink' });
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
  slide.addShape(pptx.ShapeType.arc, { x: -1.1, y: 4.4, w: 4.0, h: 4.0, adjustPoint: 0.28, line: { color: COLORS.cyan, transparency: 55, width: 2.2 } });
  slide.addShape(pptx.ShapeType.arc, { x: 10.7, y: -1.2, w: 4.0, h: 4.0, adjustPoint: 0.28, line: { color: COLORS.sky, transparency: 60, width: 1.6 } });
  addLogo(slide, 'white', 5.53, 2.10, 2.25, 0.74);
  slide.addText('Gracias', { x: 3.2, y: 3.20, w: 6.9, h: 0.66, fontFace: 'Arial', bold: true, fontSize: 28, color: COLORS.white, align: 'center', margin: 0, fit: 'shrink' });
  slide.addText(`Comunicaciones Internas · ${periodLabel()}`, { x: 2.45, y: 4.00, w: 8.4, h: 0.26, fontFace: 'Georgia', bold: true, fontSize: 15.5, color: COLORS.white, align: 'center', margin: 0, fit: 'shrink' });
  slide.addText('Informe automatizado para seguimiento ejecutivo de agenda, canales y resultados.', { x: 2.25, y: 4.42, w: 8.8, h: 0.20, fontFace: 'Arial', fontSize: 8.7, color: COLORS.sky, align: 'center', margin: 0, fit: 'shrink' });
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

fs.mkdirSync(path.dirname(outputPptxPath), { recursive: true });

pptx.writeFile({ fileName: outputPptxPath }).catch((err) => {
  console.error(err);
  process.exit(1);
});