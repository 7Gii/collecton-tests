// Creates the indicators of the muso port (targets.js / targets.extras.js of muso-mali)
// through the collecton-mcp operations. The CHT target functions are copied verbatim from
// targets.extras.js, behind an adapter (CHT report shape, getField on the flat payload,
// Luxon-like dates). Usage:
//   COLLECTON_MCP_MODE=dev node indicators.mjs --project <root> --source <muso-mali> [--plans <indicator-plans.json>]
import fs from 'node:fs';
import path from 'node:path';

const MCP = process.env.COLLECTON_MCP_DIST ?? '/Users/gilbertagbodamakou/Documents/shellit/Dev/collecton/collecton-mcp/dist';
const { loadProject } = await import(`${MCP}/project.js`);
const { assertProjectWritable } = await import(`${MCP}/guards.js`);
const metrics = await import(`${MCP}/ops/metrics.js`);
const planning = await import(`${MCP}/ops/plans.js`);

const args = Object.fromEntries(process.argv.slice(2).reduce((acc, a, i, all) => (a.startsWith('--') ? [...acc, [a.slice(2), all[i + 1]]] : acc), []));
const ROOT = args.project;
const ctx = loadProject(ROOT);
assertProjectWritable(ctx);
const planOf = args.plans ? Object.fromEntries(JSON.parse(fs.readFileSync(args.plans, 'utf8'))) : null;
const extras = fs.readFileSync(path.join(args.source, 'targets.extras.js'), 'utf8');

/** Source of a top-level `function name(...) {...}` of targets.extras.js, verbatim. */
const cht = (name) => {
  const start = extras.indexOf(`function ${name}(`);
  if (start < 0) throw new Error(`function ${name} not found in targets.extras.js`);
  const end = extras.indexOf('\n}\n', start);
  const comment = extras.lastIndexOf('\n//', start);
  const head = comment > 0 && start - comment < 200 ? extras.slice(comment + 1, start) : '';
  return `${head}${extras.slice(start, end + 2)}`;
};

const ADAPTER = `// Ported from targets.js / targets.extras.js of the CHT configuration muso-mali. The CHT
// functions below are copied verbatim; this adapter gives them the CHT shapes:
// report { _id, form, reported_date, fields } and contact { contact, reports }.

const DAY = 86400000;
let NOW = Date.now();

const Forms = {
  PATIENT_ASSESSMENT_UNDER_5: 'patient_assessment',
  PATIENT_ASSESSMENT_OVER_5: 'patient_assessment_over_5',
  TREATMENT_FOLLOWUP_UNDER_5: 'treatment_followup',
  TREATMENT_FOLLOWUP_OVER_5: 'treatment_followup_over_5',
  REFERRAL_FOLLOWUP_UNDER_5: 'referral_followup_under_5',
  REFERRAL_FOLLOWUP_OVER_5: 'referral_followup',
  MODERATE_MALNUTRITION_FOLLOWUP: 'moderate_malnutrition_followup',
  SEVERE_MALNUTRITION_FOLLOWUP: 'severe_malnutrition_followup',
};

/**
 * CHT getField(report, 'group.sub.name') on the flat Collecton payload: the key of a field
 * is its name, or <group>_<name> when the name is used twice in the form.
 */
const getField = (report, fieldPath) => {
  const payload = report?.fields;
  if (!payload) return undefined;
  const parts = String(fieldPath || '').split('.');
  for (let i = 0; i < parts.length; i++) {
    const key = parts.slice(i).join('_');
    if (payload[key] !== undefined) return payload[key];
  }
  return undefined;
};

const isDefined = (value) => value !== null && value !== undefined;

/** Luxon-like date of the CHT code (toMillis, toJSDate, diff in days, comparisons). */
const dt = (millis) => ({
  valueOf: () => millis,
  toMillis: () => millis,
  toJSDate: () => new Date(millis),
  diff: (other, unit) => ({ days: (millis - other.valueOf()) / DAY }),
  toFormat: () => new Date(millis).toISOString().slice(0, 10),
});

/** CHT getReportHomeVisitSubject / getReportHomeVisitDate. */
const getReportHomeVisitSubject = (report) => report && getField(report, 'visited_contact_uuid');
const getReportHomeVisitDate = (report) => {
  if (!getReportHomeVisitSubject(report)) return false;
  const visited = Date.parse(getField(report, 'visited_date') || '');
  return dt(Number.isNaN(visited) ? Number(report.reported_date || 0) : visited);
};

/** CHT getDynamicReportedDate. */
const getDynamicReportedDate = (report) => {
  const home = getReportHomeVisitDate(report);
  if (home) return home;
  const specified = Date.parse(getField(report, 's_reported.s_reported_date') || getField(report, 'supervision_date') || '');
  return dt(Number.isNaN(specified) ? Number((report && report.reported_date) || 0) : specified);
};

/** CHT getNumberOfDaySinceDate, on the app clock. */
function getNumberOfDaySinceDate(specificDate) {
  return Math.floor((NOW - specificDate) / DAY);
}

// ---- Collecton side ----

const formNameById = {};
const toCht = (r) => ({ _id: r.id, form: formNameById[r.form_id], reported_date: Number(r.created_at), fields: r.payload || {}, target: r.target_id });

/** Muso month of the app clock: day 1 to the last day (UTC), as the collection period. */
const musoMonth = (now) => {
  const d = new Date(now);
  return [Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), 1), Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 1) - 1];
};
const inMonth = (millis, [from, to]) => millis >= from && millis <= to;

const parseAttributes = (row) => {
  try { return typeof row?.attributes === 'string' ? JSON.parse(row.attributes) : (row?.attributes || {}); } catch { return {}; }
};

/** CHT basicTemplate: the person of the report is alive (muted has no equivalent). */
const aliveCache = new Map();
const isAliveId = async (app, personId) => {
  if (!personId) return true;
  if (!aliveCache.has(personId)) {
    const [row] = await app.find('persons', [{ field: 'id', value: personId }]);
    const deceased = parseAttributes(row).person_is_deceased;
    aliveCache.set(personId, !(deceased === true || deceased === 'true'));
  }
  return aliveCache.get(personId);
};

/** Prepares a computation: clock, form names, CHT reports of the month for the given forms. */
const prepare = async (app, forms) => {
  NOW = app.now;
  aliveCache.clear();
  const all = app.getReports();
  const idsByName = {};
  for (const name of Object.values(Forms)) {
    const sample = app.getReports({ form: name })[0];
    if (sample) { formNameById[sample.form_id] = name; idsByName[name] = sample.form_id; }
  }
  const reports = all.map(toCht).filter(r => r.form);
  const month = musoMonth(app.now);
  const targets = [];
  for (const r of reports) {
    if (!forms.includes(r.form)) continue;
    if (!inMonth(getDynamicReportedDate(r).toMillis(), month)) continue;
    if (!(await isAliveId(app, r.target))) continue;
    targets.push(r);
  }
  const byPerson = new Map();
  for (const r of reports) {
    if (!byPerson.has(r.target)) byPerson.set(r.target, []);
    byPerson.get(r.target).push(r);
  }
  const contactOf = (r) => ({ contact: { _id: r.target }, reports: byPerson.get(r.target) || [] });
  return { reports, targets, contactOf, month };
};

/** CHT context: user.role === "chw" (the role of the user's contact) -> Collecton role chw_uhc. */
export function showIndicator(app) {
  return app.currentUser?.role?.name === 'chw_uhc';
}
`;

const FUNCTIONS = [
  'isDangerSignNotReferred', 'isFeverWithoutTdrChw', 'malariaNoCta', 'incorrectDosageOfAlu', 'diarrheaWithoutZincErrorChw',
  'incorrectDosageOfZinc', 'pneumoniaWithoutAmoxicillinChw', 'incorrectDosageOfAmoxicillin', 'incorrectDosageOfVitaminA',
  'incorrectDosageOfAlbendazole', 'incorrectDosageOfFerAndAcideFolique', 'mamWithIncompletTreatmentFirstDayChw',
  'masWithoutComplicationIncompleteTreatmentFirstDayChw', 'incorrectDosageOfParacetamol', 'noMalariaFollowup',
  'noDiarrheaFollowup', 'noAriFollowup', 'noCoughFollowup', 'noMultipleTreatmentFollowup', 'aggravatedWithoutReferral',
];
const CHT_FUNCTIONS = `// ---- CHT functions of targets.extras.js (verbatim) ----\n\n${FUNCTIONS.map(cht).join('\n')}`;

const BOTH = "[Forms.PATIENT_ASSESSMENT_UNDER_5, Forms.PATIENT_ASSESSMENT_OVER_5]";
const UNDER5 = "[Forms.PATIENT_ASSESSMENT_UNDER_5]";
const errorTarget = (name, title, forms, passesIf) => ({
  name, title, type: 'percentage', goal: 0,
  value: `/** % of the evaluations of the month where the error is present (CHT passesIf, goal 0). */
export async function value(app) {
  const { targets, contactOf } = await prepare(app, ${forms});
  const passesIf = (c, f) => ${passesIf};
  return { numerator: targets.filter(f => passesIf(contactOf(f), f)).length, denominator: targets.length };
}`,
});

const indicators = [
  {
    name: 'home_visits', title: 'Visite à domicile', type: 'number',
    value: `/** Home visits of the month: distinct (family, day) pairs (CHT idType family~date). */
export async function value(app) {
  const { reports, month } = await prepare(app, []);
  const visits = new Set();
  for (const r of reports) {
    if (!getReportHomeVisitSubject(r)) continue;
    const date = getReportHomeVisitDate(r);
    if (!inMonth(date.toMillis(), month) || !(await isAliveId(app, r.target))) continue;
    visits.add(\`\${getField(r, 'visited_contact_uuid')}~\${date.toFormat('yyyy-MM-dd')}\`);
  }
  return visits.size;
}`,
  },
  {
    name: '2-home-visits-per-fam', title: '% de ménages ayant reçu au moins 2 VAD', type: 'percentage', goal: 100,
    value: `/**
 * Families with at least 2 home visit days in the month, over every family of the device
 * (CHT groupBy family, passesIfGroupCount >= 2, a failing instance per family).
 */
export async function value(app) {
  const { reports, month } = await prepare(app, []);
  const [level] = await app.find('org_unit_levels', [{ field: 'name', value: 'c50_family' }]);
  const families = level ? await app.find('org_units', [{ field: 'level_id', value: level.id }]) : [];
  const days = new Map(families.map(f => [f.id, new Set()]));
  for (const r of reports) {
    if (!getReportHomeVisitSubject(r)) continue;
    const date = getReportHomeVisitDate(r);
    const family = getField(r, 'visited_contact_uuid');
    if (!inMonth(date.toMillis(), month) || !days.has(family)) continue;
    days.get(family).add(date.toFormat('yyyy-MM-dd'));
  }
  return { numerator: [...days.values()].filter(d => d.size >= 2).length, denominator: families.length };
}`,
  },
  errorTarget('percent-assessments-without-protocol-errors', '% des évaluations sans erreur de protocole', BOTH, `isDangerSignNotReferred(f)||
      isDangerSignNotReferred(f)||
      isFeverWithoutTdrChw(f)||
      malariaNoCta(f)||
      incorrectDosageOfAlu(f,'proto')||
      diarrheaWithoutZincErrorChw(f)||
      incorrectDosageOfZinc(f)||
      pneumoniaWithoutAmoxicillinChw(f)||
      incorrectDosageOfAmoxicillin(f)||
      incorrectDosageOfVitaminA(f)||
      incorrectDosageOfAlbendazole(f)||
      incorrectDosageOfFerAndAcideFolique(c,f)||
      mamWithIncompletTreatmentFirstDayChw(f)||
      masWithoutComplicationIncompleteTreatmentFirstDayChw(f)||
      incorrectDosageOfParacetamol(f,'proto')||
      (noMalariaFollowup(c,f,1) || noAriFollowup(c,f,1) || noMultipleTreatmentFollowup(c,f,1))||
      (noMalariaFollowup(c,f,2) || noAriFollowup(c,f,2) || noMultipleTreatmentFollowup(c,f,2))||
      (noMalariaFollowup(c,f,2) || noAriFollowup(c,f,2) || noMultipleTreatmentFollowup(c,f,2))||
      (noMalariaFollowup(c,f,3) || noAriFollowup(c,f,3) || noMultipleTreatmentFollowup(c,f,3))||
      (noMalariaFollowup(c,f,5) || noAriFollowup(c,f,5) || noMultipleTreatmentFollowup(c,f,5))||
      (aggravatedWithoutReferral(c,f,1) || aggravatedWithoutReferral(c,f,2) || aggravatedWithoutReferral(c,f,3) || aggravatedWithoutReferral(c,f,5))`),
  errorTarget('danger-signs-without-referral-chw-error', "Signe(s) de danger sans référence ou accompagnement au centre de santé de premier contact - Erreur de l'ASC", BOTH, 'isDangerSignNotReferred(f)'),
  errorTarget('fever-present-malaria-rdt-not-performed-chw-error', "Présence de fièvre (température mesurée supérieur ou égal à 37,5), TDR paludisme pas réalisé - Erreur de l'ASC", BOTH, 'isFeverWithoutTdrChw(f)'),
  errorTarget('uncomplicated-malaria-without-act-chw-error', "Paludisme simple sans administration de CTA - Erreur de l'ASC", BOTH, 'malariaNoCta(f)'),
  errorTarget('incorrect-alu-dose-chw-error', "Dose de ALU donnée par l'ASC incorrecte - Erreur de l'ASC", BOTH, "incorrectDosageOfAlu(f,'single')"),
  errorTarget('uncomplicated-diarrhea-without-zinc-chw-error', "Diarrhée simple sans administration de zinc - Erreur de l'ASC", UNDER5, 'diarrheaWithoutZincErrorChw(f)'),
  errorTarget('incorrect-zinc-dose-chw-error', "Dose de Zinc donnée par l'ASC incorrecte - Erreur ASC", UNDER5, 'incorrectDosageOfZinc(f)'),
  errorTarget('pneumonia-without-amoxicillin-chw-error', "Pneumonie sans administration d'amoxicilline - Erreur de l'ASC", UNDER5, 'pneumoniaWithoutAmoxicillinChw(f)'),
  errorTarget('incorrect-amoxicillin-dose-chw-error', "Dose d'amoxicilline donnée par l'ASC incorrecte - Erreur ASC", UNDER5, 'incorrectDosageOfAmoxicillin(f)'),
  errorTarget('incorrect-vitamin-a-dose-chw-error', "Dose de vitamine A donnée par l'ASC incorrecte – Erreur ASC", UNDER5, 'incorrectDosageOfVitaminA(f)'),
  errorTarget('incorrect-albendazole-dose-chw-error', "Dose d'Albendazole donnée par l'ASC incorrecte - Erreur ASC", UNDER5, 'incorrectDosageOfAlbendazole(f)'),
  errorTarget('incorrect-iron-folic-acid-dose-chw-error', "Dose de Fer + Acide Folique donnée par l'ASC incorrecte - Erreur ASC", UNDER5, 'incorrectDosageOfFerAndAcideFolique(c,f)'),
  errorTarget('mam-no-complete-treatment-first-day-chw-error', 'MAM et pas de traitement complet et correct au premier jour du traitement – Erreur ASC', UNDER5, 'mamWithIncompletTreatmentFirstDayChw(f)'),
  errorTarget('sam-no-complete-treatment-first-day-chw-error', 'MAS sans complication et pas de traitement complet et correct au premier jour du traitement - Erreur ASC', UNDER5, 'masWithoutComplicationIncompleteTreatmentFirstDayChw(f)'),
  errorTarget('incorrect-paracetamol-dose-chw-error', "Dose de paracétamol donnée par l'ASC incorrecte – Erreur ASC", BOTH, "incorrectDosageOfParacetamol(f,'single')"),
  errorTarget('no-24-hour-follow-up-chw-error', 'Pas de suivi de 24H - Erreur ASC', BOTH, 'noMalariaFollowup(c,f,1) || noAriFollowup(c,f,1) || noMultipleTreatmentFollowup(c,f,1)'),
  errorTarget('no-48-hour-follow-up-chw-error', 'Pas de suivi de 48H - Erreur ASC', BOTH, 'noMalariaFollowup(c,f,2) || noAriFollowup(c,f,2) || noMultipleTreatmentFollowup(c,f,2)'),
  errorTarget('no-72-hour-follow-up-chw-error', 'Pas de suivi de 72H - Erreur ASC', BOTH, 'noMalariaFollowup(c,f,3) || noAriFollowup(c,f,3) || noMultipleTreatmentFollowup(c,f,3)'),
  errorTarget('no-5-day-follow-up-chw-error', 'Pas de suivi de 5eme jour - Erreur ASC', UNDER5, 'noMalariaFollowup(c,f,5) || noAriFollowup(c,f,5) || noMultipleTreatmentFollowup(c,f,5)'),
  errorTarget('condition-worsened-without-referral-chw-error', 'Etat aggravé à 24H ou 48H ou 72H ou 5eme jour sans référence/accompagnement au centre de santé - Erreur ASC', BOTH, 'aggravatedWithoutReferral(c,f,1) || aggravatedWithoutReferral(c,f,2) || aggravatedWithoutReferral(c,f,3) || aggravatedWithoutReferral(c,f,5)'),
];

if (args.list) {
  console.log(JSON.stringify(indicators.map(i => [i.name, i.title, i.type, i.goal ?? null])));
  process.exit(0);
}

const results = [];
for (const indicator of indicators) {
  if (planOf) planning.assertAuthorized(ROOT, planOf[indicator.name], 'indicator', indicator.title);
  const goal = indicator.goal !== undefined ? `\n/** CHT goal. */\nexport function goal() {\n  return ${indicator.goal};\n}\n` : '';
  const script = `${ADAPTER}\n${CHT_FUNCTIONS}\n\n${indicator.value}\n${goal}`;
  const { indicator: created } = metrics.createIndicator(ctx, indicator.title, script, indicator.name);
  const updated = metrics.updateIndicator(ctx, created.name, {
    type: indicator.type,
    view: 'default',
    displayExpression: 'showIndicator',
    valueExpression: 'value',
    description: [`CHT target ${indicator.name}`],
    ...(indicator.goal !== undefined ? { goalExpression: 'goal' } : {}),
  });
  if (planOf) planning.recordGenerated(ROOT, planOf[indicator.name], `indicator: ${indicator.title}`);
  results.push({ name: indicator.name, valid: updated.validation.valid, errors: updated.validation.errors });
}
console.log(JSON.stringify(results.filter(r => !r.valid)));
console.log(`${results.length} indicators, ${results.filter(r => r.valid).length} valid`);
