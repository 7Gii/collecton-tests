// Creates the task rules of the muso port (tasks.js / tasks.extras.js of muso-mali) through
// the collecton-mcp operations, each under its confirmed plan. Usage:
//   COLLECTON_MCP_MODE=dev node tasks.mjs --project <root> --plans <task-plans.json>
// task-plans.json: [[ruleName, planId], ...] as written when the plans were opened.
import fs from 'node:fs';

const MCP = process.env.COLLECTON_MCP_DIST ?? '/Users/gilbertagbodamakou/Documents/shellit/Dev/collecton/collecton-mcp/dist';
const { loadProject } = await import(`${MCP}/project.js`);
const { assertProjectWritable } = await import(`${MCP}/guards.js`);
const taskRules = await import(`${MCP}/ops/tasks.js`);
const planning = await import(`${MCP}/ops/plans.js`);

const args = Object.fromEntries(process.argv.slice(2).reduce((acc, a, i, all) => (a.startsWith('--') ? [...acc, [a.slice(2), all[i + 1]]] : acc), []));
const ROOT = args.project;
const ctx = loadProject(ROOT);
assertProjectWritable(ctx);
const planOf = Object.fromEntries(JSON.parse(fs.readFileSync(args.plans, 'utf8')).map(([name, id]) => [name, id]));

// ------------------------------------------------------------------ shared helpers

const HELPERS = `// Ported from tasks.js / tasks.extras.js / nools-extras.js of the CHT configuration muso-mali.
// CHT reads report fields with getField(report, 'a.b'); Collecton reports are flat: payload[key].

const DAY = 86400000;
const field = (report, key) => report?.payload?.[key];

/** CHT getDynamicReportedDate: the home visit date, else the report date (epoch ms). */
const dynamicDate = (report) => {
  if (!report) return 0;
  const visited = field(report, 'visited_contact_uuid') ? field(report, 'visited_date') : null;
  const specified = field(report, 's_reported_date');
  const parsed = Date.parse(visited || specified || '');
  return Number.isNaN(parsed) ? Number(report.created_at) || 0 : parsed;
};

/** CHT toDate(val, plusDays): the date of a value, plus whole days (UTC calendar day). */
const toDay = (value, plusDays = 0) => {
  const time = typeof value === 'number' ? value : Date.parse(value || '');
  if (Number.isNaN(time)) return null;
  const d = new Date(time);
  return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) + plusDays * DAY;
};

/** CHT toDateString: YYYY-MM-DD, '' for an invalid date. */
const toDateString = (value) => {
  const day = toDay(value);
  return day === null ? '' : new Date(day).toISOString().slice(0, 10);
};

/** CHT getMostRecentReport: the latest of reports (already filtered by form) by dynamic date. */
const mostRecent = (reports) => reports
  .reduce((best, r) => (!best || dynamicDate(r) > dynamicDate(best) ? r : best), null);

const formReports = (app, form) => app.getReports({ form, targetId: app.person?.id });

/** CHT latestReportHasField: the latest report of a form has field = value, more recent than a date. */
const latestHasField = (app, form, key, value, moreRecentThan) => {
  const latest = mostRecent(formReports(app, form));
  if (!latest || field(latest, key) !== value) return false;
  return !moreRecentThan || Number(latest.created_at) > moreRecentThan;
};

/** CHT isSomeReportInWindow on dynamic dates (reports already filtered by form). */
const someInWindow = (reports, start, end) => reports.some(r => {
  const date = dynamicDate(r);
  return date >= start && date <= end;
});

/** Window of the current event: [due - start, due + end + 1 day]. */
const eventWindow = (app) => {
  const due = Date.parse(app.event.dueDate);
  return [due - app.event.startDays * DAY, due + (app.event.endDays + 1) * DAY];
};

/** CHT resolveIfClosure_isReportInEventWindow(form). */
const actionFormInWindow = (app, form) => {
  const [start, end] = eventWindow(app);
  return someInWindow(formReports(app, form), start, end);
};

/** CHT dueDateClosure(days, lastElement). */
const dueAfter = (report, days, lastElement = false) => {
  if (lastElement) {
    const planned = toDay(field(report, 'assessment_date'), days);
    const reported = toDay(Number(report.created_at));
    if (planned !== null && reported !== null && planned < reported) return reported;
  }
  return dynamicDate(report) + days * DAY;
};

const isAlive = (app) => !(app.person?.attributes?.person_is_deceased === true || app.person?.attributes?.person_is_deceased === 'true');

// CHT displayForUserAt (place type of the user) and user.role (role of the user's contact):
// c40_chw_area -> chw_uhc, c30_supervisor_area -> supervisor, c20_health_area -> tb_focal_point.
const userRole = (app) => app.currentUser?.role?.name;

/** CHT indexOfEvent: position of the event in the rule, from 1, as a string. */
const eventIndex = (app, ids) => String(ids.indexOf(app.event?.id) + 1);

/** CHT modifyContentClosure_fromPatientAssessment(followupType). */
const fromPatientAssessment = (app, followupType, ids) => {
  const r = app.report;
  return {
    t_follow_up_type: followupType,
    t_treat_for_diarrhea: field(r, 'treat_diarrhea'),
    t_treat_for_malaria: field(r, 'treat_malaria'),
    t_treat_for_ari: field(r, 'treat_ari'),
    t_treat_for_cough: field(r, 'treat_cough'),
    t_fast_breathing: field(r, 'fast_breathing'),
    t_assessment_date: toDateString(field(r, 'assessment_date')),
    t_referral_date: toDateString(field(r, 'assessment_date')),
    t_observation: field(r, 'observe'),
    t_tb_case_category: field(r, 'tb_referral') === 'true' ? 'suspect' : 'contact',
    t_follow_up_count: eventIndex(app, ids),
  };
};
`;

// ------------------------------------------------------------------ rules

const rules = [];
const PA = 'patient_assessment';
const PA5 = 'patient_assessment_over_5';
const due = (days, last = false) => `  return dueAfter(app.report, ${days}, ${last});`;
const plusEvents = (days, window = () => [0, 0], lastRule = true) => days.map((d, i) => ({
  id: `+${d}`, start: window(d)[0], end: window(d)[1], body: due(d, lastRule && i === days.length - 1),
}));
const guard = (role) => `isAlive(app) && userRole(app) === '${role}'`;

const referralFromAssessment = (name, title, source, action) => rules.push({
  name, title, appliesTo: 'report', forms: [source], action, priority: 1,
  events: plusEvents([1, 2, 3]),
  appliesIf: `  return ${guard('chw_uhc')} && ['referral', 'accompany_to_cscom', 'refer_to_cscom'].some(k => field(app.report, k) === 'true');`,
  resolvedIf: `  const reported = dynamicDate(app.report);
  // If TB, submitting the diagnostic follow-up stops the referral.
  const tbDone = field(app.report, 'tb_referral') === 'true'
    && someInWindow(formReports(app, 'tb_test_result_chw'), reported, app.now);
  return tbDone
    || latestHasField(app, '${action}', 'close_out', 'true', reported)
    || latestHasField(app, '${action}', 'referral', 'true', reported)
    || dynamicDate(mostRecent(formReports(app, '${source}'))) > reported
    || actionFormInWindow(app, '${action}');`,
  prefill: `  return fromPatientAssessment(app, 'referral', ['+1', '+2', '+3']);`,
});
referralFromAssessment(`${PA5}_yields_referral_followup`, 'Suivi référence', PA5, 'referral_followup');
referralFromAssessment(`${PA}_yields_referral_followup_under_5`, 'Suivi référence (moins de 5 ans)', PA, 'referral_followup_under_5');

const referralTemplate = (source, action, title) => rules.push({
  name: `${source}_yields_${action}`, title, appliesTo: 'report', forms: [source], action, priority: 1,
  events: plusEvents([1, 2, 3], undefined, false),
  appliesIf: `  return ${guard('chw_uhc')} && ['referral', 'accompany_to_cscom', 'refer_to_cscom'].some(k => field(app.report, k) === 'true');`,
  resolvedIf: `  const reported = dynamicDate(app.report);
  return latestHasField(app, '${action}', 'close_out', 'true', reported)
    || latestHasField(app, '${action}', 'referral', 'true', reported)
    || actionFormInWindow(app, '${action}');`,
  prefill: `  const r = app.report;
  // CHT reads report.fields.t_follow_up_type (a top-level field these forms do not have).
  return {
    t_follow_up_count: (parseInt(field(r, 'follow_up_count'), 10) || 0) + 1,
    t_delivery_date: toDateString(field(r, 'delivery_date')),
    t_lmp_date: toDateString(field(r, 'lmp_date')),
    t_treat_for_malaria: field(r, 'treat_for_malaria'),
    t_assessment_date: toDateString(field(r, 'assessment_date')),
    t_follow_up_type: field(r, 't_follow_up_type'),
  };`,
});
referralTemplate('treatment_followup_over_5', 'referral_followup', 'Suivi référence');
referralTemplate('treatment_followup', 'referral_followup_under_5', 'Suivi référence (moins de 5 ans)');
referralTemplate('moderate_malnutrition_followup', 'referral_followup_under_5', 'Suivi référence (moins de 5 ans)');
referralTemplate('severe_malnutrition_followup', 'referral_followup_under_5', 'Suivi référence (moins de 5 ans)');

const selfSpawning = (form, title, prefill) => rules.push({
  name: `${form}_yields_self`, title, appliesTo: 'report', forms: [form], action: form, priority: 1,
  events: plusEvents([1, 2, 3], undefined, false),
  // No displayForUserAt in the CHT rule: every user; muted has no equivalent.
  appliesIf: `  return isAlive(app) && field(app.report, 'referral') === 'true';`,
  // CHT selfSpawningTemplate ignores the resolvedIf of its options: only later reports of the form count.
  resolvedIf: `  const reported = dynamicDate(app.report);
  const later = formReports(app, '${form}').filter(r => dynamicDate(r) > reported);
  const [start, end] = eventWindow(app);
  return someInWindow(later, start, end);`,
  prefill,
});
selfSpawning('referral_followup', 'Suivi référence', `  return { t_tb_case_category: field(app.report, 'tb_case_category') };`);
selfSpawning('referral_followup_under_5', 'Suivi référence (moins de 5 ans)', null);

const treatmentCount = `['treat_malaria', 'treat_diarrhea', 'treat_ari'].filter(k => field(app.report, k) === 'true').length`;
const treatmentResolved = (action, source) => `  const reported = dynamicDate(app.report);
  return latestHasField(app, '${action}', 'close_out', 'true', reported)
    || dynamicDate(mostRecent(formReports(app, '${source}'))) > reported
    || actionFormInWindow(app, '${action}')
    || latestHasField(app, '${action}', 'referral', 'true', reported);`;
const specific = (treatment, days) => rules.push({
  name: `${PA}_yields_treatment_followup_${treatment}`, title: 'Suivi traitement', appliesTo: 'report', forms: [PA], action: 'treatment_followup',
  events: plusEvents(days, d => (d < 5 ? [0, 0] : [1, 3])),
  appliesIf: `  return ${guard('chw_uhc')} && field(app.report, 'accompany_to_cscom') !== 'true' && field(app.report, 'refer_to_cscom') !== 'true'
    && field(app.report, 'treat_${treatment}') === 'true' && ${treatmentCount} === 1;`,
  resolvedIf: treatmentResolved('treatment_followup', PA),
  prefill: `  return fromPatientAssessment(app, '${treatment}', ${JSON.stringify(days.map(d => `+${d}`))});`,
});
specific('malaria', [1, 2, 3]);
specific('ari', [1, 2, 3, 5]);
specific('diarrhea', [5]);
rules.push({
  name: `${PA}_yields_treatment_followup`, title: 'Suivi traitement', appliesTo: 'report', forms: [PA], action: 'treatment_followup',
  events: plusEvents([1, 2, 3, 5], d => (d === 5 ? [1, 3] : [0, 0])),
  appliesIf: `  return ${guard('chw_uhc')} && field(app.report, 'accompany_to_cscom') !== 'true' && field(app.report, 'refer_to_cscom') !== 'true' && ${treatmentCount} > 1;`,
  resolvedIf: treatmentResolved('treatment_followup', PA),
  prefill: `  return fromPatientAssessment(app, 'multiple', ['+1', '+2', '+3', '+5']);`,
});
rules.push({
  name: `${PA5}_yields_treatment_followup_over_5`, title: 'Suivi traitement', appliesTo: 'report', forms: [PA5], action: 'treatment_followup_over_5',
  events: plusEvents([1, 2, 3]),
  appliesIf: `  return ${guard('chw_uhc')} && field(app.report, 'referral') !== 'true' && field(app.report, 'treat_malaria') === 'true';`,
  resolvedIf: `  const reported = dynamicDate(app.report);
  return latestHasField(app, 'treatment_followup_over_5', 'close_out', 'true', reported)
    || latestHasField(app, 'treatment_followup_over_5', 'referral', 'true', reported)
    || dynamicDate(mostRecent(formReports(app, '${PA5}'))) > reported
    || actionFormInWindow(app, 'treatment_followup_over_5');`,
  prefill: `  return fromPatientAssessment(app, 'malaria', ['+1', '+2', '+3']);`,
});

const malnutrition = (action, treatment, title, days, endStates) => {
  const ids = days.map(d => `reported_date+${d}`);
  rules.push({
    name: `${PA}_yields_${action}`, title, appliesTo: 'person', personTypes: ['patient'], action, priority: 2,
    events: days.map((d, i) => ({
      id: ids[i], start: i === 0 ? 2 : 0, end: i === days.length - 1 ? 14 : days[i + 1] - d - 1,
      body: `  return dueAfter(assessmentOf(app), ${d});`,
    })),
    extra: `/** CHT malnutritionFollowupTemplate: the latest assessment with treat_${treatment}, less than 97 days old. */
const assessmentOf = (app) => {
  const eligible = formReports(app, '${PA}')
    .filter(r => field(r, 'treat_${treatment}') === 'true')
    .filter(r => (app.now - dynamicDate(r)) / DAY < 97);
  return mostRecent(eligible);
};

/** Follow-ups of the assessment, latest first (CHT reads inputs.source_id). */
const followUpsOf = (app, assessment) => formReports(app, '${action}')
  .filter(r => field(r, 'inputs_source_id') === assessment.id)
  .sort((a, b) => dynamicDate(b) - dynamicDate(a));`,
    appliesIf: `  return ${guard('chw_uhc')} && !!assessmentOf(app);`,
    resolvedIf: `  const assessment = assessmentOf(app);
  const last = mostRecent(formReports(app, '${action}'));
  let ended = false;
  if (last && assessment && field(last, 'inputs_source_id') === assessment.id) {
    ended = ${JSON.stringify(endStates)}.includes(field(last, 'final_cat')) || field(last, 'patient_deceased') === 'true';
  }
  return ended || actionFormInWindow(app, '${action}');`,
    prefill: `  const assessment = assessmentOf(app);
  if (!assessment) return {};
  const followUps = followUpsOf(app, assessment);
  const values = (key) => followUps.map(f => field(f, key));
  const colors = values('shakir_strip_color');
  const presence = values('child_presence');
  const last2Colors = followUps.slice(-2).map(f => field(f, 'shakir_strip_color'));
  const count = (list, value) => String(list.filter(v => v === value).length);
  return {
    source_id: assessment.id,
    t_assessment_date: toDateString(field(assessment, 'assessment_date')),
    t_follow_up_count: eventIndex(app, ${JSON.stringify(ids)}),
    t_before_previous_child_presence: presence.length < 2 ? '' : presence[presence.length - 2],
    t_previous_strip: colors.length === 0 ? '' : colors[colors.length - 1],
    t_num_red_strips: count(colors, 'red'),
    t_num_child_presence_confirmation: count(presence, 'no'),
    t_num_yellow_strips: count(colors, 'yellow'),
    t_num_no_edema: count(values('has_edema'), 'no'),
    t_num_shakir_strip_length: String(followUps.slice(-2).map(f => field(f, 'shakir_strip_length')).filter(v => v >= 12.5).length),
    t_plumpyNut_dosage_last_week: values('plumbyNut_this_week').slice(-1)[0],
    t_previous_vaccine_status: values('current_vaccine_status').slice(-1)[0],
    t_num_successive_red_strips: count(last2Colors, 'red'),
    t_num_successive_yellow_strips: count(last2Colors, 'yellow'),
    t_num_successive_green_strips: count(last2Colors, 'green'),
    t_num_measles: count(values('s_vaccinated_measles'), 'yes'),
    t_num_recovered: count(values('final_cat'), 'recovered'),
    t_folic_acid_srp_given_once: values('s_iron_folic_type').includes('iron_folic_acid_srp') ? 'true' : 'false',
  };`,
  });
};
malnutrition('moderate_malnutrition_followup', 'MAM', 'Suivi MAM', [7, 14, 21, 28, 42, 56, 70, 84], ['refer_with_no_followup', 'abandoned']);
malnutrition('severe_malnutrition_followup', 'SAM_without_complication', 'Suivi MAS Sans Complication', [7, 14, 21, 28, 35, 42, 49, 56, 63, 70], ['refer_with_no_followup', 'abandoned', 'recovered']);

const tbContent = (source, withPatientDetails) => `  const r = app.report;
  return {
${withPatientDetails ? `    t_date_of_birth: app.person?.attributes?.birthdate ?? '',
    t_patient_phone: field(r, 'patient_phone'),
` : ''}    t_tb_diagnosis_en: field(r, 'tb_diagnosis_en'),
    t_tb_diagnosis_fr: field(r, 'tb_diagnosis_fr'),
    t_tb_diagnosis_bm: field(r, 'tb_diagnosis_bm'),
    t_cscom_area: field(r, 'cscom_area'),
    t_tb_case_category: 'suspect',
    t_tb_assessment_date: field(r, 'assessment_date'),
    t_starting_form_uuid: r.id,
    t_starting_form_name: '${source}',
    t_patient_name: field(r, 'patient_name'),
    t_patient_sex: field(r, 'patient_sex'),
    t_muso_id: field(r, 'muso_id'),
  };`;
for (const source of [PA5, PA]) {
  rules.push({
    name: `${source}_yields_tb_test_result_fp`, title: 'Résultat TB (Point Focal)', appliesTo: 'report', forms: [source], action: 'tb_test_result_fp', priority: 1,
    events: [{ id: 'event_1', start: 0, end: 30, body: due(0) }],
    extra: `/**
 * CHT user.cscom_area: the CSCOM of the TB focal point person of the user's health area
 * (a Collecton account is not a person; the one named like the account when several).
 */
const userCscom = (app) => {
  const focalPoints = (app.personsAt?.(app.currentUser?.orgUnitId) ?? []).filter(p => p.personType === 'tb_focal_point');
  const own = focalPoints.find(p => p.attributes?.person_name === app.currentUser?.name) ?? focalPoints[0];
  return own?.attributes?.cscom_area;
};`,
    appliesIf: `  return ${guard('tb_focal_point')} && field(app.report, 'tb_referral') === 'true' && field(app.report, 'cscom_area') === userCscom(app);`,
    resolvedIf: `  return actionFormInWindow(app, 'tb_test_result_fp');`,
    prefill: tbContent(source, true),
  });
  rules.push({
    name: `${source}_yields_tb_test_result_chw`, title: 'Résultat TB', appliesTo: 'report', forms: [source], action: 'tb_test_result_chw', priority: 1,
    // CHT { days: 4 }: from the raw report date, not the dynamic one.
    events: [{ id: 'event_1', start: 0, end: 180, dueDays: 4 }],
    appliesIf: `  return ${guard('chw_uhc')} && field(app.report, 'tb_referral') === 'true';`,
    resolvedIf: `  return actionFormInWindow(app, 'tb_test_result_chw');`,
    prefill: tbContent(source, false),
  });
}

// ------------------------------------------------------------------ write

const fnName = (id) => `due_${id.replace(/[^A-Za-z0-9]/g, '_')}`;
const results = [];
for (const rule of rules) {
  const planId = planOf[rule.name];
  if (!planId) throw new Error(`No plan for ${rule.name}`);
  planning.assertAuthorized(ROOT, planId, 'task', rule.title);
  const parts = [HELPERS];
  if (rule.extra) parts.push(rule.extra);
  parts.push(`export function appliesIf(app) {\n${rule.appliesIf}\n}`);
  parts.push(`export function resolvedIf(app) {\n${rule.resolvedIf}\n}`);
  if (rule.prefill) parts.push(`/** CHT modifyContent: values handed to the form (app.taskInputs). */\nexport function prefill(app) {\n${rule.prefill}\n}`);
  for (const event of rule.events) if (event.body) parts.push(`export function ${fnName(event.id)}(app) {\n${event.body}\n}`);
  const created = taskRules.createTask(ctx, {
    title: rule.title,
    name: rule.name,
    appliesTo: rule.appliesTo,
    forms: rule.forms,
    personTypes: rule.personTypes,
    actionForm: rule.action,
    events: rule.events.map(e => ({
      id: e.id,
      ...(e.body ? { dueDate: fnName(e.id) } : { dueDays: e.dueDays }),
      startDays: e.start,
      endDays: e.end,
    })),
    appliesIf: 'appliesIf',
    resolvedIf: 'resolvedIf',
    ...(rule.priority ? { priority: rule.priority } : {}),
    ...(rule.prefill ? { prefill: 'prefill' } : {}),
    logicScript: `${parts.join('\n\n')}\n`,
  });
  planning.recordGenerated(ROOT, planId, `task: ${rule.title}`);
  results.push({ name: rule.name, valid: created.validation.valid, errors: created.validation.errors, warnings: created.validation.warnings.length });
}
console.log(JSON.stringify(results.filter(r => !r.valid)));
console.log(`${results.length} rules, ${results.filter(r => r.valid).length} valid`);
