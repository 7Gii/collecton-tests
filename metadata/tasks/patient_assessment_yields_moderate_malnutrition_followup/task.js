// Ported from tasks.js / tasks.extras.js / nools-extras.js of the CHT configuration muso-mali.
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


/** CHT malnutritionFollowupTemplate: the latest assessment with treat_MAM, less than 97 days old. */
const assessmentOf = (app) => {
  const eligible = formReports(app, 'patient_assessment')
    .filter(r => field(r, 'treat_MAM') === 'true')
    .filter(r => (app.now - dynamicDate(r)) / DAY < 97);
  return mostRecent(eligible);
};

/** Follow-ups of the assessment, latest first (CHT reads inputs.source_id). */
const followUpsOf = (app, assessment) => formReports(app, 'moderate_malnutrition_followup')
  .filter(r => field(r, 'inputs_source_id') === assessment.id)
  .sort((a, b) => dynamicDate(b) - dynamicDate(a));

export function appliesIf(app) {
  return isAlive(app) && userRole(app) === 'chw_uhc' && !!assessmentOf(app);
}

export function resolvedIf(app) {
  const assessment = assessmentOf(app);
  const last = mostRecent(formReports(app, 'moderate_malnutrition_followup'));
  let ended = false;
  if (last && assessment && field(last, 'inputs_source_id') === assessment.id) {
    ended = ["refer_with_no_followup","abandoned"].includes(field(last, 'final_cat')) || field(last, 'patient_deceased') === 'true';
  }
  return ended || actionFormInWindow(app, 'moderate_malnutrition_followup');
}

/** CHT modifyContent: values handed to the form (app.taskInputs). */
export function prefill(app) {
  const assessment = assessmentOf(app);
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
    t_follow_up_count: eventIndex(app, ["reported_date+7","reported_date+14","reported_date+21","reported_date+28","reported_date+42","reported_date+56","reported_date+70","reported_date+84"]),
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
  };
}

export function due_reported_date_7(app) {
  return dueAfter(assessmentOf(app), 7);
}

export function due_reported_date_14(app) {
  return dueAfter(assessmentOf(app), 14);
}

export function due_reported_date_21(app) {
  return dueAfter(assessmentOf(app), 21);
}

export function due_reported_date_28(app) {
  return dueAfter(assessmentOf(app), 28);
}

export function due_reported_date_42(app) {
  return dueAfter(assessmentOf(app), 42);
}

export function due_reported_date_56(app) {
  return dueAfter(assessmentOf(app), 56);
}

export function due_reported_date_70(app) {
  return dueAfter(assessmentOf(app), 70);
}

export function due_reported_date_84(app) {
  return dueAfter(assessmentOf(app), 84);
}
