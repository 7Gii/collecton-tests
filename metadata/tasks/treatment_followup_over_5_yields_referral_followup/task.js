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


export function appliesIf(app) {
  return isAlive(app) && userRole(app) === 'chw_uhc' && ['referral', 'accompany_to_cscom', 'refer_to_cscom'].some(k => field(app.report, k) === 'true');
}

export function resolvedIf(app) {
  const reported = dynamicDate(app.report);
  return latestHasField(app, 'referral_followup', 'close_out', 'true', reported)
    || latestHasField(app, 'referral_followup', 'referral', 'true', reported)
    || actionFormInWindow(app, 'referral_followup');
}

/** CHT modifyContent: values handed to the form (app.taskInputs). */
export function prefill(app) {
  const r = app.report;
  // CHT reads report.fields.t_follow_up_type (a top-level field these forms do not have).
  return {
    t_follow_up_count: (parseInt(field(r, 'follow_up_count'), 10) || 0) + 1,
    t_delivery_date: toDateString(field(r, 'delivery_date')),
    t_lmp_date: toDateString(field(r, 'lmp_date')),
    t_treat_for_malaria: field(r, 'treat_for_malaria'),
    t_assessment_date: toDateString(field(r, 'assessment_date')),
    t_follow_up_type: field(r, 't_follow_up_type'),
  };
}

export function due__1(app) {
  return dueAfter(app.report, 1, false);
}

export function due__2(app) {
  return dueAfter(app.report, 2, false);
}

export function due__3(app) {
  return dueAfter(app.report, 3, false);
}
