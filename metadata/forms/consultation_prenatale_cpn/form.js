/**
 * Logic file for the 'consultation_prenatale_cpn' form.
 * Export functions here and reference them by name in 'form.json'.
 * Named functions work for the form 'displayExpression', 'calculationExpression'
 * and 'defaultValue'; field display, required and validation rules are inline
 * '(app) => ...' expressions.
 * @param {object} app - The application execution context.
 */

const DAY_MS = 24 * 60 * 60 * 1000;
const CPN_CODES = ['cpn1', 'cpn2', 'cpn3', 'cpn4'];

function parseDate(value) {
  if (!value) return null;
  const date = new Date(String(value).slice(0, 10));
  return isNaN(date.getTime()) ? null : date;
}

// Reports of the selected woman for one form, newest first.
function reportsOf(app, form) {
  const targetId = app.person?.id;
  if (!targetId || typeof app.getReports !== 'function') return [];
  return app.getReports({ form, targetId });
}

// Opened by the 'premiere_cpn' and 'cpn_suivante' tasks only: never offered in the add menu.
export function showForm(app) {
  return false;
}

export function today(app) {
  return new Date().toISOString().slice(0, 10);
}

// Default CPN number: one more than the CPN already recorded for this woman.
export function nextCpnNumber(app) {
  const done = reportsOf(app, 'consultation_prenatale_cpn').length;
  return CPN_CODES[done] || 'cpn5_plus';
}

// Weeks of amenorrhea on the CPN date, from the LMP of the latest pregnancy diagnosis.
export function gestationalAgeWeeks(app) {
  const diagnosis = reportsOf(app, 'diagnostic_de_grossesse')
    .find(report => report.payload?.date_des_dernieres_regles_ddr);
  const lmp = parseDate(diagnosis?.payload?.date_des_dernieres_regles_ddr);
  if (!lmp) return null;
  const reference = parseDate(app.currentForm?.date_de_la_cpn) || new Date();
  const weeks = Math.floor((reference.getTime() - lmp.getTime()) / (7 * DAY_MS));
  return weeks >= 0 ? weeks : null;
}
