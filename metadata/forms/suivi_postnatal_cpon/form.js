/**
 * Logic file for the 'suivi_postnatal_cpon' form.
 * Export functions here and reference them by name in 'form.json'.
 * Named functions work for the form 'displayExpression', 'calculationExpression'
 * and 'defaultValue'; field display, required and validation rules are inline
 * '(app) => ...' expressions.
 * @param {object} app - The application execution context.
 */

const DAY_MS = 24 * 60 * 60 * 1000;

function parseDate(value) {
  if (!value) return null;
  const date = new Date(String(value).slice(0, 10));
  return isNaN(date.getTime()) ? null : date;
}

export function showForm(app) {
  return true;
}

export function today(app) {
  return new Date().toISOString().slice(0, 10);
}

// Days between the latest recorded delivery of the selected woman and the visit date.
export function daysSinceDelivery(app) {
  const targetId = app.selectedItem?.id;
  if (!targetId || typeof app.getReports !== 'function') return null;
  const delivery = app.getReports({ form: 'accouchement', targetId })
    .find(report => report.payload?.date_de_laccouchement);
  const deliveredOn = parseDate(delivery?.payload?.date_de_laccouchement);
  if (!deliveredOn) return null;
  const reference = parseDate(app.currentForm?.date_de_la_visite) || new Date();
  const days = Math.round((reference.getTime() - deliveredOn.getTime()) / DAY_MS);
  return days >= 0 ? days : null;
}

// Default visit type from the days since delivery (J3 up to day 5, J7 up to day 14, then S6).
export function defaultVisitType(app) {
  const days = daysSinceDelivery(app);
  if (days === null) return undefined;
  if (days <= 5) return 'j3';
  if (days <= 14) return 'j7';
  return 's6';
}
