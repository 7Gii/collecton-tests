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

const PERSON_TYPE = 'femme';
const AGE_MIN_YEARS = 15;
const AGE_MAX_YEARS = 40;
const POSTPARTUM_DAYS = 42;
const YEAR_MS = 365.25 * 24 * 60 * 60 * 1000;

// Living woman person; never an org unit.
function livingWoman(app) {
  const person = app.person;
  if (!person || person.personType !== PERSON_TYPE) return false;
  return person.attributes?.person_is_deceased !== true;
}

function latestReport(app, form) {
  const reports = app.getReports({ form, targetId: app.person.id }) || [];
  return reports.reduce((latest, report) => (!latest || report.created_at > latest.created_at ? report : latest), null);
}

// Pregnancy state from the reports of the woman: 'pregnant' (positive diagnosis not
// followed by a delivery), 'postpartum' (delivery less than 42 days ago) or 'none'.
function pregnancyState(app) {
  const diagnosis = latestReport(app, 'diagnostic_de_grossesse');
  const delivery = latestReport(app, 'accouchement');
  const positive = diagnosis && diagnosis.payload?.resultat_du_test_de_grossesse === 'positif';
  if (positive && (!delivery || delivery.created_at < diagnosis.created_at)) return 'pregnant';
  if (delivery && Date.now() - delivery.created_at <= POSTPARTUM_DAYS * 24 * 60 * 60 * 1000) {
    const mother = latestReport(app, 'suivi_postnatal_cpon');
    const deceased = mother && mother.created_at > delivery.created_at && mother.payload?.etat_de_la_mere === 'decedee';
    return deceased ? 'none' : 'postpartum';
  }
  return 'none';
}

// Add menu entry: delivery less than 42 days ago and mother alive. Also opened by
// the 'suivi_postnatal' task.
export function showForm(app) {
  try {
    return livingWoman(app) && pregnancyState(app) === 'postpartum';
  } catch (error) {
    return false;
  }
}

export function today(app) {
  return new Date().toISOString().slice(0, 10);
}

// Days between the latest recorded delivery of the selected woman and the visit date.
export function daysSinceDelivery(app) {
  const targetId = app.person?.id;
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
