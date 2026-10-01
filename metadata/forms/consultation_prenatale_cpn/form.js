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

// Add menu entry: ongoing pregnancy (unscheduled visit). Also opened by the
// 'premiere_cpn' and 'cpn_suivante' tasks.
export function showForm(app) {
  try {
    return livingWoman(app) && pregnancyState(app) === 'pregnant';
  } catch (error) {
    return false;
  }
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
