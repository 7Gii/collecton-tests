/**
 * Logic file for the 'diagnostic_de_grossesse' form.
 * Export functions here and reference them by name in 'form.json'.
 * Named functions work for the form 'displayExpression', 'calculationExpression'
 * and 'defaultValue'; field display, required and validation rules are inline
 * '(app) => ...' expressions.
 * @param {object} app - The application execution context.
 */

const DAY_MS = 24 * 60 * 60 * 1000;
const PREGNANCY_DAYS = 280;

function parseDate(value) {
  if (!value) return null;
  const date = new Date(String(value).slice(0, 10));
  return isNaN(date.getTime()) ? null : date;
}

function lmpDate(app) {
  if (app.currentForm?.resultat_du_test_de_grossesse !== 'positif') return null;
  return parseDate(app.currentForm?.date_des_dernieres_regles_ddr);
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

// Add menu entry: living woman aged 15 to 40 with no ongoing pregnancy or postpartum
// follow-up (those continue with the CPN, delivery and CPoN forms).
export function showForm(app) {
  try {
    if (!livingWoman(app)) return false;
    const birth = parseDate(app.person.attributes?.birthdate);
    if (!birth) return false;
    const age = (Date.now() - birth.getTime()) / YEAR_MS;
    if (age < AGE_MIN_YEARS || age > AGE_MAX_YEARS) return false;
    return pregnancyState(app) === 'none';
  } catch (error) {
    return false;
  }
}

export function today(app) {
  return new Date().toISOString().slice(0, 10);
}

// Expected delivery date: last menstrual period + 280 days (Naegele), as YYYY-MM-DD.
export function expectedDeliveryDate(app) {
  const lmp = lmpDate(app);
  if (!lmp) return null;
  return new Date(lmp.getTime() + PREGNANCY_DAYS * DAY_MS).toISOString().slice(0, 10);
}

// Completed weeks of amenorrhea on the visit date (today when the visit date is empty).
export function gestationalAgeWeeks(app) {
  const lmp = lmpDate(app);
  if (!lmp) return null;
  const reference = parseDate(app.currentForm?.date_de_la_visite) || new Date();
  const weeks = Math.floor((reference.getTime() - lmp.getTime()) / (7 * DAY_MS));
  return weeks >= 0 ? weeks : null;
}
