/**
 * Logic file for the 'accouchement' form.
 * Export functions here and reference them by name in 'form.json'.
 * Named functions work for the form 'displayExpression', 'calculationExpression'
 * and 'defaultValue'; field display, required and validation rules are inline
 * '(app) => ...' expressions.
 * @param {object} app - The application execution context.
 */

function parseDate(value) {
  if (!value) return null;
  const date = new Date(String(value).slice(0, 10));
  return isNaN(date.getTime()) ? null : date;
}

const LOW_BIRTH_WEIGHT_G = 2500;
const VERY_LOW_BIRTH_WEIGHT_G = 1500;

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

// Add menu entry: ongoing pregnancy (a delivery can happen between two CPN). Also
// opened by the 'enregistrer_accouchement' task.
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

// Birth weight class: 'very_low', 'low', 'normal', or null without a weight.
export function birthWeightClass(app) {
  const raw = app.currentForm?.poids_du_bebe_g;
  if (raw === undefined || raw === null || raw === '') return null;
  const grams = Number(raw);
  if (!Number.isFinite(grams)) return null;
  if (grams < VERY_LOW_BIRTH_WEIGHT_G) return 'very_low';
  if (grams < LOW_BIRTH_WEIGHT_G) return 'low';
  return 'normal';
}
