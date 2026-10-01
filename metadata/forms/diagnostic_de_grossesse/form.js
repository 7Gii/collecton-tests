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
const AGE_MIN_YEARS = 15;
const AGE_MAX_YEARS = 40;

function parseDate(value) {
  if (!value) return null;
  const date = new Date(String(value).slice(0, 10));
  return isNaN(date.getTime()) ? null : date;
}

function lmpDate(app) {
  if (app.currentForm?.resultat_du_test_de_grossesse !== 'positif') return null;
  return parseDate(app.currentForm?.date_des_dernieres_regles_ddr);
}

function ageInYears(birthdate) {
  const birth = parseDate(birthdate);
  if (!birth) return null;
  return (Date.now() - birth.getTime()) / (365.25 * DAY_MS);
}

// Add menu entry: living women aged 15 to 40.
export function showForm(app) {
  const person = app.person;
  if (!person || person.personType !== 'femme') return false;
  if (person.attributes?.person_is_deceased === true) return false;
  const age = ageInYears(person.attributes?.birthdate);
  return age !== null && age >= AGE_MIN_YEARS && age <= AGE_MAX_YEARS;
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
