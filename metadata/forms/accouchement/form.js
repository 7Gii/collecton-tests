/**
 * Logic file for the 'accouchement' form.
 * Export functions here and reference them by name in 'form.json'.
 * Named functions work for the form 'displayExpression', 'calculationExpression'
 * and 'defaultValue'; field display, required and validation rules are inline
 * '(app) => ...' expressions.
 * @param {object} app - The application execution context.
 */

const LOW_BIRTH_WEIGHT_G = 2500;
const VERY_LOW_BIRTH_WEIGHT_G = 1500;

export function showForm(app) {
  return true;
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
