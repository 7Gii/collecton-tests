/**
 * Logic for the 'health_area_supervisor' person type: functions of its summary cards, computed offline on
 * the details view of each person (the target of app).
 * Every function receives one argument, app (the same keys everywhere):
 *   app.person     the target person { id, personTypeId, personType, orgUnitId, attributes }
 *   app.orgUnit    the target org unit { id, levelId, level, levelName, parentId, attributes }
 *                  (one of the two, the other null; both null when there is no target)
 *   app.currentUser  { id, username, name, role: { name, permissions }, orgUnitId } or null
 *   app.getReports(filter?)  every visible report, newest first, e.g.
 *       app.getReports({ form: 'child_registration' })          by form name or id
 *       app.getReports({ targetId: app.person.id })             the target only
 *       app.getReports({ orgUnitId: id })                        an org unit and everything below it
 *       app.getReports({ from: '2026-01-01', to: '2026-02-01' }) a period (to is excluded)
 *     each report: r.form_id, r.target_id, r.target_is_person, r.created_at, r.payload
 *     (form values are strings as entered: convert numbers with Number(...))
 *   app.find(table, [{ field, value }])  rows of persons, org_units, person_types,
 *     org_unit_levels (asynchronous: await it; not usable in a form field expression)
 *   app.utils      exports of the project utils
 *   app.now        current time of the app clock (epoch ms): the chosen day in deferred
 *                  collection on the mobile app; use it instead of new Date()
 *   app.option     in the optionFilter of a select field only: the option being
 *                  checked { value, name, properties }
 *
 * A card's displayExpression returns a boolean; a field's value returns a string or a
 * number (null or '' hides the field).
 */

// Registration form of the person (attributes). Ported from forms/contact/person-create.xlsx
// and person-edit.xlsx of the CHT configuration muso-mali; field keys keep the CHT names.

const values = (app) => app.currentForm || {};
const isApproxDob = (app) => [].concat(values(app).date_of_birth_method || []).includes('approx');
// CHT phone constraint: regex(., '^(\d{2}\s?){4}$').
const PHONE = /^(\d{2}\s?){4}$/;
const phoneOk = (value) => !value || PHONE.test(String(value));

/** The exact date of birth is asked when it is known. */
export function dobIsKnown(app) {
  return !isApproxDob(app);
}

/** The age in years and months is asked when the date of birth is unknown. */
export function dobIsApprox(app) {
  return isApproxDob(app);
}

/** A date of birth cannot be in the future. */
export function dobNotInFuture(app) {
  const date = values(app).dob_calendar;
  return !date || Date.parse(date) <= app.now;
}

export function ageYearsInRange(app) {
  const years = Number(values(app).age_years);
  return years >= 0 && years <= 150;
}

export function ageMonthsInRange(app) {
  const months = Number(values(app).age_months);
  return months >= 0 && months <= 11;
}

/**
 * Native birthdate: the calendar date, or today minus the approximate age, on the day of
 * the month of today (CHT ephemeral_dob/dob_approx).
 */
export function birthdate(app) {
  const form = values(app);
  if (!isApproxDob(app)) return form.dob_calendar || '';
  const now = new Date(app.now);
  let month = now.getUTCMonth() + 1 - Number(form.age_months || 0);
  let year = now.getUTCFullYear() - Number(form.age_years || 0);
  if (month <= 0) {
    month += 12;
    year -= 1;
  }
  return `${year}-${String(month).padStart(2, '0')}-${String(now.getUTCDate()).padStart(2, '0')}`;
}

export function phoneValid(app) {
  return phoneOk(values(app).phone);
}

export function alternatePhoneValid(app) {
  return phoneOk(values(app).alternate_phone);
}

// ---- Summary cards: contact-summary.templated.js of the CHT configuration ----

const parseAttributes = (row) => {
  try {
    return typeof row.attributes === 'string' ? JSON.parse(row.attributes) : (row.attributes || {});
  } catch {
    return {};
  }
};

/** Names of the org units from orgUnitId upwards, nearest first (CHT lineage filter). */
async function lineageNames(app, orgUnitId) {
  const names = [];
  let id = orgUnitId;
  for (let depth = 0; id && depth < 10; depth++) {
    const [row] = await app.find('org_units', [{ field: 'id', value: id }]);
    if (!row) break;
    names.push(parseAttributes(row).ou_name || '');
    id = row.parent_id;
  }
  return names.filter(Boolean);
}

const DAY = 86400000;
const attrs = (app) => app.person?.attributes || {};

// CHT common.js formatRole.
const ROLE_LABELS = {
  tb_focal_point: 'TB FP',
  chw: 'ASC',
  chw_supervisor: 'SUPERVISOR',
  patient: 'PATIENT',
  site_supervisor: 'SITE SUPERVISOR',
  health_area_supervisor: 'HEALTH ZONE SUPERVISOR',
  stock_manager: 'STOCK MANAGER',
  chw_manager: 'CHW MANAGER',
};

/** CHT contact.age: years, else months, else days since the date of birth. */
export function ageLabel(app) {
  const dob = attrs(app).birthdate;
  if (!dob) return null;
  const birth = new Date(dob);
  const now = new Date(app.now);
  const days = Math.floor((now.getTime() - birth.getTime()) / DAY);
  if (Number.isNaN(days) || days < 0) return null;
  let months = (now.getUTCFullYear() - birth.getUTCFullYear()) * 12 + now.getUTCMonth() - birth.getUTCMonth();
  if (now.getUTCDate() < birth.getUTCDate()) months -= 1;
  if (months >= 12) return `${Math.floor(months / 12)} ans`;
  if (months >= 1) return `${months} mois`;
  return `${days} jours`;
}

export function phoneValue(app) {
  return attrs(app).phone || null;
}

export function sexLabel(app) {
  return { F: 'Femme', M: 'Homme' }[attrs(app).person_gender] || null;
}

export function externalIdValue(app) {
  return attrs(app).external_id || null;
}

/** CHT contact.parent: the org units the person belongs to. */
export async function parentLineage(app) {
  return (await lineageNames(app, app.person?.orgUnitId)).join(' / ') || null;
}

/** CHT contact.role, through formatRole. */
export function roleLabel(app) {
  const role = attrs(app).role;
  return role ? ROLE_LABELS[role] || role : null;
}

/** CHT contact.profile.death.title: shown for a deceased person. */
export function isDeceased(app) {
  return attrs(app).person_is_deceased === true || attrs(app).person_is_deceased === 'true';
}

/** No date of death nor death report in the scope: CHT shows 'Inconnu'. */
export function deathUnknown() {
  return 'Inconnu';
}
