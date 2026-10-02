// Form patient_assessment_over_5: ported from forms/app/patient_assessment_over_5.xlsx of the CHT configuration muso-mali.
// Field keys keep the XLSForm names; the functions below are generated from its expressions.


// ---- XPath runtime of the XLSForm port (generated, do not edit by hand) ----
// XPath 1.0 semantics over the form values: strings, numbers, booleans; dates are
// numbers of days since 1970-01-01 (decimal-date-time) or Date objects.

const DAY_MS = 86400000;
const ISO_DATE = /^\d{4}-\d{2}-\d{2}(T[\d:.]+(Z|[+-]\d{2}:?\d{2})?)?$/;

const X = {
  /** Value of a form field: '' when empty; a select_multiple stays a list. */
  v(app, key) {
    const value = (app.currentForm || {})[key];
    return value === undefined || value === null ? '' : value;
  },
  str(value) {
    if (value === undefined || value === null) return '';
    if (Array.isArray(value)) return value.join(' ');
    if (value instanceof Date) return Number.isNaN(value.getTime()) ? '' : value.toISOString();
    if (typeof value === 'number') return Number.isNaN(value) ? 'NaN' : String(value);
    if (typeof value === 'boolean') return value ? 'true' : 'false';
    return String(value);
  },
  num(value) {
    if (typeof value === 'number') return value;
    if (typeof value === 'boolean') return value ? 1 : 0;
    if (value instanceof Date) return value.getTime() / DAY_MS;
    const text = X.str(value).trim();
    if (text === '') return NaN;
    if (ISO_DATE.test(text)) return Date.parse(text.length === 10 ? `${text}T00:00:00Z` : text) / DAY_MS;
    return Number(text);
  },
  bool(value) {
    if (typeof value === 'boolean') return value;
    if (typeof value === 'number') return value !== 0 && !Number.isNaN(value);
    if (value instanceof Date) return true;
    return X.str(value).length > 0;
  },
  eq(a, b) {
    if (typeof a === 'boolean' || typeof b === 'boolean') return X.bool(a) === X.bool(b);
    if (typeof a === 'number' || typeof b === 'number') return X.num(a) === X.num(b);
    return X.str(a) === X.str(b);
  },
  ne(a, b) {
    if (typeof a === 'boolean' || typeof b === 'boolean') return X.bool(a) !== X.bool(b);
    if (typeof a === 'number' || typeof b === 'number') return X.num(a) !== X.num(b);
    return X.str(a) !== X.str(b);
  },
  lt(a, b) { return X.num(a) < X.num(b); },
  le(a, b) { return X.num(a) <= X.num(b); },
  gt(a, b) { return X.num(a) > X.num(b); },
  ge(a, b) { return X.num(a) >= X.num(b); },
  /** Stored value of a calculation: XPath string conversion, NaN as empty. */
  out(value) {
    if (typeof value === 'number' && Number.isNaN(value)) return '';
    return X.str(value);
  },
  toDate(value) {
    if (value instanceof Date) return value;
    if (typeof value === 'number') return new Date(value * DAY_MS);
    const text = X.str(value).trim();
    if (text === '') return new Date(NaN);
    if (/^-?\d+(\.\d+)?$/.test(text)) return new Date(Number(text) * DAY_MS);
    return new Date(text.length === 10 ? `${text}T00:00:00Z` : text);
  },
  list(value) {
    if (Array.isArray(value)) return value.map(String);
    return X.str(value).split(/\s+/).filter(Boolean);
  },
  choices: {},
  choiceName(list, value) {
    return (X.choices[list] || {})[X.str(value)] ?? '';
  },
  /** Property of the option checked by an option filter (XLSForm choice_filter column). */
  opt(app, name) {
    const props = app.option?.properties || {};
    return props[name] ?? (name === 'name' ? app.option?.value ?? '' : '');
  },
  /** CHT form inputs: the person, its lineage, its CHW and the user (see the port notes). */
  inp(app, path) {
    const parts = path.split('/');
    if (parts[0] === 'user') {
      const user = app.currentUser || {};
      return { contact_id: user.id, name: user.name, fullname: user.name, language: 'fr', is_in_sih: 'false' }[parts[1]] ?? '';
    }
    if (parts[0] === 'source') return 'user';
    if (parts[0] !== 'contact') return X.taskInput(app, path);
    let depth = 0;
    let i = 1;
    while (parts[i] === 'parent') { depth++; i++; }
    const rest = parts.slice(i);
    const person = app.person || null;
    if (depth === 0) return X.personField(person, rest[0]);
    const unit = (app.lineage || [])[depth - 1];
    if (!unit) return '';
    if (rest[0] === 'contact') {
      const role = { c40_chw_area: 'chw', c30_supervisor_area: 'chw_supervisor' }[unit.levelName];
      const contact = (app.personsAt ? app.personsAt(unit.id) : []).find(p => p.personType === role) || null;
      return X.personField(contact, rest[1]);
    }
    if (rest[0] === '_id') return unit.id;
    if (rest[0] === 'type' || rest[0] === 'contact_type') return unit.levelName || '';
    if (rest[0] === 'name') return unit.attributes?.ou_name ?? '';
    return unit.attributes?.[rest[0]] ?? '';
  },
  personField(person, field) {
    if (!person) return '';
    const a = person.attributes || {};
    switch (field) {
      case '_id': return person.id;
      case 'name': return a.person_name ?? '';
      case 'date_of_birth': return a.birthdate ?? '';
      case 'sex': return { F: 'female', M: 'male' }[a.person_gender] ?? '';
      case 'patient_id': return a.muso_id ?? '';
      default: return a[field] ?? '';
    }
  },
  /** Values a CHT task passed to the form (t_*): read from the task context when present. */
  taskInput(app, path) {
    const value = (app.taskInputs || {})[path];
    return value === undefined || value === null ? '' : value;
  },
  pad(n, size = 2) { return String(n).padStart(size, '0'); },
  format(value, pattern) {
    const d = X.toDate(value);
    if (Number.isNaN(d.getTime())) return '';
    const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return X.str(pattern).replace(/%([YmdeHMSnbjy3])/g, (_, c) => ({
      Y: d.getUTCFullYear(), y: X.pad(d.getUTCFullYear() % 100), m: X.pad(d.getUTCMonth() + 1), n: d.getUTCMonth() + 1,
      d: X.pad(d.getUTCDate()), e: d.getUTCDate(), H: X.pad(d.getUTCHours()), M: X.pad(d.getUTCMinutes()),
      S: X.pad(d.getUTCSeconds()), b: MONTHS[d.getUTCMonth()], 3: X.pad(d.getUTCMilliseconds(), 3),
      j: Math.floor((d - Date.UTC(d.getUTCFullYear(), 0, 1)) / DAY_MS) + 1,
    }[c]));
  },
  f: {
    concat: (app, ...args) => args.map(X.str).join(''),
    coalesce: (app, a, b) => (X.str(a) !== '' ? a : b),
    floor: (app, a) => Math.floor(X.num(a)),
    ceiling: (app, a) => Math.ceil(X.num(a)),
    round: (app, a, digits) => {
      const factor = 10 ** (digits === undefined ? 0 : X.num(digits));
      return Math.round(X.num(a) * factor) / factor;
    },
    int: (app, a) => Math.trunc(X.num(a)),
    abs: (app, a) => Math.abs(X.num(a)),
    min: (app, ...args) => Math.min(...args.map(X.num)),
    max: (app, ...args) => Math.max(...args.map(X.num)),
    number: (app, a) => X.num(a),
    string: (app, a) => X.str(a),
    boolean: (app, a) => X.bool(a),
    'string-length': (app, a) => X.str(a).length,
    substr: (app, s, start, end) => X.str(s).substring(X.num(start), end === undefined ? undefined : X.num(end)),
    substring: (app, s, start, length) => {
      const from = Math.round(X.num(start)) - 1;
      return X.str(s).substr(Math.max(0, from), length === undefined ? undefined : Math.round(X.num(length)) + Math.min(0, from));
    },
    'substring-before': (app, s, t) => { const a = X.str(s); const i = a.indexOf(X.str(t)); return i < 0 ? '' : a.slice(0, i); },
    'substring-after': (app, s, t) => { const a = X.str(s); const b = X.str(t); const i = a.indexOf(b); return i < 0 ? '' : a.slice(i + b.length); },
    contains: (app, s, t) => X.str(s).includes(X.str(t)),
    'starts-with': (app, s, t) => X.str(s).startsWith(X.str(t)),
    'normalize-space': (app, s) => X.str(s).trim().replace(/\s+/g, ' '),
    regex: (app, s, re) => new RegExp(X.str(re)).test(X.str(s)),
    selected: (app, list, value) => X.list(list).includes(X.str(value)),
    'count-selected': (app, list) => X.list(list).length,
    'selected-at': (app, list, index) => X.list(list)[X.num(index)] ?? '',
    'decimal-date-time': (app, a) => X.toDate(a).getTime() / DAY_MS,
    'decimal-date': (app, a) => X.toDate(a).getTime() / DAY_MS,
    'date-time': (app, a) => X.toDate(a),
    date: (app, a) => {
      const d = X.toDate(a);
      return Number.isNaN(d.getTime()) ? '' : `${d.getUTCFullYear()}-${X.pad(d.getUTCMonth() + 1)}-${X.pad(d.getUTCDate())}`;
    },
    'format-date': (app, d, pattern) => X.format(d, pattern),
    'format-date-time': (app, d, pattern) => X.format(d, pattern),
    today: (app) => {
      const d = new Date(app.now);
      return `${d.getUTCFullYear()}-${X.pad(d.getUTCMonth() + 1)}-${X.pad(d.getUTCDate())}`;
    },
    now: (app) => new Date(app.now),
    'difference-in-months': (app, from, to) => {
      const a = X.toDate(from);
      const b = X.toDate(to);
      if (Number.isNaN(a.getTime()) || Number.isNaN(b.getTime())) return '';
      let months = (b.getUTCFullYear() - a.getUTCFullYear()) * 12 + b.getUTCMonth() - a.getUTCMonth();
      if (b.getUTCDate() < a.getUTCDate()) months -= 1;
      return months;
    },
    once: (app, a) => a,
    uuid: () => (globalThis.crypto?.randomUUID ? globalThis.crypto.randomUUID() : `${Date.now()}-${Math.random()}`),
  },
};


X.choices = {
 "yes_no": {
  "yes": "Oui",
  "no": "Non"
 },
 "pos_neg": {
  "pos": "Positif",
  "neg": "Négatif"
 },
 "normal_fast": {
  "normal": "Normale",
  "fast": "Rapide"
 },
 "c_how_child_found": {
  "home_visit": "Au cours d’une Visite à Domicile",
  "at_chw_house": "Patient venu chez l'ASC",
  "informed_by_family": "Patient accompagné par un parent ou un membre de la communauté",
  "other": "Autre"
 },
 "c_assessment_time": {
  "c_assessment_time_1": "Matin",
  "c_assessment_time_2": "Midi",
  "c_assessment_time_3": "Soir",
  "c_assessment_time_4": "Nuit"
 },
 "c_when_illness": {
  "c_when_illness_1": "Aujourd’hui",
  "c_when_illness_2": "Au cours de la nuit",
  "c_when_illness_3": "Hier matin",
  "c_when_illness_4": "Hier soir",
  "c_when_illness_5": "Avant-hier matin",
  "c_when_illness_6": "Avant-hier soir",
  "c_when_illness_7": "Il y a trois (3) jours",
  "c_when_illness_8": "Il y a plus de trois (3) jours"
 },
 "CAT_covid19": {
  "CAT_covid19_1": "Informer l’équipe Covid-19",
  "CAT_covid19_2": "Mettre en observation",
  "CAT_covid19_3": "Informer le CSCOM"
 },
 "c_tdr_not_done": {
  "tdr_unavailable": "TDR non disponible",
  "other": "Autre"
 },
 "medication_not_given": {
  "out_of_stock": "Rupture de stock",
  "other": "Autre"
 },
 "c_malaria_act_dosage_once": {
  "2_tablets_act_once": "2 comprimés",
  "3_tablets_act_once": "3 Comprimés",
  "4_tablets_act_once": "4 comprimés"
 },
 "c_malaria_act_dosage": {
  "4_tablets_act": "2 comprimés 2 fois par jour",
  "6_tablets_act": "3 Comprimés 2 fois par jour",
  "8_tablets_act": "4 comprimés 2 fois par jour"
 },
 "c_give_paracetamol_dosage": {
  "demi_tab_paracetamol_3x_a_day": "1/2 comprimé 3 fois par jour",
  "three_quarter_tab_paracetamol_3x_a_day": "3/4 comprimé 3 fois par jour",
  "one_tab_paracetamol_3x_a_day": "1 comprimé 3 fois par jour",
  "two_tab_paracetamol_3x_a_day": "2 comprimés 3 fois par jour",
  "demi_tab_paracetamol_once": "1/2 comprimé",
  "three_quarter_tab_paracetamol_once": "3/4 comprimé",
  "one_tab_paracetamol_once": "1 comprimé",
  "two_tab_paracetamol_once": "2 comprimés"
 },
 "CAT": {
  "refer": "Référer au centre de santé",
  "watching": "Mettre en observation"
 },
 "CAT_out_of_stock": {
  "refer": "Référence",
  "continue_treatment": "Poursuivre le traitement à domicile"
 }
};

const isAlive = (person) => !(person?.attributes?.person_is_deceased === true || person?.attributes?.person_is_deceased === 'true');
const ageInYears = (person, now) => {
  const dob = person?.attributes?.birthdate;
  if (!dob) return null;
  return (now - Date.parse(dob)) / (365.25 * 86400000);
};

/**
 * CHT properties.json: a living patient of 5 or more, for a CHW (role chw_uhc, assigned to a
 * c40_chw_area). summary.muted has no equivalent (see ECARTS_CHT_COLLECTON.md).
 */
export function showForm(app) {
  const person = app.person;
  if (!person || person.personType !== 'patient' || !isAlive(person)) return false;
  if (app.currentUser?.role?.name !== 'chw_uhc') return false;
  const age = ageInYears(person, app.now);
  return age === null || age >= 5;
}

/** Group pregnancy_section */
export function grp_pregnancy_section(app) {
  return X.bool(X.eq(X.v(app, "patient_can_procreate"), "true"));
}

/** Group postpartum_section */
export function grp_postpartum_section(app) {
  return X.bool((X.bool(X.eq(X.v(app, "patient_is_pregnant"), "no")) && X.bool(X.eq(X.v(app, "amenorrhea_concept"), "false"))));
}

/** Group s_temperature */
export function grp_s_temperature(app) {
  return X.bool(X.eq(X.v(app, "amenorrhea_concept"), "false"));
}

/** Group s_temperature_retake */
export function grp_s_temperature_retake(app) {
  return X.bool((X.bool((X.bool(X.lt(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw"), 0)), 35)) || X.bool(X.gt(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw"), 0)), 41)))) && X.bool(X.eq(X.v(app, "amenorrhea_concept"), "false"))));
}

/** Group s_acc_danger_signs */
export function grp_s_acc_danger_signs(app) {
  return X.bool((X.bool(X.eq(X.v(app, "amenorrhea_concept"), "false")) && X.bool(X.eq(X.v(app, "amenorrhea_concept"), "false"))));
}

/** Group s_ref_danger_signs */
export function grp_s_ref_danger_signs(app) {
  return X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "false")) && X.bool(X.eq(X.v(app, "amenorrhea_concept"), "false"))));
}

/** Group pregnancy_danger_signs */
export function grp_pregnancy_danger_signs(app) {
  return X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "false")) && X.bool(X.eq(X.v(app, "has_first_ref_danger_sign"), "false")))) && X.bool(X.eq(X.v(app, "patient_is_pregnant"), "yes")))) && X.bool(X.eq(X.v(app, "amenorrhea_concept"), "false"))));
}

/** Group postpartum_danger_signs */
export function grp_postpartum_danger_signs(app) {
  return X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "false")) && X.bool(X.eq(X.v(app, "has_first_ref_danger_sign"), "false")))) && X.bool(X.eq(X.v(app, "patient_gave_birth_less_48_day"), "yes")))) && X.bool(X.eq(X.v(app, "amenorrhea_concept"), "false"))));
}

/** Group s_ref_other_diseases */
export function grp_s_ref_other_diseases(app) {
  return X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "false")) && X.bool(X.eq(X.v(app, "has_first_ref_danger_sign"), "false")))) && X.bool(X.eq(X.v(app, "has_pregnancy_danger_sign"), "false")))) && X.bool(X.eq(X.v(app, "has_postpartum_danger_sign"), "false")))) && X.bool(X.eq(X.v(app, "amenorrhea_concept"), "false"))));
}

/** Group s_malaria */
export function grp_s_malaria(app) {
  return X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "false")) && X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "false")))) && X.bool(X.eq(X.v(app, "has_pregnancy_danger_sign"), "false")))) && X.bool(X.eq(X.v(app, "has_postpartum_danger_sign"), "false")))) && X.bool(X.eq(X.v(app, "amenorrhea_concept"), "false"))));
}

/** Group s_malaria_indication_19 */
export function grp_s_malaria_indication_19(app) {
  return X.bool(X.eq(X.v(app, "s_malaria_tdr_done"), "yes"));
}

/** Group s_malaria_indication_20 */
export function grp_s_malaria_indication_20(app) {
  return X.bool(X.eq(X.v(app, "s_malaria_tdr_done"), "yes"));
}

/** Group s_malaria_indication_21 */
export function grp_s_malaria_indication_21(app) {
  return X.bool(X.eq(X.v(app, "s_malaria_tdr_done"), "yes"));
}

/** Group s_malaria_indication_22 */
export function grp_s_malaria_indication_22(app) {
  return X.bool(X.eq(X.v(app, "s_malaria_tdr_done"), "yes"));
}

/** Group s_malaria_indication_23 */
export function grp_s_malaria_indication_23(app) {
  return X.bool(X.eq(X.v(app, "s_malaria_tdr_done"), "yes"));
}

/** Group s_malaria_indication_24 */
export function grp_s_malaria_indication_24(app) {
  return X.bool(X.eq(X.v(app, "s_malaria_tdr_done"), "yes"));
}

/** Group s_malaria_indication_25 */
export function grp_s_malaria_indication_25(app) {
  return X.bool(X.eq(X.v(app, "s_malaria_tdr_done"), "yes"));
}

/** Group s_malaria_indication_26 */
export function grp_s_malaria_indication_26(app) {
  return X.bool(X.eq(X.v(app, "s_malaria_tdr_done"), "yes"));
}

/** Group s_malaria_indication_27 */
export function grp_s_malaria_indication_27(app) {
  return X.bool(X.eq(X.v(app, "s_malaria_tdr_done"), "yes"));
}

/** Group prescription_summary */
export function grp_prescription_summary(app) {
  return X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.gt(X.v(app, "albendazol_200_disp"), 0)) || X.bool(X.gt(X.v(app, "albendazol_400_disp"), 0)))) || X.bool(X.gt(X.v(app, "plumpy_sup_disp"), 0)))) || X.bool(X.gt(X.v(app, "plumpy_nut_disp"), 0)))) || X.bool(X.gt(X.v(app, "enriched_flour_disp"), 0)))) || X.bool(X.gt(X.v(app, "iron_folic_acid_srp_disp"), 0)))) || X.bool(X.gt(X.v(app, "iron_folic_acid_disp"), 0)))) || X.bool(X.gt(X.v(app, "amoxicillin_250_disp"), 0)))) || X.bool(X.gt(X.v(app, "vitamin_a_blue_capsule_disp"), 0)))) || X.bool(X.gt(X.v(app, "vitamin_a_red_capsule_disp"), 0)))) || X.bool(X.gt(X.v(app, "act_20120_disp"), 0)))) || X.bool(X.gt(X.v(app, "malaria_rdt_disp"), 0)))) || X.bool(X.gt(X.v(app, "artesunate_50_disp"), 0)))) || X.bool(X.gt(X.v(app, "sp_disp"), 0)))) || X.bool(X.gt(X.v(app, "zinc_20_disp"), 0)))) || X.bool(X.gt(X.v(app, "ors_disp"), 0)))) || X.bool(X.gt(X.v(app, "pregnancy_rdt_disp"), 0)))) || X.bool(X.gt(X.v(app, "sayana_press_disp"), 0)))) || X.bool(X.gt(X.v(app, "male_condom_disp"), 0)))) || X.bool(X.gt(X.v(app, "female_condom_disp"), 0)))) || X.bool(X.gt(X.v(app, "cycle_collar_disp"), 0)))) || X.bool(X.gt(X.v(app, "depo_provera_disp"), 0)))) || X.bool(X.gt(X.v(app, "postpill_disp"), 0)))) || X.bool(X.gt(X.v(app, "simple_pill_disp"), 0)))) || X.bool(X.gt(X.v(app, "low_dose_pill_disp"), 0)))) || X.bool(X.gt(X.v(app, "paracetamol_500_disp"), 0)))) && X.bool(X.eq(X.inp(app, "contact/parent/parent/stm_is_active"), "true"))));
}

/** Group contact */
export function grp_contact(app) {
  return grp_prescription_summary(app);
}

/** Group fields */
export function grp_fields(app) {
  return grp_prescription_summary(app);
}

/** Calculation of c50_patient_uuid */
export function calc_c50_patient_uuid(app) {
  return X.out(X.inp(app, "contact/_id"));
}

/** Calculation of c50_family_uuid */
export function calc_c50_family_uuid(app) {
  return X.out(X.inp(app, "contact/parent/_id"));
}

/** Calculation of c40_chw_area_uuid */
export function calc_c40_chw_area_uuid(app) {
  return X.out(X.inp(app, "contact/parent/parent/_id"));
}

/** Calculation of c30_supervisor_area_uuid */
export function calc_c30_supervisor_area_uuid(app) {
  return X.out(X.inp(app, "contact/parent/parent/parent/_id"));
}

/** Calculation of c20_health_area_uuid */
export function calc_c20_health_area_uuid(app) {
  return X.out(X.inp(app, "contact/parent/parent/parent/parent/_id"));
}

/** Calculation of c10_site_uuid */
export function calc_c10_site_uuid(app) {
  return X.out(X.inp(app, "contact/parent/parent/parent/parent/parent/_id"));
}

/** Calculation of patient_age_in_years */
export function calc_patient_age_in_years(app) {
  return X.out(X.f["floor"](app, (X.num(X.f["difference-in-months"](app, X.inp(app, "contact/date_of_birth"), X.f["now"](app))) / X.num(12))));
}

/** Default of patient_age_in_years */
export function def_patient_age_in_years(app) {
  return "0";
}

/** Calculation of patient_age_in_months */
export function calc_patient_age_in_months(app) {
  return X.out(X.f["difference-in-months"](app, X.inp(app, "contact/date_of_birth"), X.f["now"](app)));
}

/** Default of patient_age_in_months */
export function def_patient_age_in_months(app) {
  return "0";
}

/** Calculation of patient_age_in_days */
export function calc_patient_age_in_days(app) {
  return X.out(X.f["floor"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) - X.num(X.f["decimal-date-time"](app, X.inp(app, "contact/date_of_birth"))))));
}

/** Calculation of patient_age_display */
export function calc_patient_age_display(app) {
  return X.out(X.f["concat"](app, X.f["concat"](app, X.v(app, "patient_age_in_years"), (X.bool(X.eq(X.v(app, "patient_age_in_years"), 1)) ? " year and " : " years and ")), X.f["concat"](app, (X.num(X.v(app, "patient_age_in_months")) % X.num(12)), (X.bool(X.eq((X.num(X.v(app, "patient_age_in_months")) % X.num(12)), 1)) ? " month old" : " months old"))));
}

/** Calculation of patient_age_display_fr */
export function calc_patient_age_display_fr(app) {
  return X.out(X.f["concat"](app, X.f["concat"](app, X.v(app, "patient_age_in_years"), (X.bool(X.eq(X.v(app, "patient_age_in_years"), 1)) ? " an et " : " ans et ")), X.f["concat"](app, (X.num(X.v(app, "patient_age_in_months")) % X.num(12)), (X.bool(X.eq((X.num(X.v(app, "patient_age_in_months")) % X.num(12)), 1)) ? " mois" : " mois"))));
}

/** Calculation of patient_age_display_bm */
export function calc_patient_age_display_bm(app) {
  return X.out(X.f["concat"](app, X.f["concat"](app, (X.bool(X.ne(X.v(app, "patient_age_in_years"), "")) ? "San " : ""), X.v(app, "patient_age_in_years")), X.f["concat"](app, (X.bool(X.eq((X.num(X.v(app, "patient_age_in_months")) % X.num(12)), 1)) ? " ani kalo " : " ani kalo "), (X.num(X.v(app, "patient_age_in_months")) % X.num(12)))));
}

/** Calculation of patient_id */
export function calc_patient_id(app) {
  return X.out(X.inp(app, "contact/_id"));
}

/** Calculation of muso_id */
export function calc_muso_id(app) {
  return X.out(X.inp(app, "contact/external_id"));
}

/** Calculation of patient_name */
export function calc_patient_name(app) {
  return X.out(X.inp(app, "contact/name"));
}

/** Calculation of patient_sex */
export function calc_patient_sex(app) {
  return X.out(X.inp(app, "contact/sex"));
}

/** Calculation of patient_sex_en */
export function calc_patient_sex_en(app) {
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_sex"), "male")) || X.bool(X.eq(X.v(app, "patient_sex"), "homme")))) || X.bool(X.eq(X.v(app, "patient_sex"), "Homme")))) ? "Male" : (X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_sex"), "female")) || X.bool(X.eq(X.v(app, "patient_sex"), "femme")))) || X.bool(X.eq(X.v(app, "patient_sex"), "Femme")))) ? "Female" : "")));
}

/** Calculation of patient_sex_fr */
export function calc_patient_sex_fr(app) {
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_sex"), "male")) || X.bool(X.eq(X.v(app, "patient_sex"), "homme")))) || X.bool(X.eq(X.v(app, "patient_sex"), "Homme")))) ? "Homme" : (X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_sex"), "female")) || X.bool(X.eq(X.v(app, "patient_sex"), "femme")))) || X.bool(X.eq(X.v(app, "patient_sex"), "Femme")))) ? "Femme" : "")));
}

/** Calculation of patient_sex_bm */
export function calc_patient_sex_bm(app) {
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_sex"), "male")) || X.bool(X.eq(X.v(app, "patient_sex"), "homme")))) || X.bool(X.eq(X.v(app, "patient_sex"), "Homme")))) ? "Cɛ" : (X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_sex"), "female")) || X.bool(X.eq(X.v(app, "patient_sex"), "femme")))) || X.bool(X.eq(X.v(app, "patient_sex"), "Femme")))) ? "Muso" : "")));
}

/** Calculation of patient_phone */
export function calc_patient_phone(app) {
  return X.out(X.inp(app, "contact/phone"));
}

/** Calculation of cscom_area */
export function calc_cscom_area(app) {
  return X.out(X.inp(app, "contact/parent/parent/aire_de_sante"));
}

/** Calculation of chw_id */
export function calc_chw_id(app) {
  return X.out(X.inp(app, "contact/parent/parent/contact/_id"));
}

/** Calculation of chw_name */
export function calc_chw_name(app) {
  return X.out(X.inp(app, "contact/parent/parent/contact/name"));
}

/** Calculation of chw_muso_id */
export function calc_chw_muso_id(app) {
  return X.out(X.inp(app, "contact/parent/parent/contact/patient_id"));
}

/** Calculation of accompany_to_cscom */
export function calc_accompany_to_cscom(app) {
  return X.out((X.bool(X.eq(X.v(app, "has_danger_sign"), "true")) ? "true" : "false"));
}

/** Calculation of refer_to_cscom */
export function calc_refer_to_cscom(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "true")) || X.bool(X.eq(X.v(app, "has_pregnancy_danger_sign"), "true")))) || X.bool(X.eq(X.v(app, "has_postpartum_danger_sign"), "true")))) || X.bool(X.eq(X.v(app, "CAT"), "refer")))) || X.bool(X.eq(X.v(app, "CAT_pregnancy"), "refer")))) || X.bool(X.eq(X.v(app, "CAT_mrdt"), "refer")))) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false")))) ? "true" : "false"));
}

/** Calculation of referral */
export function calc_referral(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")) || X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true")))) || X.bool(X.eq(X.v(app, "has_been_accompany_to_cscom"), "true")))) || X.bool(X.eq(X.v(app, "has_pregnancy_danger_sign"), "true")))) || X.bool(X.eq(X.v(app, "has_postpartum_danger_sign"), "true")))) || X.bool(X.eq(X.v(app, "CAT_out_of_stock"), "refer")))) ? "true" : "false"));
}

/** Calculation of has_tb_sign */
export function calc_has_tb_sign(app) {
  return X.out((X.bool(X.eq(X.v(app, "type_renamed"), "contact")) ? (X.bool(X.eq(X.v(app, "s_acc_danger_sign_cough_more_than_14_days"), "yes")) ? "true" : "false") : ""));
}

/** Calculation of tb_referral */
export function calc_tb_referral(app) {
  return X.out((X.bool((X.bool(X.eq(X.v(app, "has_tb_sign"), "true")) && X.bool(X.eq(X.v(app, "s_accompany_refer_CSCOM"), "yes")))) ? "true" : "false"));
}

/** Calculation of needs_signoff */
export function calc_needs_signoff(app) {
  return X.out((X.bool(X.eq(X.v(app, "tb_referral"), "true")) ? "true" : ""));
}

/** Calculation of acc_danger_sign_hemoptysis_fr */
export function calc_acc_danger_sign_hemoptysis_fr(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_acc_danger_sign_hemoptysis"), "yes")) ? "Hémoptysie (crachats teintés de sang)" : ""));
}

/** Calculation of acc_danger_sign_hemoptysis_en */
export function calc_acc_danger_sign_hemoptysis_en(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_acc_danger_sign_hemoptysis"), "yes")) ? "Hemoptysis (blood-stained sputum)" : ""));
}

/** Calculation of acc_danger_sign_hemoptysis_bm */
export function calc_acc_danger_sign_hemoptysis_bm(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_acc_danger_sign_hemoptysis"), "yes")) ? "Da ji joliman" : ""));
}

/** Calculation of acc_danger_sign_wheezing_fr */
export function calc_acc_danger_sign_wheezing_fr(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_acc_danger_sign_wheezing"), "yes")) ? "Geignements" : ""));
}

/** Calculation of acc_danger_sign_wheezing_en */
export function calc_acc_danger_sign_wheezing_en(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_acc_danger_sign_wheezing"), "yes")) ? "Wheezing" : ""));
}

/** Calculation of acc_danger_sign_wheezing_bm */
export function calc_acc_danger_sign_wheezing_bm(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_acc_danger_sign_wheezing"), "yes")) ? "Ka ŋunan" : ""));
}

/** Calculation of acc_danger_sign_chest_indrawing_fr */
export function calc_acc_danger_sign_chest_indrawing_fr(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_acc_danger_sign_chest_indrawing"), "yes")) ? "Tirage sous-costal, détresse respiratoire (difficultés respiratoire…)" : ""));
}

/** Calculation of acc_danger_sign_chest_indrawing_en */
export function calc_acc_danger_sign_chest_indrawing_en(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_acc_danger_sign_chest_indrawing"), "yes")) ? "Chest Indrawing" : ""));
}

/** Calculation of acc_danger_sign_chest_indrawing_bm */
export function calc_acc_danger_sign_chest_indrawing_bm(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_acc_danger_sign_chest_indrawing"), "yes")) ? "Galakakolo cɛ janyanli" : ""));
}

/** Calculation of ref_danger_sign_persistent_fever_fr */
export function calc_ref_danger_sign_persistent_fever_fr(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_ref_danger_sign_persistent_fever"), "yes")) ? "Fièvre persistante (>=3 semaines)" : ""));
}

/** Calculation of ref_danger_sign_persistent_fever_en */
export function calc_ref_danger_sign_persistent_fever_en(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_ref_danger_sign_persistent_fever"), "yes")) ? "Persistent fever (>= 3 weeks)" : ""));
}

/** Calculation of ref_danger_sign_persistent_fever_bm */
export function calc_ref_danger_sign_persistent_fever_bm(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_ref_danger_sign_persistent_fever"), "yes")) ? "Farikalaya kuntaala jan min bɛ dɔgɔkun saba bɔ walima a ka can ni dɔgɔkun saba ye" : ""));
}

/** Calculation of ref_danger_sign_weight_loss_fr */
export function calc_ref_danger_sign_weight_loss_fr(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_ref_danger_sign_weight_loss"), "yes")) ? "Perte de poids ou amaigrissement visible" : ""));
}

/** Calculation of ref_danger_sign_weight_loss_en */
export function calc_ref_danger_sign_weight_loss_en(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_ref_danger_sign_weight_loss"), "yes")) ? "Visible weight loss or weight loss" : ""));
}

/** Calculation of ref_danger_sign_weight_loss_bm */
export function calc_ref_danger_sign_weight_loss_bm(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_ref_danger_sign_weight_loss"), "yes")) ? "Fasa walima fasali min bɛ dɔn ɲɛ ka yeli la" : ""));
}

/** Calculation of acc_danger_sign_cough_more_than_14_days_fr */
export function calc_acc_danger_sign_cough_more_than_14_days_fr(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_acc_danger_sign_cough_more_than_14_days"), "yes")) ? "Toux qui dure plus de 14 jours sans amélioration" : ""));
}

/** Calculation of acc_danger_sign_cough_more_than_14_days_en */
export function calc_acc_danger_sign_cough_more_than_14_days_en(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_acc_danger_sign_cough_more_than_14_days"), "yes")) ? "Cough that lasts longer than 14 days without improvement" : ""));
}

/** Calculation of acc_danger_sign_cough_more_than_14_days_bm */
export function calc_acc_danger_sign_cough_more_than_14_days_bm(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_acc_danger_sign_cough_more_than_14_days"), "yes")) ? "Sɔgɔsɔgɔ kuntaala jan min ka can ni tile 14 ye a ma nɔgɔya" : ""));
}

/** Calculation of acc_danger_sign_chest_pain_fr */
export function calc_acc_danger_sign_chest_pain_fr(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_acc_danger_sign_chest_pain"), "yes")) ? "Douleurs thoraciques" : ""));
}

/** Calculation of acc_danger_sign_chest_pain_en */
export function calc_acc_danger_sign_chest_pain_en(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_acc_danger_sign_chest_pain"), "yes")) ? "Chest pain" : ""));
}

/** Calculation of acc_danger_sign_chest_pain_bm */
export function calc_acc_danger_sign_chest_pain_bm(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_acc_danger_sign_chest_pain"), "yes")) ? "Disi dimi" : ""));
}

/** Calculation of tb_diagnosis_fr */
export function calc_tb_diagnosis_fr(app) {
  return X.out((X.bool(X.eq(X.v(app, "has_tb_sign"), "true")) ? X.f["concat"](app, X.v(app, "acc_danger_sign_cough_more_than_14_days_fr")) : ""));
}

/** Calculation of tb_diagnosis_en */
export function calc_tb_diagnosis_en(app) {
  return X.out((X.bool(X.eq(X.v(app, "has_tb_sign"), "true")) ? X.f["concat"](app, X.v(app, "acc_danger_sign_cough_more_than_14_days_en")) : ""));
}

/** Calculation of tb_diagnosis_bm */
export function calc_tb_diagnosis_bm(app) {
  return X.out((X.bool(X.eq(X.v(app, "has_tb_sign"), "true")) ? X.f["concat"](app, X.v(app, "acc_danger_sign_cough_more_than_14_days_bm")) : ""));
}

/** Calculation of tdr_done */
export function calc_tdr_done(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_malaria_tdr_done"), "yes")) ? "yes" : "no"));
}

/** Calculation of symptom_fever */
export function calc_symptom_fever(app) {
  return X.out((X.bool(X.ge(X.v(app, "s_child_temperature"), 37.5)) ? X.f["concat"](app, "Fever that began: ", X.choiceName("c_when_illness", X.v(app, "s_when_illness"))) : ""));
}

/** Calculation of tdr_result */
export function calc_tdr_result(app) {
  return X.out((X.bool((X.bool(X.eq(X.v(app, "s_malaria_tdr_result_new"), "pos")) || X.bool(X.eq(X.v(app, "s_malaria_tdr_result"), "pos")))) ? "positif" : (X.bool((X.bool(X.eq(X.v(app, "s_malaria_tdr_result_new"), "neg")) || X.bool(X.eq(X.v(app, "s_malaria_tdr_result"), "neg")))) ? "négatif" : "")));
}

/** Calculation of treat_malaria */
export function calc_treat_malaria(app) {
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "tdr_result"), "positif")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false")))) && X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")))) ? "true" : "false"));
}

/** Calculation of next_visit */
export function calc_next_visit(app) {
  return X.out((X.bool(X.eq(X.v(app, "diff_assess_date_report_date"), 0)) ? X.f["format-date-time"](app, X.f["date-time"](app, (X.num(X.f["decimal-date-time"](app, X.v(app, "assessment_date"))) + X.num(1))), "%d/%m/%Y") : X.f["format-date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app)), "%d/%m/%Y")));
}

/** Calculation of observe */
export function calc_observe(app) {
  return X.out((X.bool((X.bool(X.eq(X.v(app, "CAT"), "watching")) || X.bool(X.eq(X.v(app, "CAT_pregnancy"), "watching")))) ? "true" : "false"));
}

/** Calculation of visited_contact_uuid */
export function calc_visited_contact_uuid(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_how_child_found"), "home_visit")) ? X.inp(app, "contact/parent/_id") : ""));
}

/** Calculation of visited_date */
export function calc_visited_date(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_assess_today"), "yes")) ? X.f["format-date-time"](app, X.f["date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app))), "%Y-%m-%d") : X.v(app, "s_assess_date")));
}

/** Calculation of rc_code */
export function calc_rc_code(app) {
  return X.out((X.bool(X.eq(X.v(app, "has_danger_sign"), "true")) ? "34221" : "34221"));
}

/** Calculation of can_access_sih */
export function calc_can_access_sih(app) {
  return X.out(X.inp(app, "user/is_in_sih"));
}

/** Calculation of can_access_mrdt */
export function calc_can_access_mrdt(app) {
  return X.out("true");
}

/** Calculation of shakir */
export function calc_shakir(app) {
  return X.out((X.bool(X.eq(X.v(app, "has_danger_sign"), "true")) ? "" : ""));
}

/** Calculation of assessment_date */
export function calc_assessment_date(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_assess_today"), "yes")) ? X.f["format-date-time"](app, X.f["date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app))), "%Y-%m-%d") : X.v(app, "s_assess_date")));
}

/** Calculation of no_tdr_done */
export function calc_no_tdr_done(app) {
  return X.out(X.v(app, "s_malaria_tdr_not_done"));
}

/** Calculation of no_tdr_done_fr */
export function calc_no_tdr_done_fr(app) {
  return X.out((X.bool(X.eq(X.v(app, "no_tdr_done"), "tdr_unavailable")) ? "Pas de TDR: TDR non disponible" : X.f["concat"](app, "Pas de TDR: ", X.v(app, "s_malaria_tdr_not_done_other"))));
}

/** Calculation of no_tdr_done_en */
export function calc_no_tdr_done_en(app) {
  return X.out((X.bool(X.eq(X.v(app, "no_tdr_done"), "tdr_unavailable")) ? "No RDT: RDT unavailable" : X.f["concat"](app, "No RDT: ", X.v(app, "s_malaria_tdr_not_done_other"))));
}

/** Calculation of no_tdr_done_bm */
export function calc_no_tdr_done_bm(app) {
  return X.out((X.bool(X.eq(X.v(app, "no_tdr_done"), "tdr_unavailable")) ? "TDR ma kɛ: TDR ban nen do" : X.f["concat"](app, "TDR ma kɛ: ", X.v(app, "s_malaria_tdr_not_done_other"))));
}

/** Calculation of diff_assess_date_report_date */
export function calc_diff_assess_date_report_date(app) {
  return X.out(X.f["floor"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) - X.num(X.f["decimal-date-time"](app, X.v(app, "assessment_date"))))));
}

/** Calculation of is_tdr_unavailable */
export function calc_is_tdr_unavailable(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_malaria_tdr_not_done"), "tdr_unavailable")) ? "true" : "false"));
}

/** Calculation of c_out_of_stock */
export function calc_c_out_of_stock(app) {
  return X.out((X.bool((X.bool(X.eq(X.v(app, "s_not_give_paracetamol"), "out_of_stock")) || X.bool(X.eq(X.v(app, "s_malaria_not_give_act"), "out_of_stock")))) ? "true" : "false"));
}

/** Calculation of paracetamol_status */
export function calc_paracetamol_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_give_paracetamol"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "s_give_paracetamol"), "yes")) ? "given" : X.v(app, "s_not_give_paracetamol"))));
}

/** Calculation of act_status */
export function calc_act_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_malaria_give_act"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "s_malaria_give_act"), "yes")) ? "given" : X.v(app, "s_malaria_not_give_act"))));
}

/** Calculation of rdt_status */
export function calc_rdt_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "tdr_done"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "tdr_done"), "yes")) ? "given" : X.v(app, "no_tdr_done"))));
}

/** Calculation of drugs_not_proposed */
export function calc_drugs_not_proposed(app) {
  return X.out((X.num((X.bool(X.eq(X.v(app, "paracetamol_status"), "not_proposed")) ? 1 : 0)) + X.num((X.bool(X.eq(X.v(app, "act_status"), "not_proposed")) ? 1 : 0))));
}

/** Calculation of other_products_not_proposed */
export function calc_other_products_not_proposed(app) {
  return X.out((X.bool(X.eq(X.v(app, "rdt_status"), "not_proposed")) ? 1 : 0));
}

/** Calculation of drugs_proposed */
export function calc_drugs_proposed(app) {
  return X.out((X.num(2) - X.num(X.v(app, "drugs_not_proposed"))));
}

/** Calculation of other_products_proposed */
export function calc_other_products_proposed(app) {
  return X.out((X.num(1) - X.num(X.v(app, "other_products_not_proposed"))));
}

/** Calculation of drugs_given */
export function calc_drugs_given(app) {
  return X.out((X.num((X.bool(X.eq(X.v(app, "paracetamol_status"), "given")) ? 1 : 0)) + X.num((X.bool(X.eq(X.v(app, "act_status"), "given")) ? 1 : 0))));
}

/** Calculation of other_products_given */
export function calc_other_products_given(app) {
  return X.out((X.bool(X.eq(X.v(app, "rdt_status"), "given")) ? 1 : 0));
}

/** Calculation of drugs_not_given */
export function calc_drugs_not_given(app) {
  return X.out((X.num(X.v(app, "drugs_proposed")) - X.num(X.v(app, "drugs_given"))));
}

/** Calculation of other_products_not_given */
export function calc_other_products_not_given(app) {
  return X.out((X.num(X.v(app, "other_products_proposed")) - X.num(X.v(app, "other_products_given"))));
}

/** Calculation of drugs_not_given_out_of_stock */
export function calc_drugs_not_given_out_of_stock(app) {
  return X.out((X.num((X.bool(X.eq(X.v(app, "paracetamol_status"), "out_of_stock")) ? 1 : 0)) + X.num((X.bool(X.eq(X.v(app, "act_status"), "out_of_stock")) ? 1 : 0))));
}

/** Calculation of other_products_not_given_out_of_stock */
export function calc_other_products_not_given_out_of_stock(app) {
  return X.out((X.bool(X.eq(X.v(app, "rdt_status"), "tdr_unavailable")) ? 1 : 0));
}

/** Calculation of drugs_not_given_other */
export function calc_drugs_not_given_other(app) {
  return X.out((X.num(X.v(app, "drugs_not_given")) - X.num(X.v(app, "drugs_not_given_out_of_stock"))));
}

/** Calculation of other_products_not_given_other */
export function calc_other_products_not_given_other(app) {
  return X.out((X.num(X.v(app, "other_products_not_given")) - X.num(X.v(app, "other_products_not_given_out_of_stock"))));
}

/** Calculation of show_drugs_cat */
export function calc_show_drugs_cat(app) {
  return X.out((X.bool((X.bool(X.eq(X.f["number"](app, X.v(app, "drugs_proposed")), X.f["number"](app, X.v(app, "drugs_not_given_out_of_stock")))) || X.bool(X.eq(X.f["number"](app, X.v(app, "drugs_proposed")), X.f["number"](app, X.v(app, "drugs_given")))))) ? "false" : (X.bool((X.bool((X.bool(X.eq(X.v(app, "CAT"), "")) && X.bool(X.eq(X.v(app, "CAT_mrdt"), "")))) && X.bool(X.eq(X.v(app, "CAT_pregnancy"), "")))) ? "true" : "false")));
}

/** Calculation of has_been_accompany_to_cscom */
export function calc_has_been_accompany_to_cscom(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_accompany_refer_CSCOM"), "yes")) ? "true" : "false"));
}

/** Calculation of type_renamed */
export function calc_type_renamed(app) {
  return X.out(X.inp(app, "contact/parent/type"));
}

/** Calculation of s_fever_within_24_48 */
export function calc_s_fever_within_24_48(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_malaria_fever_within_24_48"), "yes")) ? "yes" : "no"));
}

/** Calculation of s_fever */
export function calc_s_fever(app) {
  return X.out((X.bool((X.bool(X.ge(X.v(app, "s_child_temperature"), 37.5)) || X.bool(X.eq(X.v(app, "s_fever_within_24_48"), "yes")))) ? "yes" : "no"));
}

/** Calculation of s_has_fever_now */
export function calc_s_has_fever_now(app) {
  return X.out((X.bool(X.ge(X.v(app, "s_child_temperature"), 37.5)) ? "yes" : "no"));
}

/** Calculation of s_child_temperature */
export function calc_s_child_temperature(app) {
  return X.out((X.bool((X.bool(X.lt(X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw"), 0), 35)) || X.bool(X.gt(X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw"), 0), 41)))) ? (X.num(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw_retake"), 0))) + X.num(X.f["number"](app, 0))) : (X.num(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw"), 0))) + X.num(X.f["number"](app, 0)))));
}

/** Calculation of has_first_ref_danger_sign */
export function calc_has_first_ref_danger_sign(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_ref_danger_sign_severe_jaundice"), "yes")) || X.bool(X.eq(X.v(app, "s_ref_danger_sign_dark_colored_urine"), "yes")))) || X.bool(X.eq(X.v(app, "s_ref_danger_sign_conjunctival_pallor"), "yes")))) || X.bool(X.eq(X.v(app, "s_ref_danger_sign_eodema_lower_legs"), "yes")))) || X.bool(X.eq(X.v(app, "s_ref_danger_sign_persistent_fever"), "yes")))) || X.bool(X.eq(X.v(app, "s_ref_danger_sign_weight_loss"), "yes")))) ? "true" : "false"));
}

/** Calculation of has_ref_danger_sign */
export function calc_has_ref_danger_sign(app) {
  return X.out((X.bool((X.bool(X.eq(X.v(app, "has_first_ref_danger_sign"), "true")) || X.bool(X.eq(X.v(app, "s_ref_danger_sign_other_diseases"), "yes")))) ? "true" : "false"));
}

/** Calculation of has_danger_sign */
export function calc_has_danger_sign(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_acc_danger_sign_seizure"), "yes")) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_loss_consiousness"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_unable_drink"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_confusion"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_vomit"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_chest_indrawing"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_wheezing"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_hemoptysis"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_cough_more_than_14_days"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_chest_pain"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_lathargy"), "yes")))) ? "true" : "false"));
}

/** Calculation of patient_can_procreate */
export function calc_patient_can_procreate(app) {
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_sex"), "female")) && X.bool(X.ge(X.v(app, "patient_age_in_years"), 10)))) && X.bool(X.lt(X.v(app, "patient_age_in_years"), 65)))) ? "true" : "false"));
}

/** Calculation of has_pregnancy_danger_sign */
export function calc_has_pregnancy_danger_sign(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "pregnancy_danger_signs_note"), "yes")) || X.bool(X.eq(X.v(app, "vaginal_bleeding"), "yes")))) || X.bool(X.eq(X.v(app, "pelvic_pain"), "yes")))) || X.bool(X.eq(X.v(app, "headache_swelling"), "yes")))) || X.bool(X.eq(X.v(app, "fetal_movement_loss"), "yes")))) || X.bool(X.eq(X.v(app, "severe_vomiting"), "yes")))) || X.bool(X.eq(X.v(app, "uterine_contractions"), "yes")))) || X.bool(X.eq(X.v(app, "water_breaking"), "yes")))) || X.bool(X.eq(X.v(app, "urination_burning"), "yes")))) || X.bool(X.eq(X.v(app, "vaginal_discharge_odor"), "yes")))) || X.bool(X.eq(X.v(app, "postpartum_depression"), "yes")))) || X.bool(X.eq(X.v(app, "domestic_violence_signs"), "yes")))) ? "true" : "false"));
}

/** Calculation of has_postpartum_danger_sign */
export function calc_has_postpartum_danger_sign(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "heavy_bleeding"), "yes")) || X.bool(X.eq(X.v(app, "headache_face_swelling"), "yes")))) || X.bool(X.eq(X.v(app, "excessive_fatigue"), "yes")))) || X.bool(X.eq(X.v(app, "marked_paleness"), "yes")))) || X.bool(X.eq(X.v(app, "mobility_difficulty"), "yes")))) || X.bool(X.eq(X.v(app, "severe_abdominal_pain"), "yes")))) || X.bool(X.eq(X.v(app, "urination_burning_difficulty"), "yes")))) || X.bool(X.eq(X.v(app, "foul_vaginal_discharge"), "yes")))) || X.bool(X.eq(X.v(app, "breast_pain_redness"), "yes")))) || X.bool(X.eq(X.v(app, "postpartum_depression_signs"), "yes")))) ? "true" : "false"));
}

/** Calculation of amenorrhea_concept */
export function calc_amenorrhea_concept(app) {
  return X.out((X.bool(X.eq(X.v(app, "patient_missed_period"), "yes")) ? "true" : "false"));
}

/** Relevance of s_assess_date */
export function rel_s_assess_date(app) {
  return X.bool(X.eq(X.v(app, "s_assess_today"), "no"));
}

/** Constraint of s_assess_date */
export function val_s_assess_date(app) {
  if (X.str(X.v(app, "s_assess_date")) === '') return true;
  return X.bool(X.le(X.f["floor"](app, X.f["decimal-date-time"](app, X.v(app, "s_assess_date"))), X.f["floor"](app, X.f["decimal-date-time"](app, X.f["now"](app)))));
}

/** Relevance of s_how_child_found_other */
export function rel_s_how_child_found_other(app) {
  return X.bool(X.eq(X.v(app, "s_how_child_found"), "other"));
}

/** Relevance of patient_is_pregnant */
export function rel_patient_is_pregnant(app) {
  return grp_pregnancy_section(app);
}

/** Relevance of patient_missed_period */
export function rel_patient_missed_period(app) {
  return grp_pregnancy_section(app) && X.bool(X.eq(X.v(app, "patient_is_pregnant"), "no"));
}

/** Relevance of patient_gave_birth_less_48_day */
export function rel_patient_gave_birth_less_48_day(app) {
  return grp_postpartum_section(app);
}

/** Relevance of s_child_temperature_pre_chw */
export function rel_s_child_temperature_pre_chw(app) {
  return grp_s_temperature(app);
}

/** Constraint of s_child_temperature_pre_chw */
export function val_s_child_temperature_pre_chw(app) {
  if (X.str(X.v(app, "s_child_temperature_pre_chw")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "s_child_temperature_pre_chw"), 35)) && X.bool(X.le(X.v(app, "s_child_temperature_pre_chw"), 45))));
}

/** Relevance of s_malaria_fever_within_24_48 */
export function rel_s_malaria_fever_within_24_48(app) {
  return grp_s_temperature(app) && X.bool((X.bool((X.bool((X.bool(X.ne(X.v(app, "s_child_temperature"), 0)) && X.bool(X.ge(X.v(app, "s_child_temperature"), 35)))) && X.bool(X.lt(X.v(app, "s_child_temperature"), 37.5)))) && X.bool(X.eq(X.v(app, "amenorrhea_concept"), "false"))));
}

/** Relevance of retake_temperature */
export function rel_retake_temperature(app) {
  return grp_s_temperature_retake(app);
}

/** Relevance of s_child_temperature_pre_chw_retake */
export function rel_s_child_temperature_pre_chw_retake(app) {
  return grp_s_temperature_retake(app);
}

/** Constraint of s_child_temperature_pre_chw_retake */
export function val_s_child_temperature_pre_chw_retake(app) {
  if (X.str(X.v(app, "s_child_temperature_pre_chw_retake")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "s_child_temperature_pre_chw_retake"), 35)) && X.bool(X.le(X.v(app, "s_child_temperature_pre_chw_retake"), 45))));
}

/** Relevance of s_thermometer_note */
export function rel_s_thermometer_note(app) {
  return X.bool(X.eq(X.v(app, "amenorrhea_concept"), "false"));
}

/** Calculation of s_malaria_sign_fever */
export function calc_s_malaria_sign_fever(app) {
  return X.out(X.v(app, "s_fever"));
}

/** Relevance of s_acc_danger_signs_note */
export function rel_s_acc_danger_signs_note(app) {
  return grp_s_acc_danger_signs(app);
}

/** Relevance of s_acc_danger_sign_seizure */
export function rel_s_acc_danger_sign_seizure(app) {
  return grp_s_acc_danger_signs(app);
}

/** Relevance of s_acc_danger_sign_loss_consiousness */
export function rel_s_acc_danger_sign_loss_consiousness(app) {
  return grp_s_acc_danger_signs(app);
}

/** Relevance of s_acc_danger_sign_unable_drink */
export function rel_s_acc_danger_sign_unable_drink(app) {
  return grp_s_acc_danger_signs(app);
}

/** Relevance of s_acc_danger_sign_confusion */
export function rel_s_acc_danger_sign_confusion(app) {
  return grp_s_acc_danger_signs(app);
}

/** Relevance of s_acc_danger_sign_vomit */
export function rel_s_acc_danger_sign_vomit(app) {
  return grp_s_acc_danger_signs(app);
}

/** Relevance of s_acc_danger_sign_chest_indrawing */
export function rel_s_acc_danger_sign_chest_indrawing(app) {
  return grp_s_acc_danger_signs(app);
}

/** Relevance of s_acc_danger_sign_wheezing */
export function rel_s_acc_danger_sign_wheezing(app) {
  return grp_s_acc_danger_signs(app);
}

/** Relevance of s_acc_danger_sign_cough_more_than_14_days */
export function rel_s_acc_danger_sign_cough_more_than_14_days(app) {
  return grp_s_acc_danger_signs(app);
}

/** Relevance of s_acc_danger_sign_chest_pain */
export function rel_s_acc_danger_sign_chest_pain(app) {
  return grp_s_acc_danger_signs(app);
}

/** Relevance of s_acc_danger_sign_hemoptysis */
export function rel_s_acc_danger_sign_hemoptysis(app) {
  return grp_s_acc_danger_signs(app) && X.bool(X.eq(X.v(app, "type_renamed"), "contact"));
}

/** Relevance of s_acc_danger_sign_lathargy */
export function rel_s_acc_danger_sign_lathargy(app) {
  return grp_s_acc_danger_signs(app);
}

/** Relevance of s_ref_danger_signs_note */
export function rel_s_ref_danger_signs_note(app) {
  return grp_s_ref_danger_signs(app);
}

/** Relevance of s_ref_danger_sign_severe_jaundice */
export function rel_s_ref_danger_sign_severe_jaundice(app) {
  return grp_s_ref_danger_signs(app);
}

/** Relevance of s_ref_danger_sign_dark_colored_urine */
export function rel_s_ref_danger_sign_dark_colored_urine(app) {
  return grp_s_ref_danger_signs(app);
}

/** Relevance of s_ref_danger_sign_conjunctival_pallor */
export function rel_s_ref_danger_sign_conjunctival_pallor(app) {
  return grp_s_ref_danger_signs(app);
}

/** Relevance of s_ref_danger_sign_eodema_lower_legs */
export function rel_s_ref_danger_sign_eodema_lower_legs(app) {
  return grp_s_ref_danger_signs(app);
}

/** Relevance of s_ref_danger_sign_persistent_fever */
export function rel_s_ref_danger_sign_persistent_fever(app) {
  return grp_s_ref_danger_signs(app) && X.bool(X.eq(X.v(app, "type_renamed"), "contact"));
}

/** Relevance of s_ref_danger_sign_weight_loss */
export function rel_s_ref_danger_sign_weight_loss(app) {
  return grp_s_ref_danger_signs(app) && X.bool(X.eq(X.v(app, "type_renamed"), "contact"));
}

/** Relevance of pregnancy_danger_signs_note */
export function rel_pregnancy_danger_signs_note(app) {
  return grp_pregnancy_danger_signs(app);
}

/** Relevance of vaginal_bleeding */
export function rel_vaginal_bleeding(app) {
  return grp_pregnancy_danger_signs(app);
}

/** Relevance of pelvic_pain */
export function rel_pelvic_pain(app) {
  return grp_pregnancy_danger_signs(app);
}

/** Relevance of headache_swelling */
export function rel_headache_swelling(app) {
  return grp_pregnancy_danger_signs(app);
}

/** Relevance of fetal_movement_loss */
export function rel_fetal_movement_loss(app) {
  return grp_pregnancy_danger_signs(app);
}

/** Relevance of severe_vomiting */
export function rel_severe_vomiting(app) {
  return grp_pregnancy_danger_signs(app);
}

/** Relevance of uterine_contractions */
export function rel_uterine_contractions(app) {
  return grp_pregnancy_danger_signs(app);
}

/** Relevance of water_breaking */
export function rel_water_breaking(app) {
  return grp_pregnancy_danger_signs(app);
}

/** Relevance of urination_burning */
export function rel_urination_burning(app) {
  return grp_pregnancy_danger_signs(app);
}

/** Relevance of vaginal_discharge_odor */
export function rel_vaginal_discharge_odor(app) {
  return grp_pregnancy_danger_signs(app);
}

/** Relevance of postpartum_depression */
export function rel_postpartum_depression(app) {
  return grp_pregnancy_danger_signs(app);
}

/** Relevance of domestic_violence_signs */
export function rel_domestic_violence_signs(app) {
  return grp_pregnancy_danger_signs(app);
}

/** Relevance of postpartum_danger_signs_note */
export function rel_postpartum_danger_signs_note(app) {
  return grp_postpartum_danger_signs(app);
}

/** Relevance of heavy_bleeding */
export function rel_heavy_bleeding(app) {
  return grp_postpartum_danger_signs(app);
}

/** Relevance of headache_face_swelling */
export function rel_headache_face_swelling(app) {
  return grp_postpartum_danger_signs(app);
}

/** Relevance of excessive_fatigue */
export function rel_excessive_fatigue(app) {
  return grp_postpartum_danger_signs(app);
}

/** Relevance of marked_paleness */
export function rel_marked_paleness(app) {
  return grp_postpartum_danger_signs(app);
}

/** Relevance of mobility_difficulty */
export function rel_mobility_difficulty(app) {
  return grp_postpartum_danger_signs(app);
}

/** Relevance of severe_abdominal_pain */
export function rel_severe_abdominal_pain(app) {
  return grp_postpartum_danger_signs(app);
}

/** Relevance of urination_burning_difficulty */
export function rel_urination_burning_difficulty(app) {
  return grp_postpartum_danger_signs(app);
}

/** Relevance of foul_vaginal_discharge */
export function rel_foul_vaginal_discharge(app) {
  return grp_postpartum_danger_signs(app);
}

/** Relevance of breast_pain_redness */
export function rel_breast_pain_redness(app) {
  return grp_postpartum_danger_signs(app);
}

/** Relevance of postpartum_depression_signs */
export function rel_postpartum_depression_signs(app) {
  return grp_postpartum_danger_signs(app);
}

/** Relevance of s_ref_danger_sign_other_diseases */
export function rel_s_ref_danger_sign_other_diseases(app) {
  return grp_s_ref_other_diseases(app);
}

/** Relevance of s_ref_danger_sign_other */
export function rel_s_ref_danger_sign_other(app) {
  return grp_s_ref_other_diseases(app) && X.bool(X.eq(X.v(app, "s_ref_danger_sign_other_diseases"), "yes"));
}

/** Relevance of has_mal_danger_sign */
export function rel_has_mal_danger_sign(app) {
  return grp_s_malaria(app);
}

/** Calculation of has_mal_danger_sign */
export function calc_has_mal_danger_sign(app) {
  if (!rel_has_mal_danger_sign(app)) return '';
  return X.out((X.bool(X.eq(X.v(app, "s_malaria_sign_fever"), "yes")) ? "true" : "false"));
}

/** Relevance of s_malaria_tdr_done */
export function rel_s_malaria_tdr_done(app) {
  return grp_s_malaria(app) && X.bool(X.eq(X.v(app, "has_mal_danger_sign"), "true"));
}

/** Relevance of s_malaria_tdr_result */
export function rel_s_malaria_tdr_result(app) {
  return grp_s_malaria(app) && X.bool((X.bool(X.eq(X.v(app, "s_malaria_tdr_done"), "yes")) && X.bool(X.ne(X.v(app, "can_access_mrdt"), "true"))));
}

/** Relevance of s_malaria_tdr_not_done */
export function rel_s_malaria_tdr_not_done(app) {
  return grp_s_malaria(app) && X.bool(X.eq(X.v(app, "s_malaria_tdr_done"), "no"));
}

/** Relevance of s_malaria_tdr_not_done_other */
export function rel_s_malaria_tdr_not_done_other(app) {
  return grp_s_malaria(app) && X.bool(X.eq(X.v(app, "s_malaria_tdr_not_done"), "other"));
}

/** Relevance of n_malaria_tdr_refer_cscom */
export function rel_n_malaria_tdr_refer_cscom(app) {
  return grp_s_malaria(app) && X.bool(X.eq(X.v(app, "s_malaria_tdr_not_done"), "tdr_unavailable"));
}

/** Relevance of CAT */
export function rel_CAT(app) {
  return grp_s_malaria(app) && X.bool((X.bool(X.eq(X.v(app, "s_malaria_tdr_result"), "neg")) || X.bool((X.bool(X.eq(X.v(app, "s_malaria_tdr_done"), "no")) && X.bool(X.eq(X.v(app, "s_malaria_tdr_not_done"), "other"))))));
}

/** Relevance of r_note_disinfect */
export function rel_r_note_disinfect(app) {
  return grp_s_malaria_indication_19(app);
}

/** Relevance of note_malaria_2 */
export function rel_note_malaria_2(app) {
  return grp_s_malaria_indication_20(app);
}

/** Relevance of note_malaria_3 */
export function rel_note_malaria_3(app) {
  return grp_s_malaria_indication_21(app);
}

/** Relevance of note_malaria_4 */
export function rel_note_malaria_4(app) {
  return grp_s_malaria_indication_22(app);
}

/** Relevance of note_malaria_5 */
export function rel_note_malaria_5(app) {
  return grp_s_malaria_indication_23(app);
}

/** Relevance of note_malaria_6 */
export function rel_note_malaria_6(app) {
  return grp_s_malaria_indication_24(app);
}

/** Relevance of note_malaria_7 */
export function rel_note_malaria_7(app) {
  return grp_s_malaria_indication_25(app);
}

/** Relevance of note_malaria_8 */
export function rel_note_malaria_8(app) {
  return grp_s_malaria_indication_26(app);
}

/** Relevance of evaluation_timer */
export function rel_evaluation_timer(app) {
  return grp_s_malaria_indication_26(app);
}

/** Relevance of note_danger_signs_pos */
export function rel_note_danger_signs_pos(app) {
  return grp_s_malaria_indication_27(app);
}

/** Relevance of note_danger_signs_neg */
export function rel_note_danger_signs_neg(app) {
  return grp_s_malaria_indication_27(app);
}

/** Relevance of s_malaria_tdr_result_new */
export function rel_s_malaria_tdr_result_new(app) {
  return grp_s_malaria_indication_27(app);
}

/** Relevance of CAT_mrdt */
export function rel_CAT_mrdt(app) {
  return grp_s_malaria_indication_27(app) && X.bool(X.eq(X.v(app, "s_malaria_tdr_result_new"), "neg"));
}

/** Relevance of interpretation_malaria_mrdt */
export function rel_interpretation_malaria_mrdt(app) {
  return grp_s_malaria_indication_27(app);
}

/** Calculation of interpretation_malaria_mrdt */
export function calc_interpretation_malaria_mrdt(app) {
  if (!rel_interpretation_malaria_mrdt(app)) return '';
  return X.out((X.bool(X.eq(X.v(app, "s_malaria_tdr_result_new"), "pos")) ? "Positif (+)" : "Négatif (-)"));
}

/** Calculation of medication_1 */
export function calc_medication_1(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_malaria_give_act"), "yes")) ? X.f["concat"](app, (X.bool(X.eq(X.inp(app, "user/language"), "bm")) ? "Artésunate + Lumefantrine (ALU) 20/120 mg" : "Artésunate + Lumefantrine (ALU) 20/120 mg"), " ", X.choiceName("c_malaria_act_dosage", X.v(app, "s_malaria_act_dosage")), " ") : ""));
}

/** Calculation of medication_2 */
export function calc_medication_2(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_give_paracetamol"), "yes")) ? X.f["concat"](app, (X.bool(X.eq(X.inp(app, "user/language"), "bm")) ? "Parasetamɔli" : "Paracétamol"), " ", X.v(app, "paracetamol_dosage")) : ""));
}

/** Calculation of paracetamol_dosage */
export function calc_paracetamol_dosage(app) {
  return X.out((X.bool(X.v(app, "s_give_paracetamol_dosage")) ? X.choiceName("c_give_paracetamol_dosage", X.v(app, "s_give_paracetamol_dosage")) : ""));
}

/** Calculation of medications */
export function calc_medications(app) {
  return X.out(X.f["concat"](app, X.v(app, "medication_1"), X.f["concat"](app, (X.bool(X.ne(X.v(app, "medication_1"), "")) ? ", " : ""), X.v(app, "medication_2"))));
}

/** Relevance of r_symptoms */
export function rel_r_symptoms(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "referral"), "true")) || X.bool(X.eq(X.v(app, "treat_malaria"), "true")))) || X.bool(X.eq(X.v(app, "CAT"), "watching"))));
}

/** Relevance of r_amenorrhoea_concept */
export function rel_r_amenorrhoea_concept(app) {
  return X.bool(X.eq(X.v(app, "amenorrhea_concept"), "true"));
}

/** Relevance of r_amenorrhoea_concept_1 */
export function rel_r_amenorrhoea_concept_1(app) {
  return X.bool(X.eq(X.v(app, "amenorrhea_concept"), "true"));
}

/** Relevance of pregnancy_gard */
export function rel_pregnancy_gard(app) {
  return X.bool(X.eq(X.v(app, "amenorrhea_concept"), "true"));
}

/** Constraint of pregnancy_gard */
export function val_pregnancy_gard(app) {
  if (X.str(X.v(app, "pregnancy_gard")) === '') return true;
  return X.bool((X.bool(X.eq(X.v(app, "pregnancy_gard"), "")) && X.bool(X.ne(X.v(app, "pregnancy_gard"), ""))));
}

/** Relevance of r_signs_tb */
export function rel_r_signs_tb(app) {
  return X.bool(X.eq(X.v(app, "has_tb_sign"), "true"));
}

/** Relevance of r_signs_tb_list */
export function rel_r_signs_tb_list(app) {
  return X.bool(X.eq(X.v(app, "has_tb_sign"), "true"));
}

/** Relevance of r_danger_signs */
export function rel_r_danger_signs(app) {
  return X.bool(X.eq(X.v(app, "has_mal_danger_sign"), "true"));
}

/** Relevance of r_malaria_fever_within24_48 */
export function rel_r_malaria_fever_within24_48(app) {
  return X.bool((X.bool(X.eq(X.v(app, "has_mal_danger_sign"), "true")) && X.bool(X.eq(X.v(app, "s_fever"), "yes"))));
}

/** Relevance of r_accom_ref_signs */
export function rel_r_accom_ref_signs(app) {
  return X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "true")) || X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "true"))));
}

/** Relevance of n_acc_danger_sign_seizure */
export function rel_n_acc_danger_sign_seizure(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_seizure"), "yes"));
}

/** Relevance of n_acc_danger_sign_loss_consiousness */
export function rel_n_acc_danger_sign_loss_consiousness(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_loss_consiousness"), "yes"));
}

/** Relevance of n_acc_danger_sign_unable_drink */
export function rel_n_acc_danger_sign_unable_drink(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_unable_drink"), "yes"));
}

/** Relevance of n_acc_danger_sign_confusion */
export function rel_n_acc_danger_sign_confusion(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_confusion"), "yes"));
}

/** Relevance of n_acc_danger_sign_vomit */
export function rel_n_acc_danger_sign_vomit(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_vomit"), "yes"));
}

/** Relevance of n_acc_danger_sign_chest_indrawing */
export function rel_n_acc_danger_sign_chest_indrawing(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_chest_indrawing"), "yes"));
}

/** Relevance of n_acc_danger_sign_wheezing */
export function rel_n_acc_danger_sign_wheezing(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_wheezing"), "yes"));
}

/** Relevance of n_acc_danger_sign_chest_pain */
export function rel_n_acc_danger_sign_chest_pain(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_chest_pain"), "yes"));
}

/** Relevance of n_acc_danger_sign_hemoptysis */
export function rel_n_acc_danger_sign_hemoptysis(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_hemoptysis"), "yes"));
}

/** Relevance of n_acc_danger_sign_lathargy */
export function rel_n_acc_danger_sign_lathargy(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_lathargy"), "yes"));
}

/** Relevance of n_ref_danger_sign_severe_jaundice */
export function rel_n_ref_danger_sign_severe_jaundice(app) {
  return X.bool(X.eq(X.v(app, "s_ref_danger_sign_severe_jaundice"), "yes"));
}

/** Relevance of n_ref_danger_sign_dark_colored_urine */
export function rel_n_ref_danger_sign_dark_colored_urine(app) {
  return X.bool(X.eq(X.v(app, "s_ref_danger_sign_dark_colored_urine"), "yes"));
}

/** Relevance of n_ref_danger_sign_conjunctival_pallor */
export function rel_n_ref_danger_sign_conjunctival_pallor(app) {
  return X.bool(X.eq(X.v(app, "s_ref_danger_sign_conjunctival_pallor"), "yes"));
}

/** Relevance of n_ref_danger_sign_eodema_lower_legs */
export function rel_n_ref_danger_sign_eodema_lower_legs(app) {
  return X.bool(X.eq(X.v(app, "s_ref_danger_sign_eodema_lower_legs"), "yes"));
}

/** Relevance of n_ref_danger_sign_persistent_fever */
export function rel_n_ref_danger_sign_persistent_fever(app) {
  return X.bool(X.eq(X.v(app, "s_ref_danger_sign_persistent_fever"), "yes"));
}

/** Relevance of n_ref_danger_sign_weight_loss */
export function rel_n_ref_danger_sign_weight_loss(app) {
  return X.bool(X.eq(X.v(app, "s_ref_danger_sign_weight_loss"), "yes"));
}

/** Relevance of n_ref_danger_sign_other_diseases */
export function rel_n_ref_danger_sign_other_diseases(app) {
  return X.bool(X.eq(X.v(app, "s_ref_danger_sign_other_diseases"), "yes"));
}

/** Relevance of n_pregnancy_danger_signs */
export function rel_n_pregnancy_danger_signs(app) {
  return X.bool(X.eq(X.v(app, "has_pregnancy_danger_sign"), "true"));
}

/** Relevance of n_vaginal_bleeding */
export function rel_n_vaginal_bleeding(app) {
  return X.bool(X.eq(X.v(app, "vaginal_bleeding"), "yes"));
}

/** Relevance of n_pelvic_pain */
export function rel_n_pelvic_pain(app) {
  return X.bool(X.eq(X.v(app, "pelvic_pain"), "yes"));
}

/** Relevance of n_headache_swelling */
export function rel_n_headache_swelling(app) {
  return X.bool(X.eq(X.v(app, "headache_swelling"), "yes"));
}

/** Relevance of n_fetal_movement_loss */
export function rel_n_fetal_movement_loss(app) {
  return X.bool(X.eq(X.v(app, "fetal_movement_loss"), "yes"));
}

/** Relevance of n_severe_vomiting */
export function rel_n_severe_vomiting(app) {
  return X.bool(X.eq(X.v(app, "severe_vomiting"), "yes"));
}

/** Relevance of n_uterine_contractions */
export function rel_n_uterine_contractions(app) {
  return X.bool(X.eq(X.v(app, "uterine_contractions"), "yes"));
}

/** Relevance of n_water_breaking */
export function rel_n_water_breaking(app) {
  return X.bool(X.eq(X.v(app, "water_breaking"), "yes"));
}

/** Relevance of n_urination_burning */
export function rel_n_urination_burning(app) {
  return X.bool(X.eq(X.v(app, "urination_burning"), "yes"));
}

/** Relevance of n_vaginal_discharge_odor */
export function rel_n_vaginal_discharge_odor(app) {
  return X.bool(X.eq(X.v(app, "vaginal_discharge_odor"), "yes"));
}

/** Relevance of n_postpartum_depression */
export function rel_n_postpartum_depression(app) {
  return X.bool(X.eq(X.v(app, "postpartum_depression"), "yes"));
}

/** Relevance of n_domestic_violence_signs */
export function rel_n_domestic_violence_signs(app) {
  return X.bool(X.eq(X.v(app, "domestic_violence_signs"), "yes"));
}

/** Relevance of n_postpartum_danger_signs */
export function rel_n_postpartum_danger_signs(app) {
  return X.bool(X.eq(X.v(app, "has_postpartum_danger_sign"), "true"));
}

/** Relevance of n_heavy_bleeding */
export function rel_n_heavy_bleeding(app) {
  return X.bool(X.eq(X.v(app, "heavy_bleeding"), "yes"));
}

/** Relevance of n_headache_face_swelling */
export function rel_n_headache_face_swelling(app) {
  return X.bool(X.eq(X.v(app, "headache_face_swelling"), "yes"));
}

/** Relevance of n_excessive_fatigue */
export function rel_n_excessive_fatigue(app) {
  return X.bool(X.eq(X.v(app, "excessive_fatigue"), "yes"));
}

/** Relevance of n_marked_paleness */
export function rel_n_marked_paleness(app) {
  return X.bool(X.eq(X.v(app, "marked_paleness"), "yes"));
}

/** Relevance of n_mobility_difficulty */
export function rel_n_mobility_difficulty(app) {
  return X.bool(X.eq(X.v(app, "mobility_difficulty"), "yes"));
}

/** Relevance of n_severe_abdominal_pain */
export function rel_n_severe_abdominal_pain(app) {
  return X.bool(X.eq(X.v(app, "severe_abdominal_pain"), "yes"));
}

/** Relevance of n_urination_burning_difficulty */
export function rel_n_urination_burning_difficulty(app) {
  return X.bool(X.eq(X.v(app, "urination_burning_difficulty"), "yes"));
}

/** Relevance of n_foul_vaginal_discharge */
export function rel_n_foul_vaginal_discharge(app) {
  return X.bool(X.eq(X.v(app, "foul_vaginal_discharge"), "yes"));
}

/** Relevance of n_breast_pain_redness */
export function rel_n_breast_pain_redness(app) {
  return X.bool(X.eq(X.v(app, "breast_pain_redness"), "yes"));
}

/** Relevance of n_postpartum_depression_signs */
export function rel_n_postpartum_depression_signs(app) {
  return X.bool(X.eq(X.v(app, "postpartum_depression_signs"), "yes"));
}

/** Relevance of r_diagnosis */
export function rel_r_diagnosis(app) {
  return X.bool(X.eq(X.v(app, "amenorrhea_concept"), "false"));
}

/** Relevance of r_signs_malaria */
export function rel_r_signs_malaria(app) {
  return X.bool(X.eq(X.v(app, "tdr_result"), "positif"));
}

/** Relevance of r_no_tdr_result */
export function rel_r_no_tdr_result(app) {
  return X.bool((X.bool(X.eq(X.v(app, "tdr_done"), "no")) && X.bool(X.ne(X.v(app, "no_tdr_done"), ""))));
}

/** Relevance of r_referral */
export function rel_r_referral(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false")))) && X.bool(X.eq(X.v(app, "tb_referral"), "false"))));
}

/** Relevance of r_referral_tb */
export function rel_r_referral_tb(app) {
  return X.bool(X.eq(X.v(app, "has_tb_sign"), "true"));
}

/** Relevance of r_accompany */
export function rel_r_accompany(app) {
  return X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true"));
}

/** Relevance of n_tdr_refer_cscom */
export function rel_n_tdr_refer_cscom(app) {
  return X.bool(X.eq(X.v(app, "s_malaria_tdr_not_done"), "tdr_unavailable"));
}

/** Relevance of r_special_instructions */
export function rel_r_special_instructions(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true")) || X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")))) || X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true"))));
}

/** Relevance of refer_to_cscom_tb */
export function rel_refer_to_cscom_tb(app) {
  return X.bool(X.eq(X.v(app, "has_tb_sign"), "true"));
}

/** Relevance of s_dtn_danger_signs_1 */
export function rel_s_dtn_danger_signs_1(app) {
  return X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true"));
}

/** Relevance of s_dtn_danger_signs_2 */
export function rel_s_dtn_danger_signs_2(app) {
  return X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "true")) || X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "true")))) || X.bool(X.eq(X.v(app, "has_pregnancy_danger_sign"), "true")))) || X.bool(X.eq(X.v(app, "has_postpartum_danger_sign"), "true"))));
}

/** Relevance of s_pregnancy_danger_signs */
export function rel_s_pregnancy_danger_signs(app) {
  return X.bool(X.eq(X.v(app, "has_pregnancy_danger_sign"), "true"));
}

/** Relevance of s_postpartum_danger_signs */
export function rel_s_postpartum_danger_signs(app) {
  return X.bool(X.eq(X.v(app, "has_postpartum_danger_sign"), "true"));
}

/** Relevance of r_write_tb_on_sheet */
export function rel_r_write_tb_on_sheet(app) {
  return X.bool(X.eq(X.v(app, "has_tb_sign"), "true"));
}

/** Relevance of r_treatment */
export function rel_r_treatment(app) {
  return X.bool((X.bool(X.eq(X.v(app, "treat_malaria"), "true")) || X.bool(X.eq(X.v(app, "s_fever"), "yes"))));
}

/** Relevance of s_dnt_malaria */
export function rel_s_dnt_malaria(app) {
  return X.bool((X.bool(X.eq(X.v(app, "tdr_result"), "positif")) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of r_malaria_act */
export function rel_r_malaria_act(app) {
  return X.bool((X.bool(X.eq(X.v(app, "tdr_result"), "positif")) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of r_malaria_act_2 */
export function rel_r_malaria_act_2(app) {
  return X.bool((X.bool((X.bool(X.lt(X.v(app, "patient_age_in_years"), 7)) && X.bool(X.eq(X.v(app, "tdr_result"), "positif")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of r_malaria_act_3 */
export function rel_r_malaria_act_3(app) {
  return X.bool((X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_years"), 7)) && X.bool(X.lt(X.v(app, "patient_age_in_years"), 11)))) && X.bool(X.eq(X.v(app, "tdr_result"), "positif")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of r_malaria_act_4 */
export function rel_r_malaria_act_4(app) {
  return X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_years"), 11)) && X.bool(X.eq(X.v(app, "tdr_result"), "positif")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_malaria_give_act */
export function rel_s_malaria_give_act(app) {
  return X.bool((X.bool(X.eq(X.v(app, "tdr_result"), "positif")) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_malaria_act_dosage */
export function rel_s_malaria_act_dosage(app) {
  return X.bool(X.eq(X.v(app, "s_malaria_give_act"), "yes"));
}

/** Relevance of s_malaria_not_give_act */
export function rel_s_malaria_not_give_act(app) {
  return X.bool(X.eq(X.v(app, "s_malaria_give_act"), "no"));
}

/** Relevance of s_malaria_not_give_act_other */
export function rel_s_malaria_not_give_act_other(app) {
  return X.bool(X.eq(X.v(app, "s_malaria_not_give_act"), "other"));
}

/** Relevance of n_malaria_act_shortage */
export function rel_n_malaria_act_shortage(app) {
  return X.bool(X.eq(X.v(app, "s_malaria_not_give_act"), "out_of_stock"));
}

/** Relevance of s_dnt_fever */
export function rel_s_dnt_fever(app) {
  return X.bool(X.eq(X.v(app, "s_fever"), "yes"));
}

/** Relevance of r_fever_paracetamol */
export function rel_r_fever_paracetamol(app) {
  return X.bool(X.eq(X.v(app, "s_fever"), "yes"));
}

/** Relevance of r_fever_paracetamol_2 */
export function rel_r_fever_paracetamol_2(app) {
  return X.bool((X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_years"), 7)) && X.bool(X.lt(X.v(app, "patient_age_in_years"), 12)))) && X.bool(X.eq(X.v(app, "s_fever"), "yes")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")) || X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true"))))));
}

/** Relevance of r_fever_paracetamol_2_no_reference */
export function rel_r_fever_paracetamol_2_no_reference(app) {
  return X.bool((X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_years"), 7)) && X.bool(X.lt(X.v(app, "patient_age_in_years"), 12)))) && X.bool(X.eq(X.v(app, "s_fever"), "yes")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of r_fever_paracetamol_3 */
export function rel_r_fever_paracetamol_3(app) {
  return X.bool((X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_years"), 12)) && X.bool(X.lt(X.v(app, "patient_age_in_years"), 15)))) && X.bool(X.eq(X.v(app, "s_fever"), "yes")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")) || X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true"))))));
}

/** Relevance of r_fever_paracetamol_3_no_reference */
export function rel_r_fever_paracetamol_3_no_reference(app) {
  return X.bool((X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_years"), 12)) && X.bool(X.lt(X.v(app, "patient_age_in_years"), 15)))) && X.bool(X.eq(X.v(app, "s_fever"), "yes")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of r_fever_paracetamol_4 */
export function rel_r_fever_paracetamol_4(app) {
  return X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_years"), 15)) && X.bool(X.eq(X.v(app, "s_fever"), "yes")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")) || X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true"))))));
}

/** Relevance of r_fever_paracetamol_4_no_reference */
export function rel_r_fever_paracetamol_4_no_reference(app) {
  return X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_years"), 15)) && X.bool(X.eq(X.v(app, "s_fever"), "yes")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of r_fever_paracetamol_5 */
export function rel_r_fever_paracetamol_5(app) {
  return X.bool((X.bool((X.bool(X.lt(X.v(app, "patient_age_in_years"), 7)) && X.bool(X.eq(X.v(app, "s_fever"), "yes")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")) || X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true"))))));
}

/** Relevance of r_fever_paracetamol_5_no_reference */
export function rel_r_fever_paracetamol_5_no_reference(app) {
  return X.bool((X.bool((X.bool(X.lt(X.v(app, "patient_age_in_years"), 7)) && X.bool(X.eq(X.v(app, "s_fever"), "yes")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_give_paracetamol */
export function rel_s_give_paracetamol(app) {
  return X.bool(X.eq(X.v(app, "s_fever"), "yes"));
}

/** Calculation of r_paracetamol_once_filter */
export function calc_r_paracetamol_once_filter(app) {
  return X.out((X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")) || X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true")))) ? "once" : "many"));
}

/** Relevance of s_give_paracetamol_dosage */
export function rel_s_give_paracetamol_dosage(app) {
  return X.bool(X.eq(X.v(app, "s_give_paracetamol"), "yes"));
}

/** Options of s_give_paracetamol_dosage */
export function filt_s_give_paracetamol_dosage(app) {
  return X.bool(X.eq(X.opt(app, "filter1"), X.v(app, "r_paracetamol_once_filter")));
}

/** Relevance of s_not_give_paracetamol */
export function rel_s_not_give_paracetamol(app) {
  return X.bool(X.eq(X.v(app, "s_give_paracetamol"), "no"));
}

/** Relevance of s_not_give_paracetamol_other */
export function rel_s_not_give_paracetamol_other(app) {
  return X.bool(X.eq(X.v(app, "s_not_give_paracetamol"), "other"));
}

/** Relevance of s_on_observation */
export function rel_s_on_observation(app) {
  return X.bool((X.bool(X.eq(X.v(app, "CAT"), "watching")) || X.bool(X.eq(X.v(app, "CAT_pregnancy"), "watching"))));
}

/** Relevance of n_paracetamol_shortage */
export function rel_n_paracetamol_shortage(app) {
  return X.bool(X.eq(X.v(app, "s_not_give_paracetamol"), "out_of_stock"));
}

/** Relevance of CAT_pregnancy */
export function rel_CAT_pregnancy(app) {
  return X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "false")) && X.bool(X.eq(X.v(app, "s_fever_within_24_48"), "no")))) && X.bool(X.eq(X.v(app, "s_fever"), "no")))) && X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "false")))) && X.bool(X.eq(X.v(app, "has_pregnancy_danger_sign"), "false")))) && X.bool(X.eq(X.v(app, "has_postpartum_danger_sign"), "false")))) && X.bool(X.eq(X.v(app, "amenorrhea_concept"), "false"))));
}

/** Relevance of CAT_out_of_stock */
export function rel_CAT_out_of_stock(app) {
  return X.bool(X.eq(X.v(app, "show_drugs_cat"), "true"));
}

/** Relevance of s_accompany_refer_CSCOM */
export function rel_s_accompany_refer_CSCOM(app) {
  return X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "referral"), "true")) || X.bool(X.eq(X.v(app, "CAT_out_of_stock"), "refer")))) || X.bool(X.eq(X.v(app, "CAT"), "refer")))) || X.bool(X.eq(X.v(app, "CAT_mrdt"), "refer")))) || X.bool(X.eq(X.v(app, "CAT_pregnancy"), "refer")))) || X.bool(X.eq(X.v(app, "is_tdr_unavailable"), "true")))) || X.bool(X.eq(X.v(app, "has_pregnancy_danger_sign"), "true")))) || X.bool(X.eq(X.v(app, "has_postpartum_danger_sign"), "true"))));
}

/** Relevance of r_follow_up */
export function rel_r_follow_up(app) {
  return X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "referral"), "true")) || X.bool(X.eq(X.v(app, "treat_malaria"), "true")))) || X.bool(X.eq(X.v(app, "CAT"), "watching")))) || X.bool(X.eq(X.v(app, "CAT_pregnancy"), "watching"))));
}

/** Relevance of r_next_visit */
export function rel_r_next_visit(app) {
  return X.bool((X.bool(X.eq(X.v(app, "diff_assess_date_report_date"), 0)) && X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "referral"), "true")) || X.bool(X.eq(X.v(app, "treat_malaria"), "true")))) || X.bool(X.eq(X.v(app, "CAT"), "watching")))) || X.bool(X.eq(X.v(app, "CAT_pregnancy"), "watching"))))));
}

/** Relevance of r_next_visit_remainder */
export function rel_r_next_visit_remainder(app) {
  return X.bool((X.bool(X.ne(X.v(app, "diff_assess_date_report_date"), 0)) && X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "referral"), "true")) || X.bool(X.eq(X.v(app, "treat_malaria"), "true")))) || X.bool(X.eq(X.v(app, "CAT"), "watching")))) || X.bool(X.eq(X.v(app, "CAT_pregnancy"), "watching"))))));
}

/** Calculation of act_20120_disp */
export function calc_act_20120_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_malaria_give_act"), "yes")) ? (X.bool(X.eq(X.v(app, "s_malaria_act_dosage"), "4_tablets_act")) ? (X.num(4) * X.num(3)) : (X.bool(X.eq(X.v(app, "s_malaria_act_dosage"), "6_tablets_act")) ? (X.num(6) * X.num(3)) : (X.bool(X.eq(X.v(app, "s_malaria_act_dosage"), "8_tablets_act")) ? (X.num(8) * X.num(3)) : 0))) : 0));
}

/** Calculation of malaria_rdt_disp */
export function calc_malaria_rdt_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "tdr_done"), "yes")) ? 1 : 0));
}

/** Calculation of paracetamol_500_disp */
export function calc_paracetamol_500_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_give_paracetamol_dosage"), "two_tab_paracetamol_once")) ? 2 : (X.bool(X.eq(X.v(app, "s_give_paracetamol_dosage"), "one_tab_paracetamol_once")) ? 1 : (X.bool(X.eq(X.v(app, "s_give_paracetamol_dosage"), "three_quarter_tab_paracetamol_once")) ? 1 : (X.bool(X.eq(X.v(app, "s_give_paracetamol_dosage"), "demi_tab_paracetamol_once")) ? 1 : (X.bool(X.eq(X.v(app, "s_give_paracetamol_dosage"), "two_tab_paracetamol_3x_a_day")) ? (X.num(2) * X.num(3)) : (X.bool(X.eq(X.v(app, "s_give_paracetamol_dosage"), "one_tab_paracetamol_3x_a_day")) ? (X.num(1) * X.num(3)) : (X.bool(X.eq(X.v(app, "s_give_paracetamol_dosage"), "three_quarter_tab_paracetamol_3x_a_day")) ? (X.num((X.num(3) / X.num(4))) * X.num(3)) : (X.bool(X.eq(X.v(app, "s_give_paracetamol_dosage"), "demi_tab_paracetamol_3x_a_day")) ? (X.num((X.num(1) / X.num(2))) * X.num(3)) : 0)))))))));
}

/** Relevance of stock_monitoring_reported_today */
export function rel_stock_monitoring_reported_today(app) {
  return grp_prescription_summary(app);
}

/** Calculation of stock_monitoring_reported_today */
export function calc_stock_monitoring_reported_today(app) {
  if (!rel_stock_monitoring_reported_today(app)) return '';
  return X.out(X.v(app, "s_assess_today"));
}

/** Relevance of stock_monitoring_reported_date */
export function rel_stock_monitoring_reported_date(app) {
  return grp_prescription_summary(app);
}

/** Calculation of stock_monitoring_reported_date */
export function calc_stock_monitoring_reported_date(app) {
  if (!rel_stock_monitoring_reported_date(app)) return '';
  return X.out(X.v(app, "assessment_date"));
}

/** Relevance of place_id */
export function rel_place_id(app) {
  return grp_prescription_summary(app);
}

/** Calculation of place_id */
export function calc_place_id(app) {
  if (!rel_place_id(app)) return '';
  return X.out(X.v(app, "c40_chw_area_uuid"));
}

/** Relevance of type */
export function rel_type(app) {
  return grp_prescription_summary(app);
}

/** Calculation of type */
export function calc_type(app) {
  if (!rel_type(app)) return '';
  return X.out("data_record");
}

/** Relevance of created_from */
export function rel_created_from(app) {
  return grp_prescription_summary(app);
}

/** Calculation of created_from */
export function calc_created_from(app) {
  if (!rel_created_from(app)) return '';
  return X.out(X.v(app, "created_from"));
}

/** Relevance of created_from_name */
export function rel_created_from_name(app) {
  return grp_prescription_summary(app);
}

/** Calculation of created_from_name */
export function calc_created_from_name(app) {
  if (!rel_created_from_name(app)) return '';
  return X.out("patient_assessment_over_5");
}

/** Relevance of content_type */
export function rel_content_type(app) {
  return grp_prescription_summary(app);
}

/** Calculation of content_type */
export function calc_content_type(app) {
  if (!rel_content_type(app)) return '';
  return X.out("xml");
}

/** Relevance of form */
export function rel_form(app) {
  return grp_prescription_summary(app);
}

/** Calculation of form */
export function calc_form(app) {
  if (!rel_form(app)) return '';
  return X.out("prescription_summary");
}

/** Relevance of _id */
export function rel__id(app) {
  return grp_contact(app);
}

/** Calculation of _id */
export function calc__id(app) {
  if (!rel__id(app)) return '';
  return X.out(X.v(app, "c40_chw_area_uuid"));
}

/** Relevance of albendazol_200 */
export function rel_albendazol_200(app) {
  return grp_fields(app);
}

/** Calculation of albendazol_200 */
export function calc_albendazol_200(app) {
  if (!rel_albendazol_200(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "albendazol_200_disp")) - X.num(X.f["floor"](app, X.v(app, "albendazol_200_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "albendazol_200_disp"))) + X.num(1)) : X.v(app, "albendazol_200_disp")));
}

/** Relevance of albendazol_400 */
export function rel_albendazol_400(app) {
  return grp_fields(app);
}

/** Calculation of albendazol_400 */
export function calc_albendazol_400(app) {
  if (!rel_albendazol_400(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "albendazol_400_disp")) - X.num(X.f["floor"](app, X.v(app, "albendazol_400_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "albendazol_400_disp"))) + X.num(1)) : X.v(app, "albendazol_400_disp")));
}

/** Relevance of plumpy_sup */
export function rel_plumpy_sup(app) {
  return grp_fields(app);
}

/** Calculation of plumpy_sup */
export function calc_plumpy_sup(app) {
  if (!rel_plumpy_sup(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "plumpy_sup_disp")) - X.num(X.f["floor"](app, X.v(app, "plumpy_sup_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "plumpy_sup_disp"))) + X.num(1)) : X.v(app, "plumpy_sup_disp")));
}

/** Relevance of plumpy_nut */
export function rel_plumpy_nut(app) {
  return grp_fields(app);
}

/** Calculation of plumpy_nut */
export function calc_plumpy_nut(app) {
  if (!rel_plumpy_nut(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "plumpy_nut_disp")) - X.num(X.f["floor"](app, X.v(app, "plumpy_nut_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "plumpy_nut_disp"))) + X.num(1)) : X.v(app, "plumpy_nut_disp")));
}

/** Relevance of enriched_flour */
export function rel_enriched_flour(app) {
  return grp_fields(app);
}

/** Calculation of enriched_flour */
export function calc_enriched_flour(app) {
  if (!rel_enriched_flour(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "enriched_flour_disp")) - X.num(X.f["floor"](app, X.v(app, "enriched_flour_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "enriched_flour_disp"))) + X.num(1)) : X.v(app, "enriched_flour_disp")));
}

/** Relevance of iron_folic_acid_srp */
export function rel_iron_folic_acid_srp(app) {
  return grp_fields(app);
}

/** Calculation of iron_folic_acid_srp */
export function calc_iron_folic_acid_srp(app) {
  if (!rel_iron_folic_acid_srp(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "iron_folic_acid_srp_disp")) - X.num(X.f["floor"](app, X.v(app, "iron_folic_acid_srp_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "iron_folic_acid_srp_disp"))) + X.num(1)) : X.v(app, "iron_folic_acid_srp_disp")));
}

/** Relevance of iron_folic_acid */
export function rel_iron_folic_acid(app) {
  return grp_fields(app);
}

/** Calculation of iron_folic_acid */
export function calc_iron_folic_acid(app) {
  if (!rel_iron_folic_acid(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "iron_folic_acid_disp")) - X.num(X.f["floor"](app, X.v(app, "iron_folic_acid_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "iron_folic_acid_disp"))) + X.num(1)) : X.v(app, "iron_folic_acid_disp")));
}

/** Relevance of amoxicillin_250 */
export function rel_amoxicillin_250(app) {
  return grp_fields(app);
}

/** Calculation of amoxicillin_250 */
export function calc_amoxicillin_250(app) {
  if (!rel_amoxicillin_250(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "amoxicillin_250_disp")) - X.num(X.f["floor"](app, X.v(app, "amoxicillin_250_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "amoxicillin_250_disp"))) + X.num(1)) : X.v(app, "amoxicillin_250_disp")));
}

/** Relevance of vitamin_a_blue_capsule */
export function rel_vitamin_a_blue_capsule(app) {
  return grp_fields(app);
}

/** Calculation of vitamin_a_blue_capsule */
export function calc_vitamin_a_blue_capsule(app) {
  if (!rel_vitamin_a_blue_capsule(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "vitamin_a_blue_capsule_disp")) - X.num(X.f["floor"](app, X.v(app, "vitamin_a_blue_capsule_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "vitamin_a_blue_capsule_disp"))) + X.num(1)) : X.v(app, "vitamin_a_blue_capsule_disp")));
}

/** Relevance of vitamin_a_red_capsule */
export function rel_vitamin_a_red_capsule(app) {
  return grp_fields(app);
}

/** Calculation of vitamin_a_red_capsule */
export function calc_vitamin_a_red_capsule(app) {
  if (!rel_vitamin_a_red_capsule(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "vitamin_a_red_capsule_disp")) - X.num(X.f["floor"](app, X.v(app, "vitamin_a_red_capsule_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "vitamin_a_red_capsule_disp"))) + X.num(1)) : X.v(app, "vitamin_a_red_capsule_disp")));
}

/** Relevance of act_20120 */
export function rel_act_20120(app) {
  return grp_fields(app);
}

/** Calculation of act_20120 */
export function calc_act_20120(app) {
  if (!rel_act_20120(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "act_20120_disp")) - X.num(X.f["floor"](app, X.v(app, "act_20120_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "act_20120_disp"))) + X.num(1)) : X.v(app, "act_20120_disp")));
}

/** Relevance of malaria_rdt */
export function rel_malaria_rdt(app) {
  return grp_fields(app);
}

/** Calculation of malaria_rdt */
export function calc_malaria_rdt(app) {
  if (!rel_malaria_rdt(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "malaria_rdt_disp")) - X.num(X.f["floor"](app, X.v(app, "malaria_rdt_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "malaria_rdt_disp"))) + X.num(1)) : X.v(app, "malaria_rdt_disp")));
}

/** Relevance of artesunate_50 */
export function rel_artesunate_50(app) {
  return grp_fields(app);
}

/** Calculation of artesunate_50 */
export function calc_artesunate_50(app) {
  if (!rel_artesunate_50(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "artesunate_50_disp")) - X.num(X.f["floor"](app, X.v(app, "artesunate_50_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "artesunate_50_disp"))) + X.num(1)) : X.v(app, "artesunate_50_disp")));
}

/** Relevance of artesunate_100 */
export function rel_artesunate_100(app) {
  return grp_fields(app);
}

/** Calculation of artesunate_100 */
export function calc_artesunate_100(app) {
  if (!rel_artesunate_100(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "artesunate_100_disp")) - X.num(X.f["floor"](app, X.v(app, "artesunate_100_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "artesunate_100_disp"))) + X.num(1)) : X.v(app, "artesunate_100_disp")));
}

/** Relevance of artesunate_200 */
export function rel_artesunate_200(app) {
  return grp_fields(app);
}

/** Calculation of artesunate_200 */
export function calc_artesunate_200(app) {
  if (!rel_artesunate_200(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "artesunate_200_disp")) - X.num(X.f["floor"](app, X.v(app, "artesunate_200_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "artesunate_200_disp"))) + X.num(1)) : X.v(app, "artesunate_200_disp")));
}

/** Relevance of arthemeter_40 */
export function rel_arthemeter_40(app) {
  return grp_fields(app);
}

/** Calculation of arthemeter_40 */
export function calc_arthemeter_40(app) {
  if (!rel_arthemeter_40(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "arthemeter_40_disp")) - X.num(X.f["floor"](app, X.v(app, "arthemeter_40_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "arthemeter_40_disp"))) + X.num(1)) : X.v(app, "arthemeter_40_disp")));
}

/** Relevance of sp */
export function rel_sp(app) {
  return grp_fields(app);
}

/** Calculation of sp */
export function calc_sp(app) {
  if (!rel_sp(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "sp_disp")) - X.num(X.f["floor"](app, X.v(app, "sp_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "sp_disp"))) + X.num(1)) : X.v(app, "sp_disp")));
}

/** Relevance of zinc_20 */
export function rel_zinc_20(app) {
  return grp_fields(app);
}

/** Calculation of zinc_20 */
export function calc_zinc_20(app) {
  if (!rel_zinc_20(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "zinc_20_disp")) - X.num(X.f["floor"](app, X.v(app, "zinc_20_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "zinc_20_disp"))) + X.num(1)) : X.v(app, "zinc_20_disp")));
}

/** Relevance of ors */
export function rel_ors(app) {
  return grp_fields(app);
}

/** Calculation of ors */
export function calc_ors(app) {
  if (!rel_ors(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "ors_disp")) - X.num(X.f["floor"](app, X.v(app, "ors_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "ors_disp"))) + X.num(1)) : X.v(app, "ors_disp")));
}

/** Relevance of pregnancy_rdt */
export function rel_pregnancy_rdt(app) {
  return grp_fields(app);
}

/** Calculation of pregnancy_rdt */
export function calc_pregnancy_rdt(app) {
  if (!rel_pregnancy_rdt(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "pregnancy_rdt_disp")) - X.num(X.f["floor"](app, X.v(app, "pregnancy_rdt_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "pregnancy_rdt_disp"))) + X.num(1)) : X.v(app, "pregnancy_rdt_disp")));
}

/** Relevance of sayana_press */
export function rel_sayana_press(app) {
  return grp_fields(app);
}

/** Calculation of sayana_press */
export function calc_sayana_press(app) {
  if (!rel_sayana_press(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "sayana_press_disp")) - X.num(X.f["floor"](app, X.v(app, "sayana_press_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "sayana_press_disp"))) + X.num(1)) : X.v(app, "sayana_press_disp")));
}

/** Relevance of male_condom */
export function rel_male_condom(app) {
  return grp_fields(app);
}

/** Calculation of male_condom */
export function calc_male_condom(app) {
  if (!rel_male_condom(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "male_condom_disp")) - X.num(X.f["floor"](app, X.v(app, "male_condom_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "male_condom_disp"))) + X.num(1)) : X.v(app, "male_condom_disp")));
}

/** Relevance of female_condom */
export function rel_female_condom(app) {
  return grp_fields(app);
}

/** Calculation of female_condom */
export function calc_female_condom(app) {
  if (!rel_female_condom(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "female_condom_disp")) - X.num(X.f["floor"](app, X.v(app, "female_condom_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "female_condom_disp"))) + X.num(1)) : X.v(app, "female_condom_disp")));
}

/** Relevance of cycle_collar */
export function rel_cycle_collar(app) {
  return grp_fields(app);
}

/** Calculation of cycle_collar */
export function calc_cycle_collar(app) {
  if (!rel_cycle_collar(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "cycle_collar_disp")) - X.num(X.f["floor"](app, X.v(app, "cycle_collar_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "cycle_collar_disp"))) + X.num(1)) : X.v(app, "cycle_collar_disp")));
}

/** Relevance of depo_provera */
export function rel_depo_provera(app) {
  return grp_fields(app);
}

/** Calculation of depo_provera */
export function calc_depo_provera(app) {
  if (!rel_depo_provera(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "depo_provera_disp")) - X.num(X.f["floor"](app, X.v(app, "depo_provera_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "depo_provera_disp"))) + X.num(1)) : X.v(app, "depo_provera_disp")));
}

/** Relevance of postpill */
export function rel_postpill(app) {
  return grp_fields(app);
}

/** Calculation of postpill */
export function calc_postpill(app) {
  if (!rel_postpill(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "postpill_disp")) - X.num(X.f["floor"](app, X.v(app, "postpill_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "postpill_disp"))) + X.num(1)) : X.v(app, "postpill_disp")));
}

/** Relevance of simple_pill */
export function rel_simple_pill(app) {
  return grp_fields(app);
}

/** Calculation of simple_pill */
export function calc_simple_pill(app) {
  if (!rel_simple_pill(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "simple_pill_disp")) - X.num(X.f["floor"](app, X.v(app, "simple_pill_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "simple_pill_disp"))) + X.num(1)) : X.v(app, "simple_pill_disp")));
}

/** Relevance of low_dose_pill */
export function rel_low_dose_pill(app) {
  return grp_fields(app);
}

/** Calculation of low_dose_pill */
export function calc_low_dose_pill(app) {
  if (!rel_low_dose_pill(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "low_dose_pill_disp")) - X.num(X.f["floor"](app, X.v(app, "low_dose_pill_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "low_dose_pill_disp"))) + X.num(1)) : X.v(app, "low_dose_pill_disp")));
}

/** Relevance of paracetamol_500 */
export function rel_paracetamol_500(app) {
  return grp_fields(app);
}

/** Calculation of paracetamol_500 */
export function calc_paracetamol_500(app) {
  if (!rel_paracetamol_500(app)) return '';
  return X.out((X.bool(X.gt((X.num(X.v(app, "paracetamol_500_disp")) - X.num(X.f["floor"](app, X.v(app, "paracetamol_500_disp")))), 0)) ? (X.num(X.f["floor"](app, X.v(app, "paracetamol_500_disp"))) + X.num(1)) : X.v(app, "paracetamol_500_disp")));
}

/** Task input source_id */
export function calc_inputs_source_id(app) {
  return X.out(X.inp(app, "source_id"));
}
