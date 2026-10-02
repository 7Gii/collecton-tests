// Form moderate_malnutrition_followup: ported from forms/app/moderate_malnutrition_followup.xlsx of the CHT configuration muso-mali.
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
 "c_followup_task_chw": {
  "c_followup_task_chw_1": "Suivi 1",
  "c_followup_task_chw_2": "Suivi 2",
  "c_followup_task_chw_3": "Suivi 3",
  "c_followup_task_chw_4": "Suivi 4",
  "c_followup_task_chw_5": "Suivi 5",
  "c_followup_task_chw_6": "Suivi 6",
  "c_followup_task_chw_7": "Suivi 7",
  "c_followup_task_chw_8": "Suivi 8"
 },
 "c_child_presence_confirmation": {
  "yes": "Oui",
  "no": "Non",
  "deceased": "Décédé"
 },
 "c_shakir_strip_color": {
  "green": "Vert",
  "yellow": "Jaune",
  "red": "Rouge"
 },
 "normal_fast": {
  "normal": "Normale",
  "fast": "Rapide"
 },
 "c_dewormer_type": {
  "200mg": "200 mg",
  "400mg": "400 mg"
 },
 "c_dewormer_dosage": {
  "half_tab": "1/2 comprimé",
  "one_tab": "1 comprimé",
  "two_tab": "2 comprimés"
 },
 "c_iron_folic_dosage": {
  "half_tab": "1 comprimé de fer + acide folique comprimé 200 mg",
  "one_tab": "1/2 comprimé de fer + acide folique comprimé 200 mg",
  "2_5ml": "2,5 ml de Fer + acide folique sirop",
  "5ml": "5 ml de Fer + acide folique sirop"
 },
 "c_vitaminA_dosage": {
  "blue_cap": "Capsule bleue",
  "red_cap": "Capsule rouge"
 },
 "medication_not_given": {
  "out_of_stock": "Rupture de stock",
  "other": "Autre"
 },
 "c_why_close_out": {
  "deceased": "Décès",
  "referred": "Référé au CSCOM",
  "abandoned": "Abandon"
 },
 "CAT": {
  "continue_followup": "Poursuivre le suivi",
  "recovered": "Guéri",
  "abandoned": "Abandon",
  "refer_with_followup": "Référer au CSCOM sans arrêt de suivi",
  "refer_with_no_followup": "Référer au CSCOM avec arrêt de Suivi"
 },
 "plumpy_not_given": {
  "out_of_stock": "Rupture de stock",
  "child_rejected_plumpy": "Enfant ne veut pas du plumpy",
  "other": "Autre"
 },
 "enriched_flour_not_given": {
  "out_of_stock": "Rupture de stock",
  "child_rejected_enriched_flour": "L'enfant ne veut pas la farine enrichie",
  "other": "Autre"
 },
 "iron_folic_type": {
  "iron_folic_acid_srp": "fer + acide folique sirop",
  "iron_folic_acid": "fer + acide folique comprimé 200 mg"
 }
};

/** CHT properties.json: not offered in the add menu, opened by its task only. */
export function showForm(app) {
  return false;
}

/** Group s_assessment */
export function grp_s_assessment(app) {
  return X.bool(X.eq(X.v(app, "s_child_presence_confirmation"), "yes"));
}

/** Group s_treatment_given */
export function grp_s_treatment_given(app) {
  return X.bool(X.eq(X.v(app, "s_child_presence_confirmation"), "yes"));
}

/** Group medical_complications */
export function grp_medical_complications(app) {
  return X.bool(X.eq(X.v(app, "s_child_presence_confirmation"), "yes"));
}

/** Group s_respiration */
export function grp_s_respiration(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_child_presence_confirmation"), "yes")) && X.bool(X.eq(X.v(app, "s_malnutrition_cough"), "yes"))));
}

/** Group s_weight */
export function grp_s_weight(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "s_child_presence_confirmation"), "yes")) && X.bool(X.ne(X.v(app, "s_shakir_strip_color"), "red")))) && X.bool(X.ne(X.v(app, "s_malnutrition_edema"), "yes"))));
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

/** Calculation of patient_uuid */
export function calc_patient_uuid(app) {
  return X.out(X.inp(app, "contact/_id"));
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

/** Calculation of next_visit */
export function calc_next_visit(app) {
  return X.out((X.bool(X.lt(X.f["number"](app, X.v(app, "days_since_admission")), X.f["number"](app, "28"))) ? "Follow up in 1 week" : (X.bool(X.lt(X.f["number"](app, X.v(app, "days_since_admission")), X.f["number"](app, "84"))) ? "Follow up in 2 weeks" : "")));
}

/** Calculation of next_visit_fr */
export function calc_next_visit_fr(app) {
  return X.out((X.bool(X.lt(X.f["number"](app, X.v(app, "days_since_admission")), X.f["number"](app, "28"))) ? "Le prochain suivi sera dans 7 jours" : (X.bool(X.lt(X.f["number"](app, X.v(app, "days_since_admission")), X.f["number"](app, "84"))) ? "Le prochain suivi sera dans 14 jours" : "")));
}

/** Calculation of next_visit_date */
export function calc_next_visit_date(app) {
  return X.out((X.bool(X.lt(X.f["number"](app, X.v(app, "days_since_admission")), X.f["number"](app, "28"))) ? X.f["format-date-time"](app, X.f["decimal-date-time"](app, (X.num(X.f["now"](app)) + X.num(7))), "%d/%m/%Y") : X.f["format-date-time"](app, X.f["decimal-date-time"](app, (X.num(X.f["now"](app)) + X.num(14))), "%d/%m/%Y")));
}

/** Calculation of follow_up_count */
export function calc_follow_up_count(app) {
  return X.out(X.inp(app, "t_follow_up_count"));
}

/** Calculation of follow_up_count_number */
export function calc_follow_up_count_number(app) {
  return X.out(X.f["number"](app, X.inp(app, "t_follow_up_count")));
}

/** Calculation of next_followup_visit */
export function calc_next_followup_visit(app) {
  return X.out((X.bool(X.eq(X.v(app, "c_selected_task_number"), "1")) ? X.f["format-date"](app, X.f["date"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) + X.num(7))), "%e/%n/%Y") : (X.bool(X.eq(X.v(app, "c_selected_task_number"), "2")) ? X.f["format-date"](app, X.f["date"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) + X.num(14))), "%e/%n/%Y") : (X.bool(X.eq(X.v(app, "c_selected_task_number"), "3")) ? X.f["format-date"](app, X.f["date"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) + X.num(21))), "%e/%n/%Y") : (X.bool(X.eq(X.v(app, "c_selected_task_number"), "4")) ? X.f["format-date"](app, X.f["date"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) + X.num(28))), "%e/%n/%Y") : (X.bool(X.eq(X.v(app, "c_selected_task_number"), "5")) ? X.f["format-date"](app, X.f["date"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) + X.num(42))), "%e/%n/%Y") : (X.bool(X.eq(X.v(app, "c_selected_task_number"), "6")) ? X.f["format-date"](app, X.f["date"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) + X.num(56))), "%e/%n/%Y") : (X.bool(X.eq(X.v(app, "c_selected_task_number"), "7")) ? X.f["format-date"](app, X.f["date"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) + X.num(70))), "%e/%n/%Y") : (X.bool(X.eq(X.v(app, "c_selected_task_number"), "8")) ? X.f["format-date"](app, X.f["date"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) + X.num(84))), "%e/%n/%Y") : "")))))))));
}

/** Calculation of follow_up_count_label_en */
export function calc_follow_up_count_label_en(app) {
  return X.out(X.f["concat"](app, "follow up ", X.v(app, "follow_up_count_number")));
}

/** Calculation of follow_up_count_label_fr */
export function calc_follow_up_count_label_fr(app) {
  return X.out(X.f["concat"](app, "suivi ", X.v(app, "follow_up_count_number")));
}

/** Calculation of fast_breathing */
export function calc_fast_breathing(app) {
  return X.out((X.bool((X.bool((X.bool(X.ge(X.f["coalesce"](app, X.v(app, "s_ari_respiration_rate"), 0), 50)) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)))) && X.bool(X.lt(X.v(app, "patient_age_in_years"), 1)))) ? "true" : (X.bool((X.bool(X.ge(X.f["coalesce"](app, X.v(app, "s_ari_respiration_rate"), 0), 40)) && X.bool(X.ge(X.v(app, "patient_age_in_years"), 1)))) ? "true" : "false")));
}

/** Calculation of fast_breathing_label_fr */
export function calc_fast_breathing_label_fr(app) {
  return X.out((X.bool(X.eq(X.v(app, "fast_breathing"), "true")) ? "Rapide" : "Normale"));
}

/** Calculation of fast_breathing_label_bm */
export function calc_fast_breathing_label_bm(app) {
  return X.out((X.bool(X.eq(X.v(app, "fast_breathing"), "true")) ? "A ka teli" : "A man teli(A kaɲi)"));
}

/** Calculation of fast_breathing_label_en */
export function calc_fast_breathing_label_en(app) {
  return X.out((X.bool(X.eq(X.v(app, "fast_breathing"), "true")) ? "Fast" : "Normal"));
}

/** Calculation of num_yellow_strips */
export function calc_num_yellow_strips(app) {
  return X.out(X.inp(app, "t_num_successive_yellow_strips"));
}

/** Calculation of num_green_strips */
export function calc_num_green_strips(app) {
  return X.out(X.inp(app, "t_num_successive_green_strips"));
}

/** Calculation of num_red_strips */
export function calc_num_red_strips(app) {
  return X.out(X.inp(app, "t_num_successive_red_strips"));
}

/** Calculation of before_previous_child_presence */
export function calc_before_previous_child_presence(app) {
  return X.out(X.inp(app, "t_before_previous_child_presence"));
}

/** Calculation of previous_strip */
export function calc_previous_strip(app) {
  return X.out(X.inp(app, "t_previous_strip"));
}

/** Calculation of num_recovered */
export function calc_num_recovered(app) {
  return X.out(X.inp(app, "t_num_recovered"));
}

/** Calculation of child_presence */
export function calc_child_presence(app) {
  return X.out(X.v(app, "s_child_presence_confirmation"));
}

/** Calculation of num_child_presence_confirmation */
export function calc_num_child_presence_confirmation(app) {
  return X.out(X.inp(app, "t_num_child_presence_confirmation"));
}

/** Calculation of shakir_strip_color */
export function calc_shakir_strip_color(app) {
  return X.out(X.v(app, "s_shakir_strip_color"));
}

/** Calculation of has_edema */
export function calc_has_edema(app) {
  return X.out(X.v(app, "s_malnutrition_edema"));
}

/** Calculation of num_no_edema */
export function calc_num_no_edema(app) {
  return X.out(X.inp(app, "t_num_no_edema"));
}

/** Calculation of referral_with_followup */
export function calc_referral_with_followup(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_other_signs"), "yes")) || X.bool(X.eq(X.v(app, "s_diarrhea_stools"), "yes")))) || X.bool(X.eq(X.v(app, "s_weight_loss"), "yes")))) || X.bool(X.eq(X.v(app, "fast_breathing"), "true")))) && X.bool(X.eq(X.v(app, "is_ending_followup"), "false")))) && X.bool(X.eq(X.v(app, "patient_deceased"), "false")))) && X.bool(X.eq(X.v(app, "referral_with_no_followup"), "false")))) ? "true" : "false"));
}

/** Calculation of referral_with_no_followup */
export function calc_referral_with_no_followup(app) {
  return X.out((X.bool(X.lt(X.v(app, "c_selected_task_number"), 8)) ? X.v(app, "referral_with_no_followup_1") : (X.bool((X.bool((X.bool(X.eq(X.v(app, "referral_with_no_followup_1"), "true")) || X.bool(X.eq(X.v(app, "referral_with_no_followup_2"), "true")))) || X.bool(X.eq(X.v(app, "patient_recovered"), "false")))) ? "true" : "false")));
}

/** Calculation of abandon_followup */
export function calc_abandon_followup(app) {
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "num_child_presence_confirmation"), "2")) && X.bool(X.f["selected"](app, X.v(app, "s_child_presence_confirmation"), "no")))) && X.bool(X.eq(X.v(app, "patient_deceased"), "false")))) ? "true" : "false"));
}

/** Calculation of patient_recovered */
export function calc_patient_recovered(app) {
  return X.out((X.bool(X.ge(X.v(app, "c_selected_task_number"), 3)) ? (X.bool(X.lt(X.v(app, "c_selected_task_number"), 8)) ? X.v(app, "patient_recovered_1") : (X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_recovered_1"), "true")) || X.bool(X.eq(X.v(app, "patient_recovered_2"), "true")))) || X.bool(X.eq(X.v(app, "patient_recovered_3"), "true")))) || X.bool(X.eq(X.v(app, "patient_recovered_4"), "true")))) || X.bool(X.eq(X.v(app, "patient_recovered_5"), "true")))) || X.bool(X.eq(X.v(app, "patient_recovered_6"), "true")))) || X.bool(X.eq(X.v(app, "patient_recovered_7"), "true")))) ? "true" : "false")) : "false"));
}

/** Calculation of patient_deceased */
export function calc_patient_deceased(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_child_presence_confirmation"), "deceased")) ? "true" : "false"));
}

/** Calculation of patient_should_continue_followup */
export function calc_patient_should_continue_followup(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "referral_with_followup"), "false")) && X.bool(X.eq(X.v(app, "referral_with_no_followup"), "false")))) && X.bool(X.eq(X.v(app, "abandon_followup"), "false")))) && X.bool(X.eq(X.v(app, "patient_recovered"), "false")))) && X.bool(X.eq(X.v(app, "patient_deceased"), "false")))) && X.bool(X.eq(X.v(app, "is_ending_followup"), "false")))) && X.bool(X.eq(X.f["coalesce"](app, X.v(app, "s_malnutrition_edema"), "no"), "no")))) && X.bool(X.eq(X.f["coalesce"](app, X.v(app, "s_other_signs"), "no"), "no")))) ? "true" : "false"));
}

/** Calculation of referral */
export function calc_referral(app) {
  return X.out((X.bool((X.bool(X.eq(X.v(app, "cat"), "refer_with_followup")) || X.bool(X.eq(X.v(app, "cat"), "refer_with_no_followup")))) ? "true" : "false"));
}

/** Calculation of final_cat */
export function calc_final_cat(app) {
  return X.out(X.v(app, "cat"));
}

/** Calculation of assessment_date */
export function calc_assessment_date(app) {
  return X.out(X.inp(app, "t_assessment_date"));
}

/** Calculation of days_since_admission */
export function calc_days_since_admission(app) {
  return X.out(X.f["floor"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) - X.num(X.f["decimal-date-time"](app, X.inp(app, "t_assessment_date"))))));
}

/** Calculation of dewormer_dosage_label */
export function calc_dewormer_dosage_label(app) {
  return X.out(X.choiceName("c_dewormer_dosage", X.v(app, "s_given_dewormer_dosage")));
}

/** Calculation of vitaminA_dosage_label */
export function calc_vitaminA_dosage_label(app) {
  return X.out(X.choiceName("c_vitaminA_dosage", X.v(app, "s_vitaminA_dosage")));
}

/** Calculation of iron_folic_dosage_label */
export function calc_iron_folic_dosage_label(app) {
  return X.out(X.choiceName("c_iron_folic_dosage", X.v(app, "s_iron_folic_dosage")));
}

/** Calculation of s_give_vitaminea_dosage_value_en */
export function calc_s_give_vitaminea_dosage_value_en(app) {
  return X.out((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 6)) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 12)))) ? "blue capsule, 100 000 UI" : (X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 12)) && X.bool(X.le(X.v(app, "patient_age_in_months"), 59)))) ? "red capsule, 200 000 UI" : "")));
}

/** Calculation of s_give_vitaminea_dosage_value_fr */
export function calc_s_give_vitaminea_dosage_value_fr(app) {
  return X.out((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 6)) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 12)))) ? "capsule bleue, 100 000 UI" : (X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 12)) && X.bool(X.le(X.v(app, "patient_age_in_months"), 59)))) ? "capsule rouge, 200 000 UI" : "")));
}

/** Calculation of s_give_vitaminea_dosage_value_bm */
export function calc_s_give_vitaminea_dosage_value_bm(app) {
  return X.out((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 6)) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 12)))) ? "furakisɛ bulama 1, 100 000 UI" : (X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 12)) && X.bool(X.le(X.v(app, "patient_age_in_months"), 59)))) ? "furakisɛ bileman 1, 200 000 UI" : "")));
}

/** Calculation of chw_selected_computed_cat */
export function calc_chw_selected_computed_cat(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "cat"), "continue_followup")) && X.bool(X.eq(X.v(app, "patient_should_continue_followup"), "true")))) || X.bool((X.bool(X.eq(X.v(app, "cat"), "abandoned")) && X.bool(X.eq(X.v(app, "abandon_followup"), "true")))))) || X.bool((X.bool(X.eq(X.v(app, "cat"), "recovered")) && X.bool(X.eq(X.v(app, "patient_recovered"), "true")))))) || X.bool((X.bool(X.eq(X.v(app, "cat"), "refer_with_followup")) && X.bool(X.eq(X.v(app, "referral_with_followup"), "true")))))) || X.bool((X.bool(X.eq(X.v(app, "cat"), "refer_with_no_followup")) && X.bool(X.eq(X.v(app, "referral_with_no_followup"), "true")))))) || X.bool((X.bool((X.bool((X.bool((X.bool(X.ne(X.v(app, "cat"), "continue_followup")) && X.bool(X.ne(X.v(app, "cat"), "abandoned")))) && X.bool(X.ne(X.v(app, "cat"), "recovered")))) && X.bool(X.ne(X.v(app, "cat"), "refer_with_followup")))) && X.bool(X.ne(X.v(app, "cat"), "refer_with_no_followup")))))) ? "true" : "false"));
}

/** Calculation of show_continue_value */
export function calc_show_continue_value(app) {
  return X.out((X.bool(X.eq(X.v(app, "is_ending_followup"), "false")) ? "continue_on" : "off"));
}

/** Calculation of show_recovered_value */
export function calc_show_recovered_value(app) {
  return X.out((X.bool(X.gt(X.v(app, "c_selected_task_number"), 2)) ? "recovered_on" : "off"));
}

/** Calculation of show_refer_value */
export function calc_show_refer_value(app) {
  return X.out((X.bool(X.eq(X.v(app, "is_ending_followup"), "false")) ? "refer_on" : "off"));
}

/** Calculation of show_abandoned_value */
export function calc_show_abandoned_value(app) {
  return X.out((X.bool((X.bool(X.gt(X.v(app, "c_selected_task_number"), 2)) && X.bool(X.lt(X.v(app, "c_selected_task_number"), 8)))) ? "abandoned_on" : "off"));
}

/** Calculation of is_ending_followup */
export function calc_is_ending_followup(app) {
  return X.out((X.bool(X.eq(X.v(app, "c_selected_task_number"), 8)) ? "true" : "false"));
}

/** Calculation of c_selected_task_number */
export function calc_c_selected_task_number(app) {
  return X.out(X.f["number"](app, X.v(app, "follow_up_count_number")));
}

/** Calculation of referral_with_no_followup_1 */
export function calc_referral_with_no_followup_1(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "yes")) || X.bool(X.eq(X.v(app, "s_malnutrition_sores"), "yes")))) || X.bool(X.f["selected"](app, X.v(app, "s_shakir_strip_color"), "red")))) || X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "num_yellow_strips"), "2")) && X.bool(X.f["selected"](app, X.v(app, "s_shakir_strip_color"), "yellow")))) && X.bool(X.eq(X.v(app, "patient_recovered"), "false")))) && X.bool(X.eq(X.v(app, "patient_deceased"), "false")))))) ? "true" : "false"));
}

/** Calculation of referral_with_no_followup_2 */
export function calc_referral_with_no_followup_2(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_other_signs"), "yes")) || X.bool(X.eq(X.v(app, "s_diarrhea_stools"), "yes")))) || X.bool(X.eq(X.v(app, "s_weight_loss"), "yes")))) || X.bool(X.eq(X.v(app, "fast_breathing"), "true")))) && X.bool(X.eq(X.v(app, "patient_deceased"), "false")))) && X.bool(X.eq(X.v(app, "patient_recovered"), "false")))) ? "true" : "false"));
}

/** Calculation of patient_recovered_1 */
export function calc_patient_recovered_1(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "num_green_strips"), "2")) && X.bool(X.f["selected"](app, X.v(app, "s_shakir_strip_color"), "green")))) && X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "no")))) && X.bool(X.eq(X.v(app, "s_other_signs"), "no")))) && X.bool(X.eq(X.v(app, "s_diarrhea_stools"), "no")))) && X.bool(X.eq(X.f["coalesce"](app, X.v(app, "s_weight_loss"), "no"), "no")))) && X.bool(X.eq(X.v(app, "patient_deceased"), "false")))) ? "true" : "false"));
}

/** Calculation of patient_recovered_2 */
export function calc_patient_recovered_2(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "previous_strip"), "green")) && X.bool(X.f["selected"](app, X.v(app, "s_shakir_strip_color"), "green")))) && X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "no")))) && X.bool(X.eq(X.v(app, "s_other_signs"), "no")))) && X.bool(X.eq(X.v(app, "s_diarrhea_stools"), "no")))) && X.bool(X.eq(X.f["coalesce"](app, X.v(app, "s_weight_loss"), "no"), "no")))) && X.bool(X.eq(X.v(app, "patient_deceased"), "false")))) ? "true" : "false"));
}

/** Calculation of patient_recovered_3 */
export function calc_patient_recovered_3(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.f["selected"](app, X.v(app, "s_shakir_strip_color"), "green")) && X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "no")))) && X.bool(X.eq(X.v(app, "s_other_signs"), "no")))) && X.bool(X.eq(X.v(app, "s_diarrhea_stools"), "no")))) && X.bool(X.eq(X.f["coalesce"](app, X.v(app, "s_weight_loss"), "no"), "no")))) && X.bool(X.eq(X.v(app, "patient_deceased"), "false")))) ? "true" : "false"));
}

/** Calculation of patient_recovered_4 */
export function calc_patient_recovered_4(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.f["selected"](app, X.v(app, "s_shakir_strip_color"), "yellow")) && X.bool(X.gt(X.f["number"](app, X.v(app, "num_recovered")), 1)))) && X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "no")))) && X.bool(X.eq(X.v(app, "s_other_signs"), "no")))) && X.bool(X.eq(X.v(app, "s_diarrhea_stools"), "no")))) && X.bool(X.eq(X.f["coalesce"](app, X.v(app, "s_weight_loss"), "no"), "no")))) && X.bool(X.eq(X.v(app, "patient_deceased"), "false")))) ? "true" : "false"));
}

/** Calculation of patient_recovered_5 */
export function calc_patient_recovered_5(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.f["selected"](app, X.v(app, "s_shakir_strip_color"), "yellow")) && X.bool(X.eq(X.v(app, "previous_strip"), "green")))) && X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "no")))) && X.bool(X.eq(X.v(app, "s_other_signs"), "no")))) && X.bool(X.eq(X.v(app, "s_diarrhea_stools"), "no")))) && X.bool(X.eq(X.f["coalesce"](app, X.v(app, "s_weight_loss"), "no"), "no")))) && X.bool(X.eq(X.v(app, "patient_deceased"), "false")))) ? "true" : "false"));
}

/** Calculation of patient_recovered_6 */
export function calc_patient_recovered_6(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.f["selected"](app, X.v(app, "s_shakir_strip_color"), "yellow")) && X.bool(X.eq(X.v(app, "previous_strip"), "yellow")))) && X.bool(X.gt(X.f["number"](app, X.v(app, "num_recovered")), 1)))) && X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "no")))) && X.bool(X.eq(X.v(app, "s_other_signs"), "no")))) && X.bool(X.eq(X.v(app, "s_diarrhea_stools"), "no")))) && X.bool(X.eq(X.f["coalesce"](app, X.v(app, "s_weight_loss"), "no"), "no")))) && X.bool(X.eq(X.v(app, "patient_deceased"), "false")))) ? "true" : "false"));
}

/** Calculation of patient_recovered_7 */
export function calc_patient_recovered_7(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "before_previous_child_presence"), "yes")) && X.bool(X.eq(X.v(app, "s_child_presence_confirmation"), "no")))) && X.bool(X.gt(X.f["number"](app, X.v(app, "num_recovered")), 1)))) && X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "no")))) && X.bool(X.eq(X.v(app, "s_other_signs"), "no")))) && X.bool(X.eq(X.v(app, "s_diarrhea_stools"), "no")))) && X.bool(X.eq(X.f["coalesce"](app, X.v(app, "s_weight_loss"), "no"), "no")))) && X.bool(X.eq(X.v(app, "patient_deceased"), "false")))) ? "true" : "false"));
}

/** Calculation of enriched_flour_dosage_label */
export function calc_enriched_flour_dosage_label(app) {
  return X.out(X.v(app, "s_give_enriched_flour_dosage"));
}

/** Calculation of r_followup_date */
export function calc_r_followup_date(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_followup_today"), "yes")) ? X.f["format-date-time"](app, X.f["date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app))), "%Y-%m-%d") : X.v(app, "s_followup_date")));
}

/** Calculation of use_plumpy_with_mam */
export function calc_use_plumpy_with_mam(app) {
  return X.out((X.bool(X.eq(X.inp(app, "contact/parent/parent/mam_intrant"), "")) ? "true" : (X.bool(X.eq(X.inp(app, "contact/parent/parent/mam_intrant"), "plumpy")) ? "true" : "false")));
}

/** Calculation of folic_acid_srp_given_once */
export function calc_folic_acid_srp_given_once(app) {
  return X.out(X.inp(app, "t_folic_acid_srp_given_once"));
}

/** Relevance of s_followup_date */
export function rel_s_followup_date(app) {
  return X.bool(X.eq(X.v(app, "s_followup_today"), "no"));
}

/** Constraint of s_followup_date */
export function val_s_followup_date(app) {
  if (X.str(X.v(app, "s_followup_date")) === '') return true;
  return X.bool(X.le(X.f["floor"](app, X.f["decimal-date-time"](app, X.v(app, "s_followup_date"))), X.f["floor"](app, X.f["decimal-date-time"](app, X.f["now"](app)))));
}

/** Relevance of s_child_receive_inputs */
export function rel_s_child_receive_inputs(app) {
  return X.bool(X.eq(X.v(app, "s_child_presence_confirmation"), "no"));
}

/** Relevance of s_shakir_strip_color */
export function rel_s_shakir_strip_color(app) {
  return grp_s_assessment(app);
}

/** Relevance of s_nutri_shakir_strip_length_green */
export function rel_s_nutri_shakir_strip_length_green(app) {
  return grp_s_assessment(app) && X.bool(X.eq(X.v(app, "s_shakir_strip_color"), "green"));
}

/** Constraint of s_nutri_shakir_strip_length_green */
export function val_s_nutri_shakir_strip_length_green(app) {
  if (X.str(X.v(app, "s_nutri_shakir_strip_length_green")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "s_nutri_shakir_strip_length_green"), 12.5)) && X.bool(X.le(X.v(app, "s_nutri_shakir_strip_length_green"), 26.5))));
}

/** Relevance of s_nutri_shakir_strip_length_yellow */
export function rel_s_nutri_shakir_strip_length_yellow(app) {
  return grp_s_assessment(app) && X.bool(X.eq(X.v(app, "s_shakir_strip_color"), "yellow"));
}

/** Constraint of s_nutri_shakir_strip_length_yellow */
export function val_s_nutri_shakir_strip_length_yellow(app) {
  if (X.str(X.v(app, "s_nutri_shakir_strip_length_yellow")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "s_nutri_shakir_strip_length_yellow"), 11.5)) && X.bool(X.lt(X.v(app, "s_nutri_shakir_strip_length_yellow"), 12.5))));
}

/** Relevance of s_nutri_shakir_strip_length_red */
export function rel_s_nutri_shakir_strip_length_red(app) {
  return grp_s_assessment(app) && X.bool(X.eq(X.v(app, "s_shakir_strip_color"), "red"));
}

/** Constraint of s_nutri_shakir_strip_length_red */
export function val_s_nutri_shakir_strip_length_red(app) {
  if (X.str(X.v(app, "s_nutri_shakir_strip_length_red")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "s_nutri_shakir_strip_length_red"), 5.5)) && X.bool(X.lt(X.v(app, "s_nutri_shakir_strip_length_red"), 11.5))));
}

/** Relevance of s_nutri_shakir_strip_length */
export function rel_s_nutri_shakir_strip_length(app) {
  return grp_s_assessment(app);
}

/** Calculation of s_nutri_shakir_strip_length */
export function calc_s_nutri_shakir_strip_length(app) {
  if (!rel_s_nutri_shakir_strip_length(app)) return '';
  return X.out((X.bool(X.ne(X.f["coalesce"](app, X.v(app, "s_nutri_shakir_strip_length_green"), 0), "")) ? X.f["coalesce"](app, X.v(app, "s_nutri_shakir_strip_length_green"), 0) : (X.bool(X.ne(X.f["coalesce"](app, X.v(app, "s_nutri_shakir_strip_length_yellow"), 0), "")) ? X.f["coalesce"](app, X.v(app, "s_nutri_shakir_strip_length_yellow"), 0) : X.v(app, "s_nutri_shakir_strip_length_red"))));
}

/** Relevance of child_weight */
export function rel_child_weight(app) {
  return grp_s_assessment(app);
}

/** Constraint of child_weight */
export function val_child_weight(app) {
  if (X.str(X.v(app, "child_weight")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "child_weight"), 1)) && X.bool(X.le(X.v(app, "child_weight"), 30))));
}

/** Relevance of r_enriched_flour */
export function rel_r_enriched_flour(app) {
  return grp_s_treatment_given(app) && X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false"));
}

/** Relevance of s_give_enriched_flour */
export function rel_s_give_enriched_flour(app) {
  return grp_s_treatment_given(app) && X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false"));
}

/** Relevance of s_give_enriched_flour_dosage */
export function rel_s_give_enriched_flour_dosage(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_give_enriched_flour"), "yes"))));
}

/** Constraint of s_give_enriched_flour_dosage */
export function val_s_give_enriched_flour_dosage(app) {
  if (X.str(X.v(app, "s_give_enriched_flour_dosage")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "s_give_enriched_flour_dosage"), 1)) && X.bool(X.le(X.v(app, "s_give_enriched_flour_dosage"), 99))));
}

/** Relevance of s_not_give_enriched_flour */
export function rel_s_not_give_enriched_flour(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_give_enriched_flour"), "no"))));
}

/** Relevance of s_not_give_enriched_flour_other */
export function rel_s_not_give_enriched_flour_other(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_not_give_enriched_flour"), "other"))));
}

/** Relevance of n_enriched_flour_shortage */
export function rel_n_enriched_flour_shortage(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_not_give_enriched_flour"), "out_of_stock"))));
}

/** Relevance of r_plumpy_sup_1 */
export function rel_r_plumpy_sup_1(app) {
  return grp_s_treatment_given(app) && X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true"));
}

/** Relevance of s_give_plumpy_sup */
export function rel_s_give_plumpy_sup(app) {
  return grp_s_treatment_given(app) && X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true"));
}

/** Relevance of s_given_plumpy_sup_dosage */
export function rel_s_given_plumpy_sup_dosage(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "s_give_plumpy_sup"), "yes"))));
}

/** Constraint of s_given_plumpy_sup_dosage */
export function val_s_given_plumpy_sup_dosage(app) {
  if (X.str(X.v(app, "s_given_plumpy_sup_dosage")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "s_given_plumpy_sup_dosage"), 1)) && X.bool(X.le(X.v(app, "s_given_plumpy_sup_dosage"), 99))));
}

/** Relevance of s_not_give_plumpy_sup */
export function rel_s_not_give_plumpy_sup(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "s_give_plumpy_sup"), "no"))));
}

/** Relevance of s_not_give_plumpy_sup_other */
export function rel_s_not_give_plumpy_sup_other(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "s_not_give_plumpy_sup"), "other"))));
}

/** Relevance of s_give_dewormer */
export function rel_s_give_dewormer(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 12)))) && X.bool(X.eq(X.v(app, "c_selected_task_number"), 1))));
}

/** Relevance of s_dewormer_type */
export function rel_s_dewormer_type(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "s_give_dewormer"), "yes"))));
}

/** Relevance of n_dewormer_200g_one_tab */
export function rel_n_dewormer_200g_one_tab(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 12)))) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 24)))) && X.bool(X.eq(X.v(app, "c_selected_task_number"), 1)))) && X.bool(X.eq(X.v(app, "s_dewormer_type"), "200mg"))));
}

/** Relevance of n_dewormer_200g_two_tab */
export function rel_n_dewormer_200g_two_tab(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 24)))) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 60)))) && X.bool(X.eq(X.v(app, "c_selected_task_number"), 1)))) && X.bool(X.eq(X.v(app, "s_dewormer_type"), "200mg"))));
}

/** Relevance of n_dewormer_400g_half_tab */
export function rel_n_dewormer_400g_half_tab(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 12)))) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 24)))) && X.bool(X.eq(X.v(app, "c_selected_task_number"), 1)))) && X.bool(X.eq(X.v(app, "s_dewormer_type"), "400mg"))));
}

/** Relevance of n_dewormer_400g_one_tab */
export function rel_n_dewormer_400g_one_tab(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 24)))) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 60)))) && X.bool(X.eq(X.v(app, "c_selected_task_number"), 1)))) && X.bool(X.eq(X.v(app, "s_dewormer_type"), "400mg"))));
}

/** Relevance of s_given_dewormer_dosage */
export function rel_s_given_dewormer_dosage(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "s_give_dewormer"), "yes")))) && X.bool(X.ne(X.v(app, "s_dewormer_type"), ""))));
}

/** Options of s_given_dewormer_dosage */
export function filt_s_given_dewormer_dosage(app) {
  return X.bool(X.eq(X.opt(app, "filter1"), X.v(app, "s_dewormer_type")));
}

/** Relevance of s_not_give_dewormer */
export function rel_s_not_give_dewormer(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "s_give_dewormer"), "no"))));
}

/** Relevance of s_not_give_dewormer_other */
export function rel_s_not_give_dewormer_other(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "s_not_give_dewormer"), "other"))));
}

/** Relevance of s_give_vitamina_dosage */
export function rel_s_give_vitamina_dosage(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "c_selected_task_number"), 3))));
}

/** Relevance of s_give_vitaminA */
export function rel_s_give_vitaminA(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "c_selected_task_number"), 3))));
}

/** Relevance of s_vitaminA_dosage */
export function rel_s_vitaminA_dosage(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "s_give_vitaminA"), "yes"))));
}

/** Relevance of s_not_give_vitaminA */
export function rel_s_not_give_vitaminA(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "s_give_vitaminA"), "no"))));
}

/** Relevance of s_not_give_vitaminA_other */
export function rel_s_not_give_vitaminA_other(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "s_not_give_vitaminA"), "other"))));
}

/** Relevance of s_give_iron_folic */
export function rel_s_give_iron_folic(app) {
  return grp_s_treatment_given(app) && X.bool(X.gt(X.v(app, "c_selected_task_number"), 1));
}

/** Relevance of s_not_give_iron_folic */
export function rel_s_not_give_iron_folic(app) {
  return grp_s_treatment_given(app) && X.bool(X.eq(X.v(app, "s_give_iron_folic"), "no"));
}

/** Relevance of s_not_give_iron_folic_other */
export function rel_s_not_give_iron_folic_other(app) {
  return grp_s_treatment_given(app) && X.bool(X.eq(X.v(app, "s_not_give_iron_folic"), "other"));
}

/** Relevance of s_iron_folic_type */
export function rel_s_iron_folic_type(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.gt(X.v(app, "c_selected_task_number"), 1)) && X.bool(X.eq(X.v(app, "s_give_iron_folic"), "yes"))));
}

/** Relevance of s_give_iron_folic_srp_dosage_1 */
export function rel_s_give_iron_folic_srp_dosage_1(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool((X.bool(X.gt(X.v(app, "c_selected_task_number"), 1)) && X.bool(X.lt(X.v(app, "child_weight"), 10)))) && X.bool(X.eq(X.v(app, "s_iron_folic_type"), "iron_folic_acid_srp"))));
}

/** Relevance of s_give_iron_folic_srp_dosage_2 */
export function rel_s_give_iron_folic_srp_dosage_2(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool((X.bool(X.gt(X.v(app, "c_selected_task_number"), 1)) && X.bool(X.ge(X.v(app, "child_weight"), 10)))) && X.bool(X.eq(X.v(app, "s_iron_folic_type"), "iron_folic_acid_srp"))));
}

/** Relevance of s_give_iron_folic_cp_dosage_1 */
export function rel_s_give_iron_folic_cp_dosage_1(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool((X.bool(X.gt(X.v(app, "c_selected_task_number"), 1)) && X.bool(X.lt(X.v(app, "child_weight"), 10)))) && X.bool(X.eq(X.v(app, "s_iron_folic_type"), "iron_folic_acid"))));
}

/** Relevance of s_give_iron_folic_cp_dosage_2 */
export function rel_s_give_iron_folic_cp_dosage_2(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool((X.bool(X.gt(X.v(app, "c_selected_task_number"), 1)) && X.bool(X.ge(X.v(app, "child_weight"), 10)))) && X.bool(X.eq(X.v(app, "s_iron_folic_type"), "iron_folic_acid"))));
}

/** Relevance of s_iron_folic_dosage */
export function rel_s_iron_folic_dosage(app) {
  return grp_s_treatment_given(app) && X.bool((X.bool(X.eq(X.v(app, "s_give_iron_folic"), "yes")) && X.bool(X.ne(X.v(app, "s_iron_folic_type"), ""))));
}

/** Options of s_iron_folic_dosage */
export function filt_s_iron_folic_dosage(app) {
  return X.bool(X.eq(X.opt(app, "rescat"), X.v(app, "s_iron_folic_type")));
}

/** Relevance of s_malnutrition_edema */
export function rel_s_malnutrition_edema(app) {
  return grp_medical_complications(app);
}

/** Relevance of s_malnutrition_sores */
export function rel_s_malnutrition_sores(app) {
  return grp_medical_complications(app) && X.bool((X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "yes")) || X.bool((X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "no")) && X.bool(X.eq(X.v(app, "s_shakir_strip_color"), "Red"))))));
}

/** Relevance of s_other_signs */
export function rel_s_other_signs(app) {
  return grp_medical_complications(app);
}

/** Relevance of s_diarrhea_stools */
export function rel_s_diarrhea_stools(app) {
  return grp_medical_complications(app);
}

/** Relevance of s_malnutrition_cough */
export function rel_s_malnutrition_cough(app) {
  return grp_medical_complications(app);
}

/** Relevance of breath_timer */
export function rel_breath_timer(app) {
  return grp_s_respiration(app);
}

/** Relevance of s_ari_respiration_rate */
export function rel_s_ari_respiration_rate(app) {
  return grp_s_respiration(app);
}

/** Constraint of s_ari_respiration_rate */
export function val_s_ari_respiration_rate(app) {
  if (X.str(X.v(app, "s_ari_respiration_rate")) === '') return true;
  return X.bool(X.gt(X.v(app, "s_ari_respiration_rate"), 0));
}

/** Relevance of s_ari_resp_estimation */
export function rel_s_ari_resp_estimation(app) {
  return grp_s_respiration(app) && X.bool(X.ne(X.v(app, "s_ari_respiration_rate"), ""));
}

/** Relevance of s_weight_loss */
export function rel_s_weight_loss(app) {
  return grp_s_weight(app);
}

/** Relevance of r_diagnosis */
export function rel_r_diagnosis(app) {
  return X.bool(X.eq(X.v(app, "patient_deceased"), "false"));
}

/** Relevance of n_shakir_red */
export function rel_n_shakir_red(app) {
  return X.bool(X.eq(X.v(app, "s_shakir_strip_color"), "red"));
}

/** Relevance of n_shakir_yellow */
export function rel_n_shakir_yellow(app) {
  return X.bool(X.eq(X.v(app, "s_shakir_strip_color"), "yellow"));
}

/** Relevance of n_malnutrition_edema */
export function rel_n_malnutrition_edema(app) {
  return X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "yes"));
}

/** Relevance of r_malnutrition_sores */
export function rel_r_malnutrition_sores(app) {
  return X.bool(X.eq(X.v(app, "s_malnutrition_sores"), "yes"));
}

/** Relevance of n_other_signs */
export function rel_n_other_signs(app) {
  return X.bool(X.eq(X.v(app, "s_other_signs"), "yes"));
}

/** Relevance of n_diarrhea_stools */
export function rel_n_diarrhea_stools(app) {
  return X.bool(X.eq(X.v(app, "s_diarrhea_stools"), "yes"));
}

/** Relevance of n_ari_resp_estimation */
export function rel_n_ari_resp_estimation(app) {
  return X.bool(X.eq(X.v(app, "fast_breathing"), "true"));
}

/** Relevance of n_weight_loss */
export function rel_n_weight_loss(app) {
  return X.bool(X.eq(X.v(app, "s_weight_loss"), "yes"));
}

/** Relevance of n_shakir_yellow_3_times */
export function rel_n_shakir_yellow_3_times(app) {
  return X.bool((X.bool(X.eq(X.v(app, "num_yellow_strips"), "2")) && X.bool(X.f["selected"](app, X.v(app, "s_shakir_strip_color"), "yellow"))));
}

/** Relevance of n_abandon */
export function rel_n_abandon(app) {
  return X.bool(X.eq(X.v(app, "abandon_followup"), "true"));
}

/** Relevance of r_referral */
export function rel_r_referral(app) {
  return X.bool((X.bool(X.eq(X.v(app, "referral_with_followup"), "true")) || X.bool(X.eq(X.v(app, "referral_with_no_followup"), "true"))));
}

/** Relevance of r_referral_1 */
export function rel_r_referral_1(app) {
  return X.bool(X.eq(X.v(app, "referral_with_followup"), "true"));
}

/** Relevance of r_referral_2 */
export function rel_r_referral_2(app) {
  return X.bool(X.eq(X.v(app, "referral_with_no_followup"), "true"));
}

/** Relevance of r_referral_3 */
export function rel_r_referral_3(app) {
  return X.bool(X.eq(X.v(app, "abandon_followup"), "true"));
}

/** Relevance of r_referral_4 */
export function rel_r_referral_4(app) {
  return X.bool(X.eq(X.v(app, "patient_recovered"), "true"));
}

/** Relevance of r_referral_5 */
export function rel_r_referral_5(app) {
  return X.bool(X.eq(X.v(app, "patient_should_continue_followup"), "true"));
}

/** Relevance of cat */
export function rel_cat(app) {
  return X.bool(X.eq(X.v(app, "patient_deceased"), "false"));
}

/** Options of cat */
export function filt_cat(app) {
  return X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.opt(app, "rescat"), "all")) || X.bool(X.eq(X.opt(app, "rescat"), X.v(app, "show_continue_value"))))) || X.bool(X.eq(X.opt(app, "rescat"), X.v(app, "show_recovered_value"))))) || X.bool(X.eq(X.opt(app, "rescat"), X.v(app, "show_refer_value"))))) || X.bool(X.eq(X.opt(app, "rescat"), X.v(app, "show_abandoned_value")))));
}

/** Relevance of why_different_cat_note */
export function rel_why_different_cat_note(app) {
  return X.bool(X.eq(X.v(app, "chw_selected_computed_cat"), "false"));
}

/** Relevance of why_different_cat */
export function rel_why_different_cat(app) {
  return X.bool(X.eq(X.v(app, "chw_selected_computed_cat"), "false"));
}

/** Constraint of why_different_cat */
export function val_why_different_cat(app) {
  if (X.str(X.v(app, "why_different_cat")) === '') return true;
  return X.bool(X.eq(X.v(app, "why_different_cat"), "yes"));
}

/** Relevance of specify_why_different_cat */
export function rel_specify_why_different_cat(app) {
  return X.bool(X.eq(X.v(app, "why_different_cat"), "yes"));
}

/** Relevance of r_treatment */
export function rel_r_treatment(app) {
  return X.bool(X.eq(X.v(app, "referral"), "false"));
}

/** Relevance of r_give_plumpy_sup */
export function rel_r_give_plumpy_sup(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "referral"), "false")))) && X.bool(X.eq(X.v(app, "s_give_plumpy_sup"), "yes"))));
}

/** Relevance of r_not_give_plumpy_sup */
export function rel_r_not_give_plumpy_sup(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "referral"), "false")))) && X.bool(X.eq(X.v(app, "s_give_plumpy_sup"), "no"))));
}

/** Relevance of r_give_dewormer */
export function rel_r_give_dewormer(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "referral"), "false")))) && X.bool(X.eq(X.v(app, "s_give_dewormer"), "yes"))));
}

/** Relevance of r_not_give_dewormer */
export function rel_r_not_give_dewormer(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "referral"), "false")))) && X.bool(X.eq(X.v(app, "s_give_dewormer"), "no"))));
}

/** Relevance of r_give_vitaminA */
export function rel_r_give_vitaminA(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "referral"), "false")))) && X.bool(X.eq(X.v(app, "s_give_vitaminA"), "yes"))));
}

/** Relevance of r_not_give_vitaminA */
export function rel_r_not_give_vitaminA(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "referral"), "false")))) && X.bool(X.eq(X.v(app, "s_give_vitaminA"), "no"))));
}

/** Relevance of r_give_enriched_flour */
export function rel_r_give_enriched_flour(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "referral"), "false")))) && X.bool(X.eq(X.v(app, "s_give_enriched_flour"), "yes"))));
}

/** Relevance of r_not_give_enriched_flour */
export function rel_r_not_give_enriched_flour(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "referral"), "false")))) && X.bool(X.eq(X.v(app, "s_give_enriched_flour"), "no"))));
}

/** Relevance of r_give_iron_folic */
export function rel_r_give_iron_folic(app) {
  return X.bool((X.bool(X.eq(X.v(app, "referral"), "false")) && X.bool(X.eq(X.v(app, "s_give_iron_folic"), "yes"))));
}

/** Relevance of r_not_give_iron_folic */
export function rel_r_not_give_iron_folic(app) {
  return X.bool((X.bool(X.eq(X.v(app, "referral"), "false")) && X.bool(X.eq(X.v(app, "s_give_iron_folic"), "no"))));
}

/** Relevance of r_instructions */
export function rel_r_instructions(app) {
  return X.bool(X.ne(X.v(app, "cat"), "continue_followup"));
}

/** Relevance of r_close_out_recovered */
export function rel_r_close_out_recovered(app) {
  return X.bool(X.eq(X.v(app, "cat"), "recovered"));
}

/** Relevance of r_referral_no_followup */
export function rel_r_referral_no_followup(app) {
  return X.bool(X.eq(X.v(app, "cat"), "refer_with_no_followup"));
}

/** Relevance of r_referral_with_followup */
export function rel_r_referral_with_followup(app) {
  return X.bool(X.eq(X.v(app, "cat"), "refer_with_followup"));
}

/** Relevance of r_close_out_abandoned */
export function rel_r_close_out_abandoned(app) {
  return X.bool(X.eq(X.v(app, "cat"), "abandoned"));
}

/** Relevance of r_close_out_deceased */
export function rel_r_close_out_deceased(app) {
  return X.bool(X.eq(X.v(app, "s_child_presence_confirmation"), "deceased"));
}

/** Relevance of r_continue_followup */
export function rel_r_continue_followup(app) {
  return X.bool(X.eq(X.v(app, "s_child_receive_inputs"), "yes"));
}

/** Relevance of r_come_back */
export function rel_r_come_back(app) {
  return X.bool(X.eq(X.v(app, "s_child_receive_inputs"), "no"));
}

/** Relevance of s_next_followup_visit */
export function rel_s_next_followup_visit(app) {
  return X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "cat"), "continue_followup")) || X.bool(X.eq(X.v(app, "cat"), "refer_with_followup")))) || X.bool(X.eq(X.v(app, "cat"), "recovered")))) && X.bool(X.eq(X.v(app, "is_ending_followup"), "false"))));
}

/** Relevance of s_no_next_visit */
export function rel_s_no_next_visit(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "cat"), "refer_with_no_followup")) || X.bool(X.eq(X.v(app, "cat"), "abandoned")))) || X.bool(X.eq(X.v(app, "s_child_presence_confirmation"), "deceased"))));
}

/** Relevance of s_next_visit_within_3days */
export function rel_s_next_visit_within_3days(app) {
  return X.bool(X.eq(X.v(app, "s_child_receive_inputs"), "no"));
}

/** Calculation of albendazol_200_disp */
export function calc_albendazol_200_disp(app) {
  return X.out((X.bool((X.bool(X.eq(X.v(app, "s_give_dewormer"), "yes")) && X.bool(X.eq(X.v(app, "s_dewormer_type"), "200mg")))) ? (X.bool(X.eq(X.v(app, "s_given_dewormer_dosage"), "two_tab")) ? 2 : 1) : 0));
}

/** Calculation of albendazol_400_disp */
export function calc_albendazol_400_disp(app) {
  return X.out((X.bool((X.bool(X.eq(X.v(app, "s_give_dewormer"), "yes")) && X.bool(X.eq(X.v(app, "s_dewormer_type"), "400mg")))) ? 1 : 0));
}

/** Calculation of plumpy_sup_disp */
export function calc_plumpy_sup_disp(app) {
  return X.out(X.f["coalesce"](app, X.v(app, "s_given_plumpy_sup_dosage"), 0));
}

/** Calculation of enriched_flour_disp */
export function calc_enriched_flour_disp(app) {
  return X.out(X.f["coalesce"](app, X.v(app, "s_give_enriched_flour_dosage"), 0));
}

/** Calculation of iron_folic_acid_srp_disp */
export function calc_iron_folic_acid_srp_disp(app) {
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_give_iron_folic"), "yes")) && X.bool(X.eq(X.v(app, "s_iron_folic_type"), "iron_folic_acid_srp")))) && X.bool(X.eq(X.v(app, "folic_acid_srp_given_once"), "false")))) ? 1 : 0));
}

/** Calculation of iron_folic_acid_disp */
export function calc_iron_folic_acid_disp(app) {
  return X.out((X.bool((X.bool(X.eq(X.v(app, "s_give_iron_folic"), "yes")) && X.bool(X.eq(X.v(app, "s_iron_folic_type"), "iron_folic_acid")))) ? 1 : 0));
}

/** Calculation of vitamin_a_blue_capsule_disp */
export function calc_vitamin_a_blue_capsule_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_vitaminA_dosage"), "blue_cap")) ? 1 : 0));
}

/** Calculation of vitamin_a_red_capsule_disp */
export function calc_vitamin_a_red_capsule_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_vitaminA_dosage"), "red_cap")) ? 1 : 0));
}

/** Relevance of stock_monitoring_reported_date */
export function rel_stock_monitoring_reported_date(app) {
  return grp_prescription_summary(app);
}

/** Calculation of stock_monitoring_reported_date */
export function calc_stock_monitoring_reported_date(app) {
  if (!rel_stock_monitoring_reported_date(app)) return '';
  return X.out(X.v(app, "r_followup_date"));
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
  return X.out("moderate_malnutrition_followup");
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

/** Task input t_follow_up_type */
export function calc_inputs_t_follow_up_type(app) {
  return X.out(X.inp(app, "t_follow_up_type"));
}

/** Task input t_follow_up_count */
export function calc_inputs_t_follow_up_count(app) {
  return X.out(X.inp(app, "t_follow_up_count"));
}

/** Task input t_previous_strip */
export function calc_inputs_t_previous_strip(app) {
  return X.out(X.inp(app, "t_previous_strip"));
}

/** Task input t_num_recovered */
export function calc_inputs_t_num_recovered(app) {
  return X.out(X.inp(app, "t_num_recovered"));
}

/** Task input t_before_previous_child_presence */
export function calc_inputs_t_before_previous_child_presence(app) {
  return X.out(X.inp(app, "t_before_previous_child_presence"));
}

/** Task input t_num_successive_green_strips */
export function calc_inputs_t_num_successive_green_strips(app) {
  return X.out(X.inp(app, "t_num_successive_green_strips"));
}

/** Task input t_num_successive_yellow_strips */
export function calc_inputs_t_num_successive_yellow_strips(app) {
  return X.out(X.inp(app, "t_num_successive_yellow_strips"));
}

/** Task input t_num_successive_red_strips */
export function calc_inputs_t_num_successive_red_strips(app) {
  return X.out(X.inp(app, "t_num_successive_red_strips"));
}

/** Task input t_num_yellow_strips */
export function calc_inputs_t_num_yellow_strips(app) {
  return X.out(X.inp(app, "t_num_yellow_strips"));
}

/** Task input t_num_child_presence_confirmation */
export function calc_inputs_t_num_child_presence_confirmation(app) {
  return X.out(X.inp(app, "t_num_child_presence_confirmation"));
}

/** Task input t_num_no_edema */
export function calc_inputs_t_num_no_edema(app) {
  return X.out(X.inp(app, "t_num_no_edema"));
}

/** Task input t_folic_acid_srp_given_once */
export function calc_inputs_t_folic_acid_srp_given_once(app) {
  return X.out(X.inp(app, "t_folic_acid_srp_given_once"));
}

/** Task input t_assessment_date */
export function calc_inputs_t_assessment_date(app) {
  return X.out(X.inp(app, "t_assessment_date"));
}
