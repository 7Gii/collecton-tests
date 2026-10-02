// Form referral_followup_under_5: ported from forms/app/referral_followup_under_5.xlsx of the CHT configuration muso-mali.
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
 "c_task_to_perform": {
  "follow_up": "Suivi",
  "close_out": "Clôture"
 },
 "c_how_assess": {
  "c_how_assess_1": "Présence avec touche",
  "c_how_assess_2": "Présence sans touche"
 },
 "c_how_assess_no_ppe": {
  "c_how_assess_no_ppe_1": "Présence sans touche",
  "c_how_assess_no_ppe_2": "Par téléphone"
 },
 "no_cscom_reason": {
  "parent_refusal": "Refus d'un parent",
  "lack_time": "Manque de temps",
  "other": "Autre"
 },
 "c_how_disease_progressing": {
  "aggravated": "Aggravée",
  "no_change": "Pas de changement",
  "improved": "Améliorée",
  "cured": "Guéri"
 },
 "c_why_close_out": {
  "patient_died": "Le patient est décédé",
  "refusal_of_parent": "Refus des parents pour diverses raisons",
  "chw_unavailable": "ASC non disponible"
 },
 "c_where_child_died": {
  "home": "Domicile",
  "cscom": "CSCOM",
  "csref": "CSREF",
  "hospital": "Hôpital",
  "other": "Autre"
 },
 "c_give_paracetamol_dosage": {
  "quarter_tab_paracetamol_once": "¼ comprimé",
  "half_tab_paracetamol_once": "½ comprimé",
  "quarter_tab_paracetamol_3x_a_day": "¼ comprimé 3 fois par jour",
  "half_tab_paracetamol_3x_a_day": "½ comprimé 3 fois par jour"
 },
 "medication_not_given": {
  "out_of_stock": "Rupture de stock",
  "other": "Autre"
 }
};

/** CHT properties.json: not offered in the add menu, opened by its task only. */
export function showForm(app) {
  return false;
}

/** Group s_registration_mode */
export function grp_s_registration_mode(app) {
  return X.bool(X.eq(X.v(app, "s_task_to_perform"), "follow_up"));
}

/** Group s_patient_available */
export function grp_s_patient_available(app) {
  return X.bool(X.eq(X.v(app, "s_task_to_perform"), "follow_up"));
}

/** Group s_followup_ref */
export function grp_s_followup_ref(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_task_to_perform"), "follow_up")) && X.bool(X.eq(X.v(app, "s_patient_present"), "yes"))));
}

/** Group s_temperature */
export function grp_s_temperature(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_task_to_perform"), "follow_up")) && X.bool(X.eq(X.v(app, "s_patient_present"), "yes"))));
}

/** Group s_temperature_retake */
export function grp_s_temperature_retake(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "s_patient_present"), "yes")) && X.bool(X.lt(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw"), 0)), 35)))) || X.bool(X.gt(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw"), 0)), 41))));
}

/** Group s_disease_progression */
export function grp_s_disease_progression(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_task_to_perform"), "follow_up")) && X.bool(X.eq(X.v(app, "s_patient_present"), "yes"))));
}

/** Group s_close_out */
export function grp_s_close_out(app) {
  return X.bool(X.f["selected"](app, X.v(app, "s_task_to_perform"), "close_out"));
}

/** Group group_diagnosis */
export function grp_group_diagnosis(app) {
  return X.bool(X.ne(X.v(app, "s_why_close_out"), "patient_died"));
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
  return X.out(X.f["floor"](app, (X.num(X.f["difference-in-months"](app, X.v(app, "patient_dob"), X.f["now"](app))) / X.num(12))));
}

/** Default of patient_age_in_years */
export function def_patient_age_in_years(app) {
  return "0";
}

/** Calculation of patient_age_in_months */
export function calc_patient_age_in_months(app) {
  return X.out(X.f["difference-in-months"](app, X.v(app, "patient_dob"), X.f["now"](app)));
}

/** Default of patient_age_in_months */
export function def_patient_age_in_months(app) {
  return "0";
}

/** Calculation of patient_age_in_days */
export function calc_patient_age_in_days(app) {
  return X.out(X.f["floor"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) - X.num(X.f["decimal-date-time"](app, X.v(app, "patient_dob"))))));
}

/** Calculation of patient_age_display_bm */
export function calc_patient_age_display_bm(app) {
  return X.out(X.f["concat"](app, X.f["concat"](app, (X.bool(X.ne(X.v(app, "patient_age_in_years"), "")) ? "San " : ""), X.v(app, "patient_age_in_years")), X.f["concat"](app, (X.bool(X.eq((X.num(X.v(app, "patient_age_in_months")) % X.num(12)), 1)) ? " ani kalo " : " ani kalo "), (X.num(X.v(app, "patient_age_in_months")) % X.num(12)))));
}

/** Calculation of patient_sex */
export function calc_patient_sex(app) {
  return X.out(X.inp(app, "contact/sex"));
}

/** Calculation of patient_sex_bm */
export function calc_patient_sex_bm(app) {
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_sex"), "male")) || X.bool(X.eq(X.v(app, "patient_sex"), "homme")))) || X.bool(X.eq(X.v(app, "patient_sex"), "Homme")))) ? "Cɛ" : (X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_sex"), "female")) || X.bool(X.eq(X.v(app, "patient_sex"), "femme")))) || X.bool(X.eq(X.v(app, "patient_sex"), "Femme")))) ? "Muso" : "")));
}

/** Calculation of patient_age_display */
export function calc_patient_age_display(app) {
  return X.out((X.bool(X.gt(X.v(app, "patient_age_in_days"), 28)) ? X.f["concat"](app, X.f["concat"](app, X.v(app, "patient_age_in_years"), (X.bool(X.eq(X.v(app, "patient_age_in_years"), 1)) ? " year and " : " years and ")), X.f["concat"](app, (X.num(X.v(app, "patient_age_in_months")) % X.num(12)), (X.bool(X.eq((X.num(X.v(app, "patient_age_in_months")) % X.num(12)), 1)) ? " month old" : " months old"))) : X.f["concat"](app, X.v(app, "patient_age_in_days"), (X.bool(X.eq(X.v(app, "patient_age_in_days"), 1)) ? " day old" : " days old"))));
}

/** Calculation of patient_id */
export function calc_patient_id(app) {
  return X.out(X.inp(app, "contact/_id"));
}

/** Calculation of patient_name */
export function calc_patient_name(app) {
  return X.out(X.inp(app, "contact/name"));
}

/** Calculation of patient_dob */
export function calc_patient_dob(app) {
  return X.out(X.f["coalesce"](app, X.inp(app, "t_delivery_date"), X.inp(app, "contact/date_of_birth")));
}

/** Calculation of follow_up_count */
export function calc_follow_up_count(app) {
  return X.out(X.inp(app, "t_follow_up_count"));
}

/** Calculation of tb_case_category */
export function calc_tb_case_category(app) {
  return X.out(X.inp(app, "t_tb_case_category"));
}

/** Calculation of close_out */
export function calc_close_out(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_task_to_perform"), "close_out")) ? "true" : "false"));
}

/** Calculation of next_visit */
export function calc_next_visit(app) {
  return X.out(X.f["date-time"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) + X.num(1))));
}

/** Calculation of next_visit_show */
export function calc_next_visit_show(app) {
  return X.out(X.f["format-date-time"](app, X.f["date-time"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) + X.num(1))), "%d/%m/%Y"));
}

/** Calculation of referral */
export function calc_referral(app) {
  return X.out((X.bool((X.bool(X.eq(X.v(app, "s_how_disease_progressing"), "aggravated")) || X.bool(X.eq(X.v(app, "s_how_disease_progressing"), "no_change")))) ? "true" : "false"));
}

/** Calculation of next_followup_visit */
export function calc_next_followup_visit(app) {
  return X.out((X.bool(X.eq(X.v(app, "follow_up_count"), "1")) ? X.f["format-date"](app, X.f["date"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) + X.num(1))), "%e/%n/%Y") : (X.bool(X.eq(X.v(app, "follow_up_count"), "2")) ? X.f["format-date"](app, X.f["date"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) + X.num(2))), "%e/%n/%Y") : (X.bool(X.eq(X.v(app, "follow_up_count"), "3")) ? X.f["format-date"](app, X.f["date"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) + X.num(3))), "%e/%n/%Y") : (X.bool(X.eq(X.v(app, "follow_up_count"), "4")) ? X.f["format-date"](app, X.f["date"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) + X.num(5))), "%e/%n/%Y") : "")))));
}

/** Calculation of follow_up_count_label_en */
export function calc_follow_up_count_label_en(app) {
  return X.out((X.bool(X.eq(X.v(app, "follow_up_count"), "1")) ? "24 Hours Followup" : (X.bool(X.eq(X.v(app, "follow_up_count"), "2")) ? "48 Hours Followup" : (X.bool(X.eq(X.v(app, "follow_up_count"), "3")) ? "72 Hours Followup" : (X.bool(X.eq(X.v(app, "follow_up_count"), "4")) ? "5 Days Followup" : "")))));
}

/** Calculation of follow_up_count_label_fr */
export function calc_follow_up_count_label_fr(app) {
  return X.out((X.bool(X.eq(X.v(app, "follow_up_count"), "1")) ? "Suivi 24 Heures" : (X.bool(X.eq(X.v(app, "follow_up_count"), "2")) ? "Suivi 48 Heures" : (X.bool(X.eq(X.v(app, "follow_up_count"), "3")) ? "Suivi 72 Heures" : (X.bool(X.eq(X.v(app, "follow_up_count"), "4")) ? "Suivi 5 jours" : "")))));
}

/** Calculation of follow_up_count_number */
export function calc_follow_up_count_number(app) {
  return X.out(X.f["number"](app, X.inp(app, "t_follow_up_count")));
}

/** Calculation of s_has_fever_now */
export function calc_s_has_fever_now(app) {
  return X.out((X.bool(X.ge(X.f["coalesce"](app, X.v(app, "s_child_temperature"), 0), 37.5)) ? "yes" : "no"));
}

/** Calculation of followup_date */
export function calc_followup_date(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_followup_today"), "yes")) ? X.f["format-date-time"](app, X.f["date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app))), "%Y-%m-%d") : X.v(app, "s_followup_date")));
}

/** Relevance of s_have_ppe */
export function rel_s_have_ppe(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_task_to_perform"), "follow_up")) && X.bool(X.eq(X.v(app, "tb_case_category"), "suspect"))));
}

/** Relevance of n_ppe_patient_touch_warning */
export function rel_n_ppe_patient_touch_warning(app) {
  return X.bool(X.eq(X.v(app, "s_have_ppe"), "no"));
}

/** Relevance of n_task_without_touch_advice */
export function rel_n_task_without_touch_advice(app) {
  return X.bool(X.eq(X.v(app, "s_have_ppe"), "yes"));
}

/** Relevance of n_wear_ppe_protection_rules */
export function rel_n_wear_ppe_protection_rules(app) {
  return X.bool(X.eq(X.v(app, "s_have_ppe"), "yes"));
}

/** Relevance of n_hand_wash_before_after */
export function rel_n_hand_wash_before_after(app) {
  return X.bool(X.eq(X.v(app, "s_have_ppe"), "yes"));
}

/** Relevance of n_cough_hygiene_disinfect */
export function rel_n_cough_hygiene_disinfect(app) {
  return X.bool(X.eq(X.v(app, "s_have_ppe"), "yes"));
}

/** Relevance of s_followup_today */
export function rel_s_followup_today(app) {
  return grp_s_registration_mode(app);
}

/** Relevance of s_followup_date */
export function rel_s_followup_date(app) {
  return grp_s_registration_mode(app) && X.bool(X.eq(X.v(app, "s_followup_today"), "no"));
}

/** Constraint of s_followup_date */
export function val_s_followup_date(app) {
  if (X.str(X.v(app, "s_followup_date")) === '') return true;
  return X.bool(X.le(X.f["floor"](app, X.f["decimal-date-time"](app, X.v(app, "s_followup_date"))), X.f["floor"](app, X.f["decimal-date-time"](app, X.f["now"](app)))));
}

/** Relevance of s_patient_present */
export function rel_s_patient_present(app) {
  return grp_s_patient_available(app);
}

/** Relevance of s_has_consult_cscom */
export function rel_s_has_consult_cscom(app) {
  return grp_s_followup_ref(app);
}

/** Relevance of s_why_no_cscom */
export function rel_s_why_no_cscom(app) {
  return grp_s_followup_ref(app) && X.bool(X.eq(X.v(app, "s_has_consult_cscom"), "no"));
}

/** Relevance of s_other */
export function rel_s_other(app) {
  return grp_s_followup_ref(app) && X.bool(X.eq(X.v(app, "s_why_no_cscom"), "other"));
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

/** Calculation of s_child_temperature */
export function calc_s_child_temperature(app) {
  return X.out((X.bool((X.bool(X.lt(X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw"), 0), 35)) || X.bool(X.gt(X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw"), 0), 41)))) ? (X.num(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw_retake"), 0))) + X.num(X.f["number"](app, 0))) : (X.num(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw"), 0))) + X.num(X.f["number"](app, 0)))));
}

/** Relevance of s_how_disease_progressing */
export function rel_s_how_disease_progressing(app) {
  return grp_s_disease_progression(app);
}

/** Relevance of s_new_danger_sign */
export function rel_s_new_danger_sign(app) {
  return grp_s_disease_progression(app) && X.bool((X.bool(X.eq(X.v(app, "s_how_disease_progressing"), "cured")) || X.bool(X.eq(X.v(app, "s_how_disease_progressing"), "improved"))));
}

/** Relevance of s_why_close_out */
export function rel_s_why_close_out(app) {
  return grp_s_close_out(app);
}

/** Relevance of s_when_child_died */
export function rel_s_when_child_died(app) {
  return grp_s_close_out(app) && X.bool(X.f["selected"](app, X.v(app, "s_why_close_out"), "patient_died"));
}

/** Constraint of s_when_child_died */
export function val_s_when_child_died(app) {
  if (X.str(X.v(app, "s_when_child_died")) === '') return true;
  return X.bool((X.bool(X.le(X.f["decimal-date-time"](app, X.v(app, "s_when_child_died")), X.f["decimal-date-time"](app, X.f["now"](app)))) && X.bool(X.lt((X.num(X.f["decimal-date-time"](app, X.f["now"](app))) - X.num(X.f["decimal-date-time"](app, X.v(app, "s_when_child_died")))), 365))));
}

/** Relevance of s_where_child_died */
export function rel_s_where_child_died(app) {
  return grp_s_close_out(app) && X.bool(X.f["selected"](app, X.v(app, "s_why_close_out"), "patient_died"));
}

/** Relevance of other */
export function rel_other(app) {
  return grp_s_close_out(app) && X.bool(X.eq(X.v(app, "s_where_child_died"), "other"));
}

/** Relevance of patient_summary */
export function rel_patient_summary(app) {
  return grp_group_diagnosis(app);
}

/** Relevance of patient_info */
export function rel_patient_info(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.inp(app, "t_delivery_date"), ""));
}

/** Relevance of r_accompany */
export function rel_r_accompany(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.v(app, "referral"), "true"));
}

/** Relevance of special_instructions */
export function rel_special_instructions(app) {
  return grp_group_diagnosis(app);
}

/** Relevance of r_accompany_text */
export function rel_r_accompany_text(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.v(app, "referral"), "true"));
}

/** Relevance of s_accompany_refer_CSCOM */
export function rel_s_accompany_refer_CSCOM(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.v(app, "referral"), "true"));
}

/** Relevance of s_advices_1 */
export function rel_s_advices_1(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.v(app, "s_new_danger_sign"), "no"));
}

/** Relevance of s_advices_4 */
export function rel_s_advices_4(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.v(app, "s_new_danger_sign"), "yes"));
}

/** Relevance of n_present */
export function rel_n_present(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.v(app, "s_patient_present"), "no"));
}

/** Relevance of r_treatment */
export function rel_r_treatment(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.v(app, "s_has_fever_now"), "yes"));
}

/** Relevance of s_dnt_fever */
export function rel_s_dnt_fever(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.v(app, "s_has_fever_now"), "yes"));
}

/** Relevance of r_fever_paracetamol */
export function rel_r_fever_paracetamol(app) {
  return grp_group_diagnosis(app) && X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)) && X.bool(X.eq(X.v(app, "s_has_fever_now"), "yes"))));
}

/** Relevance of r_fever_paracetamol_1 */
export function rel_r_fever_paracetamol_1(app) {
  return grp_group_diagnosis(app) && X.bool((X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)) && X.bool(X.lt(X.v(app, "patient_age_in_years"), 1)))) && X.bool(X.eq(X.v(app, "s_has_fever_now"), "yes")))) && X.bool(X.eq(X.v(app, "referral"), "true"))));
}

/** Relevance of r_fever_paracetamol_1_no_reference */
export function rel_r_fever_paracetamol_1_no_reference(app) {
  return grp_group_diagnosis(app) && X.bool((X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)) && X.bool(X.lt(X.v(app, "patient_age_in_years"), 1)))) && X.bool(X.eq(X.v(app, "s_has_fever_now"), "yes")))) && X.bool(X.eq(X.v(app, "referral"), "false"))));
}

/** Relevance of r_fever_paracetamol_2 */
export function rel_r_fever_paracetamol_2(app) {
  return grp_group_diagnosis(app) && X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_years"), 1)) && X.bool(X.eq(X.v(app, "s_has_fever_now"), "yes")))) && X.bool(X.eq(X.v(app, "referral"), "true"))));
}

/** Relevance of r_fever_paracetamol_2_no_reference */
export function rel_r_fever_paracetamol_2_no_reference(app) {
  return grp_group_diagnosis(app) && X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_years"), 1)) && X.bool(X.eq(X.v(app, "s_has_fever_now"), "yes")))) && X.bool(X.eq(X.v(app, "referral"), "false"))));
}

/** Relevance of s_give_paracetamol */
export function rel_s_give_paracetamol(app) {
  return grp_group_diagnosis(app) && X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)) && X.bool(X.eq(X.v(app, "s_has_fever_now"), "yes"))));
}

/** Relevance of r_paracetamol_once_filter */
export function rel_r_paracetamol_once_filter(app) {
  return grp_group_diagnosis(app);
}

/** Calculation of r_paracetamol_once_filter */
export function calc_r_paracetamol_once_filter(app) {
  if (!rel_r_paracetamol_once_filter(app)) return '';
  return X.out((X.bool(X.eq(X.v(app, "referral"), "true")) ? "once" : "many"));
}

/** Relevance of s_give_paracetamol_dosage */
export function rel_s_give_paracetamol_dosage(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.v(app, "s_give_paracetamol"), "yes"));
}

/** Options of s_give_paracetamol_dosage */
export function filt_s_give_paracetamol_dosage(app) {
  return X.bool(X.eq(X.opt(app, "filter1"), X.v(app, "r_paracetamol_once_filter")));
}

/** Relevance of s_not_give_paracetamol */
export function rel_s_not_give_paracetamol(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.v(app, "s_give_paracetamol"), "no"));
}

/** Relevance of s_not_give_paracetamol_other */
export function rel_s_not_give_paracetamol_other(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.v(app, "s_not_give_paracetamol"), "other"));
}

/** Relevance of n_paracetamol_shortage */
export function rel_n_paracetamol_shortage(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.v(app, "s_not_give_paracetamol"), "out_of_stock"));
}

/** Relevance of s_close_out_note */
export function rel_s_close_out_note(app) {
  return grp_group_diagnosis(app) && X.bool(X.ne(X.v(app, "s_why_close_out"), ""));
}

/** Relevance of r_next_visit */
export function rel_r_next_visit(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.v(app, "close_out"), "false"));
}

/** Relevance of s_ds_ref_danger_sign */
export function rel_s_ds_ref_danger_sign(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.v(app, "close_out"), "false"));
}

/** Relevance of s_new_referral_cycle */
export function rel_s_new_referral_cycle(app) {
  return grp_group_diagnosis(app) && X.bool(X.eq(X.v(app, "s_accompany_refer_CSCOM"), "yes"));
}

/** Calculation of paracetamol_500_disp */
export function calc_paracetamol_500_disp(app) {
  return X.out((X.bool(X.ne(X.v(app, "s_give_paracetamol_dosage"), "")) ? (X.bool(X.eq(X.v(app, "s_give_paracetamol_dosage"), "quarter_tab_paracetamol_once")) ? (X.num(1) / X.num(4)) : (X.num(1) / X.num(2))) : (X.bool(X.ne(X.v(app, "s_give_paracetamol_dosage"), "")) ? (X.bool(X.eq(X.v(app, "s_give_paracetamol_dosage"), "quarter_tab_paracetamol_3x_a_day")) ? (X.num((X.num(1) / X.num(4))) * X.num(3)) : (X.num((X.num(1) / X.num(2))) * X.num(3))) : 0)));
}

/** Relevance of stock_monitoring_reported_today */
export function rel_stock_monitoring_reported_today(app) {
  return grp_prescription_summary(app);
}

/** Calculation of stock_monitoring_reported_today */
export function calc_stock_monitoring_reported_today(app) {
  if (!rel_stock_monitoring_reported_today(app)) return '';
  return X.out(X.v(app, "s_followup_today"));
}

/** Relevance of stock_monitoring_reported_date */
export function rel_stock_monitoring_reported_date(app) {
  return grp_prescription_summary(app);
}

/** Calculation of stock_monitoring_reported_date */
export function calc_stock_monitoring_reported_date(app) {
  if (!rel_stock_monitoring_reported_date(app)) return '';
  return X.out(X.v(app, "followup_date"));
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
  return X.out("referral_followup_under_5");
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

/** Task input t_follow_up_count */
export function calc_inputs_t_follow_up_count(app) {
  return X.out(X.inp(app, "t_follow_up_count"));
}

/** Task input t_tb_case_category */
export function calc_inputs_t_tb_case_category(app) {
  return X.out(X.inp(app, "t_tb_case_category"));
}

/** Task input t_delivery_date */
export function calc_inputs_t_delivery_date(app) {
  return X.out(X.inp(app, "t_delivery_date"));
}
