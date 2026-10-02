// Form tb_test_result_chw: ported from forms/app/tb_test_result_chw.xlsx of the CHT configuration muso-mali.
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
 "c_task_to_perform": {
  "follow_up": "Suivi",
  "close_out": "Clôture"
 },
 "c_why_close_out": {
  "patient_died": "Le patient est décédé",
  "transfer": "Transfert du patient",
  "refusal_of_parent": "Refus des parents pour divers raisons",
  "loss_of_sight": "Perte de vue",
  "not_evaluated": "Non évalué",
  "hospital": "Patient hospitalisé"
 },
 "yes_no": {
  "yes": "Oui",
  "no": "Non"
 },
 "treatment_action": {
  "chemoprevention": "Chimio-prévention",
  "test_tb": "Tester pour TB",
  "clinical_followup": "Suivi clinique simple"
 },
 "no_chemoprev_reason": {
  "stock_out": "Rupture",
  "another_appointment_made": "Un autre rendez vous à été donné au cas contact",
  "other": "Autre"
 },
 "tb_result": {
  "pos": "Positif",
  "neg": "Négatif"
 },
 "no_treatment_reason": {
  "stock_out": "Rupture",
  "other": "Autre"
 },
 "contraception_method": {
  "iud": "DIU (Stérilet)",
  "implants": "Implants",
  "sterilization_f": "Stérilisation féminine",
  "simple_pill": "Pilule simple",
  "mini_pill": "Pilule minidosée",
  "e_iud": "DIU d'urgence",
  "condom_m": "Préservatif masculin",
  "lam": "Méthode MAMA",
  "condom_f": "Préservatif Féminin",
  "injectable": "Injectable",
  "collier": "Collier",
  "sayana_press": "Sayana Press",
  "morning_pill": "Pillule du lendemain"
 }
};

/** CHT properties.json: not offered in the add menu, opened by its task only. */
export function showForm(app) {
  return false;
}

/** Group g_tdo_observer */
export function grp_g_tdo_observer(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_treatment_action"), "chemoprevention")) || X.bool(X.eq(X.v(app, "tb_test_result"), "pos"))));
}

/** Group g_contraceptive_history */
export function grp_g_contraceptive_history(app) {
  return X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "tb_test_result"), "pos")) && X.bool(X.eq(X.v(app, "patient_sex_normalized"), "female")))) && X.bool(X.ge(X.v(app, "patient_age_in_years"), 10)))) && X.bool(X.le(X.v(app, "patient_age_in_years"), 65))));
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
  return X.out(X.f["concat"](app, X.f["concat"](app, X.v(app, "patient_age_in_years"), (X.bool(X.eq(X.v(app, "patient_age_in_years"), 1)) ? " an et " : " ans et ")), X.f["concat"](app, (X.num(X.v(app, "patient_age_in_months")) % X.num(12)), " mois")));
}

/** Calculation of patient_age_display_bm */
export function calc_patient_age_display_bm(app) {
  return X.out(X.f["concat"](app, X.f["concat"](app, "San ", X.v(app, "patient_age_in_years"), " ani "), X.f["concat"](app, " Kalo ", (X.num(X.v(app, "patient_age_in_months")) % X.num(12)))));
}

/** Calculation of patient_sex */
export function calc_patient_sex(app) {
  return X.out(X.inp(app, "contact/sex"));
}

/** Calculation of patient_sex_normalized */
export function calc_patient_sex_normalized(app) {
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_sex"), "male")) || X.bool(X.eq(X.v(app, "patient_sex"), "homme")))) || X.bool(X.eq(X.v(app, "patient_sex"), "Homme")))) ? "male" : (X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_sex"), "female")) || X.bool(X.eq(X.v(app, "patient_sex"), "femme")))) || X.bool(X.eq(X.v(app, "patient_sex"), "Femme")))) ? "female" : "")));
}

/** Calculation of patient_sex_fr */
export function calc_patient_sex_fr(app) {
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_sex"), "male")) || X.bool(X.eq(X.v(app, "patient_sex"), "homme")))) || X.bool(X.eq(X.v(app, "patient_sex"), "Homme")))) ? "Homme" : (X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_sex"), "female")) || X.bool(X.eq(X.v(app, "patient_sex"), "femme")))) || X.bool(X.eq(X.v(app, "patient_sex"), "Femme")))) ? "Femme" : "")));
}

/** Calculation of patient_sex_bm */
export function calc_patient_sex_bm(app) {
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_sex"), "male")) || X.bool(X.eq(X.v(app, "patient_sex"), "homme")))) || X.bool(X.eq(X.v(app, "patient_sex"), "Homme")))) ? "Cɛ" : (X.bool((X.bool((X.bool(X.eq(X.v(app, "patient_sex"), "female")) || X.bool(X.eq(X.v(app, "patient_sex"), "femme")))) || X.bool(X.eq(X.v(app, "patient_sex"), "Femme")))) ? "Muso" : "")));
}

/** Calculation of patient_id */
export function calc_patient_id(app) {
  return X.out(X.inp(app, "contact/_id"));
}

/** Calculation of muso_id */
export function calc_muso_id(app) {
  return X.out(X.inp(app, "t_muso_id"));
}

/** Calculation of patient_name */
export function calc_patient_name(app) {
  return X.out(X.inp(app, "contact/name"));
}

/** Calculation of patient_phone */
export function calc_patient_phone(app) {
  return X.out(X.inp(app, "contact/phone"));
}

/** Calculation of family_name */
export function calc_family_name(app) {
  return X.out(X.inp(app, "contact/parent/name"));
}

/** Calculation of tb_case_category */
export function calc_tb_case_category(app) {
  return X.out(X.inp(app, "t_tb_case_category"));
}

/** Calculation of tb_assessment_date */
export function calc_tb_assessment_date(app) {
  return X.out(X.inp(app, "t_tb_assessment_date"));
}

/** Calculation of tb_assessment_date_label */
export function calc_tb_assessment_date_label(app) {
  return X.out(X.f["format-date-time"](app, X.f["date-time"](app, X.f["decimal-date-time"](app, X.inp(app, "t_tb_assessment_date"))), "%d/%m/%Y"));
}

/** Calculation of tb_test_result */
export function calc_tb_test_result(app) {
  return X.out(X.v(app, "s_tb_test_result"));
}

/** Calculation of tb_test_result_display */
export function calc_tb_test_result_display(app) {
  return X.out((X.bool(X.eq(X.v(app, "tb_test_result"), "pos")) ? "Positive" : (X.bool(X.eq(X.v(app, "tb_test_result"), "neg")) ? "Negative" : "No result")));
}

/** Calculation of tb_test_result_display_fr */
export function calc_tb_test_result_display_fr(app) {
  return X.out((X.bool(X.eq(X.v(app, "tb_test_result"), "pos")) ? "Positif" : (X.bool(X.eq(X.v(app, "tb_test_result"), "neg")) ? "Négatif" : "Pas de résultat")));
}

/** Calculation of tb_test_result_display_bm */
export function calc_tb_test_result_display_bm(app) {
  return X.out((X.bool(X.eq(X.v(app, "tb_test_result"), "pos")) ? "Positif" : (X.bool(X.eq(X.v(app, "tb_test_result"), "neg")) ? "Négatif" : "Pas de résultat")));
}

/** Calculation of s_tb_test_result_display */
export function calc_s_tb_test_result_display(app) {
  return X.out(X.choiceName("tb_result", X.v(app, "s_tb_test_result")));
}

/** Calculation of needs_signoff */
export function calc_needs_signoff(app) {
  return X.out("true");
}

/** Calculation of treatment_duration_in_days */
export function calc_treatment_duration_in_days(app) {
  return X.out((X.bool((X.bool(X.ne(X.v(app, "s_treatment_start_date"), "")) && X.bool(X.ne(X.v(app, "s_treatment_start_date"), "0")))) ? X.f["floor"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) - X.num(X.f["decimal-date-time"](app, X.v(app, "s_treatment_start_date"))))) : 0));
}

/** Calculation of cscom_area */
export function calc_cscom_area(app) {
  return X.out(X.inp(app, "t_cscom_area"));
}

/** Calculation of chw_name */
export function calc_chw_name(app) {
  return X.out(X.inp(app, "contact/parent/parent/contact/name"));
}

/** Calculation of chw_muso_id */
export function calc_chw_muso_id(app) {
  return X.out(X.inp(app, "contact/parent/parent/contact/patient_id"));
}

/** Calculation of starting_form_uuid */
export function calc_starting_form_uuid(app) {
  return X.out(X.inp(app, "t_starting_form_uuid"));
}

/** Calculation of starting_form_name */
export function calc_starting_form_name(app) {
  return X.out(X.inp(app, "t_starting_form_name"));
}

/** Relevance of s_have_ppe */
export function rel_s_have_ppe(app) {
  return X.bool(X.eq(X.v(app, "s_task_to_perform"), "follow_up"));
}

/** Relevance of n_epi_no_touch_note */
export function rel_n_epi_no_touch_note(app) {
  return X.bool(X.eq(X.v(app, "s_have_ppe"), "no"));
}

/** Relevance of n_epi_reminder */
export function rel_n_epi_reminder(app) {
  return X.bool(X.eq(X.v(app, "s_have_ppe"), "yes"));
}

/** Relevance of n_wash_hands */
export function rel_n_wash_hands(app) {
  return X.bool(X.eq(X.v(app, "s_have_ppe"), "yes"));
}

/** Relevance of n_hygiene_note */
export function rel_n_hygiene_note(app) {
  return X.bool(X.eq(X.v(app, "s_have_ppe"), "yes"));
}

/** Relevance of s_followup_date */
export function rel_s_followup_date(app) {
  return X.bool(X.eq(X.v(app, "s_followup_today"), "no"));
}

/** Constraint of s_followup_date */
export function val_s_followup_date(app) {
  if (X.str(X.v(app, "s_followup_date")) === '') return true;
  return X.bool(X.lt(X.v(app, "s_followup_date"), X.f["today"](app)));
}

/** Relevance of s_why_close_out */
export function rel_s_why_close_out(app) {
  return X.bool(X.eq(X.v(app, "s_task_to_perform"), "close_out"));
}

/** Relevance of n_death_note */
export function rel_n_death_note(app) {
  return X.bool(X.eq(X.v(app, "s_why_close_out"), "patient_died"));
}

/** Options of s_treatment_action */
export function filt_s_treatment_action(app) {
  return X.bool((X.bool(X.eq(X.opt(app, "filter1"), X.v(app, "tb_case_category"))) || X.bool(X.eq(X.opt(app, "filter1"), "all"))));
}

/** Relevance of s_contact_case_received_chemoprev */
export function rel_s_contact_case_received_chemoprev(app) {
  return X.bool(X.eq(X.v(app, "s_treatment_action"), "chemoprevention"));
}

/** Relevance of s_no_chemoprev_reason */
export function rel_s_no_chemoprev_reason(app) {
  return X.bool(X.eq(X.v(app, "s_contact_case_received_chemoprev"), "no"));
}

/** Relevance of s_no_chemoprev_reason_other */
export function rel_s_no_chemoprev_reason_other(app) {
  return X.bool(X.eq(X.v(app, "s_no_chemoprev_reason"), "other"));
}

/** Relevance of s_chemoprev_start_date */
export function rel_s_chemoprev_start_date(app) {
  return X.bool(X.eq(X.v(app, "s_contact_case_received_chemoprev"), "yes"));
}

/** Constraint of s_chemoprev_start_date */
export function val_s_chemoprev_start_date(app) {
  if (X.str(X.v(app, "s_chemoprev_start_date")) === '') return true;
  return X.bool((X.bool(X.le(X.f["floor"](app, X.f["decimal-date-time"](app, X.v(app, "s_chemoprev_start_date"))), X.f["floor"](app, X.f["decimal-date-time"](app, X.f["today"](app))))) && X.bool(X.ge(X.f["floor"](app, X.f["decimal-date-time"](app, X.v(app, "s_chemoprev_start_date"))), X.f["floor"](app, X.f["decimal-date-time"](app, X.v(app, "tb_assessment_date")))))));
}

/** Relevance of s_tb_test_result */
export function rel_s_tb_test_result(app) {
  return X.bool(X.eq(X.v(app, "s_treatment_action"), "test_tb"));
}

/** Relevance of s_cdt_registry_number */
export function rel_s_cdt_registry_number(app) {
  return X.bool(X.eq(X.v(app, "s_treatment_action"), "test_tb"));
}

/** Relevance of s_received_treatment */
export function rel_s_received_treatment(app) {
  return X.bool(X.eq(X.v(app, "tb_test_result"), "pos"));
}

/** Relevance of s_no_treatment_reason */
export function rel_s_no_treatment_reason(app) {
  return X.bool(X.eq(X.v(app, "s_received_treatment"), "no"));
}

/** Relevance of s_no_treatment_reason_other */
export function rel_s_no_treatment_reason_other(app) {
  return X.bool(X.eq(X.v(app, "s_no_treatment_reason"), "other"));
}

/** Relevance of s_treatment_start_date */
export function rel_s_treatment_start_date(app) {
  return X.bool(X.eq(X.v(app, "s_received_treatment"), "yes"));
}

/** Constraint of s_treatment_start_date */
export function val_s_treatment_start_date(app) {
  if (X.str(X.v(app, "s_treatment_start_date")) === '') return true;
  return X.bool((X.bool(X.le(X.f["floor"](app, X.f["decimal-date-time"](app, X.v(app, "s_treatment_start_date"))), X.f["floor"](app, X.f["decimal-date-time"](app, X.f["today"](app))))) && X.bool(X.ge(X.f["floor"](app, X.f["decimal-date-time"](app, X.v(app, "s_treatment_start_date"))), X.f["floor"](app, X.f["decimal-date-time"](app, X.v(app, "tb_assessment_date")))))));
}

/** Relevance of n_observer */
export function rel_n_observer(app) {
  return grp_g_tdo_observer(app);
}

/** Relevance of s_observer_name */
export function rel_s_observer_name(app) {
  return grp_g_tdo_observer(app);
}

/** Relevance of s_observer_phone */
export function rel_s_observer_phone(app) {
  return grp_g_tdo_observer(app);
}

/** Constraint of s_observer_phone */
export function val_s_observer_phone(app) {
  if (X.str(X.v(app, "s_observer_phone")) === '') return true;
  return X.bool(X.f["regex"](app, X.v(app, "s_observer_phone"), "^(\\d{2}\\s?){4}$"));
}

/** Relevance of s_contraceptive */
export function rel_s_contraceptive(app) {
  return grp_g_contraceptive_history(app);
}

/** Relevance of s_contraceptive_method */
export function rel_s_contraceptive_method(app) {
  return grp_g_contraceptive_history(app) && X.bool(X.eq(X.v(app, "s_contraceptive"), "yes"));
}

/** Relevance of n_summary_result */
export function rel_n_summary_result(app) {
  return X.bool(X.ne(X.v(app, "tb_test_result"), ""));
}

/** Relevance of n_chemoprev_instruction */
export function rel_n_chemoprev_instruction(app) {
  return X.bool(X.eq(X.v(app, "s_treatment_action"), "chemoprevention"));
}

/** Relevance of n_negative_instruction */
export function rel_n_negative_instruction(app) {
  return X.bool(X.eq(X.v(app, "tb_test_result"), "neg"));
}

/** Relevance of n_no_treatment_instruction */
export function rel_n_no_treatment_instruction(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_received_treatment"), "no")) || X.bool(X.eq(X.v(app, "s_contact_case_received_chemoprev"), "no"))));
}

/** Relevance of n_refer_cscom */
export function rel_n_refer_cscom(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_received_treatment"), "no")) || X.bool(X.eq(X.v(app, "s_contact_case_received_chemoprev"), "no"))));
}

/** Relevance of n_contact_followup_note */
export function rel_n_contact_followup_note(app) {
  return X.bool(X.eq(X.v(app, "tb_test_result"), "pos"));
}

/** Relevance of n_contraception_counseling */
export function rel_n_contraception_counseling(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_contraceptive"), "yes")) && X.bool((X.bool(X.eq(X.v(app, "s_contraceptive_method"), "simple_pill")) || X.bool(X.eq(X.v(app, "s_contraceptive_method"), "mini_pill"))))));
}

/** Relevance of n_refer_cscom2 */
export function rel_n_refer_cscom2(app) {
  return X.bool(X.eq(X.v(app, "s_received_treatment"), "no"));
}

/** Relevance of s_accompany_refer_cscom */
export function rel_s_accompany_refer_cscom(app) {
  return X.bool(X.eq(X.v(app, "s_received_treatment"), "no"));
}

/** Relevance of n_dayly_followup_note */
export function rel_n_dayly_followup_note(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_received_treatment"), "yes")) && X.bool(X.lt(X.v(app, "treatment_duration_in_days"), 14))));
}

/** Relevance of n_chemoprev_followup_note */
export function rel_n_chemoprev_followup_note(app) {
  return X.bool(X.eq(X.v(app, "s_treatment_action"), "chemoprevention"));
}

/** Relevance of n_weekly_followup_note */
export function rel_n_weekly_followup_note(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_received_treatment"), "yes")) && X.bool(X.ge(X.v(app, "treatment_duration_in_days"), 14))));
}

/** Relevance of n_no_followup_note */
export function rel_n_no_followup_note(app) {
  return X.bool((X.bool(X.eq(X.v(app, "tb_test_result"), "neg")) || X.bool(X.eq(X.v(app, "s_treatment_action"), "clinical_followup"))));
}

/** Relevance of n_closure_note */
export function rel_n_closure_note(app) {
  return X.bool(X.eq(X.v(app, "s_task_to_perform"), "close_out"));
}

/** Relevance of n_cscom_treatment */
export function rel_n_cscom_treatment(app) {
  return X.bool(X.eq(X.v(app, "s_received_treatment"), "yes"));
}

/** Task input source_id */
export function calc_inputs_source_id(app) {
  return X.out(X.inp(app, "source_id"));
}

/** Task input t_cscom_area */
export function calc_inputs_t_cscom_area(app) {
  return X.out(X.inp(app, "t_cscom_area"));
}

/** Task input t_tb_diagnosis_en */
export function calc_inputs_t_tb_diagnosis_en(app) {
  return X.out(X.inp(app, "t_tb_diagnosis_en"));
}

/** Task input t_tb_diagnosis_fr */
export function calc_inputs_t_tb_diagnosis_fr(app) {
  return X.out(X.inp(app, "t_tb_diagnosis_fr"));
}

/** Task input t_tb_diagnosis_bm */
export function calc_inputs_t_tb_diagnosis_bm(app) {
  return X.out(X.inp(app, "t_tb_diagnosis_bm"));
}

/** Task input t_tb_case_category */
export function calc_inputs_t_tb_case_category(app) {
  return X.out(X.inp(app, "t_tb_case_category"));
}

/** Task input t_tb_assessment_date */
export function calc_inputs_t_tb_assessment_date(app) {
  return X.out(X.inp(app, "t_tb_assessment_date"));
}

/** Task input t_muso_id */
export function calc_inputs_t_muso_id(app) {
  return X.out(X.inp(app, "t_muso_id"));
}

/** Task input t_starting_form_uuid */
export function calc_inputs_t_starting_form_uuid(app) {
  return X.out(X.inp(app, "t_starting_form_uuid"));
}

/** Task input t_starting_form_name */
export function calc_inputs_t_starting_form_name(app) {
  return X.out(X.inp(app, "t_starting_form_name"));
}
