// Form patient_assessment: ported from forms/app/patient_assessment.xlsx of the CHT configuration muso-mali.
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
  "accompanied_by_parents": "Enfant accompagné par ses parents chez l’ASC",
  "accompanied_by_community_member": "Enfant et ses parents accompagnés chez l’ASC par un autre membre de la communauté",
  "other": "Autre"
 },
 "c_assessment_time": {
  "c_assessment_time_1": "Matin",
  "c_assessment_time_2": "Midi",
  "c_assessment_time_3": "Après-midi",
  "c_assessment_time_4": "Soir"
 },
 "c_when_illness": {
  "c_when_illness_1": "Aujourd’hui",
  "c_when_illness_2": "Au cours de la nuit",
  "c_when_illness_3": "Hier matin",
  "c_when_illness_4": "Hier soir",
  "c_when_illness_5": "Avant-hier matin",
  "c_when_illness_6": "Avant-hier soir",
  "c_when_illness_7": "Il y'a trois (3) jours",
  "c_when_illness_8": "Il y'a plus de trois (3) jours"
 },
 "CAT_covid19": {
  "CAT_covid19_1": "Informer l’équipe Covid-19",
  "CAT_covid19_2": "Mettre en observation",
  "CAT_covid19_3": "Informer le CSCOM"
 },
 "c_acc_child_danger_signs": {
  "seizure": "Convulsions",
  "loss_of_consciousness": "Perte de conscience",
  "unable_to_drink": "Impossibilité de manger ou de boire",
  "disorientation_confusion": "Confusion ou Désorientation",
  "vomit_everything": "Vomit tout qu’il avale",
  "chest_indrawing": "Tirage sous-costal",
  "wheezing": "Geignements",
  "bleeding_that_does_not_stop": "Foo",
  "lethargy": "Léthargie (bouge seulement quand il est stimulé)"
 },
 "c_tdr_not_done": {
  "tdr_unavailable": "TDR non disponible",
  "other": "Autre"
 },
 "medication_not_given": {
  "out_of_stock": "Rupture de stock",
  "other": "Autre"
 },
 "c_ref_child_danger_signs": {
  "severe_jaundice": "Ictère sévère ( jaunisse corps/ plantes des pieds)",
  "dark_colored_urine": "Urines foncées ( couleur Coca-cola)",
  "newborn_sick": "L'enfant est malade selon sa maman",
  "conjunctival_pallor": "Pâleur conjonctivo-palmo-plantaire",
  "red_shakir_strip": "Bande de Shakir au rouge ou amaigrissement visible",
  "not_gaining_weight": "Notion de perte de poids pour les enfants de moins de 6 mois",
  "eodema_lower_legs": "Oedème des membres inférieurs",
  "diarrhea_more_than_14_days": "Diarrhée qui dure plus de 14 jours",
  "other_diseases_newborn": "Autres maladies rencontrées chez le nouveau né",
  "cough_more_than_14_days": "Toux qui dure plus de 14 jours sans amélioration",
  "blood_in_stools": "Présence de sang dans les selles"
 },
 "c_malaria_signs": {
  "fever": "Notion de fièvre dans les 24-48 heures",
  "vomiting": "Vomissment",
  "sluggish": "Amorphe (perte de dynamisme)",
  "chills": "Frissons",
  "loss_of_appetite": "Perte d’appetit"
 },
 "c_malaria_act_dosage": {
  "2_tablets_act": "1 comprimé 2 fois par jour",
  "4_tablets_act": "2 comprimés 2 fois par jour"
 },
 "c_diarrhea_zinc_dosage": {
  "c_diarrhea_zinc_dosage_2": "Un demi (1/2) comprimé de 20 mg pendant 10 jours",
  "c_diarrhea_zinc_dosage_3": "Un (1) comprimé de 20 mg pendant 10 jours"
 },
 "c_diarrhea_ors_dosage": {
  "1_packet_ors": "1 sachet",
  "2_packets_ors": "2 sachets",
  "3_packets_ors": "3 sachets",
  "4_packets_ors": "4 sachets"
 },
 "c_ari_give_amox_dosage": {
  "5_ml_amox": "5 ml 2 fois par jour",
  "10_ml_amox": "10 ml 2 fois par jour"
 },
 "c_nutri_color_shakir": {
  "Green": "Vert",
  "Yellow": "Jaune",
  "Red": "Rouge"
 },
 "c_nutri_give_enriched_flour_dosage": {
  "1_packet_flour": "1 sachet",
  "2_packets_flour": "2 sachets",
  "3_packets_flour": "3 sachets",
  "4_packets_flour": "4 sachets"
 },
 "c_nutri_give_vitaminA_dosage": {
  "blue_capsule_vita": "Une (1) capsule bleue (100 000 UI)",
  "red_capsule_vita": "Une (1) capsule rouge (200 000 UI)"
 },
 "c_nutri_give_albendazole_type": {
  "200_mg_albendazole": "Albendazole 200 mg",
  "400_mg_albendazole": "Albendazole 400 mg"
 },
 "c_nutri_give_albendazole_dosage": {
  "half_tab_albendazole": "½ comprimé",
  "1_tab_albendazole": "1 comprimé",
  "2_tab_albendazole": "2 comprimés"
 },
 "c_give_paracetamol_dosage": {
  "quarter_tab_paracetamol_once": "¼ comprimé",
  "half_tab_paracetamol_once": "½ comprimé",
  "quarter_tab_paracetamol_3x_a_day": "¼ comprimé 3 fois par jour",
  "half_tab_paracetamol_3x_a_day": "½ comprimé 3 fois par jour"
 },
 "CAT": {
  "refer": "Référer au CSCOM",
  "watching": "Mettre en observation"
 },
 "c_not_use_shakir_strip": {
  "out_of_stock": "Rupture de stock",
  "already_taken_in_charge": "Enfant est déjà pris en charge pour malnutrition",
  "other": "Autre"
 },
 "c_malnutrition_appetite": {
  "good": "Bon",
  "meduim": "Moyen",
  "low": "Faible"
 },
 "c_first_malnutrition_screening": {
  "chw_during_assessment": "ASC lors de l'évaluation du malade",
  "chw_during_screening": "ASC lors du dépistage systématique",
  "mother_other_member": "Maman ou autre membre de la famille",
  "mass_screening_community": "Dépistage de masse dans la communauté",
  "cscom": "CSCOM",
  "other": "Autre"
 },
 "CAT_out_of_stock": {
  "refer": "Référence",
  "continue_treatment": "Poursuivre le traitement à domicile"
 },
 "c_albendazole_dosage": {
  "one_of_200mg": "1 comprimé de 200 mg",
  "one_of_400mg": "1 comprimé de 400 mg"
 },
 "c_vitamina_dosage": {
  "one_of_100000_ui_blue": "1 capsule de 100 000 UI (1 capsule bleu)",
  "one_of_200000_ui_red": "1 capsule de 200 000 UI (1 capsule rouge)"
 },
 "c_antimalarial_suppository_type": {
  "artesunate_suppository": "Artésunate suppo",
  "arthemeter_suppository": "Arthéméter suppo"
 },
 "c_artesunate_suppository_type": {
  "artesunate_suppository_50mg": "Artésunate suppo 50 mg",
  "artesunate_suppository_100mg": "Artésunate suppo 100 mg",
  "artesunate_suppository_200mg": "Artésunate suppo 200 mg"
 },
 "c_give_artesunate_dosage": {
  "1_artesunate_50mg": "1 suppositoire de 50 mg",
  "2_artesunate_50mg": "2 suppositoires de 50 mg",
  "4_artesunate_50mg": "4 suppositoires de 50 mg",
  "1_artesunate_100mg": "1 suppositoire de 100 mg",
  "2_artesunate_100mg": "2 suppositoires de 100 mg",
  "4_artesunate_100mg": "4 suppositoires de 100 mg",
  "1_artesunate_200mg": "1 suppositoire de 200 mg",
  "2_artesunate_200mg": "2 suppositoires de 200 mg",
  "4_artesunate_200mg": "4 suppositoires de 200 mg"
 },
 "c_antimalarial_not_given": {
  "child_has_diarrhea": "L'enfant fait de la diarrhée",
  "other": "Autre"
 },
 "c_give_artemether_dosage": {
  "1_artemether_40mg": "1 suppositoire de 40 mg",
  "2_artemether_40mg": "2 suppositoires de 40 mg"
 }
};

const isAlive = (person) => !(person?.attributes?.person_is_deceased === true || person?.attributes?.person_is_deceased === 'true');
const ageInYears = (person, now) => {
  const dob = person?.attributes?.birthdate;
  if (!dob) return null;
  return (now - Date.parse(dob)) / (365.25 * 86400000);
};

/**
 * CHT properties.json: a living patient under 5, for a CHW (role chw_uhc, assigned to a
 * c40_chw_area). summary.muted has no equivalent (see ECARTS_CHT_COLLECTON.md).
 */
export function showForm(app) {
  const person = app.person;
  if (!person || person.personType !== 'patient' || !isAlive(person)) return false;
  if (app.currentUser?.role?.name !== 'chw_uhc') return false;
  const age = ageInYears(person, app.now);
  return age === null || age < 5;
}

/** Group s_temperature_retake */
export function grp_s_temperature_retake(app) {
  return X.bool((X.bool(X.lt(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw"), 0)), 35)) || X.bool(X.gt(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw"), 0)), 41))));
}

/** Group s_malaria_indication_1 */
export function grp_s_malaria_indication_1(app) {
  return X.bool(X.eq(X.v(app, "s_acc_tdr_done"), "yes"));
}

/** Group s_malaria_indication_2 */
export function grp_s_malaria_indication_2(app) {
  return X.bool(X.eq(X.v(app, "s_acc_tdr_done"), "yes"));
}

/** Group s_malaria_indication_3 */
export function grp_s_malaria_indication_3(app) {
  return X.bool(X.eq(X.v(app, "s_acc_tdr_done"), "yes"));
}

/** Group s_malaria_indication_4 */
export function grp_s_malaria_indication_4(app) {
  return X.bool(X.eq(X.v(app, "s_acc_tdr_done"), "yes"));
}

/** Group s_malaria_indication_5 */
export function grp_s_malaria_indication_5(app) {
  return X.bool(X.eq(X.v(app, "s_acc_tdr_done"), "yes"));
}

/** Group s_malaria_indication_6 */
export function grp_s_malaria_indication_6(app) {
  return X.bool(X.eq(X.v(app, "s_acc_tdr_done"), "yes"));
}

/** Group s_malaria_indication_7 */
export function grp_s_malaria_indication_7(app) {
  return X.bool(X.eq(X.v(app, "s_acc_tdr_done"), "yes"));
}

/** Group s_malaria_indication_8 */
export function grp_s_malaria_indication_8(app) {
  return X.bool(X.eq(X.v(app, "s_acc_tdr_done"), "yes"));
}

/** Group s_malaria_indication_9 */
export function grp_s_malaria_indication_9(app) {
  return X.bool(X.eq(X.v(app, "s_acc_tdr_done"), "yes"));
}

/** Group s_ref_danger_signs */
export function grp_s_ref_danger_signs(app) {
  return X.bool(X.eq(X.v(app, "has_danger_sign"), "false"));
}

/** Group s_ref_other_diseases */
export function grp_s_ref_other_diseases(app) {
  return X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "false")) && X.bool(X.eq(X.v(app, "has_first_ref_danger_sign"), "false"))));
}

/** Group s_malnutrition */
export function grp_s_malnutrition(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "false")) && X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "false")))) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 6))));
}

/** Group s_malnutrition_observation */
export function grp_s_malnutrition_observation(app) {
  return X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "false")) && X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "false")))) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 6)))) && X.bool((X.bool((X.bool(X.eq(X.v(app, "s_use_shakir_strip"), "yes")) || X.bool(X.eq(X.v(app, "s_not_use_shakir_strip"), "other")))) || X.bool(X.eq(X.v(app, "s_not_use_shakir_strip"), "out_of_stock"))))));
}

/** Group s_malnutrition_weight */
export function grp_s_malnutrition_weight(app) {
  return X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true"));
}

/** Group s_malnutrition_mRDT */
export function grp_s_malnutrition_mRDT(app) {
  return X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "false")) && X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "false")))) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 6)))) && X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true"))));
}

/** Group s_malaria_indication_10 */
export function grp_s_malaria_indication_10(app) {
  return X.bool(X.eq(X.v(app, "s_malnutrition_tdr_done"), "yes"));
}

/** Group s_malaria_indication_11 */
export function grp_s_malaria_indication_11(app) {
  return X.bool(X.eq(X.v(app, "s_malnutrition_tdr_done"), "yes"));
}

/** Group s_malaria_indication_12 */
export function grp_s_malaria_indication_12(app) {
  return X.bool(X.eq(X.v(app, "s_malnutrition_tdr_done"), "yes"));
}

/** Group s_malaria_indication_13 */
export function grp_s_malaria_indication_13(app) {
  return X.bool(X.eq(X.v(app, "s_malnutrition_tdr_done"), "yes"));
}

/** Group s_malaria_indication_14 */
export function grp_s_malaria_indication_14(app) {
  return X.bool(X.eq(X.v(app, "s_malnutrition_tdr_done"), "yes"));
}

/** Group s_malaria_indication_15 */
export function grp_s_malaria_indication_15(app) {
  return X.bool(X.eq(X.v(app, "s_malnutrition_tdr_done"), "yes"));
}

/** Group s_malaria_indication_16 */
export function grp_s_malaria_indication_16(app) {
  return X.bool(X.eq(X.v(app, "s_malnutrition_tdr_done"), "yes"));
}

/** Group s_malaria_indication_17 */
export function grp_s_malaria_indication_17(app) {
  return X.bool(X.eq(X.v(app, "s_malnutrition_tdr_done"), "yes"));
}

/** Group s_malaria_indication_18 */
export function grp_s_malaria_indication_18(app) {
  return X.bool(X.eq(X.v(app, "s_malnutrition_tdr_done"), "yes"));
}

/** Group s_screening */
export function grp_s_screening(app) {
  return X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "false")) && X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "false")))) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 6)))) && X.bool((X.bool((X.bool(X.eq(X.v(app, "has_MAM"), "true")) || X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")))) || X.bool(X.eq(X.v(app, "has_SAM_with_complication"), "true"))))));
}

/** Group s_malaria */
export function grp_s_malaria(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "false")) && X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "false")))) && X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 6)))) || X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 6)) && X.bool((X.bool(X.eq(X.v(app, "s_not_use_shakir_strip"), "already_taken_in_charge")) || X.bool((X.bool(X.eq(X.v(app, "has_SAM_with_complication"), "false")) && X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "false"))))))))))));
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

/** Group s_diarrhea */
export function grp_s_diarrhea(app) {
  return X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "false")) && X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "false")))) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)))) && X.bool(X.ne(X.v(app, "has_SAM_with_complication"), "true"))));
}

/** Group s_acute_respiratory_infection */
export function grp_s_acute_respiratory_infection(app) {
  return X.bool((X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)) && X.bool(X.ne(X.v(app, "has_SAM_with_complication"), "true")))) && X.bool(X.eq(X.v(app, "has_danger_sign"), "false")))) && X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "false"))));
}

/** Group group_breathing */
export function grp_group_breathing(app) {
  return grp_s_acute_respiratory_infection(app) && X.bool((X.bool((X.bool(X.eq(X.v(app, "s_ari_have_cough"), "yes")) && X.bool(X.eq(X.v(app, "has_danger_sign"), "false")))) && X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "false"))));
}

/** Group group_breathing_frequency */
export function grp_group_breathing_frequency(app) {
  return grp_s_acute_respiratory_infection(app) && X.bool(X.eq(X.v(app, "s_ari_have_cough"), "yes"));
}

/** Group s_CAT */
export function grp_s_CAT(app) {
  return X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "false")) && X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "false")))) && X.bool((X.bool((X.bool(X.eq(X.v(app, "has_mal_danger_sign"), "false")) || X.bool(X.eq(X.v(app, "tdr_result"), "négatif")))) || X.bool(X.eq(X.v(app, "s_malaria_tdr_not_done"), "other")))))) && X.bool(X.eq(X.v(app, "s_diarrhea_stools_a_day"), "no")))) && X.bool(X.eq(X.v(app, "treat_malnutrition"), "false")))) && X.bool(X.eq(X.v(app, "s_ari_have_cough"), "no"))));
}

/** Group referral_and_protocol */
export function grp_referral_and_protocol(app) {
  return X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "true")) || X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "true"))));
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

/** Calculation of chw_area_name */
export function calc_chw_area_name(app) {
  return X.out(X.inp(app, "contact/parent/parent/name"));
}

/** Calculation of accompany_to_cscom */
export function calc_accompany_to_cscom(app) {
  return X.out((X.bool(X.eq(X.v(app, "has_danger_sign"), "true")) ? "true" : "false"));
}

/** Calculation of refer_to_cscom */
export function calc_refer_to_cscom(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "true")) || X.bool(X.eq(X.v(app, "has_SAM_with_complication"), "true")))) || X.bool(X.eq(X.v(app, "new_CAT"), "refer")))) || X.bool(X.eq(X.v(app, "MAM_MAS_has_diarrhea_pneumonia"), "true")))) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false")))) ? "true" : "false"));
}

/** Calculation of tdr_done */
export function calc_tdr_done(app) {
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_acc_tdr_done"), "yes")) || X.bool(X.eq(X.v(app, "s_malaria_tdr_done"), "yes")))) || X.bool(X.eq(X.v(app, "s_malnutrition_tdr_done"), "yes")))) ? "yes" : "no"));
}

/** Calculation of symptom_fever */
export function calc_symptom_fever(app) {
  return X.out((X.bool(X.ge(X.v(app, "s_child_temperature"), 38)) ? X.f["concat"](app, "Fever that began: ", X.choiceName("c_when_illness", X.v(app, "s_when_illness"))) : ""));
}

/** Calculation of tdr_result */
export function calc_tdr_result(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_malaria_tdr_result_new"), "pos")) || X.bool(X.eq(X.v(app, "s_malaria_tdr_result"), "pos")))) || X.bool(X.eq(X.v(app, "s_acc_tdr_result"), "pos")))) || X.bool(X.eq(X.v(app, "s_acc_tdr_result_new"), "pos")))) || X.bool(X.eq(X.v(app, "s_malnutrition_tdr_result"), "pos")))) || X.bool(X.eq(X.v(app, "s_malnutrition_tdr_result_new"), "pos")))) ? "positif" : (X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_malaria_tdr_result_new"), "neg")) || X.bool(X.eq(X.v(app, "s_malaria_tdr_result"), "neg")))) || X.bool(X.eq(X.v(app, "s_acc_tdr_result"), "neg")))) || X.bool(X.eq(X.v(app, "s_acc_tdr_result_new"), "neg")))) || X.bool(X.eq(X.v(app, "s_malnutrition_tdr_result"), "neg")))) || X.bool(X.eq(X.v(app, "s_malnutrition_tdr_result_new"), "neg")))) ? "négatif" : "")));
}

/** Calculation of symptom_cough */
export function calc_symptom_cough(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_ari_have_cough"), "yes")) ? X.f["concat"](app, "Cough", X.f["concat"](app, " that began: ", X.choiceName("c_when_illness", X.v(app, "s_when_illness")))) : ""));
}

/** Calculation of treat_malaria */
export function calc_treat_malaria(app) {
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "tdr_result"), "positif")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false")))) && X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")))) ? "true" : "false"));
}

/** Calculation of treat_diarrhea */
export function calc_treat_diarrhea(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_diarrhea_stools_a_day"), "yes")) ? "true" : "false"));
}

/** Calculation of treat_ari */
export function calc_treat_ari(app) {
  return X.out((X.bool((X.bool(X.eq(X.v(app, "fast_breathing"), "true")) && X.bool(X.eq(X.v(app, "s_ari_have_cough"), "yes")))) ? "true" : "false"));
}

/** Calculation of treat_cough */
export function calc_treat_cough(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_ari_have_cough"), "yes")) ? "true" : "false"));
}

/** Calculation of treat_malnutrition */
export function calc_treat_malnutrition(app) {
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_MAM"), "true")) || X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")))) || X.bool(X.eq(X.v(app, "has_SAM_with_complication"), "true")))) ? "true" : "false"));
}

/** Calculation of treat_MAM */
export function calc_treat_MAM(app) {
  return X.out((X.bool(X.eq(X.v(app, "has_MAM"), "true")) ? "true" : "false"));
}

/** Calculation of treat_MAM_with_flour */
export function calc_treat_MAM_with_flour(app) {
  return X.out((X.bool(X.eq(X.v(app, "has_MAM"), "true")) ? "true" : "false"));
}

/** Calculation of treat_MAM_with_plumpy */
export function calc_treat_MAM_with_plumpy(app) {
  return X.out((X.bool(X.eq(X.v(app, "has_MAM"), "true")) ? "true" : "false"));
}

/** Calculation of treat_SAM_without_complication */
export function calc_treat_SAM_without_complication(app) {
  return X.out((X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")) ? "true" : "false"));
}

/** Calculation of next_visit_covid_malaria_mult */
export function calc_next_visit_covid_malaria_mult(app) {
  return X.out((X.bool(X.eq(X.v(app, "diff_assess_date_report_date"), 0)) ? X.f["format-date-time"](app, X.f["date-time"](app, (X.num(X.f["decimal-date-time"](app, X.v(app, "assessment_date"))) + X.num(1))), "%d/%m/%Y") : X.f["format-date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app)), "%d/%m/%Y")));
}

/** Calculation of next_visit_diarrhea */
export function calc_next_visit_diarrhea(app) {
  return X.out((X.bool(X.le(X.v(app, "diff_assess_date_report_date"), 5)) ? X.f["format-date-time"](app, X.f["date-time"](app, (X.num(X.f["decimal-date-time"](app, X.v(app, "assessment_date"))) + X.num(5))), "%d/%m/%Y") : X.f["format-date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app)), "%d/%m/%Y")));
}

/** Calculation of next_visit_ari */
export function calc_next_visit_ari(app) {
  return X.out((X.bool((X.bool(X.eq(X.v(app, "diff_assess_date_report_date"), 0)) || X.bool(X.eq(X.v(app, "diff_assess_date_report_date"), 4)))) ? X.f["format-date-time"](app, X.f["date-time"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) + X.num(1))), "%d/%m/%Y") : X.f["format-date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app)), "%d/%m/%Y")));
}

/** Calculation of next_visit_malnutrition */
export function calc_next_visit_malnutrition(app) {
  return X.out((X.bool(X.lt(X.v(app, "diff_assess_date_report_date"), 7)) ? X.f["format-date-time"](app, X.f["date-time"](app, (X.num(X.f["decimal-date-time"](app, X.v(app, "assessment_date"))) + X.num(7))), "%d/%m/%Y") : (X.bool(X.eq(X.v(app, "diff_assess_date_report_date"), 7)) ? X.f["format-date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app)), "%d/%m/%Y") : (X.bool((X.bool(X.gt(X.v(app, "diff_assess_date_report_date"), 7)) && X.bool(X.lt(X.v(app, "diff_assess_date_report_date"), 14)))) ? X.f["format-date-time"](app, X.f["date-time"](app, (X.num(X.f["decimal-date-time"](app, X.v(app, "assessment_date"))) + X.num(14))), "%d/%m/%Y") : (X.bool(X.eq(X.v(app, "diff_assess_date_report_date"), 14)) ? X.f["format-date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app)), "%d/%m/%Y") : (X.bool((X.bool(X.gt(X.v(app, "diff_assess_date_report_date"), 14)) && X.bool(X.lt(X.v(app, "diff_assess_date_report_date"), 21)))) ? X.f["format-date-time"](app, X.f["date-time"](app, (X.num(X.f["decimal-date-time"](app, X.v(app, "assessment_date"))) + X.num(21))), "%d/%m/%Y") : (X.bool(X.eq(X.v(app, "diff_assess_date_report_date"), 21)) ? X.f["format-date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app)), "%d/%m/%Y") : (X.bool((X.bool(X.gt(X.v(app, "diff_assess_date_report_date"), 21)) && X.bool(X.lt(X.v(app, "diff_assess_date_report_date"), 28)))) ? X.f["format-date-time"](app, X.f["date-time"](app, (X.num(X.f["decimal-date-time"](app, X.v(app, "assessment_date"))) + X.num(28))), "%d/%m/%Y") : (X.bool(X.eq(X.v(app, "diff_assess_date_report_date"), 28)) ? X.f["format-date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app)), "%d/%m/%Y") : (X.bool((X.bool(X.gt(X.v(app, "diff_assess_date_report_date"), 28)) && X.bool(X.lt(X.v(app, "diff_assess_date_report_date"), 42)))) ? X.f["format-date-time"](app, X.f["date-time"](app, (X.num(X.f["decimal-date-time"](app, X.v(app, "assessment_date"))) + X.num(42))), "%d/%m/%Y") : (X.bool(X.eq(X.v(app, "diff_assess_date_report_date"), 42)) ? X.f["format-date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app)), "%d/%m/%Y") : (X.bool((X.bool(X.gt(X.v(app, "diff_assess_date_report_date"), 42)) && X.bool(X.lt(X.v(app, "diff_assess_date_report_date"), 56)))) ? X.f["format-date-time"](app, X.f["date-time"](app, (X.num(X.f["decimal-date-time"](app, X.v(app, "assessment_date"))) + X.num(56))), "%d/%m/%Y") : (X.bool(X.eq(X.v(app, "diff_assess_date_report_date"), 56)) ? X.f["format-date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app)), "%d/%m/%Y") : (X.bool((X.bool(X.gt(X.v(app, "diff_assess_date_report_date"), 56)) && X.bool(X.lt(X.v(app, "diff_assess_date_report_date"), 70)))) ? X.f["format-date-time"](app, X.f["date-time"](app, (X.num(X.f["decimal-date-time"](app, X.v(app, "assessment_date"))) + X.num(70))), "%d/%m/%Y") : (X.bool(X.eq(X.v(app, "diff_assess_date_report_date"), 70)) ? X.f["format-date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app)), "%d/%m/%Y") : (X.bool((X.bool(X.gt(X.v(app, "diff_assess_date_report_date"), 70)) && X.bool(X.lt(X.v(app, "diff_assess_date_report_date"), 84)))) ? X.f["format-date-time"](app, X.f["date-time"](app, (X.num(X.f["decimal-date-time"](app, X.v(app, "assessment_date"))) + X.num(84))), "%d/%m/%Y") : X.f["format-date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app)), "%d/%m/%Y")))))))))))))))));
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

/** Calculation of referral */
export function calc_referral(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")) || X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true")))) || X.bool(X.eq(X.v(app, "has_SAM_with_complication"), "true")))) || X.bool(X.eq(X.v(app, "CAT_out_of_stock"), "refer")))) || X.bool(X.eq(X.v(app, "has_been_accompany_to_cscom"), "true")))) || X.bool(X.eq(X.v(app, "MAM_MAS_has_diarrhea_pneumonia"), "true")))) ? "true" : "false"));
}

/** Calculation of MAM_MAS_has_diarrhea_pneumonia */
export function calc_MAM_MAS_has_diarrhea_pneumonia(app) {
  return X.out((X.bool((X.bool(X.eq(X.v(app, "s_not_use_shakir_strip"), "already_taken_in_charge")) && X.bool((X.bool(X.eq(X.v(app, "treat_diarrhea"), "true")) || X.bool(X.eq(X.v(app, "treat_ari"), "true")))))) ? "true" : "false"));
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
  return X.out((X.bool(X.eq(X.v(app, "has_tb_sign"), "true")) ? X.v(app, "acc_danger_sign_cough_more_than_14_days_fr") : ""));
}

/** Calculation of tb_diagnosis_en */
export function calc_tb_diagnosis_en(app) {
  return X.out((X.bool(X.eq(X.v(app, "has_tb_sign"), "true")) ? X.f["concat"](app, X.v(app, "acc_danger_sign_cough_more_than_14_days_en")) : ""));
}

/** Calculation of tb_diagnosis_bm */
export function calc_tb_diagnosis_bm(app) {
  return X.out((X.bool(X.eq(X.v(app, "has_tb_sign"), "true")) ? X.f["concat"](app, X.v(app, "acc_danger_sign_cough_more_than_14_days_bm")) : ""));
}

/** Calculation of can_access_sih */
export function calc_can_access_sih(app) {
  return X.out(X.inp(app, "user/is_in_sih"));
}

/** Calculation of visited_date */
export function calc_visited_date(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_assess_today"), "yes")) ? X.f["format-date-time"](app, X.f["date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app))), "%Y-%m-%d") : X.v(app, "s_assess_date")));
}

/** Calculation of visited_contact_uuid */
export function calc_visited_contact_uuid(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_how_child_found"), "home_visit")) ? X.inp(app, "contact/parent/_id") : ""));
}

/** Calculation of rc_code */
export function calc_rc_code(app) {
  return X.out((X.bool(X.eq(X.v(app, "has_danger_sign"), "true")) ? "34221" : "34221"));
}

/** Calculation of can_access_mrdt */
export function calc_can_access_mrdt(app) {
  return X.out("true");
}

/** Calculation of shakir */
export function calc_shakir(app) {
  return X.out((X.bool(X.v(app, "s_nutri_color_shakir")) ? X.v(app, "s_nutri_color_shakir_fr") : ""));
}

/** Calculation of assessment_date */
export function calc_assessment_date(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_assess_today"), "yes")) ? X.f["format-date-time"](app, X.f["date-time"](app, X.f["decimal-date-time"](app, X.f["now"](app))), "%Y-%m-%d") : X.v(app, "s_assess_date")));
}

/** Calculation of no_tdr_done */
export function calc_no_tdr_done(app) {
  return X.out((X.bool(X.ne(X.v(app, "s_acc_tdr_not_done"), "")) ? X.v(app, "s_acc_tdr_not_done") : (X.bool(X.ne(X.v(app, "s_malnutrition_tdr_not_done"), "")) ? X.v(app, "s_malnutrition_tdr_not_done") : X.v(app, "s_malaria_tdr_not_done"))));
}

/** Calculation of no_tdr_done_other */
export function calc_no_tdr_done_other(app) {
  return X.out((X.bool(X.ne(X.v(app, "s_acc_tdr_not_done"), "")) ? X.v(app, "s_acc_tdr_not_done_other") : (X.bool(X.ne(X.v(app, "s_malnutrition_tdr_not_done"), "")) ? X.v(app, "s_malnutrition_tdr_not_done_other") : X.v(app, "s_malaria_tdr_not_done_other"))));
}

/** Calculation of no_tdr_done_fr */
export function calc_no_tdr_done_fr(app) {
  return X.out((X.bool(X.eq(X.v(app, "no_tdr_done"), "tdr_unavailable")) ? "Pas de TDR: TDR non disponible" : X.f["concat"](app, "Pas de TDR: ", X.v(app, "no_tdr_done_other"))));
}

/** Calculation of no_tdr_done_en */
export function calc_no_tdr_done_en(app) {
  return X.out((X.bool(X.eq(X.v(app, "no_tdr_done"), "tdr_unavailable")) ? "No RDT: RDT unavailable" : X.f["concat"](app, "No RDT: ", X.v(app, "no_tdr_done_other"))));
}

/** Calculation of no_tdr_done_bm */
export function calc_no_tdr_done_bm(app) {
  return X.out((X.bool(X.eq(X.v(app, "no_tdr_done"), "tdr_unavailable")) ? "TDR ma kɛ: TDR ban nen do" : X.f["concat"](app, "TDR ma kɛ: ", X.v(app, "no_tdr_done_other"))));
}

/** Calculation of diff_assess_date_report_date */
export function calc_diff_assess_date_report_date(app) {
  return X.out(X.f["floor"](app, (X.num(X.f["decimal-date-time"](app, X.f["now"](app))) - X.num(X.f["decimal-date-time"](app, X.v(app, "assessment_date"))))));
}

/** Calculation of is_tdr_unavailable */
export function calc_is_tdr_unavailable(app) {
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_acc_tdr_not_done"), "tdr_unavailable")) || X.bool(X.eq(X.v(app, "s_malaria_tdr_not_done"), "tdr_unavailable")))) || X.bool(X.eq(X.v(app, "s_malnutrition_tdr_not_done"), "tdr_unavailable")))) ? "true" : "false"));
}

/** Calculation of c_out_of_stock */
export function calc_c_out_of_stock(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_not_give_plumpy_sup"), "out_of_stock")) || X.bool(X.eq(X.v(app, "s_not_give_plumpy_nut"), "out_of_stock")))) || X.bool(X.eq(X.v(app, "s_not_give_enriched_flour"), "out_of_stock")))) || X.bool(X.eq(X.v(app, "s_not_give_vitamina"), "out_of_stock")))) || X.bool(X.eq(X.v(app, "s_not_give_albendazole"), "out_of_stock")))) || X.bool(X.eq(X.v(app, "s_not_give_paracetamol"), "out_of_stock")))) || X.bool(X.eq(X.v(app, "s_malaria_not_give_act"), "out_of_stock")))) || X.bool(X.eq(X.v(app, "s_diarrhea_not_give_zinc"), "out_of_stock")))) || X.bool(X.eq(X.v(app, "s_diarrhea_not_give_ors"), "out_of_stock")))) || X.bool(X.eq(X.v(app, "s_ari_not_give_amox"), "out_of_stock")))) || X.bool(X.eq(X.v(app, "s_not_use_shakir_strip"), "out_of_stock")))) || X.bool(X.eq(X.v(app, "s_antimalarial_suppository_not_give"), "out_of_stock")))) ? "true" : "false"));
}

/** Calculation of plumpy_sup_status */
export function calc_plumpy_sup_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_give_plumpy_sup"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "s_give_plumpy_sup"), "yes")) ? "given" : X.v(app, "s_not_give_plumpy_sup"))));
}

/** Calculation of plumpy_nut_status */
export function calc_plumpy_nut_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_give_plumpy_nut"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "s_give_plumpy_nut"), "yes")) ? "given" : X.v(app, "s_not_give_plumpy_nut"))));
}

/** Calculation of enriched_flour_status */
export function calc_enriched_flour_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_give_enriched_flour"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "s_give_enriched_flour"), "yes")) ? "given" : X.v(app, "s_not_give_enriched_flour"))));
}

/** Calculation of vitamina_status */
export function calc_vitamina_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_give_vitamina"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "s_give_vitamina"), "yes")) ? "given" : X.v(app, "s_not_give_vitamina"))));
}

/** Calculation of albendazole_status */
export function calc_albendazole_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_give_albendazole"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "s_give_albendazole"), "yes")) ? "given" : X.v(app, "s_not_give_albendazole"))));
}

/** Calculation of paracetamol_status */
export function calc_paracetamol_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_give_paracetamol"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "s_give_paracetamol"), "yes")) ? "given" : X.v(app, "s_not_give_paracetamol"))));
}

/** Calculation of antimalarial_suppository_status */
export function calc_antimalarial_suppository_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_did_place_antimalarial_suppository"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "s_did_place_antimalarial_suppository"), "yes")) ? "given" : X.v(app, "s_antimalarial_suppository_not_give"))));
}

/** Calculation of act_status */
export function calc_act_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_malaria_give_act"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "s_malaria_give_act"), "yes")) ? "given" : X.v(app, "s_malaria_not_give_act"))));
}

/** Calculation of zinc_status */
export function calc_zinc_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_diarrhea_give_zinc"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "s_diarrhea_give_zinc"), "yes")) ? "given" : X.v(app, "s_diarrhea_not_give_zinc"))));
}

/** Calculation of ors_status */
export function calc_ors_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_diarrhea_give_ors"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "s_diarrhea_give_ors"), "yes")) ? "given" : X.v(app, "s_diarrhea_not_give_ors"))));
}

/** Calculation of amox_status */
export function calc_amox_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_ari_give_amox"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "s_ari_give_amox"), "yes")) ? "given" : X.v(app, "s_ari_not_give_amox"))));
}

/** Calculation of rdt_status */
export function calc_rdt_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "tdr_done"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "tdr_done"), "yes")) ? "given" : X.v(app, "no_tdr_done"))));
}

/** Calculation of shakir_strip_status */
export function calc_shakir_strip_status(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_use_shakir_strip"), "")) ? "not_proposed" : (X.bool(X.eq(X.v(app, "s_use_shakir_strip"), "yes")) ? "given" : X.v(app, "s_not_use_shakir_strip"))));
}

/** Calculation of drugs_not_proposed */
export function calc_drugs_not_proposed(app) {
  return X.out((X.num((X.num((X.num((X.num((X.num((X.num((X.num((X.num((X.num((X.num((X.bool(X.eq(X.v(app, "plumpy_sup_status"), "not_proposed")) ? 1 : 0)) + X.num((X.bool(X.eq(X.v(app, "plumpy_nut_status"), "not_proposed")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "enriched_flour_status"), "not_proposed")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "vitamina_status"), "not_proposed")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "albendazole_status"), "not_proposed")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "paracetamol_status"), "not_proposed")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "act_status"), "not_proposed")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "zinc_status"), "not_proposed")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "ors_status"), "not_proposed")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "antimalarial_suppository_status"), "not_proposed")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "amox_status"), "not_proposed")) ? 1 : 0))));
}

/** Calculation of other_products_not_proposed */
export function calc_other_products_not_proposed(app) {
  return X.out((X.num((X.bool(X.eq(X.v(app, "rdt_status"), "not_proposed")) ? 1 : 0)) + X.num((X.bool(X.eq(X.v(app, "shakir_strip_status"), "not_proposed")) ? 1 : 0))));
}

/** Calculation of drugs_proposed */
export function calc_drugs_proposed(app) {
  return X.out((X.num(11) - X.num(X.v(app, "drugs_not_proposed"))));
}

/** Calculation of other_products_proposed */
export function calc_other_products_proposed(app) {
  return X.out((X.num(2) - X.num(X.v(app, "other_products_not_proposed"))));
}

/** Calculation of drugs_given */
export function calc_drugs_given(app) {
  return X.out((X.num((X.num((X.num((X.num((X.num((X.num((X.num((X.num((X.num((X.num((X.bool(X.eq(X.v(app, "plumpy_sup_status"), "given")) ? 1 : 0)) + X.num((X.bool(X.eq(X.v(app, "plumpy_nut_status"), "given")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "enriched_flour_status"), "given")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "vitamina_status"), "given")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "albendazole_status"), "given")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "paracetamol_status"), "given")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "act_status"), "given")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "zinc_status"), "given")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "ors_status"), "given")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "antimalarial_suppository_status"), "given")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "amox_status"), "given")) ? 1 : 0))));
}

/** Calculation of other_products_given */
export function calc_other_products_given(app) {
  return X.out((X.num((X.bool(X.eq(X.v(app, "rdt_status"), "given")) ? 1 : 0)) + X.num((X.bool(X.eq(X.v(app, "shakir_strip_status"), "given")) ? 1 : 0))));
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
  return X.out((X.num((X.num((X.num((X.num((X.num((X.num((X.num((X.num((X.num((X.num((X.bool(X.eq(X.v(app, "plumpy_sup_status"), "out_of_stock")) ? 1 : 0)) + X.num((X.bool(X.eq(X.v(app, "plumpy_nut_status"), "out_of_stock")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "enriched_flour_status"), "out_of_stock")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "vitamina_status"), "out_of_stock")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "albendazole_status"), "out_of_stock")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "paracetamol_status"), "out_of_stock")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "act_status"), "out_of_stock")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "zinc_status"), "out_of_stock")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "ors_status"), "out_of_stock")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "antimalarial_suppository_status"), "out_of_stock")) ? 1 : 0)))) + X.num((X.bool(X.eq(X.v(app, "amox_status"), "out_of_stock")) ? 1 : 0))));
}

/** Calculation of other_products_not_given_out_of_stock */
export function calc_other_products_not_given_out_of_stock(app) {
  return X.out((X.num((X.bool(X.eq(X.v(app, "rdt_status"), "tdr_unavailable")) ? 1 : 0)) + X.num((X.bool(X.eq(X.v(app, "shakir_strip_status"), "out_of_stock")) ? 1 : 0))));
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
  return X.out((X.bool((X.bool(X.eq(X.f["number"](app, X.v(app, "drugs_proposed")), X.f["number"](app, X.v(app, "drugs_not_given_out_of_stock")))) || X.bool(X.eq(X.f["number"](app, X.v(app, "drugs_proposed")), X.f["number"](app, X.v(app, "drugs_given")))))) ? "false" : (X.bool(X.eq(X.v(app, "new_CAT"), "")) ? "true" : "false")));
}

/** Calculation of has_been_accompany_to_cscom */
export function calc_has_been_accompany_to_cscom(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_accompany_refer_CSCOM"), "yes")) ? "true" : "false"));
}

/** Calculation of why_shakir_not_used */
export function calc_why_shakir_not_used(app) {
  return X.out(X.choiceName("c_not_use_shakir_strip", X.v(app, "s_not_use_shakir_strip")));
}

/** Calculation of why_shakir_not_used_label */
export function calc_why_shakir_not_used_label(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_not_use_shakir_strip"), "other")) ? X.v(app, "s_not_use_shakir_strip_other") : X.v(app, "why_shakir_not_used")));
}

/** Calculation of used_mam_intrant */
export function calc_used_mam_intrant(app) {
  return X.out(X.inp(app, "contact/parent/parent/mam_intrant"));
}

/** Calculation of use_plumpy_with_mam */
export function calc_use_plumpy_with_mam(app) {
  return X.out((X.bool(X.eq(X.inp(app, "contact/parent/parent/mam_intrant"), "")) ? "true" : (X.bool(X.eq(X.inp(app, "contact/parent/parent/mam_intrant"), "plumpy")) ? "true" : "false")));
}

/** Calculation of type_renamed */
export function calc_type_renamed(app) {
  return X.out(X.inp(app, "contact/parent/type"));
}

/** Calculation of s_child_temperature */
export function calc_s_child_temperature(app) {
  return X.out((X.bool((X.bool(X.lt(X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw"), 0), 35)) || X.bool(X.gt(X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw"), 0), 41)))) ? (X.num(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw_retake"), 0))) + X.num(X.f["number"](app, 0))) : (X.num(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_temperature_pre_chw"), 0))) + X.num(X.f["number"](app, 0)))));
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

/** Calculation of has_first_ref_danger_sign */
export function calc_has_first_ref_danger_sign(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_ref_danger_sign_newborn_sick"), "yes")) || X.bool(X.eq(X.v(app, "s_ref_danger_sign_not_gaining_weight"), "yes")))) || X.bool(X.eq(X.v(app, "s_ref_danger_sign_diarrhea_more_than_14_days"), "yes")))) || X.bool(X.eq(X.v(app, "s_ref_danger_sign_persistent_fever"), "yes")))) || X.bool(X.eq(X.v(app, "s_ref_danger_sign_blood_in_stools"), "yes")))) ? "true" : "false"));
}

/** Calculation of has_ref_danger_sign */
export function calc_has_ref_danger_sign(app) {
  return X.out((X.bool((X.bool(X.eq(X.v(app, "has_first_ref_danger_sign"), "true")) || X.bool(X.eq(X.v(app, "s_ref_danger_sign_other_diseases_newborn"), "yes")))) ? "true" : "false"));
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

/** Relevance of s_acc_danger_sign_cough_more_than_14_days */
export function rel_s_acc_danger_sign_cough_more_than_14_days(app) {
  return X.bool(X.gt(X.v(app, "patient_age_in_days"), 14));
}

/** Relevance of s_acc_danger_sign_hemoptysis */
export function rel_s_acc_danger_sign_hemoptysis(app) {
  return X.bool(X.eq(X.v(app, "type_renamed"), "contact"));
}

/** Calculation of has_danger_sign */
export function calc_has_danger_sign(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_acc_danger_sign_seizure"), "yes")) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_loss_consiousness"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_unable_drink"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_confusion"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_vomit"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_chest_indrawing"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_wheezing"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_hemoptysis"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_bleeding"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_lathargy"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_cough_more_than_14_days"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_severe_jaundice"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_dark_colored_urine"), "yes")))) || X.bool(X.eq(X.v(app, "s_acc_danger_sign_conjunctival_pallor"), "yes")))) ? "true" : "false"));
}

/** Relevance of s_acc_tdr_done */
export function rel_s_acc_tdr_done(app) {
  return X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "true")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 2))));
}

/** Relevance of s_acc_tdr_result */
export function rel_s_acc_tdr_result(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_acc_tdr_done"), "yes")) && X.bool(X.ne(X.v(app, "can_access_mrdt"), "true"))));
}

/** Relevance of s_acc_tdr_not_done */
export function rel_s_acc_tdr_not_done(app) {
  return X.bool(X.eq(X.v(app, "s_acc_tdr_done"), "no"));
}

/** Relevance of s_acc_tdr_not_done_other */
export function rel_s_acc_tdr_not_done_other(app) {
  return X.bool(X.eq(X.v(app, "s_acc_tdr_not_done"), "other"));
}

/** Relevance of n_acc_tdr_refer_cscom */
export function rel_n_acc_tdr_refer_cscom(app) {
  return X.bool(X.eq(X.v(app, "s_acc_tdr_not_done"), "tdr_unavailable"));
}

/** Relevance of s_malaria_indication_1_r_note_disinfect */
export function rel_s_malaria_indication_1_r_note_disinfect(app) {
  return grp_s_malaria_indication_1(app);
}

/** Relevance of s_malaria_indication_2_note_malaria_2 */
export function rel_s_malaria_indication_2_note_malaria_2(app) {
  return grp_s_malaria_indication_2(app);
}

/** Relevance of s_malaria_indication_3_note_malaria_3 */
export function rel_s_malaria_indication_3_note_malaria_3(app) {
  return grp_s_malaria_indication_3(app);
}

/** Relevance of s_malaria_indication_4_note_malaria_4 */
export function rel_s_malaria_indication_4_note_malaria_4(app) {
  return grp_s_malaria_indication_4(app);
}

/** Relevance of s_malaria_indication_5_note_malaria_5 */
export function rel_s_malaria_indication_5_note_malaria_5(app) {
  return grp_s_malaria_indication_5(app);
}

/** Relevance of s_malaria_indication_6_note_malaria_6 */
export function rel_s_malaria_indication_6_note_malaria_6(app) {
  return grp_s_malaria_indication_6(app);
}

/** Relevance of s_malaria_indication_7_note_malaria_7 */
export function rel_s_malaria_indication_7_note_malaria_7(app) {
  return grp_s_malaria_indication_7(app);
}

/** Relevance of s_malaria_indication_8_note_malaria_8 */
export function rel_s_malaria_indication_8_note_malaria_8(app) {
  return grp_s_malaria_indication_8(app);
}

/** Relevance of s_malaria_indication_8_evaluation_timer */
export function rel_s_malaria_indication_8_evaluation_timer(app) {
  return grp_s_malaria_indication_8(app);
}

/** Relevance of s_malaria_indication_9_note_danger_signs_pos */
export function rel_s_malaria_indication_9_note_danger_signs_pos(app) {
  return grp_s_malaria_indication_9(app);
}

/** Relevance of s_malaria_indication_9_note_danger_signs_neg */
export function rel_s_malaria_indication_9_note_danger_signs_neg(app) {
  return grp_s_malaria_indication_9(app);
}

/** Relevance of s_acc_tdr_result_new */
export function rel_s_acc_tdr_result_new(app) {
  return grp_s_malaria_indication_9(app);
}

/** Relevance of interpretation_mrdt */
export function rel_interpretation_mrdt(app) {
  return grp_s_malaria_indication_9(app);
}

/** Calculation of interpretation_mrdt */
export function calc_interpretation_mrdt(app) {
  if (!rel_interpretation_mrdt(app)) return '';
  return X.out((X.bool(X.eq(X.v(app, "s_acc_tdr_result_new"), "pos")) ? "Positif (+)" : (X.bool(X.eq(X.inp(app, "user/language"), "fr")) ? "Négatif (-)" : "Negatif (-)")));
}

/** Relevance of s_ref_danger_signs_note */
export function rel_s_ref_danger_signs_note(app) {
  return grp_s_ref_danger_signs(app);
}

/** Relevance of s_ref_danger_sign_newborn_sick */
export function rel_s_ref_danger_sign_newborn_sick(app) {
  return grp_s_ref_danger_signs(app);
}

/** Relevance of s_ref_danger_sign_not_gaining_weight */
export function rel_s_ref_danger_sign_not_gaining_weight(app) {
  return grp_s_ref_danger_signs(app);
}

/** Relevance of s_ref_danger_sign_diarrhea_more_than_14_days */
export function rel_s_ref_danger_sign_diarrhea_more_than_14_days(app) {
  return grp_s_ref_danger_signs(app) && X.bool(X.gt(X.v(app, "patient_age_in_days"), 14));
}

/** Relevance of s_ref_danger_sign_blood_in_stools */
export function rel_s_ref_danger_sign_blood_in_stools(app) {
  return grp_s_ref_danger_signs(app);
}

/** Relevance of s_ref_danger_sign_persistent_fever */
export function rel_s_ref_danger_sign_persistent_fever(app) {
  return grp_s_ref_danger_signs(app) && X.bool(X.eq(X.v(app, "type_renamed"), "contact"));
}

/** Relevance of s_ref_danger_sign_other_diseases_newborn */
export function rel_s_ref_danger_sign_other_diseases_newborn(app) {
  return grp_s_ref_other_diseases(app);
}

/** Relevance of s_ref_danger_sign_other */
export function rel_s_ref_danger_sign_other(app) {
  return grp_s_ref_other_diseases(app) && X.bool(X.eq(X.v(app, "s_ref_danger_sign_other_diseases_newborn"), "yes"));
}

/** Relevance of s_use_shakir_strip */
export function rel_s_use_shakir_strip(app) {
  return grp_s_malnutrition(app);
}

/** Relevance of s_nutri_color_shakir */
export function rel_s_nutri_color_shakir(app) {
  return grp_s_malnutrition(app) && X.bool(X.eq(X.v(app, "s_use_shakir_strip"), "yes"));
}

/** Relevance of s_nutri_color_shakir_fr */
export function rel_s_nutri_color_shakir_fr(app) {
  return grp_s_malnutrition(app) && X.bool(X.eq(X.v(app, "s_use_shakir_strip"), "yes"));
}

/** Calculation of s_nutri_color_shakir_fr */
export function calc_s_nutri_color_shakir_fr(app) {
  if (!rel_s_nutri_color_shakir_fr(app)) return '';
  return X.out((X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Green")) ? "Verte" : (X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Yellow")) ? "Jaune" : (X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Red")) ? "Rouge" : ""))));
}

/** Relevance of s_nutri_color_shakir_bm */
export function rel_s_nutri_color_shakir_bm(app) {
  return grp_s_malnutrition(app) && X.bool(X.eq(X.v(app, "s_use_shakir_strip"), "yes"));
}

/** Calculation of s_nutri_color_shakir_bm */
export function calc_s_nutri_color_shakir_bm(app) {
  if (!rel_s_nutri_color_shakir_bm(app)) return '';
  return X.out((X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Green")) ? "ɲuguji" : (X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Yellow")) ? "Nɛrɛ" : (X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Red")) ? "Bileman" : ""))));
}

/** Relevance of s_nutri_shakir_strip_length_green */
export function rel_s_nutri_shakir_strip_length_green(app) {
  return grp_s_malnutrition(app) && X.bool((X.bool(X.eq(X.v(app, "s_use_shakir_strip"), "yes")) && X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Green"))));
}

/** Constraint of s_nutri_shakir_strip_length_green */
export function val_s_nutri_shakir_strip_length_green(app) {
  if (X.str(X.v(app, "s_nutri_shakir_strip_length_green")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "s_nutri_shakir_strip_length_green"), 12.5)) && X.bool(X.le(X.v(app, "s_nutri_shakir_strip_length_green"), 26.5))));
}

/** Relevance of s_nutri_shakir_strip_length_yellow */
export function rel_s_nutri_shakir_strip_length_yellow(app) {
  return grp_s_malnutrition(app) && X.bool((X.bool(X.eq(X.v(app, "s_use_shakir_strip"), "yes")) && X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Yellow"))));
}

/** Constraint of s_nutri_shakir_strip_length_yellow */
export function val_s_nutri_shakir_strip_length_yellow(app) {
  if (X.str(X.v(app, "s_nutri_shakir_strip_length_yellow")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "s_nutri_shakir_strip_length_yellow"), 11.5)) && X.bool(X.lt(X.v(app, "s_nutri_shakir_strip_length_yellow"), 12.5))));
}

/** Relevance of s_nutri_shakir_strip_length_red */
export function rel_s_nutri_shakir_strip_length_red(app) {
  return grp_s_malnutrition(app) && X.bool((X.bool(X.eq(X.v(app, "s_use_shakir_strip"), "yes")) && X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Red"))));
}

/** Constraint of s_nutri_shakir_strip_length_red */
export function val_s_nutri_shakir_strip_length_red(app) {
  if (X.str(X.v(app, "s_nutri_shakir_strip_length_red")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "s_nutri_shakir_strip_length_red"), 5.5)) && X.bool(X.lt(X.v(app, "s_nutri_shakir_strip_length_red"), 11.5))));
}

/** Relevance of s_nutri_shakir_strip_length */
export function rel_s_nutri_shakir_strip_length(app) {
  return grp_s_malnutrition(app);
}

/** Calculation of s_nutri_shakir_strip_length */
export function calc_s_nutri_shakir_strip_length(app) {
  if (!rel_s_nutri_shakir_strip_length(app)) return '';
  return X.out((X.bool(X.ne(X.f["coalesce"](app, X.v(app, "s_nutri_shakir_strip_length_green"), 0), "")) ? X.f["coalesce"](app, X.v(app, "s_nutri_shakir_strip_length_green"), 0) : (X.bool(X.ne(X.f["coalesce"](app, X.v(app, "s_nutri_shakir_strip_length_yellow"), 0), "")) ? X.f["coalesce"](app, X.v(app, "s_nutri_shakir_strip_length_yellow"), 0) : X.f["coalesce"](app, X.v(app, "s_nutri_shakir_strip_length_red"), 0))));
}

/** Relevance of s_not_use_shakir_strip */
export function rel_s_not_use_shakir_strip(app) {
  return grp_s_malnutrition(app) && X.bool(X.eq(X.v(app, "s_use_shakir_strip"), "no"));
}

/** Relevance of s_not_use_shakir_strip_other */
export function rel_s_not_use_shakir_strip_other(app) {
  return grp_s_malnutrition(app) && X.bool(X.eq(X.v(app, "s_not_use_shakir_strip"), "other"));
}

/** Relevance of n_shakir_strip_refer_cscom */
export function rel_n_shakir_strip_refer_cscom(app) {
  return grp_s_malnutrition(app) && X.bool(X.eq(X.v(app, "s_not_use_shakir_strip"), "out_of_stock"));
}

/** Relevance of s_malnutrition_edema */
export function rel_s_malnutrition_edema(app) {
  return grp_s_malnutrition_observation(app);
}

/** Relevance of n_nutri_color_shakir_green */
export function rel_n_nutri_color_shakir_green(app) {
  return grp_s_malnutrition_observation(app) && X.bool((X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Green")) && X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "yes"))));
}

/** Relevance of n_nutri_color_shakir_yellow */
export function rel_n_nutri_color_shakir_yellow(app) {
  return grp_s_malnutrition_observation(app) && X.bool((X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Yellow")) && X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "yes"))));
}

/** Relevance of n_nutri_color_shakir_red */
export function rel_n_nutri_color_shakir_red(app) {
  return grp_s_malnutrition_observation(app) && X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Red"));
}

/** Relevance of has_MAM */
export function rel_has_MAM(app) {
  return grp_s_malnutrition_observation(app);
}

/** Calculation of has_MAM */
export function calc_has_MAM(app) {
  if (!rel_has_MAM(app)) return '';
  return X.out((X.bool((X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Yellow")) && X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "no")))) ? "true" : "false"));
}

/** Relevance of has_SAM */
export function rel_has_SAM(app) {
  return grp_s_malnutrition_observation(app);
}

/** Calculation of has_SAM */
export function calc_has_SAM(app) {
  if (!rel_has_SAM(app)) return '';
  return X.out((X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "yes")) ? "true" : "false"));
}

/** Relevance of s_malnutrition_appetite */
export function rel_s_malnutrition_appetite(app) {
  return grp_s_malnutrition_observation(app) && X.bool((X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "yes")) || X.bool((X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "no")) && X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Red"))))));
}

/** Relevance of s_malnutrition_sores */
export function rel_s_malnutrition_sores(app) {
  return grp_s_malnutrition_observation(app) && X.bool((X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "yes")) || X.bool((X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "no")) && X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Red"))))));
}

/** Relevance of has_SAM_without_complication */
export function rel_has_SAM_without_complication(app) {
  return grp_s_malnutrition_observation(app);
}

/** Calculation of has_SAM_without_complication */
export function calc_has_SAM_without_complication(app) {
  if (!rel_has_SAM_without_complication(app)) return '';
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_malnutrition_sores"), "no")) && X.bool((X.bool(X.eq(X.v(app, "s_malnutrition_appetite"), "meduim")) || X.bool(X.eq(X.v(app, "s_malnutrition_appetite"), "good")))))) && X.bool((X.bool(X.eq(X.v(app, "has_SAM"), "true")) || X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Red")))))) ? "true" : "false"));
}

/** Relevance of has_SAM_with_complication */
export function rel_has_SAM_with_complication(app) {
  return grp_s_malnutrition_observation(app);
}

/** Calculation of has_SAM_with_complication */
export function calc_has_SAM_with_complication(app) {
  if (!rel_has_SAM_with_complication(app)) return '';
  return X.out((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_malnutrition_appetite"), "low")) || X.bool(X.eq(X.v(app, "s_malnutrition_sores"), "yes")))) && X.bool((X.bool(X.eq(X.v(app, "has_SAM"), "true")) || X.bool((X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Red")) && X.bool(X.eq(X.v(app, "has_SAM"), "false")))))))) ? "true" : "false"));
}

/** Relevance of s_child_weight */
export function rel_s_child_weight(app) {
  return grp_s_malnutrition_weight(app) && X.bool((X.bool((X.bool(X.eq(X.v(app, "s_use_shakir_strip"), "yes")) || X.bool(X.eq(X.v(app, "s_not_use_shakir_strip"), "out_of_stock")))) || X.bool(X.eq(X.v(app, "s_not_use_shakir_strip"), "other"))));
}

/** Constraint of s_child_weight */
export function val_s_child_weight(app) {
  if (X.str(X.v(app, "s_child_weight")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "s_child_weight"), 3)) && X.bool(X.le(X.v(app, "s_child_weight"), 50))));
}

/** Relevance of s_malnutrition_tdr_done */
export function rel_s_malnutrition_tdr_done(app) {
  return grp_s_malnutrition_mRDT(app);
}

/** Relevance of s_malnutrition_tdr_result */
export function rel_s_malnutrition_tdr_result(app) {
  return grp_s_malnutrition_mRDT(app) && X.bool((X.bool(X.eq(X.v(app, "s_malnutrition_tdr_done"), "yes")) && X.bool(X.ne(X.v(app, "can_access_mrdt"), "true"))));
}

/** Relevance of s_malnutrition_tdr_not_done */
export function rel_s_malnutrition_tdr_not_done(app) {
  return grp_s_malnutrition_mRDT(app) && X.bool(X.eq(X.v(app, "s_malnutrition_tdr_done"), "no"));
}

/** Relevance of s_malnutrition_tdr_not_done_other */
export function rel_s_malnutrition_tdr_not_done_other(app) {
  return grp_s_malnutrition_mRDT(app) && X.bool(X.eq(X.v(app, "s_malnutrition_tdr_not_done"), "other"));
}

/** Relevance of n_malnutrition_tdr_refer_cscom */
export function rel_n_malnutrition_tdr_refer_cscom(app) {
  return grp_s_malnutrition_mRDT(app) && X.bool(X.eq(X.v(app, "s_malnutrition_tdr_not_done"), "tdr_unavailable"));
}

/** Relevance of s_malaria_indication_10_r_note_disinfect */
export function rel_s_malaria_indication_10_r_note_disinfect(app) {
  return grp_s_malaria_indication_10(app);
}

/** Relevance of s_malaria_indication_11_note_malaria_2 */
export function rel_s_malaria_indication_11_note_malaria_2(app) {
  return grp_s_malaria_indication_11(app);
}

/** Relevance of s_malaria_indication_12_note_malaria_3 */
export function rel_s_malaria_indication_12_note_malaria_3(app) {
  return grp_s_malaria_indication_12(app);
}

/** Relevance of s_malaria_indication_13_note_malaria_4 */
export function rel_s_malaria_indication_13_note_malaria_4(app) {
  return grp_s_malaria_indication_13(app);
}

/** Relevance of s_malaria_indication_14_note_malaria_5 */
export function rel_s_malaria_indication_14_note_malaria_5(app) {
  return grp_s_malaria_indication_14(app);
}

/** Relevance of s_malaria_indication_15_note_malaria_6 */
export function rel_s_malaria_indication_15_note_malaria_6(app) {
  return grp_s_malaria_indication_15(app);
}

/** Relevance of s_malaria_indication_16_note_malaria_7 */
export function rel_s_malaria_indication_16_note_malaria_7(app) {
  return grp_s_malaria_indication_16(app);
}

/** Relevance of s_malaria_indication_17_note_malaria_8 */
export function rel_s_malaria_indication_17_note_malaria_8(app) {
  return grp_s_malaria_indication_17(app);
}

/** Relevance of s_malaria_indication_17_evaluation_timer */
export function rel_s_malaria_indication_17_evaluation_timer(app) {
  return grp_s_malaria_indication_17(app);
}

/** Relevance of s_malaria_indication_18_note_danger_signs_pos */
export function rel_s_malaria_indication_18_note_danger_signs_pos(app) {
  return grp_s_malaria_indication_18(app);
}

/** Relevance of s_malaria_indication_18_note_danger_signs_neg */
export function rel_s_malaria_indication_18_note_danger_signs_neg(app) {
  return grp_s_malaria_indication_18(app);
}

/** Relevance of s_malnutrition_tdr_result_new */
export function rel_s_malnutrition_tdr_result_new(app) {
  return grp_s_malaria_indication_18(app);
}

/** Relevance of interpretation_malnutrition_mrdt */
export function rel_interpretation_malnutrition_mrdt(app) {
  return grp_s_malaria_indication_18(app);
}

/** Calculation of interpretation_malnutrition_mrdt */
export function calc_interpretation_malnutrition_mrdt(app) {
  if (!rel_interpretation_malnutrition_mrdt(app)) return '';
  return X.out((X.bool(X.eq(X.v(app, "s_malnutrition_tdr_result_new"), "pos")) ? "Positif (+)" : (X.bool(X.eq(X.inp(app, "user/language"), "fr")) ? "Négatif (-)" : "Negatif (-)")));
}

/** Relevance of s_first_malnutrition_screening */
export function rel_s_first_malnutrition_screening(app) {
  return grp_s_screening(app);
}

/** Relevance of s_first_malnutrition_screening_other */
export function rel_s_first_malnutrition_screening_other(app) {
  return grp_s_screening(app) && X.bool(X.eq(X.v(app, "s_first_malnutrition_screening"), "other"));
}

/** Relevance of s_malaria_sign_fever */
export function rel_s_malaria_sign_fever(app) {
  return grp_s_malaria(app);
}

/** Calculation of s_malaria_sign_fever */
export function calc_s_malaria_sign_fever(app) {
  if (!rel_s_malaria_sign_fever(app)) return '';
  return X.out(X.v(app, "s_fever"));
}

/** Relevance of s_malaria_fever_within_24_48 */
export function rel_s_malaria_fever_within_24_48(app) {
  return grp_s_malaria(app) && X.bool(X.lt(X.v(app, "s_child_temperature"), 37.5));
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

/** Relevance of n_refer_tdr_refer_cscom */
export function rel_n_refer_tdr_refer_cscom(app) {
  return grp_s_malaria(app) && X.bool(X.eq(X.v(app, "s_malaria_tdr_not_done"), "tdr_unavailable"));
}

/** Relevance of s_malaria_indication_19_r_note_disinfect */
export function rel_s_malaria_indication_19_r_note_disinfect(app) {
  return grp_s_malaria_indication_19(app);
}

/** Relevance of s_malaria_indication_20_note_malaria_2 */
export function rel_s_malaria_indication_20_note_malaria_2(app) {
  return grp_s_malaria_indication_20(app);
}

/** Relevance of s_malaria_indication_21_note_malaria_3 */
export function rel_s_malaria_indication_21_note_malaria_3(app) {
  return grp_s_malaria_indication_21(app);
}

/** Relevance of s_malaria_indication_22_note_malaria_4 */
export function rel_s_malaria_indication_22_note_malaria_4(app) {
  return grp_s_malaria_indication_22(app);
}

/** Relevance of s_malaria_indication_23_note_malaria_5 */
export function rel_s_malaria_indication_23_note_malaria_5(app) {
  return grp_s_malaria_indication_23(app);
}

/** Relevance of s_malaria_indication_24_note_malaria_6 */
export function rel_s_malaria_indication_24_note_malaria_6(app) {
  return grp_s_malaria_indication_24(app);
}

/** Relevance of s_malaria_indication_25_note_malaria_7 */
export function rel_s_malaria_indication_25_note_malaria_7(app) {
  return grp_s_malaria_indication_25(app);
}

/** Relevance of s_malaria_indication_26_note_malaria_8 */
export function rel_s_malaria_indication_26_note_malaria_8(app) {
  return grp_s_malaria_indication_26(app);
}

/** Relevance of s_malaria_indication_26_evaluation_timer */
export function rel_s_malaria_indication_26_evaluation_timer(app) {
  return grp_s_malaria_indication_26(app);
}

/** Relevance of s_malaria_indication_27_note_danger_signs_pos */
export function rel_s_malaria_indication_27_note_danger_signs_pos(app) {
  return grp_s_malaria_indication_27(app);
}

/** Relevance of s_malaria_indication_27_note_danger_signs_neg */
export function rel_s_malaria_indication_27_note_danger_signs_neg(app) {
  return grp_s_malaria_indication_27(app);
}

/** Relevance of s_malaria_tdr_result_new */
export function rel_s_malaria_tdr_result_new(app) {
  return grp_s_malaria_indication_27(app);
}

/** Relevance of interpretation_malaria_mrdt */
export function rel_interpretation_malaria_mrdt(app) {
  return grp_s_malaria_indication_27(app);
}

/** Calculation of interpretation_malaria_mrdt */
export function calc_interpretation_malaria_mrdt(app) {
  if (!rel_interpretation_malaria_mrdt(app)) return '';
  return X.out((X.bool(X.eq(X.v(app, "s_malaria_tdr_result_new"), "pos")) ? "Positif (+)" : (X.bool(X.eq(X.inp(app, "user/language"), "fr")) ? "Négatif (-)" : "Negatif (-)")));
}

/** Relevance of s_diarrhea_stools_a_day */
export function rel_s_diarrhea_stools_a_day(app) {
  return grp_s_diarrhea(app);
}

/** Relevance of s_ari_have_cough */
export function rel_s_ari_have_cough(app) {
  return grp_s_acute_respiratory_infection(app);
}

/** Relevance of breath_timer */
export function rel_breath_timer(app) {
  return grp_group_breathing(app);
}

/** Relevance of s_ari_respiration_rate */
export function rel_s_ari_respiration_rate(app) {
  return grp_group_breathing(app);
}

/** Constraint of s_ari_respiration_rate */
export function val_s_ari_respiration_rate(app) {
  if (X.str(X.v(app, "s_ari_respiration_rate")) === '') return true;
  return X.bool(X.gt(X.v(app, "s_ari_respiration_rate"), 0));
}

/** Relevance of s_ari_resp_estimation */
export function rel_s_ari_resp_estimation(app) {
  return grp_group_breathing_frequency(app) && X.bool(X.ne(X.v(app, "s_ari_respiration_rate"), ""));
}

/** Relevance of n_no_signs */
export function rel_n_no_signs(app) {
  return grp_s_CAT(app) && X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "false")) && X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "false")))) && X.bool(X.eq(X.v(app, "has_mal_danger_sign"), "false")))) && X.bool(X.eq(X.v(app, "s_diarrhea_stools_a_day"), "no")))) && X.bool(X.eq(X.v(app, "treat_malnutrition"), "false")))) && X.bool(X.eq(X.v(app, "s_ari_have_cough"), "no"))));
}

/** Relevance of new_CAT */
export function rel_new_CAT(app) {
  return grp_s_CAT(app);
}

/** Relevance of on_observation */
export function rel_on_observation(app) {
  return grp_s_CAT(app) && X.bool(X.eq(X.v(app, "new_CAT"), "watching"));
}

/** Calculation of medication_1 */
export function calc_medication_1(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_give_paracetamol"), "yes")) ? X.f["concat"](app, (X.bool(X.eq(X.inp(app, "user/language"), "bm")) ? "Parasetamɔli" : "Paracétamol"), " ", X.choiceName("c_give_paracetamol_dosage", X.v(app, "s_give_paracetamol_dosage")), " ") : ""));
}

/** Calculation of medication_2 */
export function calc_medication_2(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_malaria_give_act"), "yes")) ? X.f["concat"](app, (X.bool(X.eq(X.inp(app, "user/language"), "bm")) ? "Artésunate + Lumefantrine (ALU) 20/120 mg" : "Artésunate + Lumefantrine (ALU) 20/120 mg"), " ", X.choiceName("c_malaria_act_dosage", X.v(app, "s_malaria_act_dosage")), " ") : ""));
}

/** Calculation of medication_3 */
export function calc_medication_3(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_diarrhea_give_zinc"), "yes")) ? X.f["concat"](app, (X.bool(X.eq(X.inp(app, "user/language"), "bm")) ? "Zinc" : "Zinc"), " ", X.choiceName("c_diarrhea_zinc_dosage", X.v(app, "s_diarrhea_zinc_dosage")), " ") : ""));
}

/** Calculation of medication_4 */
export function calc_medication_4(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_diarrhea_give_ors"), "yes")) ? X.f["concat"](app, (X.bool(X.eq(X.inp(app, "user/language"), "bm")) ? "SRO" : "SRO"), " ") : ""));
}

/** Calculation of medication_5 */
export function calc_medication_5(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_ari_give_amox"), "yes")) ? X.f["concat"](app, (X.bool(X.eq(X.inp(app, "user/language"), "bm")) ? "Amɔkisilini" : "Amoxiciline"), " ", X.choiceName("c_ari_give_amox_dosage", X.v(app, "s_ari_give_amox_dosage")), " ") : ""));
}

/** Calculation of medication_8 */
export function calc_medication_8(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_did_place_antimalarial_suppository"), "yes")) ? X.f["concat"](app, (X.bool(X.eq(X.inp(app, "user/language"), "bm")) ? "Aritezunati supozituwari" : "Artésunate suppositoire"), " ", X.choiceName("c_give_artesunate_dosage", X.v(app, "s_artesunate_suppository_dosage")), " ") : ""));
}

/** Calculation of medication_10 */
export function calc_medication_10(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_rp_did_admin_gentamycine"), "yes")) ? X.f["concat"](app, (X.bool(X.eq(X.inp(app, "user/language"), "bm")) ? "Gentamicini " : " Gentamycine"), " ") : ""));
}

/** Calculation of medication_11 */
export function calc_medication_11(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_rp_did_admin_ampiciline"), "yes")) ? X.f["concat"](app, (X.bool(X.eq(X.inp(app, "user/language"), "bm")) ? "Ampicilini " : "Ampiciline"), " ") : ""));
}

/** Calculation of medication_12 */
export function calc_medication_12(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_ari_receive_honey"), "yes")) ? X.f["concat"](app, (X.bool(X.eq(X.inp(app, "user/language"), "bm")) ? "Miel" : "Miel"), " ") : ""));
}

/** Calculation of medication_13 */
export function calc_medication_13(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_ari_receive_eucalyptus"), "yes")) ? X.f["concat"](app, (X.bool(X.eq(X.inp(app, "user/language"), "bm")) ? "Thé de feuilles d Eucalyptus" : "Thé de feuilles d Eucalyptus"), " ") : ""));
}

/** Calculation of medication_14 */
export function calc_medication_14(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_ari_receive_lemon"), "yes")) ? X.f["concat"](app, (X.bool(X.eq(X.inp(app, "user/language"), "bm")) ? "Citron" : "Citron"), " ") : ""));
}

/** Calculation of medications */
export function calc_medications(app) {
  return X.out(X.f["concat"](app, X.v(app, "medication_1"), X.f["concat"](app, (X.bool(X.ne(X.v(app, "medication_1"), "")) ? ", " : ""), X.v(app, "medication_2")), X.f["concat"](app, (X.bool(X.ne(X.v(app, "medication_2"), "")) ? ", " : ""), X.v(app, "medication_3")), X.f["concat"](app, (X.bool(X.ne(X.v(app, "medication_3"), "")) ? ", " : ""), X.v(app, "medication_4")), X.f["concat"](app, (X.bool(X.ne(X.v(app, "medication_4"), "")) ? ", " : ""), X.v(app, "medication_5")), X.f["concat"](app, (X.bool(X.ne(X.v(app, "medication_5"), "")) ? ", " : ""), X.v(app, "medication_8")), X.f["concat"](app, (X.bool(X.ne(X.v(app, "medication_8"), "")) ? ", " : ""), X.v(app, "medication_10")), X.f["concat"](app, (X.bool(X.ne(X.v(app, "medication_10"), "")) ? ", " : ""), X.v(app, "medication_11")), X.f["concat"](app, (X.bool(X.ne(X.v(app, "medication_11"), "")) ? ", " : ""), X.v(app, "medication_12")), X.f["concat"](app, (X.bool(X.ne(X.v(app, "medication_12"), "")) ? ", " : ""), X.v(app, "medication_13")), X.f["concat"](app, (X.bool(X.ne(X.v(app, "medication_13"), "")) ? ", " : ""), X.v(app, "medication_14"))));
}

/** Relevance of r_symptoms */
export function rel_r_symptoms(app) {
  return X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "true")) || X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "true")))) || X.bool(X.eq(X.v(app, "s_ari_have_cough"), "yes")))) || X.bool(X.eq(X.v(app, "s_fever"), "yes")))) || X.bool((X.bool(X.eq(X.v(app, "has_MAM"), "true")) || X.bool(X.eq(X.v(app, "has_SAM"), "true"))))));
}

/** Relevance of r_fever_within24_48 */
export function rel_r_fever_within24_48(app) {
  return X.bool(X.eq(X.v(app, "s_fever"), "yes"));
}

/** Relevance of r_cough */
export function rel_r_cough(app) {
  return X.bool(X.eq(X.v(app, "s_ari_have_cough"), "yes"));
}

/** Relevance of r_danger_signs */
export function rel_r_danger_signs(app) {
  return X.bool(X.eq(X.v(app, "has_mal_danger_sign"), "true"));
}

/** Relevance of r_malaria_fever_within24_48 */
export function rel_r_malaria_fever_within24_48(app) {
  return X.bool((X.bool(X.eq(X.v(app, "has_mal_danger_sign"), "true")) && X.bool(X.eq(X.v(app, "s_fever"), "yes"))));
}

/** Calculation of has_malnutrition_summary */
export function calc_has_malnutrition_summary(app) {
  return X.out((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_use_shakir_strip"), "no")) || X.bool((X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Green")) && X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "yes")))))) || X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Yellow")))) || X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Red")))) || X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "yes")))) || X.bool(X.eq(X.v(app, "s_malnutrition_appetite"), "low")))) || X.bool(X.eq(X.v(app, "s_malnutrition_sores"), "yes")))) ? "true" : "false"));
}

/** Relevance of r_malnutrition_signs */
export function rel_r_malnutrition_signs(app) {
  return X.bool(X.eq(X.v(app, "has_malnutrition_summary"), "true"));
}

/** Relevance of r_no_shakir_strip */
export function rel_r_no_shakir_strip(app) {
  return X.bool(X.eq(X.v(app, "s_use_shakir_strip"), "no"));
}

/** Relevance of r_shakir_color_green */
export function rel_r_shakir_color_green(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Green")) && X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "yes"))));
}

/** Relevance of r_shakir_color_yellow */
export function rel_r_shakir_color_yellow(app) {
  return X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Yellow"));
}

/** Relevance of r_shakir_color_red */
export function rel_r_shakir_color_red(app) {
  return X.bool(X.eq(X.v(app, "s_nutri_color_shakir"), "Red"));
}

/** Relevance of r_malnutrition_edema */
export function rel_r_malnutrition_edema(app) {
  return X.bool(X.eq(X.v(app, "s_malnutrition_edema"), "yes"));
}

/** Relevance of r_malnutrition_appetite */
export function rel_r_malnutrition_appetite(app) {
  return X.bool(X.eq(X.v(app, "s_malnutrition_appetite"), "low"));
}

/** Relevance of r_malnutrition_sores */
export function rel_r_malnutrition_sores(app) {
  return X.bool(X.eq(X.v(app, "s_malnutrition_sores"), "yes"));
}

/** Relevance of r_signs_tb */
export function rel_r_signs_tb(app) {
  return X.bool(X.eq(X.v(app, "has_tb_sign"), "true"));
}

/** Relevance of r_signs_tb_list */
export function rel_r_signs_tb_list(app) {
  return X.bool(X.eq(X.v(app, "has_tb_sign"), "true"));
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

/** Relevance of n_acc_danger_sign_hemoptysis */
export function rel_n_acc_danger_sign_hemoptysis(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_hemoptysis"), "yes"));
}

/** Relevance of n_acc_danger_sign_chest_pain */
export function rel_n_acc_danger_sign_chest_pain(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_chest_pain"), "yes"));
}

/** Relevance of n_acc_danger_sign_bleeding */
export function rel_n_acc_danger_sign_bleeding(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_bleeding"), "yes"));
}

/** Relevance of n_acc_danger_sign_lathargy */
export function rel_n_acc_danger_sign_lathargy(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_lathargy"), "yes"));
}

/** Relevance of n_acc_danger_sign_severe_jaundice */
export function rel_n_acc_danger_sign_severe_jaundice(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_severe_jaundice"), "yes"));
}

/** Relevance of n_acc_danger_sign_dark_colored_urine */
export function rel_n_acc_danger_sign_dark_colored_urine(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_dark_colored_urine"), "yes"));
}

/** Relevance of n_acc_danger_sign_conjunctival_pallor */
export function rel_n_acc_danger_sign_conjunctival_pallor(app) {
  return X.bool(X.eq(X.v(app, "s_acc_danger_sign_conjunctival_pallor"), "yes"));
}

/** Relevance of n_ref_danger_sign_newborn_sick */
export function rel_n_ref_danger_sign_newborn_sick(app) {
  return X.bool(X.eq(X.v(app, "s_ref_danger_sign_newborn_sick"), "yes"));
}

/** Relevance of n_ref_danger_sign_not_gaining_weight */
export function rel_n_ref_danger_sign_not_gaining_weight(app) {
  return X.bool(X.eq(X.v(app, "s_ref_danger_sign_not_gaining_weight"), "yes"));
}

/** Relevance of n_ref_danger_sign_diarrhea_more_than_14_days */
export function rel_n_ref_danger_sign_diarrhea_more_than_14_days(app) {
  return X.bool(X.eq(X.v(app, "s_ref_danger_sign_diarrhea_more_than_14_days"), "yes"));
}

/** Relevance of n_ref_danger_sign_blood_in_stools */
export function rel_n_ref_danger_sign_blood_in_stools(app) {
  return X.bool(X.eq(X.v(app, "s_ref_danger_sign_blood_in_stools"), "yes"));
}

/** Relevance of n_ref_danger_sign_persistent_fever */
export function rel_n_ref_danger_sign_persistent_fever(app) {
  return X.bool(X.eq(X.v(app, "s_ref_danger_sign_persistent_fever"), "yes"));
}

/** Relevance of n_ref_danger_sign_other */
export function rel_n_ref_danger_sign_other(app) {
  return X.bool(X.eq(X.v(app, "s_ref_danger_sign_other_diseases_newborn"), "yes"));
}

/** Relevance of r_MAM */
export function rel_r_MAM(app) {
  return X.bool(X.eq(X.v(app, "has_MAM"), "true"));
}

/** Relevance of r_SAM_without_complication */
export function rel_r_SAM_without_complication(app) {
  return X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true"));
}

/** Relevance of r_SAM_with_complication */
export function rel_r_SAM_with_complication(app) {
  return X.bool(X.eq(X.v(app, "has_SAM_with_complication"), "true"));
}

/** Relevance of r_signs_malaria */
export function rel_r_signs_malaria(app) {
  return X.bool(X.eq(X.v(app, "tdr_result"), "positif"));
}

/** Relevance of r_no_tdr_result */
export function rel_r_no_tdr_result(app) {
  return X.bool((X.bool(X.eq(X.v(app, "tdr_result"), "")) && X.bool((X.bool(X.eq(X.v(app, "has_mal_danger_sign"), "true")) || X.bool(X.eq(X.v(app, "has_danger_sign"), "true"))))));
}

/** Relevance of r_symptom_diarrhea */
export function rel_r_symptom_diarrhea(app) {
  return X.bool(X.eq(X.v(app, "s_diarrhea_stools_a_day"), "yes"));
}

/** Relevance of s_ari_fast_resp_frequency_50 */
export function rel_s_ari_fast_resp_frequency_50(app) {
  return X.bool(X.eq(X.v(app, "fast_breathing"), "true"));
}

/** Relevance of s_ari_normal_resp_frequency_40 */
export function rel_s_ari_normal_resp_frequency_40(app) {
  return X.bool((X.bool(X.eq(X.v(app, "fast_breathing"), "false")) && X.bool(X.gt(X.f["coalesce"](app, X.v(app, "s_ari_respiration_rate"), 0), 0))));
}

/** Relevance of r_referral */
export function rel_r_referral(app) {
  return X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")) || X.bool(X.eq(X.v(app, "MAM_MAS_has_diarrhea_pneumonia"), "true")))) || X.bool((X.bool(X.lt(X.v(app, "patient_age_in_months"), 2)) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false")))))) && X.bool(X.eq(X.v(app, "tb_referral"), "false"))));
}

/** Relevance of r_referral_tb */
export function rel_r_referral_tb(app) {
  return X.bool(X.eq(X.v(app, "has_tb_sign"), "true"));
}

/** Relevance of n_tdr_refer_cscom */
export function rel_n_tdr_refer_cscom(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "s_acc_tdr_not_done"), "tdr_unavailable")) || X.bool(X.eq(X.v(app, "s_malnutrition_tdr_not_done"), "tdr_unavailable")))) || X.bool(X.eq(X.v(app, "s_malaria_tdr_not_done"), "tdr_unavailable"))));
}

/** Relevance of r_accompany */
export function rel_r_accompany(app) {
  return X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true"));
}

/** Relevance of s_SAM_with_complication */
export function rel_s_SAM_with_complication(app) {
  return X.bool(X.eq(X.v(app, "has_SAM_with_complication"), "true"));
}

/** Relevance of summary_special_instructions */
export function rel_summary_special_instructions(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "has_SAM_with_complication"), "true")) || X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true")))) || X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")) && X.bool(X.ne(X.f["number"](app, X.v(app, "drugs_proposed")), X.f["number"](app, X.v(app, "drugs_not_given_out_of_stock"))))))));
}

/** Relevance of refer_to_cscom_tb */
export function rel_refer_to_cscom_tb(app) {
  return X.bool(X.eq(X.v(app, "has_tb_sign"), "true"));
}

/** Relevance of r_under_2mo */
export function rel_r_under_2mo(app) {
  return X.bool((X.bool((X.bool(X.lt(X.v(app, "patient_age_in_months"), 2)) && X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")))) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))));
}

/** Relevance of s_dtn_danger_signs_1 */
export function rel_s_dtn_danger_signs_1(app) {
  return X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true"));
}

/** Relevance of r_referral_2 */
export function rel_r_referral_2(app) {
  return X.bool(X.eq(X.v(app, "has_SAM_with_complication"), "true"));
}

/** Relevance of s_dtn_danger_signs_2 */
export function rel_s_dtn_danger_signs_2(app) {
  return X.bool((X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true")) || X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")) && X.bool(X.eq(X.v(app, "has_danger_sign"), "true"))))));
}

/** Relevance of r_write_tb_on_sheet */
export function rel_r_write_tb_on_sheet(app) {
  return X.bool(X.eq(X.v(app, "has_tb_sign"), "true"));
}

/** Relevance of s_SAM_without_complication */
export function rel_s_SAM_without_complication(app) {
  return X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true"));
}

/** Relevance of s_MAM */
export function rel_s_MAM(app) {
  return X.bool(X.eq(X.v(app, "has_MAM"), "true"));
}

/** Relevance of r_treatment */
export function rel_r_treatment(app) {
  return X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "treat_diarrhea"), "true")) || X.bool(X.eq(X.v(app, "treat_malaria"), "true")))) || X.bool(X.eq(X.v(app, "treat_ari"), "true")))) || X.bool(X.eq(X.v(app, "s_fever"), "yes")))) || X.bool(X.eq(X.v(app, "treat_malnutrition"), "true"))));
}

/** Relevance of s_dnt_malnutrition */
export function rel_s_dnt_malnutrition(app) {
  return X.bool((X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")) || X.bool(X.eq(X.v(app, "has_MAM"), "true"))));
}

/** Relevance of s_dnt_MAM */
export function rel_s_dnt_MAM(app) {
  return X.bool(X.eq(X.v(app, "has_MAM"), "true"));
}

/** Relevance of s_MAM_RUTF */
export function rel_s_MAM_RUTF(app) {
  return X.bool(X.eq(X.v(app, "has_MAM"), "true"));
}

/** Relevance of r_plumpy_sup */
export function rel_r_plumpy_sup(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "has_MAM"), "true"))));
}

/** Relevance of r_plumpy_sup_1 */
export function rel_r_plumpy_sup_1(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "has_MAM"), "true"))));
}

/** Relevance of s_give_plumpy_sup */
export function rel_s_give_plumpy_sup(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "has_MAM"), "true"))));
}

/** Relevance of s_give_plumpy_sup_dosage */
export function rel_s_give_plumpy_sup_dosage(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "s_give_plumpy_sup"), "yes"))));
}

/** Constraint of s_give_plumpy_sup_dosage */
export function val_s_give_plumpy_sup_dosage(app) {
  if (X.str(X.v(app, "s_give_plumpy_sup_dosage")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "s_give_plumpy_sup_dosage"), 1)) && X.bool(X.le(X.v(app, "s_give_plumpy_sup_dosage"), 99))));
}

/** Relevance of s_not_give_plumpy_sup */
export function rel_s_not_give_plumpy_sup(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "s_give_plumpy_sup"), "no"))));
}

/** Relevance of s_not_give_plumpy_sup_other */
export function rel_s_not_give_plumpy_sup_other(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "s_not_give_plumpy_sup"), "other"))));
}

/** Relevance of n_plumpy_nut_shortage */
export function rel_n_plumpy_nut_shortage(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "true")) && X.bool(X.eq(X.v(app, "s_not_give_plumpy_nut"), "out_of_stock"))));
}

/** Relevance of r_enriched_flour */
export function rel_r_enriched_flour(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "has_MAM"), "true"))));
}

/** Relevance of r_enriched_flour_note */
export function rel_r_enriched_flour_note(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "has_MAM"), "true"))));
}

/** Relevance of s_give_enriched_flour */
export function rel_s_give_enriched_flour(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "has_MAM"), "true"))));
}

/** Relevance of s_give_enriched_flour_dosage */
export function rel_s_give_enriched_flour_dosage(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_give_enriched_flour"), "yes"))));
}

/** Constraint of s_give_enriched_flour_dosage */
export function val_s_give_enriched_flour_dosage(app) {
  if (X.str(X.v(app, "s_give_enriched_flour_dosage")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "s_give_enriched_flour_dosage"), 1)) && X.bool(X.le(X.v(app, "s_give_enriched_flour_dosage"), 99))));
}

/** Relevance of s_not_give_enriched_flour */
export function rel_s_not_give_enriched_flour(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_give_enriched_flour"), "no"))));
}

/** Relevance of s_not_give_enriched_flour_other */
export function rel_s_not_give_enriched_flour_other(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_not_give_enriched_flour"), "other"))));
}

/** Relevance of n_enriched_flour_shortage */
export function rel_n_enriched_flour_shortage(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_not_give_enriched_flour"), "out_of_stock"))));
}

/** Relevance of r_vitamina */
export function rel_r_vitamina(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "has_MAM"), "true"))));
}

/** Calculation of s_vitamina_dosage */
export function calc_s_vitamina_dosage(app) {
  return X.out((X.bool(X.lt(X.v(app, "patient_age_in_months"), 6)) ? "" : (X.bool(X.lt(X.v(app, "patient_age_in_months"), 12)) ? (X.bool(X.eq(X.inp(app, "user/language"), "en")) ? "1 capsule of 100,000 IU (1 blue capsule)" : (X.bool(X.eq(X.inp(app, "user/language"), "fr")) ? "1 capsule de 100 000 UI (1 capsule bleu)" : "Furakisɛ bulama 1 (100 000 UI)  siɲe kelen")) : (X.bool(X.eq(X.inp(app, "user/language"), "en")) ? "1 capsule of 200,000 IU (1 red capsule)" : (X.bool(X.eq(X.inp(app, "user/language"), "fr")) ? "1 capsule de 200 000 UI (1 capsule rouge)" : "Furakisɛ bilenma 1 (200 000 UI)  siɲe kelen")))));
}

/** Relevance of r_vitamina_1 */
export function rel_r_vitamina_1(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "has_MAM"), "true"))));
}

/** Relevance of s_give_vitamina */
export function rel_s_give_vitamina(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "has_MAM"), "true"))));
}

/** Relevance of s_give_vitamina_dosage */
export function rel_s_give_vitamina_dosage(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_give_vitamina"), "yes"))));
}

/** Relevance of s_not_give_vitamina */
export function rel_s_not_give_vitamina(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_give_vitamina"), "no"))));
}

/** Relevance of s_not_give_vitamina_other */
export function rel_s_not_give_vitamina_other(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_not_give_vitamina"), "other"))));
}

/** Relevance of n_vitamina_shortage */
export function rel_n_vitamina_shortage(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_not_give_vitamina"), "out_of_stock"))));
}

/** Relevance of r_albendazole */
export function rel_r_albendazole(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "has_MAM"), "true")))) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 12))));
}

/** Calculation of s_albendazole_dosage */
export function calc_s_albendazole_dosage(app) {
  return X.out((X.bool(X.lt(X.v(app, "patient_age_in_months"), 12)) ? "" : (X.bool(X.lt(X.v(app, "patient_age_in_months"), 24)) ? (X.bool(X.eq(X.v(app, "s_give_albendazole_type"), "200_mg_albendazole")) ? (X.bool(X.eq(X.inp(app, "user/language"), "en")) ? "1 tablet of 200mg" : (X.bool(X.eq(X.inp(app, "user/language"), "fr")) ? "1 comprimé de 200 mg" : "200 mg furakisɛ 1")) : (X.bool(X.eq(X.inp(app, "user/language"), "en")) ? "½  tablet of 400mg" : (X.bool(X.eq(X.inp(app, "user/language"), "fr")) ? "½ comprimé de 400 mg" : "400 mg furakisɛ ½"))) : (X.bool(X.eq(X.v(app, "s_give_albendazole_type"), "200_mg_albendazole")) ? (X.bool(X.eq(X.inp(app, "user/language"), "en")) ? "2 tablets of 200mg" : (X.bool(X.eq(X.inp(app, "user/language"), "fr")) ? "2 comprimés de 200 mg" : "200 mg furakisɛ 2")) : (X.bool(X.eq(X.inp(app, "user/language"), "en")) ? "1 tablet of 400mg" : (X.bool(X.eq(X.inp(app, "user/language"), "fr")) ? "1 comprimé de 400 mg" : "400 mg furakisɛ 1"))))));
}

/** Relevance of s_give_albendazole */
export function rel_s_give_albendazole(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "has_MAM"), "true")))) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 12))));
}

/** Relevance of s_give_albendazole_type */
export function rel_s_give_albendazole_type(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_give_albendazole"), "yes"))));
}

/** Relevance of r_albendazole_1 */
export function rel_r_albendazole_1(app) {
  return X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "has_MAM"), "true")))) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 12)))) && X.bool(X.ne(X.v(app, "s_give_albendazole_type"), ""))));
}

/** Relevance of s_given_albendazole_dosage */
export function rel_s_given_albendazole_dosage(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_give_albendazole"), "yes")))) && X.bool(X.ne(X.v(app, "s_give_albendazole_type"), ""))));
}

/** Options of s_given_albendazole_dosage */
export function filt_s_given_albendazole_dosage(app) {
  return X.bool((X.bool(X.eq(X.opt(app, "filter1"), X.v(app, "s_give_albendazole_type"))) || X.bool(X.eq(X.opt(app, "filter1"), "all"))));
}

/** Relevance of s_not_give_albendazole */
export function rel_s_not_give_albendazole(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_give_albendazole"), "no"))));
}

/** Relevance of s_not_give_albendazole_other */
export function rel_s_not_give_albendazole_other(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_not_give_albendazole"), "other"))));
}

/** Relevance of n__shoalbendazolertage */
export function rel_n__shoalbendazolertage(app) {
  return X.bool((X.bool(X.eq(X.v(app, "use_plumpy_with_mam"), "false")) && X.bool(X.eq(X.v(app, "s_not_give_albendazole"), "out_of_stock"))));
}

/** Relevance of n_plumpy_sup_shortage */
export function rel_n_plumpy_sup_shortage(app) {
  return X.bool(X.eq(X.v(app, "s_not_give_plumpy_sup"), "out_of_stock"));
}

/** Relevance of s_dnt_SAM_without_complication */
export function rel_s_dnt_SAM_without_complication(app) {
  return X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true"));
}

/** Relevance of s_SAM_RUTF */
export function rel_s_SAM_RUTF(app) {
  return X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true"));
}

/** Relevance of r_plumpy_nut */
export function rel_r_plumpy_nut(app) {
  return X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true"));
}

/** Relevance of r_plumpy_nut_1 */
export function rel_r_plumpy_nut_1(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")) && X.bool(X.ge(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_weight"), 0)), 3.5)))) && X.bool(X.le(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_weight"), 0)), 3.9))));
}

/** Relevance of r_plumpy_nut_2 */
export function rel_r_plumpy_nut_2(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")) && X.bool(X.ge(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_weight"), 0)), 4)))) && X.bool(X.le(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_weight"), 0)), 5.4))));
}

/** Relevance of r_plumpy_nut_3 */
export function rel_r_plumpy_nut_3(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")) && X.bool(X.ge(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_weight"), 0)), 5.5)))) && X.bool(X.le(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_weight"), 0)), 6.9))));
}

/** Relevance of r_plumpy_nut_4 */
export function rel_r_plumpy_nut_4(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")) && X.bool(X.ge(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_weight"), 0)), 7)))) && X.bool(X.le(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_weight"), 0)), 8.4))));
}

/** Relevance of r_plumpy_nut_5 */
export function rel_r_plumpy_nut_5(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")) && X.bool(X.ge(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_weight"), 0)), 8.5)))) && X.bool(X.le(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_weight"), 0)), 9.4))));
}

/** Relevance of r_plumpy_nut_6 */
export function rel_r_plumpy_nut_6(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")) && X.bool(X.ge(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_weight"), 0)), 9.5)))) && X.bool(X.le(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_weight"), 0)), 10.4))));
}

/** Relevance of r_plumpy_nut_7 */
export function rel_r_plumpy_nut_7(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")) && X.bool(X.ge(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_weight"), 0)), 10.5)))) && X.bool(X.le(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_weight"), 0)), 11.9))));
}

/** Relevance of r_plumpy_nut_8 */
export function rel_r_plumpy_nut_8(app) {
  return X.bool((X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")) && X.bool(X.ge(X.f["number"](app, X.f["coalesce"](app, X.v(app, "s_child_weight"), 0)), 12))));
}

/** Relevance of s_give_plumpy_nut */
export function rel_s_give_plumpy_nut(app) {
  return X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true"));
}

/** Relevance of s_give_plumpy_nut_dosage */
export function rel_s_give_plumpy_nut_dosage(app) {
  return X.bool(X.eq(X.v(app, "s_give_plumpy_nut"), "yes"));
}

/** Constraint of s_give_plumpy_nut_dosage */
export function val_s_give_plumpy_nut_dosage(app) {
  if (X.str(X.v(app, "s_give_plumpy_nut_dosage")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "s_give_plumpy_nut_dosage"), 1)) && X.bool(X.le(X.v(app, "s_give_plumpy_nut_dosage"), 99))));
}

/** Relevance of s_not_give_plumpy_nut */
export function rel_s_not_give_plumpy_nut(app) {
  return X.bool(X.eq(X.v(app, "s_give_plumpy_nut"), "no"));
}

/** Relevance of s_not_give_plumpy_nut_other */
export function rel_s_not_give_plumpy_nut_other(app) {
  return X.bool(X.eq(X.v(app, "s_not_give_plumpy_nut"), "other"));
}

/** Relevance of s_dnt_ari */
export function rel_s_dnt_ari(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_ari_have_cough"), "yes")) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_dnt_ari_note_1 */
export function rel_s_dnt_ari_note_1(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "s_ari_have_cough"), "yes")) && X.bool(X.eq(X.v(app, "fast_breathing"), "false")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_ari_receive_honey */
export function rel_s_ari_receive_honey(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "s_ari_have_cough"), "yes")) && X.bool(X.eq(X.v(app, "fast_breathing"), "false")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_ari_receive_eucalyptus */
export function rel_s_ari_receive_eucalyptus(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "s_ari_have_cough"), "yes")) && X.bool(X.eq(X.v(app, "fast_breathing"), "false")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_ari_receive_lemon */
export function rel_s_ari_receive_lemon(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "s_ari_have_cough"), "yes")) && X.bool(X.eq(X.v(app, "fast_breathing"), "false")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of r_ari_amoxicillin */
export function rel_r_ari_amoxicillin(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "fast_breathing"), "true")) || X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_dnt_ari_note_2 */
export function rel_s_dnt_ari_note_2(app) {
  return X.bool((X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)) && X.bool(X.lt(X.v(app, "patient_age_in_years"), 1)))) && X.bool((X.bool(X.eq(X.v(app, "fast_breathing"), "true")) || X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")))))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_dnt_ari_note_3 */
export function rel_s_dnt_ari_note_3(app) {
  return X.bool((X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_years"), 1)) && X.bool(X.lt(X.v(app, "patient_age_in_years"), 5)))) && X.bool((X.bool(X.eq(X.v(app, "fast_breathing"), "true")) || X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")))))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_ari_give_amox */
export function rel_s_ari_give_amox(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "fast_breathing"), "true")) || X.bool(X.eq(X.v(app, "has_SAM_without_complication"), "true")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_ari_give_amox_dosage */
export function rel_s_ari_give_amox_dosage(app) {
  return X.bool(X.eq(X.v(app, "s_ari_give_amox"), "yes"));
}

/** Relevance of s_ari_not_give_amox */
export function rel_s_ari_not_give_amox(app) {
  return X.bool(X.eq(X.v(app, "s_ari_give_amox"), "no"));
}

/** Relevance of s_ari_not_give_amox_other */
export function rel_s_ari_not_give_amox_other(app) {
  return X.bool(X.eq(X.v(app, "s_ari_not_give_amox"), "other"));
}

/** Relevance of n_amoxicillin_shortage */
export function rel_n_amoxicillin_shortage(app) {
  return X.bool(X.eq(X.v(app, "s_ari_not_give_amox"), "out_of_stock"));
}

/** Relevance of s_dnt_fever */
export function rel_s_dnt_fever(app) {
  return X.bool(X.eq(X.v(app, "s_fever"), "yes"));
}

/** Relevance of r_fever_paracetamol */
export function rel_r_fever_paracetamol(app) {
  return X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)) && X.bool(X.eq(X.v(app, "s_fever"), "yes"))));
}

/** Relevance of r_fever_paracetamol_1 */
export function rel_r_fever_paracetamol_1(app) {
  return X.bool((X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)) && X.bool(X.lt(X.v(app, "patient_age_in_years"), 1)))) && X.bool(X.eq(X.v(app, "s_fever"), "yes")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")) || X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true"))))));
}

/** Relevance of r_fever_paracetamol_1_no_reference */
export function rel_r_fever_paracetamol_1_no_reference(app) {
  return X.bool((X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)) && X.bool(X.lt(X.v(app, "patient_age_in_years"), 1)))) && X.bool(X.eq(X.v(app, "s_fever"), "yes")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of r_fever_paracetamol_2 */
export function rel_r_fever_paracetamol_2(app) {
  return X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_years"), 1)) && X.bool(X.eq(X.v(app, "s_fever"), "yes")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")) || X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true"))))));
}

/** Relevance of r_fever_paracetamol_2_no_reference */
export function rel_r_fever_paracetamol_2_no_reference(app) {
  return X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_years"), 1)) && X.bool(X.eq(X.v(app, "s_fever"), "yes")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_give_paracetamol */
export function rel_s_give_paracetamol(app) {
  return X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)) && X.bool(X.eq(X.v(app, "s_fever"), "yes"))));
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

/** Relevance of n_paracetamol_shortage */
export function rel_n_paracetamol_shortage(app) {
  return X.bool(X.eq(X.v(app, "s_not_give_paracetamol"), "out_of_stock"));
}

/** Relevance of s_dnt_malaria */
export function rel_s_dnt_malaria(app) {
  return X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)) && X.bool(X.eq(X.v(app, "tdr_result"), "positif")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of r_malaria_act */
export function rel_r_malaria_act(app) {
  return X.bool((X.bool(X.eq(X.v(app, "tdr_result"), "positif")) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of r_malaria_act_1 */
export function rel_r_malaria_act_1(app) {
  return X.bool((X.bool((X.bool(X.lt(X.v(app, "patient_age_in_years"), 4)) && X.bool(X.eq(X.v(app, "tdr_result"), "positif")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of r_malaria_act_2 */
export function rel_r_malaria_act_2(app) {
  return X.bool((X.bool((X.bool((X.bool(X.ge(X.v(app, "patient_age_in_years"), 4)) && X.bool(X.lt(X.v(app, "patient_age_in_years"), 5)))) && X.bool(X.eq(X.v(app, "tdr_result"), "positif")))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
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

/** Relevance of n_act_shortage */
export function rel_n_act_shortage(app) {
  return X.bool(X.eq(X.v(app, "s_malaria_not_give_act"), "out_of_stock"));
}

/** Relevance of s_dnt_diarrhea */
export function rel_s_dnt_diarrhea(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_diarrhea_stools_a_day"), "yes")) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of r_diarrhea_zinc */
export function rel_r_diarrhea_zinc(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_diarrhea_stools_a_day"), "yes")) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_dnt_diarrhea_note_1 */
export function rel_s_dnt_diarrhea_note_1(app) {
  return X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_diarrhea_stools_a_day"), "yes")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)))) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 6)))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_dnt_diarrhea_note_2 */
export function rel_s_dnt_diarrhea_note_2(app) {
  return X.bool((X.bool((X.bool(X.eq(X.v(app, "s_diarrhea_stools_a_day"), "yes")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 6)))) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_diarrhea_give_zinc */
export function rel_s_diarrhea_give_zinc(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_diarrhea_stools_a_day"), "yes")) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_diarrhea_zinc_dosage */
export function rel_s_diarrhea_zinc_dosage(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_diarrhea_give_zinc"), "yes")) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_diarrhea_not_give_zinc */
export function rel_s_diarrhea_not_give_zinc(app) {
  return X.bool(X.eq(X.v(app, "s_diarrhea_give_zinc"), "no"));
}

/** Relevance of s_diarrhea_not_give_zinc_other */
export function rel_s_diarrhea_not_give_zinc_other(app) {
  return X.bool(X.eq(X.v(app, "s_diarrhea_not_give_zinc"), "other"));
}

/** Relevance of n_zinc_shortage */
export function rel_n_zinc_shortage(app) {
  return X.bool(X.eq(X.v(app, "s_diarrhea_not_give_zinc"), "out_of_stock"));
}

/** Relevance of r_diarrhea_ors */
export function rel_r_diarrhea_ors(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_diarrhea_stools_a_day"), "yes")) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_dnt_diarrhea_note_3 */
export function rel_s_dnt_diarrhea_note_3(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_diarrhea_stools_a_day"), "yes")) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_diarrhea_give_ors */
export function rel_s_diarrhea_give_ors(app) {
  return X.bool((X.bool(X.eq(X.v(app, "s_diarrhea_stools_a_day"), "yes")) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_give_ors_dosage */
export function rel_s_give_ors_dosage(app) {
  return X.bool(X.eq(X.v(app, "s_diarrhea_give_ors"), "yes"));
}

/** Constraint of s_give_ors_dosage */
export function val_s_give_ors_dosage(app) {
  if (X.str(X.v(app, "s_give_ors_dosage")) === '') return true;
  return X.bool((X.bool(X.ge(X.v(app, "s_give_ors_dosage"), 1)) && X.bool(X.le(X.v(app, "s_give_ors_dosage"), 99))));
}

/** Relevance of s_diarrhea_not_give_ors */
export function rel_s_diarrhea_not_give_ors(app) {
  return X.bool(X.eq(X.v(app, "s_diarrhea_give_ors"), "no"));
}

/** Relevance of s_diarrhea_not_give_ors_other */
export function rel_s_diarrhea_not_give_ors_other(app) {
  return X.bool(X.eq(X.v(app, "s_diarrhea_not_give_ors"), "other"));
}

/** Relevance of n_ors_shortage */
export function rel_n_ors_shortage(app) {
  return X.bool(X.eq(X.v(app, "s_diarrhea_not_give_ors"), "out_of_stock"));
}

/** Relevance of s_referral_protocol */
export function rel_s_referral_protocol(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool(X.eq(X.v(app, "tdr_result"), "positif")) || X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "true")) || X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "true"))))));
}

/** Relevance of s_place_antimalarial_suppository */
export function rel_s_place_antimalarial_suppository(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool((X.bool(X.eq(X.v(app, "tdr_result"), "positif")) && X.bool(X.eq(X.v(app, "has_danger_sign"), "true")))) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 2))));
}

/** Relevance of s_did_place_antimalarial_suppository */
export function rel_s_did_place_antimalarial_suppository(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool((X.bool(X.eq(X.v(app, "tdr_result"), "positif")) && X.bool(X.eq(X.v(app, "has_danger_sign"), "true")))) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 2))));
}

/** Relevance of s_antimalarial_suppository_not_give */
export function rel_s_antimalarial_suppository_not_give(app) {
  return grp_referral_and_protocol(app) && X.bool(X.eq(X.v(app, "s_did_place_antimalarial_suppository"), "no"));
}

/** Relevance of s_antimalarial_suppository_not_give_other */
export function rel_s_antimalarial_suppository_not_give_other(app) {
  return grp_referral_and_protocol(app) && X.bool(X.eq(X.v(app, "s_antimalarial_suppository_not_give"), "other"));
}

/** Relevance of n_artesunate_shortage */
export function rel_n_artesunate_shortage(app) {
  return grp_referral_and_protocol(app) && X.bool(X.eq(X.v(app, "s_antimalarial_suppository_not_give"), "out_of_stock"));
}

/** Relevance of s_antimalarial_suppository_type */
export function rel_s_antimalarial_suppository_type(app) {
  return grp_referral_and_protocol(app) && X.bool(X.eq(X.v(app, "s_did_place_antimalarial_suppository"), "yes"));
}

/** Relevance of s_artesunate_suppository_type */
export function rel_s_artesunate_suppository_type(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool(X.eq(X.v(app, "s_did_place_antimalarial_suppository"), "yes")) && X.bool(X.eq(X.v(app, "s_antimalarial_suppository_type"), "artesunate_suppository"))));
}

/** Constraint of s_artesunate_suppository_type */
export function val_s_artesunate_suppository_type(app) {
  if (X.str(X.v(app, "s_artesunate_suppository_type")) === '') return true;
  return X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "s_artesunate_suppository_type"), "artesunate_suppository_50mg")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)))) || X.bool((X.bool(X.eq(X.v(app, "s_artesunate_suppository_type"), "artesunate_suppository_100mg")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 24)))))) || X.bool((X.bool(X.eq(X.v(app, "s_artesunate_suppository_type"), "artesunate_suppository_200mg")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 48))))));
}

/** Relevance of n_artesunate_suppository_50mg_1 */
export function rel_n_artesunate_suppository_50mg_1(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool((X.bool(X.eq(X.v(app, "s_artesunate_suppository_type"), "artesunate_suppository_50mg")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)))) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 24))));
}

/** Relevance of n_artesunate_suppository_50mg_2 */
export function rel_n_artesunate_suppository_50mg_2(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool((X.bool(X.eq(X.v(app, "s_artesunate_suppository_type"), "artesunate_suppository_50mg")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 24)))) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 48))));
}

/** Relevance of n_artesunate_suppository_50mg_3 */
export function rel_n_artesunate_suppository_50mg_3(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool((X.bool(X.eq(X.v(app, "s_artesunate_suppository_type"), "artesunate_suppository_50mg")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 48)))) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 60))));
}

/** Relevance of n_artesunate_suppository_100mg_1 */
export function rel_n_artesunate_suppository_100mg_1(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool((X.bool(X.eq(X.v(app, "s_artesunate_suppository_type"), "artesunate_suppository_100mg")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 24)))) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 48))));
}

/** Relevance of n_artesunate_suppository_100mg_2 */
export function rel_n_artesunate_suppository_100mg_2(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool((X.bool(X.eq(X.v(app, "s_artesunate_suppository_type"), "artesunate_suppository_100mg")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 48)))) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 60))));
}

/** Relevance of n_artesunate_suppository_200mg_1 */
export function rel_n_artesunate_suppository_200mg_1(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool((X.bool(X.eq(X.v(app, "s_artesunate_suppository_type"), "artesunate_suppository_200mg")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 48)))) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 60))));
}

/** Relevance of s_did_place_artesunate_suppository */
export function rel_s_did_place_artesunate_suppository(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool(X.eq(X.v(app, "s_did_place_antimalarial_suppository"), "yes")) && X.bool(X.eq(X.v(app, "s_antimalarial_suppository_type"), "artesunate_suppository"))));
}

/** Relevance of s_artesunate_suppository_not_give */
export function rel_s_artesunate_suppository_not_give(app) {
  return grp_referral_and_protocol(app) && X.bool(X.eq(X.v(app, "s_did_place_artesunate_suppository"), "no"));
}

/** Relevance of s_artesunate_suppository_not_give_other */
export function rel_s_artesunate_suppository_not_give_other(app) {
  return grp_referral_and_protocol(app) && X.bool(X.eq(X.v(app, "s_artesunate_suppository_not_give"), "other"));
}

/** Relevance of s_artesunate_suppository_dosage */
export function rel_s_artesunate_suppository_dosage(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool(X.eq(X.v(app, "s_did_place_artesunate_suppository"), "yes")) && X.bool(X.eq(X.v(app, "s_antimalarial_suppository_type"), "artesunate_suppository"))));
}

/** Options of s_artesunate_suppository_dosage */
export function filt_s_artesunate_suppository_dosage(app) {
  return X.bool(X.eq(X.opt(app, "filter1"), X.v(app, "s_artesunate_suppository_type")));
}

/** Relevance of n_artemether_suppository_40mg_1 */
export function rel_n_artemether_suppository_40mg_1(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool((X.bool(X.eq(X.v(app, "s_antimalarial_suppository_type"), "arthemeter_suppository")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 2)))) && X.bool(X.lt(X.v(app, "patient_age_in_months"), 12))));
}

/** Relevance of n_artemether_suppository_40mg_2 */
export function rel_n_artemether_suppository_40mg_2(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool((X.bool(X.eq(X.v(app, "s_antimalarial_suppository_type"), "arthemeter_suppository")) && X.bool(X.ge(X.v(app, "patient_age_in_months"), 12)))) && X.bool(X.le(X.v(app, "patient_age_in_months"), 60))));
}

/** Relevance of s_did_place_artemether_suppository */
export function rel_s_did_place_artemether_suppository(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool(X.eq(X.v(app, "s_did_place_antimalarial_suppository"), "yes")) && X.bool(X.eq(X.v(app, "s_antimalarial_suppository_type"), "arthemeter_suppository"))));
}

/** Relevance of s_artemether_suppository_not_give */
export function rel_s_artemether_suppository_not_give(app) {
  return grp_referral_and_protocol(app) && X.bool(X.eq(X.v(app, "s_did_place_artemether_suppository"), "no"));
}

/** Relevance of s_artemether_suppository_not_give_other */
export function rel_s_artemether_suppository_not_give_other(app) {
  return grp_referral_and_protocol(app) && X.bool(X.eq(X.v(app, "s_artemether_suppository_not_give"), "other"));
}

/** Relevance of s_artemether_suppository_dosage */
export function rel_s_artemether_suppository_dosage(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool(X.eq(X.v(app, "s_did_place_artemether_suppository"), "yes")) && X.bool(X.eq(X.v(app, "s_antimalarial_suppository_type"), "arthemeter_suppository"))));
}

/** Relevance of s_temperature_exceeds_38_enveropment */
export function rel_s_temperature_exceeds_38_enveropment(app) {
  return grp_referral_and_protocol(app) && X.bool(X.eq(X.v(app, "s_fever"), "yes"));
}

/** Relevance of s_breastfeed_a_child */
export function rel_s_breastfeed_a_child(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool(X.eq(X.v(app, "has_danger_sign"), "true")) || X.bool(X.eq(X.v(app, "has_ref_danger_sign"), "true"))));
}

/** Relevance of s_rp_admin_gentamycine */
export function rel_s_rp_admin_gentamycine(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool(X.le(X.v(app, "patient_age_in_days"), 28)) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_rp_admin_ampiciline */
export function rel_s_rp_admin_ampiciline(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool(X.le(X.v(app, "patient_age_in_days"), 28)) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_rp_did_admin_gentamycine */
export function rel_s_rp_did_admin_gentamycine(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool(X.le(X.v(app, "patient_age_in_days"), 28)) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_rp_did_admin_ampiciline */
export function rel_s_rp_did_admin_ampiciline(app) {
  return grp_referral_and_protocol(app) && X.bool((X.bool(X.le(X.v(app, "patient_age_in_days"), 28)) && X.bool((X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false"))))));
}

/** Relevance of s_rp_did_cover_clothes */
export function rel_s_rp_did_cover_clothes(app) {
  return grp_referral_and_protocol(app) && X.bool(X.eq(X.v(app, "s_fever"), "yes"));
}

/** Relevance of CAT_out_of_stock */
export function rel_CAT_out_of_stock(app) {
  return X.bool(X.eq(X.v(app, "show_drugs_cat"), "true"));
}

/** Relevance of s_accompany_refer_CSCOM */
export function rel_s_accompany_refer_CSCOM(app) {
  return X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "referral"), "true")) || X.bool(X.eq(X.v(app, "CAT_out_of_stock"), "refer")))) || X.bool(X.eq(X.v(app, "new_CAT"), "refer")))) || X.bool(X.eq(X.v(app, "is_tdr_unavailable"), "true"))));
}

/** Relevance of n_on_observation */
export function rel_n_on_observation(app) {
  return X.bool(X.eq(X.v(app, "new_CAT"), "watching"));
}

/** Relevance of s_next_covid_mal_multiple */
export function rel_s_next_covid_mal_multiple(app) {
  return X.bool((X.bool(X.eq(X.v(app, "diff_assess_date_report_date"), 0)) && X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true")) || X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")))) || X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "treat_malaria"), "true")) || X.bool((X.bool(X.eq(X.v(app, "treat_malaria"), "true")) && X.bool(X.eq(X.v(app, "treat_diarrhea"), "true")))))) || X.bool((X.bool(X.eq(X.v(app, "treat_ari"), "true")) && X.bool(X.eq(X.v(app, "treat_diarrhea"), "true")))))) || X.bool((X.bool(X.eq(X.v(app, "treat_malaria"), "true")) && X.bool(X.eq(X.v(app, "treat_ari"), "true")))))))) || X.bool(X.eq(X.v(app, "tb_referral"), "true"))))));
}

/** Relevance of s_next_covid_mal_multiple_remainder */
export function rel_s_next_covid_mal_multiple_remainder(app) {
  return X.bool((X.bool(X.ne(X.v(app, "diff_assess_date_report_date"), 0)) && X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "accompany_to_cscom"), "true")) || X.bool(X.eq(X.v(app, "refer_to_cscom"), "true")))) || X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "treat_malaria"), "true")) || X.bool((X.bool(X.eq(X.v(app, "treat_malaria"), "true")) && X.bool(X.eq(X.v(app, "treat_diarrhea"), "true")))))) || X.bool((X.bool(X.eq(X.v(app, "treat_ari"), "true")) && X.bool(X.eq(X.v(app, "treat_diarrhea"), "true")))))) || X.bool((X.bool(X.eq(X.v(app, "treat_malaria"), "true")) && X.bool(X.eq(X.v(app, "treat_ari"), "true")))))))) || X.bool(X.eq(X.v(app, "tb_referral"), "true"))))));
}

/** Relevance of s_next_followup_diarhea */
export function rel_s_next_followup_diarhea(app) {
  return X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.le(X.v(app, "diff_assess_date_report_date"), 5)) && X.bool(X.eq(X.v(app, "treat_diarrhea"), "true")))) && X.bool(X.eq(X.v(app, "treat_malaria"), "false")))) && X.bool(X.eq(X.v(app, "treat_ari"), "false")))) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false")))) && X.bool(X.eq(X.v(app, "refer_to_cscom"), "false"))));
}

/** Relevance of s_last_followup_diarhea */
export function rel_s_last_followup_diarhea(app) {
  return X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.gt(X.v(app, "diff_assess_date_report_date"), 5)) && X.bool(X.eq(X.v(app, "treat_diarrhea"), "true")))) && X.bool(X.eq(X.v(app, "treat_malaria"), "false")))) && X.bool(X.eq(X.v(app, "treat_ari"), "false")))) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false")))) && X.bool(X.eq(X.v(app, "refer_to_cscom"), "false"))));
}

/** Relevance of s_next_visit_ari */
export function rel_s_next_visit_ari(app) {
  return X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "treat_ari"), "true")) && X.bool(X.eq(X.v(app, "treat_malaria"), "false")))) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false")))) && X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")))) && X.bool(X.eq(X.v(app, "treat_diarrhea"), "false")))) && X.bool((X.bool(X.eq(X.v(app, "diff_assess_date_report_date"), 0)) || X.bool(X.eq(X.v(app, "diff_assess_date_report_date"), 4))))));
}

/** Relevance of s_next_visit_ari_remainder */
export function rel_s_next_visit_ari_remainder(app) {
  return X.bool((X.bool((X.bool((X.bool((X.bool((X.bool((X.bool(X.eq(X.v(app, "treat_ari"), "true")) && X.bool(X.eq(X.v(app, "treat_malaria"), "false")))) && X.bool(X.eq(X.v(app, "accompany_to_cscom"), "false")))) && X.bool(X.eq(X.v(app, "refer_to_cscom"), "false")))) && X.bool(X.eq(X.v(app, "treat_diarrhea"), "false")))) && X.bool(X.ne(X.v(app, "diff_assess_date_report_date"), 0)))) && X.bool(X.ne(X.v(app, "diff_assess_date_report_date"), 4))));
}

/** Relevance of s_next_followup_malnutrition */
export function rel_s_next_followup_malnutrition(app) {
  return X.bool((X.bool(X.eq(X.v(app, "treat_MAM"), "true")) || X.bool(X.eq(X.v(app, "treat_SAM_without_complication"), "true"))));
}

/** Calculation of albendazol_200_disp */
export function calc_albendazol_200_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_give_albendazole_type"), "200_mg_albendazole")) ? (X.bool(X.eq(X.v(app, "s_given_albendazole_dosage"), "2_tab_albendazole")) ? 2 : 1) : 0));
}

/** Calculation of albendazol_400_disp */
export function calc_albendazol_400_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_give_albendazole_type"), "400_mg_albendazole")) ? 1 : 0));
}

/** Calculation of plumpy_sup_disp */
export function calc_plumpy_sup_disp(app) {
  return X.out(X.f["coalesce"](app, X.v(app, "s_give_plumpy_sup_dosage"), 0));
}

/** Calculation of plumpy_nut_disp */
export function calc_plumpy_nut_disp(app) {
  return X.out(X.f["coalesce"](app, X.v(app, "s_give_plumpy_nut_dosage"), 0));
}

/** Calculation of enriched_flour_disp */
export function calc_enriched_flour_disp(app) {
  return X.out(X.f["coalesce"](app, X.v(app, "s_give_enriched_flour_dosage"), 0));
}

/** Calculation of amoxicillin_250_disp */
export function calc_amoxicillin_250_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_ari_give_amox"), "yes")) ? (X.bool(X.eq(X.v(app, "s_ari_give_amox_dosage"), "5_ml_amox")) ? 1 : 2) : 0));
}

/** Calculation of vitamin_a_blue_capsule_disp */
export function calc_vitamin_a_blue_capsule_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_give_vitamina_dosage"), "one_of_100000_ui_blue")) ? 1 : 0));
}

/** Calculation of vitamin_a_red_capsule_disp */
export function calc_vitamin_a_red_capsule_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_give_vitamina_dosage"), "one_of_200000_ui_red")) ? 1 : 0));
}

/** Calculation of act_20120_disp */
export function calc_act_20120_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_malaria_give_act"), "yes")) ? (X.bool(X.eq(X.v(app, "s_malaria_act_dosage"), "2_tablets_act")) ? (X.num(2) * X.num(3)) : (X.num(4) * X.num(3))) : 0));
}

/** Calculation of malaria_rdt_disp */
export function calc_malaria_rdt_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "tdr_done"), "yes")) ? 1 : 0));
}

/** Calculation of artesunate_50_disp */
export function calc_artesunate_50_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_artesunate_suppository_dosage"), "1_artesunate_50mg")) ? 1 : (X.bool(X.eq(X.v(app, "s_artesunate_suppository_dosage"), "2_artesunate_50mg")) ? 2 : (X.bool(X.eq(X.v(app, "s_artesunate_suppository_dosage"), "4_artesunate_50mg")) ? 4 : 0))));
}

/** Calculation of artesunate_100_disp */
export function calc_artesunate_100_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_artesunate_suppository_dosage"), "1_artesunate_100mg")) ? 1 : (X.bool(X.eq(X.v(app, "s_artesunate_suppository_dosage"), "2_artesunate_100mg")) ? 2 : (X.bool(X.eq(X.v(app, "s_artesunate_suppository_dosage"), "4_artesunate_100mg")) ? 4 : 0))));
}

/** Calculation of artesunate_200_disp */
export function calc_artesunate_200_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_artesunate_suppository_dosage"), "1_artesunate_200mg")) ? 1 : (X.bool(X.eq(X.v(app, "s_artesunate_suppository_dosage"), "2_artesunate_200mg")) ? 2 : (X.bool(X.eq(X.v(app, "s_artesunate_suppository_dosage"), "4_artesunate_200mg")) ? 4 : 0))));
}

/** Calculation of arthemeter_40_disp */
export function calc_arthemeter_40_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_artemether_suppository_dosage"), "1_artemether_40mg")) ? 1 : (X.bool(X.eq(X.v(app, "s_artemether_suppository_dosage"), "2_artemether_40mg")) ? 2 : 0)));
}

/** Calculation of zinc_20_disp */
export function calc_zinc_20_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_diarrhea_give_zinc"), "yes")) ? (X.bool(X.eq(X.v(app, "s_diarrhea_zinc_dosage"), "c_diarrhea_zinc_dosage_2")) ? 5 : 10) : 0));
}

/** Calculation of ors_disp */
export function calc_ors_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_diarrhea_give_ors"), "yes")) ? X.f["coalesce"](app, X.v(app, "s_give_ors_dosage"), 0) : 0));
}

/** Calculation of paracetamol_500_disp */
export function calc_paracetamol_500_disp(app) {
  return X.out((X.bool(X.eq(X.v(app, "s_give_paracetamol_dosage"), "quarter_tab_paracetamol_once")) ? (X.num(1) / X.num(4)) : (X.bool(X.eq(X.v(app, "s_give_paracetamol_dosage"), "half_tab_paracetamol_once")) ? (X.num(1) / X.num(2)) : (X.bool(X.eq(X.v(app, "s_give_paracetamol_dosage"), "quarter_tab_paracetamol_3x_a_day")) ? (X.num((X.num(1) / X.num(4))) * X.num(3)) : (X.bool(X.eq(X.v(app, "s_give_paracetamol_dosage"), "half_tab_paracetamol_3x_a_day")) ? (X.num((X.num(1) / X.num(2))) * X.num(3)) : 0)))));
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
  return X.out("patient_assessment");
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
