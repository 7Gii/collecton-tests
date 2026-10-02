// Ported from targets.js / targets.extras.js of the CHT configuration muso-mali. The CHT
// functions below are copied verbatim; this adapter gives them the CHT shapes:
// report { _id, form, reported_date, fields } and contact { contact, reports }.

const DAY = 86400000;
let NOW = Date.now();

const Forms = {
  PATIENT_ASSESSMENT_UNDER_5: 'patient_assessment',
  PATIENT_ASSESSMENT_OVER_5: 'patient_assessment_over_5',
  TREATMENT_FOLLOWUP_UNDER_5: 'treatment_followup',
  TREATMENT_FOLLOWUP_OVER_5: 'treatment_followup_over_5',
  REFERRAL_FOLLOWUP_UNDER_5: 'referral_followup_under_5',
  REFERRAL_FOLLOWUP_OVER_5: 'referral_followup',
  MODERATE_MALNUTRITION_FOLLOWUP: 'moderate_malnutrition_followup',
  SEVERE_MALNUTRITION_FOLLOWUP: 'severe_malnutrition_followup',
};

/**
 * CHT getField(report, 'group.sub.name') on the flat Collecton payload: the key of a field
 * is its name, or <group>_<name> when the name is used twice in the form.
 */
const getField = (report, fieldPath) => {
  const payload = report?.fields;
  if (!payload) return undefined;
  const parts = String(fieldPath || '').split('.');
  for (let i = 0; i < parts.length; i++) {
    const key = parts.slice(i).join('_');
    if (payload[key] !== undefined) return payload[key];
  }
  return undefined;
};

const isDefined = (value) => value !== null && value !== undefined;

/** Luxon-like date of the CHT code (toMillis, toJSDate, diff in days, comparisons). */
const dt = (millis) => ({
  valueOf: () => millis,
  toMillis: () => millis,
  toJSDate: () => new Date(millis),
  diff: (other, unit) => ({ days: (millis - other.valueOf()) / DAY }),
  toFormat: () => new Date(millis).toISOString().slice(0, 10),
});

/** CHT getReportHomeVisitSubject / getReportHomeVisitDate. */
const getReportHomeVisitSubject = (report) => report && getField(report, 'visited_contact_uuid');
const getReportHomeVisitDate = (report) => {
  if (!getReportHomeVisitSubject(report)) return false;
  const visited = Date.parse(getField(report, 'visited_date') || '');
  return dt(Number.isNaN(visited) ? Number(report.reported_date || 0) : visited);
};

/** CHT getDynamicReportedDate. */
const getDynamicReportedDate = (report) => {
  const home = getReportHomeVisitDate(report);
  if (home) return home;
  const specified = Date.parse(getField(report, 's_reported.s_reported_date') || getField(report, 'supervision_date') || '');
  return dt(Number.isNaN(specified) ? Number((report && report.reported_date) || 0) : specified);
};

/** CHT getNumberOfDaySinceDate, on the app clock. */
function getNumberOfDaySinceDate(specificDate) {
  return Math.floor((NOW - specificDate) / DAY);
}

// ---- Collecton side ----

const formNameById = {};
const toCht = (r) => ({ _id: r.id, form: formNameById[r.form_id], reported_date: Number(r.created_at), fields: r.payload || {}, target: r.target_id });

/** Muso month of the app clock: day 1 to the last day (UTC), as the collection period. */
const musoMonth = (now) => {
  const d = new Date(now);
  return [Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), 1), Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 1) - 1];
};
const inMonth = (millis, [from, to]) => millis >= from && millis <= to;

const parseAttributes = (row) => {
  try { return typeof row?.attributes === 'string' ? JSON.parse(row.attributes) : (row?.attributes || {}); } catch { return {}; }
};

/** CHT basicTemplate: the person of the report is alive (muted has no equivalent). */
const aliveCache = new Map();
const isAliveId = async (app, personId) => {
  if (!personId) return true;
  if (!aliveCache.has(personId)) {
    const [row] = await app.find('persons', [{ field: 'id', value: personId }]);
    const deceased = parseAttributes(row).person_is_deceased;
    aliveCache.set(personId, !(deceased === true || deceased === 'true'));
  }
  return aliveCache.get(personId);
};

/** Prepares a computation: clock, form names, CHT reports of the month for the given forms. */
const prepare = async (app, forms) => {
  NOW = app.now;
  aliveCache.clear();
  const all = app.getReports();
  const idsByName = {};
  for (const name of Object.values(Forms)) {
    const sample = app.getReports({ form: name })[0];
    if (sample) { formNameById[sample.form_id] = name; idsByName[name] = sample.form_id; }
  }
  const reports = all.map(toCht).filter(r => r.form);
  const month = musoMonth(app.now);
  const targets = [];
  for (const r of reports) {
    if (!forms.includes(r.form)) continue;
    if (!inMonth(getDynamicReportedDate(r).toMillis(), month)) continue;
    if (!(await isAliveId(app, r.target))) continue;
    targets.push(r);
  }
  const byPerson = new Map();
  for (const r of reports) {
    if (!byPerson.has(r.target)) byPerson.set(r.target, []);
    byPerson.get(r.target).push(r);
  }
  const contactOf = (r) => ({ contact: { _id: r.target }, reports: byPerson.get(r.target) || [] });
  return { reports, targets, contactOf, month };
};

/** CHT context: user.role === "chw" (the role of the user's contact) -> Collecton role chw_uhc. */
export function showIndicator(app) {
  return app.currentUser?.role?.name === 'chw_uhc';
}

// ---- CHT functions of targets.extras.js (verbatim) ----

//Danger sign(s) without referral or accompaniment to the first contact health center - CHW error
function isDangerSignNotReferred(f) {
  const hasDangerSign = (getField(f,'s_acc_danger_signs.has_danger_sign')  === 'true') || (getField(f,'has_danger_sign') === 'true');
  const refer =  getField(f,'refer_to_cscom');
  const accompany = getField(f,'accompany_to_cscom');
  return hasDangerSign &&  (refer !== 'true' && accompany !== 'true');
}
//Presence of fever (measured temperature greater than or equal to 37.5), malaria RDT not performed - CHW error
function isFeverWithoutTdrChw(f) {
  const temp = f.form === Forms.PATIENT_ASSESSMENT_UNDER_5 ? parseFloat(getField(f,'s_child_temperature')) : parseFloat(getField(f,'child_temperature'));

  const accValue = getField(f, 's_acc_danger_signs.s_acc_tdr_not_done');
  const malnutritionValue = getField(f, 's_malnutrition_mRDT.malnutrition_tdr_not_done');
  const malariaValue = getField(f, 's_malaria.malaria_tdr_not_done');

  const accNotDone = accValue && accValue !== 'tdr_unavailable';
  const malnutritionNotDone = malnutritionValue && malnutritionValue !== 'tdr_unavailable';
  const malariaNotDone = malariaValue && malariaValue !== 'tdr_unavailable';


  const hasFever = temp >= 37.5;
  const tdrResultIsEmpty = !getField(f, 'tdr_result');


  let reasonIsValid = false;
  if (f.form === Forms.PATIENT_ASSESSMENT_UNDER_5) {
    reasonIsValid = accNotDone || malariaNotDone || malnutritionNotDone;
  } else {
    const sMalariaValue = getField(f,'s_malaria.s_malaria_tdr_not_done');
    reasonIsValid = sMalariaValue && sMalariaValue !== 'tdr_unavailable';
  }

  return hasFever && tdrResultIsEmpty && reasonIsValid;
}
//Uncomplicated malaria without ACT administration - CHW error
function malariaNoCta(f) {
  if (f.form === Forms.PATIENT_ASSESSMENT_UNDER_5){
    return (
      getField(f,'treat_malaria') === 'true' &&
      getField(f,'accompany_to_cscom') !== 'true' &&
      getField(f,'refer_to_cscom') !== 'true' &&
      (getField(f,'group_diagnosis.s_malaria_give_act') === 'no' ) &&
      getField(f,'group_diagnosis.s_malaria_not_give_act') !== 'out_of_stock'
    );
  } else {
    return (
      getField(f,'treat_malaria') === 'true' &&
      getField(f,'referral') !== 'true' &&
      (getField(f,'group_diagnosis.s_malaria_give_act') === 'no') &&
      getField(f,'group_diagnosis.s_malaria_not_give_act') !== 'out_of_stock'
    );
  }
}
//Incorrect dose of ALU given by CHW - CHW error
function incorrectDosageOfAlu(f) {
  let expected = null;
  const dosage = (getField(f,'group_diagnosis.s_malaria_act_dosage') || '').trim();
  if (!isDefined(dosage)) {return false;}

  if (f.form === Forms.PATIENT_ASSESSMENT_UNDER_5){
    if (getField(f,'treat_malaria') !== 'true') {return false;}
    if (getField(f,'accompany_to_cscom') === 'true' || getField(f,'refer_to_cscom') === 'true') {return false;}
    if (getField(f,'group_diagnosis.s_malaria_give_act') !== 'yes') {return false;}
    const ageMonths = parseFloat(getField(f,'patient_age_in_months'));

    if (isNaN(ageMonths)) {return false;}

    if (ageMonths >= 2 && ageMonths < 48) {
      expected = '2_tablets_act';
    } else if (ageMonths >= 48 && ageMonths < 60) {
      expected = '4_tablets_act';
    } else {
      return false;
    }
  } else {
    if (getField(f,'treat_malaria') !== 'true') {return false;}
    if (getField(f,'referral') === 'true') {return false;}
    if (getField(f,'group_diagnosis.s_malaria_give_act') !== 'yes') {return false;}

    const ageYears = parseFloat(getField(f, 'patient_age_in_years'));
    if (isNaN(ageYears)) {return false;}

    if (ageYears >= 5 && ageYears < 7) {
      expected = '4_tablets_act';
    } else if (ageYears >= 7 && ageYears < 11) {
      expected = '6_tablets_act';
    } else if (ageYears >= 11) {
      expected = '8_tablets_act';
    } else {
      return false;
    }
  }
  return dosage !== expected;
}
//Uncomplicated diarrhea without zinc administration - CHW error
function diarrheaWithoutZincErrorChw(f) {
  const stools = getField(f, 's_diarrhea.s_diarrhea_stools_a_day');
  const accompany = getField(f, 'accompany_to_cscom');
  const refer = getField(f, 'refer_to_cscom');
  const giveZinc = getField(f, 'group_diagnosis.s_diarrhea_give_zinc');
  const notGiveZinc = getField(f, 'group_diagnosis.s_diarrhea_not_give_zinc');
  return (
    stools === 'yes' &&
    accompany !== 'true' &&
    refer !== 'true' &&
    giveZinc === 'no' &&
    notGiveZinc !== 'out_of_stock'
  );
}
//Incorrect dose of zinc given by CHW - CHW error
function incorrectDosageOfZinc(f) {
  const stools = getField(f, 's_diarrhea.s_diarrhea_stools_a_day');
  const accompany = getField(f, 'accompany_to_cscom');
  const refer = getField(f, 'refer_to_cscom');
  const giveZinc = getField(f, 'group_diagnosis.s_diarrhea_give_zinc');
  const dosage = getField(f, 'group_diagnosis.s_diarrhea_zinc_dosage');
  const ageMonths = getField(f, 'patient_age_in_months');

  return (
    stools === 'yes' &&
    accompany !== 'true' &&
    refer !== 'true' &&
    giveZinc === 'yes' &&
    isDefined(dosage) &&
    (
      (ageMonths >= 2 && ageMonths < 6 && dosage !== 'c_diarrhea_zinc_dosage_2') ||
      (ageMonths >= 6 && ageMonths < 60 && dosage !== 'c_diarrhea_zinc_dosage_3')
    )
  );
}
//Pneumonia without amoxicillin administration - CHW error
function pneumoniaWithoutAmoxicillinChw(f) {
  const fastBreathing = getField(f, 'fast_breathing');
  const cough = getField(f, 's_acute_respiratory_infection.s_ari_have_cough');
  const accompany = getField(f, 'accompany_to_cscom');
  const refer = getField(f, 'refer_to_cscom');
  const giveAmox = getField(f, 'group_diagnosis.s_ari_give_amox');
  const notGiveAmox = getField(f, 'group_diagnosis.s_ari_not_give_amox');
  return (
    fastBreathing === 'true' &&
    cough === 'yes' &&
    accompany !== 'true' &&
    refer !== 'true' &&
    giveAmox !== 'yes' &&
    notGiveAmox !== 'out_of_stock'
  );
}
//Incorrect dose of vitamin A given by CHW - CHW error
function incorrectDosageOfAmoxicillin(f) {
  const fastBreathing = getField(f, 'fast_breathing');
  const cough = getField(f, 's_acute_respiratory_infection.s_ari_have_cough');
  const accompany = getField(f, 'accompany_to_cscom');
  const refer = getField(f, 'refer_to_cscom');
  const giveAmox = getField(f, 'group_diagnosis.s_ari_give_amox');
  const dosage = getField(f, 'group_diagnosis.s_ari_give_amox_dosage');
  const ageMonths = getField(f, 'patient_age_in_months');

  const pneumoniaWithoutAmox =
    fastBreathing === 'true' &&
    cough === 'yes' &&
    accompany !== 'true' &&
    refer !== 'true' &&
    giveAmox !== 'yes';

  const incorrectDosage =
    isDefined(dosage) &&
    (
      (ageMonths >= 2 && ageMonths < 12 && dosage !== '5_ml_amox') ||
      (ageMonths >= 12 && ageMonths < 60 && dosage !== '10_ml_amox')
    );
  return pneumoniaWithoutAmox || incorrectDosage;
}
//Dose de vitamine A donnée par l'ASC incorrecte – Erreur ASC
function incorrectDosageOfVitaminA(f) {
  if (getField(f,'use_plumpy_with_mam') !== 'false') {
    return false;
  }
  if (getField(f,'s_malnutrition_observation.has_MAM') !== 'true') {
    return false;
  }

  const giveVitaminA = getField(f,'group_diagnosis.s_give_vitamina');
  const notGiveVitaminA = getField(f,'group_diagnosis.s_not_give_vitamina');
  const dosage = getField(f,'group_diagnosis.s_give_vitamina_dosage');
  const age = getField(f,'patient_age_in_months');

  if (giveVitaminA === 'no' && notGiveVitaminA !== 'out_of_stock') {
    return true;
  }

  if (giveVitaminA === 'yes') {
    if (age >= 6 && age < 12) {
      return dosage !== 'one_of_100000_ui_blue';
    } else if (age >= 12 && age < 60) {
      return dosage !== 'one_of_200000_ui_red';
    }
  }

  return false;
}
//Incorrect dose of Albendazole given by CHW - CHW error
function incorrectDosageOfAlbendazole(f) {
  if (getField(f,'use_plumpy_with_mam') !== 'false') {
    return false;
  }
  if (getField(f,'s_malnutrition_observation.has_MAM') !== 'true') {
    return false;
  }

  const giveAlbendazole = getField(f,'group_diagnosis.s_give_albendazole');
  const notGiveAlbendazole = getField(f,'group_diagnosis.s_not_give_albendazole');
  const albType = getField(f,'group_diagnosis.s_give_albendazole_type');
  const dosage = getField(f,'group_diagnosis.s_given_albendazole_dosage');
  const age = getField(f,'patient_age_in_months');

  if (giveAlbendazole === 'no' && notGiveAlbendazole !== 'out_of_stock') {
    return true;
  }


  if (giveAlbendazole === 'yes') {
    if (albType === '200_mg_albendazole') {
      if (age >= 12 && age < 24) {
        return dosage !== '1_tab_albendazole';
      } else if (age >= 24 && age < 60) {
        return dosage !== '2_tab_albendazole';
      }
    } else if (albType === '400_mg_albendazole') {
      if (age >= 12 && age < 24) {
        return dosage !== 'half_tab_albendazole';
      } else if (age >= 24 && age < 60) {
        return dosage !== '1_tab_albendazole';
      }
    }
  }

  return false;
}
//Iron + Folic Acid dose Incorrectly given by the CHW - CHW error
function incorrectDosageOfFerAndAcideFolique(contact, f) {
  const followups = contact.reports.filter(report => {
    const isCorrectForm = [Forms.MODERATE_MALNUTRITION_FOLLOWUP].includes(report.form);
    const isCorrectSource = getField(report, 'inputs.source_id') === f._id;
    if (!isCorrectForm || !isCorrectSource) {
      return false;
    }
    const type = getField(report, 's_treatment_given.s_iron_folic_type');
    const dosage = getField(report, 's_treatment_given.s_iron_folic_dosage');
    const weight = getField(report, 's_assessment.child_weight');

    if ( !isDefined(type) || !isDefined(dosage) || !isDefined(weight)) {
      return false;
    }

    if (type === 'iron_folic_acid') {
      return (weight < 10.0 && dosage !== 'one_tab') || (weight >= 10.0 && dosage !== 'half_tab');
    }

    if (type === 'iron_folic_acid_srp') {
      return (weight < 10.0 && dosage !== '2_5ml') || (weight >= 10.0 && dosage !== '5ml');
    }
    return false;
  });

  return followups.length > 0 ;
}
//MAM and no complete and correct treatment on the first day of treatment - CHW error
function mamWithIncompletTreatmentFirstDayChw(f) {
  if (getField(f,'use_plumpy_with_mam') !== 'false') {
    return false;
  }
  if (getField(f,'s_malnutrition_observation.has_MAM') !== 'true') {
    return false;
  }

  const giveFlour = getField(f,'group_diagnosis.s_give_enriched_flour');
  const notGiveFlour = getField(f,'group_diagnosis.s_not_give_enriched_flour');
  const flourError = (giveFlour === 'no' && notGiveFlour !== 'out_of_stock');

  const vitaminAError = incorrectDosageOfVitaminA(f);

  const albendazoleError = incorrectDosageOfAlbendazole(f);

  return flourError || vitaminAError || albendazoleError;
}
//Uncomplicated SAM and no complete and correct treatment on the first day of treatment - CHW error
function masWithoutComplicationIncompleteTreatmentFirstDayChw(f) {
  if (getField(f,'s_malnutrition_observation.has_SAM_without_complication') !== 'true') {
    return false;
  }

  //Vérif Plumpy
  const givePlumpy = getField(f,'group_diagnosis.s_give_plumpy_nut');
  const notGivePlumpy = getField(f,'group_diagnosis.not_give_plumpy_nut');
  const plumpyDosage = getField(f,'group_diagnosis.s_not_give_plumpy_nut');
  const weight = getField(f,'s_malnutrition_weight.s_child_weight');

  let plumpyError = false;
  if (givePlumpy === 'no' && notGivePlumpy !== 'out_of_stock') {
    plumpyError = true;
  } else if (givePlumpy === 'yes' && isDefined(weight) && isDefined(plumpyDosage)) {
    if ((weight >= 3.5 && weight < 5.5 && plumpyDosage !== 14.0) ||
        (weight >= 5.5 && weight < 8.5 && plumpyDosage !== 21.0) ||
        (weight >= 8.5 && weight < 12.0 && plumpyDosage !== 28.0) ||
        (weight >= 12.0 && plumpyDosage !== 35.0)) {
      plumpyError = true;
    }
  }

  //Amoxicillin
  const amoxError = (function(f) {
    const giveAmox = getField(f,'group_diagnosis.s_ari_give_amox');
    const notGiveAmox = getField(f,'group_diagnosis.s_ari_not_give_amox');
    const dosage = getField(f,'group_diagnosis.s_ari_give_amox_dosage');
    const age = getField(f,'patient_age_in_months');

    if (giveAmox === 'no' && notGiveAmox !== 'out_of_stock') {return false;}
    if (giveAmox === 'yes' && isDefined(dosage)) {
      if (age >= 2 && age < 12) {return dosage !== '5_ml_amox';}
      if (age >= 12 && age < 60) {return dosage !== '10_ml_amox';}
    }
    return false;
  })(f);

  //Malaria
  const malariaError = (function(f) {
    const tdrDone = getField(f,'s_malnutrition_mRDT.s_malnutrition_tdr_done');
    const tdrResult = getField(f,'s_malnutrition_mRDT.s_malnutrition_tdr_result');

    if (tdrDone === 'no' && getField(f,'s_malnutrition_mRDT.s_malnutrition_tdr_not_done') !== 'tdr_unavailable') {return false;}
    if (tdrDone === 'yes' && tdrResult === 'pos') {
      const giveAct = getField(f,'group_diagnosis.s_malaria_give_act');
      const actDosage = getField(f,'group_diagnosis.s_malaria_act_dosage');
      const age = getField(f,'patient_age_in_months');
      if (giveAct === 'no' && getField(f,'group_diagnosis.s_malaria_not_give_act') !== 'out_of_stock') {return false;}
      if (giveAct === 'yes' && isDefined(actDosage)) {
        if ((age >= 2 && age < 36 && actDosage !== '2_tablets_act') ||
            (age >= 36 && age < 60 && actDosage !== '4_tablets_act')) {return false;}
      }
    }
    return false;
  })(f);

  return plumpyError || amoxError || malariaError;
}
//Incorrect paracetamol dose given by the CHW - CHW error
function incorrectDosageOfParacetamol(f) {
  const age = f.form === Forms.PATIENT_ASSESSMENT_UNDER_5 ? getField(f,'patient_age_in_months') :  getField(f,'patient_age_in_years');
  const temp = f.form === Forms.PATIENT_ASSESSMENT_UNDER_5 ? getField(f,'s_child_temperature') : getField(f,'s_child_temperature');
  
  const refer =  getField(f,'refer_to_cscom');
  const accompany = getField(f,'accompany_to_cscom');
  const givePara = getField(f,'group_diagnosis.s_give_paracetamol');
  const notGivePara = getField(f,'group_diagnosis.s_not_give_paracetamol');
  const dosage = getField(f,'group_diagnosis.s_give_paracetamol_dosage');
  
  if (temp < 37.5 || !isDefined(givePara)) {return false;}
  if (givePara === 'no' && notGivePara !== 'out_of_stock') {return false;}
  if (givePara === 'yes') {
    const referred = (refer === 'true' || accompany === 'true');
    if (f.form === Forms.PATIENT_ASSESSMENT_UNDER_5) {
      if (referred) {
        if (age >= 2 && age < 12) {return dosage !== 'quarter_tab_paracetamol_once';}
        if (age >= 12 && age < 60) {return dosage !== 'half_tab_paracetamol_once';}
      } else {
        if (age >= 2 && age < 12) {return dosage !== 'quarter_tab_paracetamol_3x_a_day';}
        if (age >= 12 && age < 60) {return dosage !== 'half_tab_paracetamol_3x_a_day';}
      }      
    } else if (f.form === Forms.PATIENT_ASSESSMENT_OVER_5) {
      if (referred) {
        if (age >= 5 && age < 7) {return dosage !== 'demi_tab_paracetamol_once';}
        if (age >= 7 && age < 12) {return dosage !== 'three_quarter_tab_paracetamol_once';}
        if (age >= 12 && age < 16) {return dosage !== 'one_tab_paracetamol_once';}
        if (age >= 16) {return dosage !== 'two_tab_paracetamol_once';}
      } else {
        if (age >= 5 && age < 7) {return dosage !== 'demi_tab_paracetamol_3x_a_day';}
        if (age >= 7 && age < 12) {return dosage !== 'three_quarter_tab_paracetamol_3x_a_day';}
        if (age >= 12 && age < 15) {return dosage !== 'one_tab_paracetamol_3x_a_day';}
        if (age >= 15) {return dosage !== 'two_tab_paracetamol_3x_a_day';}
      }
    }
  }
  return false;
}
// No 24-hour follow-up - CHW error
function noMalariaFollowup(contact, f, followupCount) {
  if (getField(f,'treat_malaria') !== 'true') {return false;}
  const assessmentAge = getNumberOfDaySinceDate(getDynamicReportedDate(f).toJSDate());
  if (assessmentAge < followupCount){return false;}

  const treatments = contact.reports.filter(report =>
    [Forms.TREATMENT_FOLLOWUP_UNDER_5, Forms.TREATMENT_FOLLOWUP_OVER_5].includes(report.form) && 
    getField(report, 'inputs.source_id') === f._id &&
    getField(report,'treat_for_malaria') === 'true' &&
    getField(report, 'follow_up_count') === String(followupCount)
  );
  return treatments.length === 0;
}
function noDiarrheaFollowup(contact, f, followupCount) {
  if (getField(f,'treat_diarrhea') !== 'true') {return false;}
  const assessmentAge = getNumberOfDaySinceDate(getDynamicReportedDate(f).toJSDate());
  if (assessmentAge < followupCount){return false;}

  const treatments = contact.reports.filter(report =>
    [Forms.TREATMENT_FOLLOWUP_UNDER_5].includes(report.form) && 
    getField(report, 'inputs.source_id') === f._id &&
    getField(report,'treat_for_diarrhea') === 'true' &&
    getField(report, 'follow_up_count') === String(followupCount)
  );
  return treatments.length === 0 ;
}
function noAriFollowup(contact, f, followupCount) {
  if (getField(f,'treat_ari') !== 'true') {return false;}
  const assessmentAge = getNumberOfDaySinceDate(getDynamicReportedDate(f).toJSDate());
  if (assessmentAge < followupCount){return false;}

  const treatments = contact.reports.filter(report =>
    [Forms.TREATMENT_FOLLOWUP_UNDER_5].includes(report.form) && 
    getField(report, 'inputs.source_id') === f._id &&
    getField(report,'treat_for_ari') === 'true' &&
    getField(report, 'follow_up_count') === String(followupCount)
  );
  return treatments.length === 0 ;
}
function noCoughFollowup(contact, f, followupCount) {
  if (getField(f,'treat_cough') !== 'true') {return false;}
  const assessmentAge = getNumberOfDaySinceDate(getDynamicReportedDate(f).toJSDate());
  if (assessmentAge < followupCount){return false;}

  const treatments = contact.reports.filter(report =>
    [Forms.TREATMENT_FOLLOWUP_UNDER_5].includes(report.form) && 
    getField(report, 'inputs.source_id') === f._id &&
    getField(report,'treat_for_cough') === 'true' &&
    getField(report, 'follow_up_count') === String(followupCount)
  );
  return treatments.length === 0 ;
}
function noMultipleTreatmentFollowup(contact, f, followupCount) {
  let totalTreatment=0;
  totalTreatment = totalTreatment + (noMalariaFollowup(contact, f, followupCount) === true ? 1: 0);
  totalTreatment = totalTreatment + (noDiarrheaFollowup(contact, f, followupCount) === true ? 1: 0);
  totalTreatment = totalTreatment + (noAriFollowup(contact, f, followupCount) === true ? 1: 0);
  totalTreatment = totalTreatment + (noCoughFollowup(contact, f, followupCount) === true ? 1: 0);
  return totalTreatment > 1;
}
//Condition worsened at 24, 48, 72, or 5 days without referral/accompaniment to the health center - CHW error
function aggravatedWithoutReferral(contact,f, followupCount) {
  const assessmentAge = getNumberOfDaySinceDate(getDynamicReportedDate(f).toJSDate());
  if (assessmentAge < followupCount){return false;}
  
  const treatments = contact.reports.filter(report =>
    [Forms.TREATMENT_FOLLOWUP_UNDER_5, Forms.TREATMENT_FOLLOWUP_OVER_5].includes(report.form) && 
  getField(report, 'inputs.source_id') === f._id &&
  getField(report, 's_patient_available.s_patient_present') !== 'no' &&
  getField(report, 's_disease_progression.s_how_disease_progressing') === 'aggravated' &&
  getField(report, 'referral') === 'true' &&
  report.form ===  Forms.TREATMENT_FOLLOWUP_OVER_5 ? getField(f, 'group_diagnosis.s_accompany_refer_over5') : getField(f, 'group_diagnosis.s_accompany_refer_under5') === 'no' &&
  getField(f, 'follow_up_count') === String(followupCount)
  );
  return treatments.length > 0 ;
}

/** % of the evaluations of the month where the error is present (CHT passesIf, goal 0). */
export async function value(app) {
  const { targets, contactOf } = await prepare(app, [Forms.PATIENT_ASSESSMENT_UNDER_5]);
  const passesIf = (c, f) => diarrheaWithoutZincErrorChw(f);
  return { numerator: targets.filter(f => passesIf(contactOf(f), f)).length, denominator: targets.length };
}

/** CHT goal. */
export function goal() {
  return 0;
}
