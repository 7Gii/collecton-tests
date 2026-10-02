// Purge rules of muso, ported from purge.js of the CHT configuration muso-mali (rules 0 to 14,
// same lists and windows). CHT runs the purge every Saturday at 22:00 over the reports of
// each contact; Collecton calls shouldPurge for each synced submission after every sync.
// CHT reported_date is the creation date on the device (report.createdAt); userCtx.roles
// are the Collecton role names, identical (chw_uhc, tb_focal_point, stock_manager, chw_manager).

const DAY = 86400000;

// List of supervisor forms to be purged from CHWs profiles
const supervisorFormsToPurgeForChws = [
  'individual_feedback',
  'individual_feedback_confirmation',
  'supervision_calendar',
  'supervision_visit_realization',
  'supervision_with_chw_confirmation',
  'supervision_without_chw_confirmation',
  'supervision_with_chw_iccm',
  'supervision_without_chw_iccm',
  'supervision_with_chw_proccm',
  'supervision_without_chw_proccm',
];
const tenDaysForms = ['referral_followup', 'referral_followup_under_5', 'treatment_followup', 'treatment_followup_over_5'];
const fourteenDaysForms = ['covid19_referral_followup'];
const fifteenDaysForms = ['behavior_change', 'family_planning_men', 'family_planning_administred_followup_men', 'newborn_simple_care', 'pregnancy_term_followup'];
const thirtyDaysForms = [
  'DA', 'epi_daily_report', 'home_visit', 'mute_clinic', 'newborn_followup', 'NS', 'patient_assessment_over_5', 'PH',
  'postnatal_followup', 'RA', 'redo_tdr', 'self_assessment', 'unmute_clinic',
];
const forms38Months = [
  'children_vaccination', 'children_vaccination_followup', 'hpv_verification', 'anc_followup', 'prenatal_followup',
  'prenatal_sp_administration_followup',
];
const ninetyDaysForms = [
  'pnc_followup', 'stock_distribution', 'stock_receipt', 'stock_adjustment_req', 'stock_adjustment_response',
  'family_planning_administred_followup', 'family_planning_followup',
];
const oneHundredDaysForms = [
  'patient_assessment', 'moderate_malnutrition_followup', 'severe_malnutrition_followup', 'stock_management',
  'stock_distribution_response', 'stock_distribution_response_ack', 'stock_adjustment_response_alert',
  'stock_adjustment_response_ack', 'stock_site_distribution_response', 'stock_site_adjustment_response',
  'prescription_summary', 'mgr_activity_planning', 'mgr_monthly_individual_meeting_realization', 'mgr_group_meeting',
  'mgr_group_meeting_planning', 'mgr_group_meeting_realization', 'mgr_individual_meeting', 'mgr_individual_meeting_followup',
  'mgr_leave_management', 'mgr_visit_realization', 'mgr_visit_with_sup', 'mgr_visit_without_sup',
  'mgr_visit_steps_realization', 'mgr_group_meeting_followup', 'mgr_monthly_one_on_one',
];
// List of forms to not be purged from tb_focal_points profiles
const formsNotPurgedForTbFocalPoints = ['patient_assessment', 'patient_assessment_over_5', 'tb_case_eval', 'tb_test_result_fp'];
const formsNotPurgedForStockManagers = [
  'stock_receipt', 'stock_management', 'stock_distribution', 'stock_distribution_response',
  'stock_distribution_response_ack', 'stock_adjustment_req', 'stock_adjustment_response',
  'stock_adjustment_response_alert', 'stock_adjustment_response_ack', 'stock_site_distribution_response',
  'stock_site_adjustment_response',
];
const formsNotPurgedForChwManagers = [
  'mgr_activity_planning', 'mgr_group_meeting', 'mgr_group_meeting_planning', 'mgr_group_meeting_realization',
  'mgr_monthly_individual_meeting_realization', 'mgr_individual_meeting', 'mgr_individual_meeting_followup',
  'mgr_leave_management', 'mgr_visit_realization', 'mgr_visit_with_sup', 'mgr_visit_without_sup',
  'mgr_visit_steps_realization', 'mgr_group_meeting_followup', 'mgr_monthly_one_on_one', ...formsNotPurgedForStockManagers,
];

/** True when the submission is purged from the device (CHT reportsToPurge). */
export function shouldPurge(report, context) {
  const age = context.now - report.createdAt;
  const olderThan = (days) => age >= days * DAY;
  const form = report.form;
  const field = (key) => report.payload?.[key];
  const role = context.role;
  // CHT newReports: reports of the last 300 days (the rules 1 to 13 only see them).
  const isNew = !olderThan(300);

  // 0. aggressive purging but do not purge reports in forms38Months
  if (olderThan(300) && !forms38Months.includes(form)) return true;

  if (isNew) {
    // 1. 10 days forms, 2. 14 days, 3. 15 days, 4. 30 days
    if (tenDaysForms.includes(form) && olderThan(10)) return true;
    if (fourteenDaysForms.includes(form) && olderThan(14)) return true;
    if (fifteenDaysForms.includes(form) && olderThan(15)) return true;
    if (thirtyDaysForms.includes(form) && olderThan(30)) return true;
    // 4.1 children assessments not needed for the malnutrition workflow are purged at 30 days
    // (CHT reads s_malnutrition_observation.has_MAM / has_SAM_without_complication).
    if (form === 'patient_assessment' && olderThan(30)
      && field('has_MAM') !== 'true' && field('has_SAM_without_complication') !== 'true') return true;
    // 5. Pregnancy
    if (form === 'pregnancy_family_planning' && ((field('fp_method') !== '' && olderThan(30)) || olderThan(300))) return true;
    // 6. PNC followup (90 days forms)
    if (ninetyDaysForms.includes(form) && olderThan(90)) return true;
    // 7. Realization supervision
    if (form === 'supervision_visit_realization' && olderThan(120)) return true;
    // 9. Malnutrition followup and CHW manager reports (100 days forms)
    if (oneHundredDaysForms.includes(form) && olderThan(100)) return true;
  }

  // 10. Purge all supervisors forms in CHWs accounts
  if (role === 'chw_uhc' && supervisorFormsToPurgeForChws.includes(form)) return true;

  if (isNew) {
    // 11. tb_focal_point: every form older than 30 days, and every form not in its list
    if (role === 'tb_focal_point' && (olderThan(30) || !formsNotPurgedForTbFocalPoints.includes(form))) return true;
    // 12. stock_manager
    if (role === 'stock_manager' && (olderThan(30) || !formsNotPurgedForStockManagers.includes(form))) return true;
    // 13. chw_manager
    if (role === 'chw_manager' && (olderThan(30) || !formsNotPurgedForChwManagers.includes(form))) return true;
  }

  // 14. Purge all reports older than 38 months (CHT: 30 * 38 days)
  if (forms38Months.includes(form) && olderThan(30 * 38)) return true;

  return false;
}
