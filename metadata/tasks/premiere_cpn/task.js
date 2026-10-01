/**
 * First antenatal visit (CPN): one month after a pregnancy diagnosis with a
 * positive test. Resolved by any later CPN or delivery report of the woman.
 */

function laterReports(app, forms) {
  return (app.getReports({ form: forms, targetId: app.report.target_id }) || [])
    .filter(report => report.created_at > app.report.created_at);
}

export function appliesIf(app) {
  if (!app.person || app.person.attributes?.person_is_deceased === true) return false;
  return app.report.payload?.resultat_du_test_de_grossesse === 'positif';
}

export function resolvedIf(app) {
  return laterReports(app, ['consultation_prenatale_cpn', 'accouchement']).length > 0;
}
