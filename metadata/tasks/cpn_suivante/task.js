/**
 * Next antenatal visit (CPN): one month after a CPN where the delivery has not
 * happened yet. Resolved by any later CPN or delivery report of the woman.
 */

function laterReports(app, forms) {
  return (app.getReports({ form: forms, targetId: app.report.target_id }) || [])
    .filter(report => report.created_at > app.report.created_at);
}

export function appliesIf(app) {
  if (!app.person || app.person.attributes?.person_is_deceased === true) return false;
  return app.report.payload?.accouchement_survenu !== 'oui';
}

export function resolvedIf(app) {
  return laterReports(app, ['consultation_prenatale_cpn', 'accouchement']).length > 0;
}
