/**
 * Record the delivery: a CPN reported that the delivery happened, the delivery
 * form must be filled right away. Resolved by any later delivery report.
 */

export function appliesIf(app) {
  if (!app.person || app.person.attributes?.person_is_deceased === true) return false;
  return app.report.payload?.accouchement_survenu === 'oui';
}

export function resolvedIf(app) {
  return (app.getReports({ form: 'accouchement', targetId: app.report.target_id }) || [])
    .some(report => report.created_at > app.report.created_at);
}
