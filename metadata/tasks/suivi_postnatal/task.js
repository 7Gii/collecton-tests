/**
 * Postnatal follow-up (CPoN): visits at day 3, day 7 and week 6 after the
 * delivery report. An event is resolved by a CPoN report inside its window;
 * the whole rule stops once the mother is reported deceased.
 */

const DAY_MS = 24 * 60 * 60 * 1000;

export function appliesIf(app) {
  return !!app.person && app.person.attributes?.person_is_deceased !== true;
}

export function resolvedIf(app) {
  const later = (app.getReports({ form: 'suivi_postnatal_cpon', targetId: app.report.target_id }) || [])
    .filter(report => report.created_at > app.report.created_at);
  if (later.some(report => report.payload?.etat_de_la_mere === 'decedee')) return true;
  const due = new Date(app.event.dueDate).getTime();
  const windowStart = due - app.event.startDays * DAY_MS;
  const windowEnd = due + (app.event.endDays + 1) * DAY_MS;
  return later.some(report => report.created_at >= windowStart && report.created_at < windowEnd);
}
