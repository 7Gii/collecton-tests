/**
 * Logic for the 'c50_family' org unit level: functions of its summary cards, computed offline on
 * the details view of each org unit (the target of app).
 * Every function receives one argument, app (the same keys everywhere):
 *   app.person     the target person { id, personTypeId, personType, orgUnitId, attributes }
 *   app.orgUnit    the target org unit { id, levelId, level, levelName, parentId, attributes }
 *                  (one of the two, the other null; both null when there is no target)
 *   app.currentUser  { id, username, name, role: { name, permissions }, orgUnitId } or null
 *   app.getReports(filter?)  every visible report, newest first, e.g.
 *       app.getReports({ form: 'child_registration' })          by form name or id
 *       app.getReports({ targetId: app.person.id })             the target only
 *       app.getReports({ orgUnitId: id })                        an org unit and everything below it
 *       app.getReports({ from: '2026-01-01', to: '2026-02-01' }) a period (to is excluded)
 *     each report: r.form_id, r.target_id, r.target_is_person, r.created_at, r.payload
 *     (form values are strings as entered: convert numbers with Number(...))
 *   app.find(table, [{ field, value }])  rows of persons, org_units, person_types,
 *     org_unit_levels (asynchronous: await it; not usable in a form field expression)
 *   app.utils      exports of the project utils
 *   app.now        current time of the app clock (epoch ms): the chosen day in deferred
 *                  collection on the mobile app; use it instead of new Date()
 *
 * A card's displayExpression returns a boolean; a field's value returns a string or a
 * number (null or '' hides the field).
 */

// ---- Summary cards: contact-summary.templated.js of the CHT configuration ----

const parseAttributes = (row) => {
  try {
    return typeof row.attributes === 'string' ? JSON.parse(row.attributes) : (row.attributes || {});
  } catch {
    return {};
  }
};

/** Names of the org units from orgUnitId upwards, nearest first (CHT lineage filter). */
async function lineageNames(app, orgUnitId) {
  const names = [];
  let id = orgUnitId;
  for (let depth = 0; id && depth < 10; depth++) {
    const [row] = await app.find('org_units', [{ field: 'id', value: id }]);
    if (!row) break;
    names.push(parseAttributes(row).ou_name || '');
    id = row.parent_id;
  }
  return names.filter(Boolean);
}

/** CHT contact.parent: the org units above this one. */
export async function parentLineage(app) {
  return (await lineageNames(app, app.orgUnit?.parentId)).join(' / ') || null;
}

// A home visit is a report whose visited_contact_uuid is the family (CHT getReportHomeVisitSubject).
const homeVisits = (app) => app.getReports({ orgUnitId: app.orgUnit.id })
  .filter(r => r.payload?.visited_contact_uuid === app.orgUnit.id);

// Date of a home visit: visited_date, else the report date (CHT getReportHomeVisitDate).
const visitTime = (report) => {
  const visited = report.payload?.visited_date;
  return visited ? Date.parse(visited) : Number(report.created_at);
};

const formatDate = (time) => {
  const d = new Date(time);
  return `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`;
};

/** CHT contacts.results.sort.date.last.visited: the latest home visit, else 'jamais'. */
export function lastVisitDate(app) {
  const times = homeVisits(app).map(visitTime).filter(t => !Number.isNaN(t));
  return times.length ? formatDate(Math.max(...times)) : 'jamais';
}

/**
 * CHT contact.last.visit.current.month: number of distinct days with a home visit in the
 * Muso month (day 1 to the last day, as the collection period).
 */
export function visitsThisMonth(app) {
  const now = new Date(app.now);
  const days = new Set(homeVisits(app)
    .map(r => new Date(visitTime(r)))
    .filter(d => d.getUTCFullYear() === now.getUTCFullYear() && d.getUTCMonth() === now.getUTCMonth())
    .map(d => d.getUTCDate()));
  return days.size;
}
