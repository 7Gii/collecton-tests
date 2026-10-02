// Runs the task rules of a project with the collecton-lib engine on simulated reports.
import fs from 'node:fs';
import path from 'node:path';
const LIB = process.env.COLLECTON_LIB_DIST ?? '/Users/gilbertagbodamakou/Documents/shellit/Dev/collecton/collecton-lib/dist/index.js';
const { computeTasks, buildScriptApp, buildReportIndex } = await import(LIB);
const root = process.argv[2];
const load = (p) => import('data:text/javascript;base64,' + Buffer.from(fs.readFileSync(p)).toString('base64'));
const forms = fs.readdirSync(`${root}/metadata/forms`).map(n => JSON.parse(fs.readFileSync(`${root}/metadata/forms/${n}/form.json`, 'utf8')));
const formId = Object.fromEntries(forms.map(f => [f.name, f.id]));
const rules = [];
for (const n of fs.readdirSync(`${root}/metadata/tasks`)) {
  const rule = JSON.parse(fs.readFileSync(`${root}/metadata/tasks/${n}/task.json`, 'utf8'));
  rule.actionFormId = rule.actionFormId;
  rules.push({ rule, helpers: await load(`${root}/metadata/tasks/${n}/task.js`) });
}
const DAY = 86400000;
const now = Date.parse('2026-10-10T09:00:00Z');
const ago = (d) => now - d * DAY;
const day = (d) => new Date(ago(d)).toISOString().slice(0, 10);
const scenario = JSON.parse(fs.readFileSync(process.argv[3], 'utf8'));
const persons = scenario.persons.map(p => ({ ...p, created_at: ago(400), is_active: true }));
const reports = scenario.reports.map((r, i) => ({ id: r.id ?? `r${i}`, form_id: formId[r.form], target_id: r.target, target_is_person: true, created_at: ago(r.daysAgo), payload: Object.fromEntries(Object.entries(r.payload).map(([k, v]) => [k, typeof v === 'string' && v.startsWith('@') ? day(Number(v.slice(1))) : v])) }));
const index = buildReportIndex({ forms, orgUnits: [{ id: 'fam1', parentId: 'chw1' }], persons: persons.map(p => ({ id: p.id, orgUnitId: p.org_unit_id })) });
const focal = { id: 'fp', personTypeId: 'x', personType: 'tb_focal_point', orgUnitId: 'ha1', attributes: { person_name: 'TBF-Awa', cscom_area: 'Dimbal' } };
const user = scenario.user;
const { tasks, errors } = await computeTasks({
  rules, persons, orgUnits: [], reports, now,
  personTypeNameById: new Map([['pt_patient', 'patient']]),
  formIdByName: index.formIdByName,
  buildApp: (target, report) => ({
    ...buildScriptApp({ person: { id: target.row.id, personTypeId: 'pt_patient', personType: 'patient', orgUnitId: 'fam1', attributes: target.row.attributes }, reports, index, now, personsAt: id => (id === 'ha1' ? [focal] : []) }),
    currentUser: user, report,
  }),
});
console.log('errors:', errors);
for (const t of tasks) console.log(t.state.padEnd(8), t.dueDate.slice(0, 10), t.ruleName.padEnd(55), t.eventId.padEnd(18), 'p' + (t.priority ?? '-'), JSON.stringify(t.inputs ?? {}).slice(0, 150));
