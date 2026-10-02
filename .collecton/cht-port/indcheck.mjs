// Computes the indicators of a project with the collecton-lib engine on a simulated month.
import fs from 'node:fs';
const LIB = process.env.COLLECTON_LIB_DIST ?? '/Users/gilbertagbodamakou/Documents/shellit/Dev/collecton/collecton-lib/dist/index.js';
const { computeIndicatorResult, buildScriptApp, buildReportIndex } = await import(LIB);
const root = process.argv[2];
const load = (p) => import('data:text/javascript;base64,' + Buffer.from(fs.readFileSync(p)).toString('base64'));
const forms = fs.readdirSync(`${root}/metadata/forms`).map(n => JSON.parse(fs.readFileSync(`${root}/metadata/forms/${n}/form.json`, 'utf8')));
const formId = Object.fromEntries(forms.map(f => [f.name, f.id]));
const DAY = 86400000;
const now = Date.parse('2026-10-20T10:00:00Z');
const at = (d) => Date.parse(`2026-${d}T09:00:00Z`);
const r = (id, form, target, date, payload) => ({ id, form_id: formId[form], target_id: target, target_is_person: true, created_at: at(date), payload });
const reports = [
  r('a1', 'patient_assessment', 'p1', '10-02', { has_danger_sign: 'true', refer_to_cscom: 'false', accompany_to_cscom: 'false' }),
  r('a2', 'patient_assessment', 'p2', '10-05', { treat_malaria: 'true', accompany_to_cscom: 'false', refer_to_cscom: 'false', s_malaria_give_act: 'no', s_malaria_not_give_act: 'refused' }),
  r('a3', 'patient_assessment', 'p3', '10-17', { treat_malaria: 'true', accompany_to_cscom: 'false', refer_to_cscom: 'false', s_malaria_give_act: 'yes' }),
  r('a4', 'patient_assessment_over_5', 'p4', '10-03', { referral: 'false', visited_contact_uuid: 'fam1', visited_date: '2026-10-03' }),
  r('a5', 'patient_assessment', 'p5', '10-08', { visited_contact_uuid: 'fam1', visited_date: '2026-10-08' }),
  r('a6', 'patient_assessment', 'p6', '10-08', { visited_contact_uuid: 'fam2', visited_date: '2026-10-08' }),
  r('a7', 'patient_assessment', 'p1', '09-10', { has_danger_sign: 'true' }),
  r('a8', 'patient_assessment', 'pdead', '10-04', { has_danger_sign: 'true' }),
];
const persons = { pdead: { id: 'pdead', attributes: JSON.stringify({ person_is_deceased: true }) } };
const find = async (table, [c]) => {
  if (table === 'persons') return [persons[c.value] ?? { id: c.value, attributes: '{}' }];
  if (table === 'org_unit_levels') return [{ id: 'l5', name: 'c50_family' }];
  if (table === 'org_units') return [{ id: 'fam1' }, { id: 'fam2' }, { id: 'fam3' }];
  return [];
};
const index = buildReportIndex({ forms });
const app = { ...buildScriptApp({ reports, index, now, find }), currentUser: { id: 'u1', name: 'ASC', role: { name: 'chw_uhc', permissions: [] } } };
for (const n of fs.readdirSync(`${root}/metadata/indicators`).sort()) {
  const indicator = JSON.parse(fs.readFileSync(`${root}/metadata/indicators/${n}/indicator.json`, 'utf8'));
  const mod = await load(`${root}/metadata/indicators/${n}/indicator.js`);
  const result = await computeIndicatorResult(indicator, { has: (name) => typeof mod[name] === 'function', call: (name) => mod[name](app) }, (m) => console.log('  warn', m));
  console.log(n.padEnd(52), JSON.stringify(result.value), 'goal', result.goal ?? '-');
}
