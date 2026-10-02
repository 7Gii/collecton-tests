import { readFileSync } from 'node:fs';
const load = (p) => import('data:text/javascript;base64,' + Buffer.from(readFileSync(p)).toString('base64'));
const P = await load('person_types/chw/person_type.js');
const F = await load('org_unit_levels/c50_family/org_unit_level.js');
const now = Date.parse('2026-10-02T10:00:00Z');
const rows = { u5: { id: 'u5', parent_id: 'u4', attributes: '{"ou_name":"Famille Diallo"}' }, u4: { id: 'u4', parent_id: null, attributes: '{"ou_name":"Zone ASC 1"}' } };
const find = async (_t, [c]) => (rows[c.value] ? [rows[c.value]] : []);
const person = (a) => ({ now, find, person: { orgUnitId: 'u5', attributes: a } });
console.log('age 3y:', P.ageLabel(person({ birthdate: '2023-06-01' })), '| 5 mois:', P.ageLabel(person({ birthdate: '2026-04-15' })));
console.log('dob approx 2y 3m:', P.birthdate({ now, currentForm: { date_of_birth_method: ['approx'], age_years: '2', age_months: '3' } }));
console.log('dob approx 0y 11m:', P.birthdate({ now, currentForm: { date_of_birth_method: ['approx'], age_years: '0', age_months: '11' } }));
console.log('prefix ok/ko:', P.nameHasPrefix({ currentForm: { person_name: 'ASC-Awa' } }), P.nameHasPrefix({ currentForm: { person_name: 'Awa' } }));
console.log('phone ok/ko:', P.phoneValid({ currentForm: { phone: '76 12 34 56' } }), P.phoneValid({ currentForm: { phone: '7612' } }));
console.log('lineage:', await P.parentLineage(person({})), '| sex:', P.sexLabel(person({ person_gender: 'F' })), '| role:', P.roleLabel(person({ role: 'chw' })));
const reports = [
  { payload: { visited_contact_uuid: 'u5', visited_date: '2026-10-01' }, created_at: now },
  { payload: { visited_contact_uuid: 'u5' }, created_at: Date.parse('2026-10-01T15:00:00Z') },
  { payload: { visited_contact_uuid: 'u5', visited_date: '2026-09-20' }, created_at: now },
  { payload: { visited_contact_uuid: 'other' }, created_at: now },
];
const fam = { now, orgUnit: { id: 'u5' }, getReports: () => reports };
console.log('last visit:', F.lastVisitDate(fam), '| visits this month:', F.visitsThisMonth(fam), '| none:', F.lastVisitDate({ ...fam, getReports: () => [] }));
