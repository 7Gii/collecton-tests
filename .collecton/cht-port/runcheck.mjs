// Runs a generated form like the mobile app: FormEngine of collecton-lib + form.js functions.
import fs from 'node:fs';
const LIB = process.env.COLLECTON_LIB_DIST ?? '/Users/gilbertagbodamakou/Documents/shellit/Dev/collecton/collecton-lib/dist/index.js';
const { FormEngine, evaluateBoolean, buildScriptApp } = await import(LIB);
const [root, form, ...pairs] = process.argv.slice(2);
const dir = `${root}/metadata/forms/${form}`;
const def = JSON.parse(fs.readFileSync(`${dir}/form.json`, 'utf8'));
const helpers = await import('data:text/javascript;base64,' + Buffer.from(fs.readFileSync(`${dir}/form.js`)).toString('base64'));
const errors = [];
const origError = console.error;
console.error = (...a) => errors.push(a.join(' ').slice(0, 200));
const now = Date.parse('2026-10-02T10:00:00Z');
const person = { id: 'p1', personTypeId: 'pt', personType: 'patient', orgUnitId: 'fam1', attributes: { person_name: 'Awa', person_gender: 'F', birthdate: '2024-07-02', external_id: '01000118101' } };
const lineage = [
  { id: 'fam1', levelId: 'l5', level: 5, levelName: 'c50_family', parentId: 'chw1', attributes: { ou_name: 'Famille Diallo' } },
  { id: 'chw1', levelId: 'l4', level: 4, levelName: 'c40_chw_area', parentId: 'sup1', attributes: { ou_name: 'Zone ASC 1', aire_de_sante: 'Dimbal', mam_intrant: 'plumpy' } },
  { id: 'sup1', levelId: 'l3', level: 3, levelName: 'c30_supervisor_area', parentId: 'ha1', attributes: {} },
];
const chw = { id: 'u-chw', personTypeId: 'x', personType: 'chw', orgUnitId: 'chw1', attributes: { person_name: 'ASC-Moussa', muso_id: '1.0001', phone: '76123456' } };
const taskInputs = JSON.parse(process.env.TASK_INPUTS || '{}');
const app = { taskInputs, ...buildScriptApp({ person, reports: [], now, lineage, personsAt: id => (id === 'chw1' ? [chw] : []) }), currentUser: { id: 'u1', name: 'ASC-Moussa', role: { name: 'chw_uhc', permissions: [] } } };
const engine = new FormEngine();
const fields = engine.flattenFields(def.sections);
const given = Object.fromEntries(pairs.map(p => p.split('=')).map(([k, v]) => [k, v.includes(',') ? v.split(',') : v]));
const values = engine.calculateValues(fields, given, helpers, app);
const fApp = { ...app, currentForm: values };
let shown = 0;
for (const s of def.sections) for (const f of s.fields) if (evaluateBoolean(f.displayExpression, fApp, helpers, true)) shown++;
console.error = origError;
console.log('showForm:', helpers.showForm(app));
console.log('fields:', fields.length, 'shown now:', shown, 'runtime errors:', errors.length);
for (const e of [...new Set(errors)].slice(0, 8)) console.log('  ERR', e);
const pick = (process.env.KEYS || '').split(',').filter(Boolean);
for (const k of pick) console.log(' ', k, '=', JSON.stringify(values[k]));
