import fs from 'node:fs';
import { parse, toJs } from './xpath.mjs';
const runtime = fs.readFileSync('./runtime.js', 'utf8');
const X = new Function(`${runtime}; return X;`)();
X.choices = { yes_no: { true: 'Oui', false: 'Non' } };
const ctx = { self: 'age', resolve: n => n, choiceList: () => 'yes_no' };
const now = Date.parse('2026-10-02T10:00:00Z');
const run = (xp, form, extra = {}) => new Function('X', 'app', `return ${toJs(parse(xp), { ...ctx, ...extra })};`)(X, { now, currentForm: form, ...extra.app });
const cases = [
  ["${a} = 'true'", { a: 'true' }, true],
  ["${n} >= 2", { n: '3' }, true],
  ["${n} >= 2", { n: '' }, false],
  ["if(${a}='yes', 'ok', 'ko')", { a: 'no' }, 'ko'],
  ["coalesce(${a}, 'x')", { a: '' }, 'x'],
  ["selected(${m}, 'b')", { m: ['a', 'b'] }, true],
  ["selected(${m}, 'c')", { m: 'a b' }, false],
  ["floor(decimal-date-time(.)) <= floor(decimal-date-time(now()))", { age: '2026-09-30' }, true],
  ["floor(decimal-date-time(.)) <= floor(decimal-date-time(now()))", { age: '2026-12-30' }, false],
  ["difference-in-months(${dob}, today())", { dob: '2025-08-15' }, 13],
  ["floor((decimal-date-time(today()) - decimal-date-time(${dob})) div 365.25)", { dob: '2020-10-01' }, 6],
  ["${p} = ’fr’", { p: 'fr' }, true],
  ["format-date-time(date-time(decimal-date-time(${d}) + 3), \"%Y-%m-%d\")", { d: '2026-09-29' }, '2026-10-02'],
  ["jr:choice-name(${a}, '${a}')", { a: 'true' }, 'Oui'],
  ["${a} = 'true' and (${b} != 'x' or ${c} > 1)", { a: 'true', b: 'x', c: '2' }, true],
  ["not(${a} = 'true')", { a: 'true' }, false],
  ["-${n} + 5 * 2", { n: '1' }, 9],
  ["concat(${a}, '-', 3)", { a: 'A' }, 'A-3'],
  ["${t} mod 2 = 1", { t: '7' }, true],
  ["regex(., '^(\\d{2}\\s?){4}$')", { age: '76 12 34 56' }, true],
];
let failed = 0;
for (const [xp, form, expected] of cases) {
  const got = run(xp, form);
  const ok = JSON.stringify(got) === JSON.stringify(expected);
  if (!ok) failed++;
  console.log(ok ? 'ok  ' : 'FAIL', xp, '->', JSON.stringify(got));
}
console.log(failed ? `${failed} FAILED` : 'ALL OK');
