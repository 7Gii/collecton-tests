// Ports a CHT XLSForm (exported to JSON by extract.py) into a Collecton form through the
// collecton-mcp operations. Usage:
//   node convert.mjs --project <root> --form <name> --plan <planId> [--out <dir>]
// Every expression becomes a function of form.js (XPath runtime of runtime.js); what
// cannot be translated is listed in <out>/<form>.review.json, never guessed.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parse, toJs, Unsupported } from './xpath.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const MCP = process.env.COLLECTON_MCP_DIST ?? '/Users/gilbertagbodamakou/Documents/shellit/Dev/collecton/collecton-mcp/dist';
const { loadProject, listCollection } = await import(`${MCP}/project.js`);
const { assertProjectWritable } = await import(`${MCP}/guards.js`);
const forms = await import(`${MCP}/ops/forms.js`);
const planning = await import(`${MCP}/ops/plans.js`);

const args = Object.fromEntries(process.argv.slice(2).reduce((acc, a, i, all) => (a.startsWith('--') ? [...acc, [a.slice(2), all[i + 1]]] : acc), []));
const ROOT = args.project;
const FORM = args.form;
const OUT = args.out ?? path.join(HERE, 'out');
fs.mkdirSync(OUT, { recursive: true });

const ctx = loadProject(ROOT);
assertProjectWritable(ctx);
const source = JSON.parse(fs.readFileSync(path.join(HERE, 'json', `${FORM}.json`), 'utf8'));
const CONFIG = JSON.parse(fs.readFileSync(path.join(HERE, 'forms.json'), 'utf8'))[FORM];
if (!CONFIG) throw new Error(`No configuration for form ${FORM} in forms.json`);

const review = [];
const note = (row, column, reason, expression) => review.push({ field: row?.name ?? null, type: row?.type ?? null, column, reason, expression: expression ?? null });

// ------------------------------------------------------------------ survey walk

const fr = (row, base) => {
  const value = row[`${base}::fr`] ?? row[base] ?? '';
  return String(value).trim() === 'NO_LABEL' ? '' : String(value);
};
const cleanLabel = (text) => String(text)
  .replace(/<i[^>]*>\s*<\/i>/gi, '')
  .replace(/<[^>]+>/g, '')
  .replace(/[ \t]+/g, ' ')
  .trim();
const spanColor = (text) => /color\s*:\s*(red|blue|green|yellow|lime)/i.exec(String(text))?.[1]?.toLowerCase();

const items = [];
const groups = [];
const stack = [];
let inInputs = 0;
// Fields of the CHT inputs group: path below inputs/, read through X.inp (person, lineage, user).
const inputStack = [];
const inputPaths = {};
const storedInputs = [];
// Corrections of source bugs decided with the user (ECARTS_CHT_COLLECTON.md).
for (const row of source.survey) {
  const fix = CONFIG.overrides?.[row.name];
  if (fix) Object.assign(row, fix);
}
for (const row of source.survey) {
  const type = String(row.type ?? '').trim().replace(/\s+/g, ' ');
  if (!type) continue;
  if (type === 'begin group' || type === 'begin repeat') {
    if (type === 'begin repeat') note(row, 'type', 'repeat group: not supported');
    if (inInputs || (stack.length === 0 && row.name === 'inputs')) {
      if (inInputs) inputStack.push(row.name);
      inInputs++;
      continue;
    }
    const group = { name: row.name, label: cleanLabel(fr(row, 'label')), relevant: row.relevant, appearance: String(row.appearance ?? ''), parent: stack[stack.length - 1] ?? null, depth: stack.length };
    groups.push(group);
    stack.push(group);
    items.push({ kind: 'group', group });
    continue;
  }
  if (type === 'end group' || type === 'end repeat') {
    if (inInputs) { inInputs--; inputStack.pop(); continue; }
    stack.pop();
    continue;
  }
  if (inInputs) {
    const name = String(row.name);
    inputPaths[name] = name in inputPaths ? null : [...inputStack, name].join('/');
    // Values handed by the task (CHT modifyContent: source_id, t_*...) are stored, as CHT does.
    if (inputStack.length === 0 && name !== 'source' && !type.startsWith('db:')) storedInputs.push(name);
    continue;
  }
  items.push({ kind: 'field', row, type, groups: [...stack] });
}

// ------------------------------------------------------------------ keys

const sanitize = (s) => String(s).replace(/[^A-Za-z0-9_]/g, '_');
const fieldItems = items.filter(i => i.kind === 'field');
const nameCount = {};
for (const it of fieldItems) nameCount[it.row.name] = (nameCount[it.row.name] ?? 0) + 1;
const usedKeys = new Set();
for (const it of fieldItems) {
  let key = it.row.name;
  if (nameCount[key] > 1) {
    const ancestors = [...it.groups].reverse();
    for (const g of ancestors) { key = `${g.name}_${key}`; if (!usedKeys.has(key)) break; }
    let n = 2;
    while (usedKeys.has(key)) key = `${it.row.name}_${n++}`;
  }
  usedKeys.add(key);
  it.key = key;
}
const keyOfName = {};
for (const it of fieldItems) if (nameCount[it.row.name] === 1) keyOfName[it.row.name] = it.key;
for (const g of groups) g.key = sanitize(g.name) + (groups.filter(x => x.name === g.name).length > 1 ? `_${groups.indexOf(g)}` : '');

const listOf = {};
for (const it of fieldItems) {
  const m = /^select_(one|multiple)\s+(\S+)/.exec(it.type);
  if (m) { it.select = m[1]; it.list = m[2]; listOf[it.key] = m[2]; }
}

/** JavaScript of ${name}: a field of the form, else a field of the inputs group. */
const varJs = (name) => {
  if (keyOfName[name]) return `X.v(app, ${JSON.stringify(keyOfName[name])})`;
  if (nameCount[name] === undefined && inputPaths[name]) return `X.inp(app, ${JSON.stringify(inputPaths[name])})`;
  return `X.v(app, ${JSON.stringify(resolve(name))})`;
};

const resolve = (name) => {
  if (keyOfName[name]) return keyOfName[name];
  if (nameCount[name] > 1) throw new Unsupported(`\${${name}} is ambiguous (${nameCount[name]} fields)`);
  throw new Unsupported(`\${${name}} does not exist (or is an input)`);
};

// ------------------------------------------------------------------ expressions

const functions = [];
const compile = (row, column, expression, self, options = {}) => {
  if (expression === undefined || expression === null || String(expression).trim() === '') return null;
  try {
    // An Excel-style leading '=' (=TRUE()) is not part of the expression.
    return toJs(parse(String(expression).trim().replace(/^=/, '')), {
      self,
      resolve,
      varJs,
      choiceFilter: !!options.choiceFilter,
      choiceList: (name) => listOf[resolve(name)] ?? (() => { throw new Unsupported(`jr:choice-name of a non-select field ${name}`); })(),
    });
  } catch (e) {
    if (!(e instanceof Unsupported)) throw e;
    note(row, column, e.message, expression);
    return null;
  }
};
const define = (name, body, comment) => {
  functions.push(`${comment ? `/** ${comment} */\n` : ''}export function ${name}(app) {\n${body}\n}`);
  return name;
};

const groupFn = {};
for (const g of groups) {
  const own = compile({ name: g.name, type: 'group' }, 'relevant', g.relevant, null);
  const parent = g.parent ? groupFn[g.parent.key] : null;
  if (!own && !parent) continue;
  const parts = [parent ? `${parent}(app)` : null, own ? `X.bool(${own})` : null].filter(Boolean);
  groupFn[g.key] = define(`grp_${g.key}`, `  return ${parts.join(' && ')};`, `Group ${g.name}`);
}

const isTrue = (v) => ['yes', 'true', 'true()'].includes(String(v ?? '').trim().toLowerCase());
const isFalse = (v) => v === undefined || ['no', 'false', 'false()', ''].includes(String(v).trim().toLowerCase());

// ------------------------------------------------------------------ fields

const labelTemplate = (text) => String(text).replace(/\$\{([^}]+)\}/g, (_, name) => {
  try { return `{app.currentForm[${JSON.stringify(resolve(name.trim()))}] ?? ''}`; } catch { return `\${${name}}`; }
});

const plans = [];
for (const it of fieldItems) {
  const { row, type, key } = it;
  const fk = sanitize(key);
  const appearance = String(row.appearance ?? '').trim().split(/\s+/).filter(Boolean);
  const nearest = it.groups[it.groups.length - 1];
  const relParts = [nearest && groupFn[nearest.key] ? `${groupFn[nearest.key]}(app)` : null];
  const own = compile(row, 'relevant', row.relevant, key);
  if (own) relParts.push(`X.bool(${own})`);
  const rel = relParts.filter(Boolean).length ? define(`rel_${fk}`, `  return ${relParts.filter(Boolean).join(' && ')};`, `Relevance of ${key}`) : null;
  const plan = { key, row, rel };

  if (type.startsWith('db:')) { note(row, 'type', `${type} widget: no Collecton equivalent (the target of the form is the person)`); continue; }

  if (type === 'note' && !row.calculation) {
    const label = labelTemplate(cleanLabel(fr(row, 'label')));
    if (!label) continue;
    plan.display = appearance.includes('countdown-timer') ? 'timer' : 'label';
    plan.label = label;
    const heading = appearance.find(a => /^h[123]$/.test(a));
    if (heading) plan.headingLevel = Number(heading[1]);
    plan.color = appearance.find(a => ['red', 'blue', 'green', 'yellow', 'lime'].includes(a)) ?? spanColor(fr(row, 'label'));
    if (plan.display === 'timer') { plan.timerSeconds = Number(row.default) > 0 ? Number(row.default) : undefined; plan.color = undefined; plan.headingLevel = undefined; }
    if (appearance.includes('hidden')) plan.hidden = true;
    plans.push(plan);
    continue;
  }

  plan.dataType = it.select ? `select_${it.select}`
    : { string: 'string', text: 'string', integer: 'integer', int: 'integer', decimal: 'decimal', date: 'date', dateTime: 'datetime', calculate: 'string', hidden: 'string', note: 'string' }[type];
  if (!plan.dataType) { note(row, 'type', `type '${type}' not supported`); continue; }
  plan.list = it.list;
  plan.label = labelTemplate(cleanLabel(fr(row, 'label'))) || key;
  plan.hint = cleanLabel(fr(row, 'hint')) || undefined;
  const calculated = type === 'calculate' || (type === 'note' && row.calculation) || (row.calculation && isTrue(row.read_only));
  if (row.calculation && !calculated) note(row, 'calculation', 'calculation on an editable field: used as default value', row.calculation);
  if (row.calculation) {
    const js = compile(row, 'calculation', row.calculation, key);
    if (js) {
      const keep = /^\s*(once|uuid)\s*\(/.test(String(row.calculation)) ? `  const current = X.v(app, ${JSON.stringify(key)});\n  if (X.str(current) !== '') return current;\n` : '';
      const gate = rel ? `  if (!${rel}(app)) return '';\n` : '';
      plan[calculated ? 'calculation' : 'defaultValue'] = define(`calc_${fk}`, `${keep}${gate}  return X.out(${js});`, `Calculation of ${key}`);
    }
  }
  if (type === 'hidden' || appearance.includes('hidden') || type === 'note') plan.hidden = true;
  if (row.default !== undefined && !plan.defaultValue) {
    const value = String(row.default);
    if (/[$(/]/.test(value)) {
      const js = compile(row, 'default', value, key);
      if (js) plan.defaultValue = define(`def_${fk}`, `  return X.out(${js});`, `Default of ${key}`);
    } else {
      plan.defaultValue = /^[A-Za-z_$][\w$]*$/.test(value) ? value : define(`def_${fk}`, `  return ${JSON.stringify(value)};`, `Default of ${key}`);
    }
  }
  if (isTrue(row.required)) plan.required = true;
  else if (!isFalse(row.required)) {
    const js = compile(row, 'required', row.required, key);
    if (js) plan.required = define(`req_${fk}`, `  return X.bool(${js});`, `Required ${key}`);
  }
  if (row.required && fr(row, 'required_message')) plan.requiredMessage = cleanLabel(fr(row, 'required_message'));
  if (row.constraint) {
    const js = compile(row, 'constraint', row.constraint, key);
    if (js) {
      plan.validation = define(`val_${fk}`, `  if (X.str(X.v(app, ${JSON.stringify(key)})) === '') return true;\n  return X.bool(${js});`, `Constraint of ${key}`);
      plan.validationMessage = cleanLabel(fr(row, 'constraint_message')) || undefined;
    }
  }
  if (row.choice_filter) {
    const js = compile(row, 'choice_filter', row.choice_filter, key, { choiceFilter: true });
    if (js) plan.optionFilter = define(`filt_${fk}`, `  return X.bool(${js});`, `Options of ${key}`);
  }
  for (const a of appearance) {
    if (!['hidden', 'field-list', 'li', 'summary'].includes(a)) note(row, 'appearance', `appearance '${a}' ignored`);
  }
  plans.push(plan);
}

// ------------------------------------------------------------------ choices

const choicesByList = {};
for (const c of source.choices) {
  if (!c.list_name || c.name === undefined) continue;
  (choicesByList[String(c.list_name).trim()] ??= []).push(c);
}
const choiceTables = {};
for (const [list, rows] of Object.entries(choicesByList)) {
  choiceTables[list] = Object.fromEntries(rows.map(c => [String(c.name), cleanLabel(fr(c, 'label'))]));
}
const propsOf = (c) => {
  const props = {};
  for (const [k, v] of Object.entries(c)) {
    if (['list_name', 'name'].includes(k) || k.startsWith('label') || k.startsWith('media') || k.startsWith('image') || k.startsWith('audio')) continue;
    props[k] = String(v);
  }
  return Object.keys(props).length ? props : undefined;
};

// ------------------------------------------------------------------ write

const authorize = (kind, name) => { if (args.plan) planning.assertAuthorized(ROOT, args.plan, kind, name); };
const sameOptions = (a, b) => JSON.stringify(a) === JSON.stringify(b);
const optionSetFor = (list) => {
  const rows = choicesByList[list];
  if (!rows) throw new Error(`Choice list ${list} not found`);
  const options = rows.map(c => ({ name: cleanLabel(fr(c, 'label')), value: String(c.name), ...(propsOf(c) ? { properties: propsOf(c) } : {}) }));
  const shape = (opts) => opts.map(o => ({ name: o.name, value: o.value, properties: o.properties ?? null }));
  const existing = listCollection(ctx, 'optionSets').find(os => sameOptions(shape(os.options ?? []), shape(options)));
  if (existing) return existing.name;
  let name = list;
  if (listCollection(ctx, 'optionSets').some(os => os.name === name)) name = `${FORM}__${list}`;
  authorize('option_set', name);
  forms.createOptionSet(ctx, name, options, name);
  return name;
};
const dataElementFor = (plan) => {
  const optionSet = plan.list ? optionSetFor(plan.list) : null;
  const optionSetId = optionSet ? listCollection(ctx, 'optionSets').find(os => os.name === optionSet).id : null;
  const all = listCollection(ctx, 'dataElements');
  const fits = (de) => de.dataType === plan.dataType && (de.optionSetId ?? null) === optionSetId;
  // A data element name starts with a letter or a digit; the field key keeps the XLSForm name.
  const base = /^[A-Za-z0-9]/.test(plan.row.name) ? plan.row.name : `${FORM}__${plan.row.name.replace(/^[^A-Za-z0-9]+/, '')}`;
  const own = all.find(de => de.name === base);
  if (own && fits(own)) return own.name;
  const name = own ? `${FORM}__${base}` : base;
  const again = all.find(de => de.name === name);
  if (again && fits(again)) return again.name;
  authorize('data_element', plan.label);
  const { dataElement } = forms.createDataElement(ctx, { label: plan.label, name, dataType: plan.dataType, optionSetName: optionSet ?? undefined });
  // The template fills placeholder, hint and description: the mobile app would show them.
  return dataElement.name;
};

authorize('form', CONFIG.title);
forms.createForm(ctx, CONFIG.title, FORM);

const choicesJs = `X.choices = ${JSON.stringify(choiceTables, null, 1)};`;
const runtime = fs.readFileSync(path.join(HERE, 'runtime.js'), 'utf8');
const header = `// Form ${FORM}: ported from forms/app/${FORM}.xlsx of the CHT configuration muso-mali.\n// Field keys keep the XLSForm names; the functions below are generated from its expressions.\n`;
const script = [header, runtime, choicesJs, CONFIG.showForm, ...functions].join('\n\n') + '\n';
forms.setFormLogic(ctx, FORM, script);

// Sections: top-level groups (pages), and the loose top-level fields around them.
let currentSection = null;
let looseIndex = 0;
const sectionOf = (it) => {
  const top = it.groups[0];
  if (top) return top.key;
  if (!currentSection || !currentSection.startsWith('page_')) {
    currentSection = `page_${++looseIndex}`;
    forms.addFormGroup(ctx, FORM, '', { key: currentSection, fullPage: true });
  }
  return currentSection;
};
const created = new Set();
const ensureSection = (g) => {
  if (created.has(g.key)) return;
  created.add(g.key);
  forms.addFormGroup(ctx, FORM, g.label, { key: g.key, condition: groupFn[g.key], fullPage: true });
};

const titlesDone = new Set();
let fieldCount = 0;
for (const it of items) {
  if (it.kind === 'group') {
    const g = it.group;
    if (g.depth === 0) { ensureSection(g); currentSection = g.key; continue; }
    if (g.label && !g.appearance.includes('hidden') && !titlesDone.has(g.key)) {
      titlesDone.add(g.key);
      const top = g; let root = g; while (root.parent) root = root.parent;
      ensureSection(root);
      forms.addFormField(ctx, FORM, undefined, root.key, { displayType: 'label', key: `${g.key}__title`, label: g.label, headingLevel: 3 });
      if (groupFn[g.key]) forms.updateFormField(ctx, FORM, `${g.key}__title`, { displayExpression: groupFn[g.key] });
      void top;
    }
    continue;
  }
  const plan = plans.find(p => p.row === it.row);
  if (!plan) continue;
  const section = it.groups[0] ? (ensureSection(it.groups[0]), it.groups[0].key) : sectionOf(it);
  const patch = {};
  if (plan.display) {
    forms.addFormField(ctx, FORM, undefined, section, {
      displayType: plan.display, key: plan.key, label: plan.label,
      headingLevel: plan.headingLevel, color: plan.color, timerSeconds: plan.timerSeconds,
    });
  } else {
    const de = dataElementFor(plan);
    forms.addFormField(ctx, FORM, de, section, { key: plan.key, label: plan.label });
    patch.required = plan.required ?? false;
    if (plan.calculation) patch.calculationExpression = plan.calculation;
    if (plan.defaultValue) patch.defaultValue = plan.defaultValue;
    if (plan.validation) { patch.validationExpression = plan.validation; if (plan.validationMessage) patch.validationMessage = plan.validationMessage; }
    if (plan.requiredMessage) patch.requiredMessage = plan.requiredMessage;
    if (plan.optionFilter) patch.optionFilter = plan.optionFilter;
    patch.hint = plan.hint ?? '';
    patch.placeholder = '';
  }
  patch.displayExpression = plan.hidden ? 'false' : plan.rel ?? 'true';
  forms.updateFormField(ctx, FORM, plan.key, patch);
  fieldCount++;
}

// Inputs handed by the task, stored in a hidden section: inputs_<name> (read by the task rules).
if (storedInputs.length) {
  forms.addFormGroup(ctx, FORM, '', { key: 'inputs', condition: 'false', fullPage: false });
  for (const name of storedInputs) {
    const key = `inputs_${name}`;
    const fn = define(`calc_${sanitize(key)}`, `  return X.out(X.inp(app, ${JSON.stringify(name)}));`, `Task input ${name}`);
    const de = dataElementFor({ row: { name: key }, dataType: 'string', label: key });
    forms.addFormField(ctx, FORM, de, 'inputs', { key, label: key });
    forms.updateFormField(ctx, FORM, key, { calculationExpression: fn, displayExpression: 'false', hint: '', placeholder: '' });
    fieldCount++;
  }
  // The functions defined above join form.js.
  forms.setFormLogic(ctx, FORM, [header, runtime, choicesJs, CONFIG.showForm, ...functions].join('\n\n') + '\n');
}

// Translations (en, bm) of the FR texts, for set_translations at the end of the port.
const i18n = { en: {}, bm: {} };
const addTr = (row, base) => {
  const f = cleanLabel(fr(row, base));
  if (!f) return;
  for (const lang of ['en', 'bm']) { const t = row[`${base}::${lang}`]; if (t) i18n[lang][f] = cleanLabel(t); }
};
for (const r of source.survey) { addTr(r, 'label'); addTr(r, 'hint'); addTr(r, 'constraint_message'); addTr(r, 'required_message'); }
for (const c of source.choices) addTr(c, 'label');

fs.writeFileSync(path.join(OUT, `${FORM}.review.json`), JSON.stringify(review, null, 1));
fs.writeFileSync(path.join(OUT, `${FORM}.i18n.json`), JSON.stringify(i18n, null, 1));
console.log(JSON.stringify({ form: FORM, fields: fieldCount, functions: functions.length, review: review.length }));
