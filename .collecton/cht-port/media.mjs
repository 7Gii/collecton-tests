// Icons and form images of the CHT configuration, carried to the Collecton project:
// - form icons (forms/app/<form>.properties.json "icon", resolved by resources.json)
//   copied to assets/icons and set on each form;
// - icons of the contact types (app_settings.json contact_types) set on the org unit
//   levels (same id) and on the person types (CHT "person" type);
// - images of the form notes (forms/app/<form>-media/images, media::image column)
//   copied to assets/images and set on the fields (FormField.image).
// Idempotent. Usage: node media.mjs <CHT config folder> <Collecton project folder>

import fs from 'node:fs';
import path from 'node:path';

const [chtRoot, projectRoot] = process.argv.slice(2);
if (!chtRoot || !projectRoot) {
  console.error('Usage: node media.mjs <CHT config folder> <Collecton project folder>');
  process.exit(1);
}
const readJson = file => JSON.parse(fs.readFileSync(file, 'utf8'));
const writeJson = (file, value) => fs.writeFileSync(file, `${JSON.stringify(value, null, 2)}\n`);
const resources = readJson(path.join(chtRoot, 'resources.json'));
const iconsDir = path.join(projectRoot, 'assets', 'icons');
const imagesDir = path.join(projectRoot, 'assets', 'images');
fs.mkdirSync(iconsDir, { recursive: true });
fs.mkdirSync(imagesDir, { recursive: true });

/** Copies the CHT resource of an icon name to assets/icons; returns the Collecton icon value. */
const copyIcon = name => {
  const file = name && resources[name];
  const source = file && path.join(chtRoot, 'resources', file);
  if (!source || !fs.existsSync(source)) return null;
  fs.copyFileSync(source, path.join(iconsDir, file));
  return file;
};

const report = { forms: [], levels: [], personTypes: [], images: [], missing: [] };
const metadata = path.join(projectRoot, 'metadata');

// Form icons and note images.
const walk = (nodes, visit) => nodes.forEach(node => (Array.isArray(node.fields) && !node.dataType ? walk(node.fields, visit) : visit(node)));
for (const name of fs.readdirSync(path.join(metadata, 'forms'))) {
  const formFile = path.join(metadata, 'forms', name, 'form.json');
  const propsFile = path.join(chtRoot, 'forms', 'app', `${name}.properties.json`);
  if (!fs.existsSync(formFile) || !fs.existsSync(propsFile)) continue;
  const form = readJson(formFile);
  const icon = copyIcon(readJson(propsFile).icon);
  if (icon) { form.icon = icon; report.forms.push(`${name}: ${icon}`); }
  else report.missing.push(`form ${name}: icon '${readJson(propsFile).icon}'`);

  const surveyFile = path.join(path.dirname(new URL(import.meta.url).pathname), 'json', `${name}.json`);
  const mediaDir = path.join(chtRoot, 'forms', 'app', `${name}-media`, 'images');
  if (fs.existsSync(surveyFile)) {
    const imageByName = new Map();
    for (const row of readJson(surveyFile).survey ?? []) {
      const file = row['media::image::fr'] ?? row['media::image'] ?? row['media::image::en'] ?? row.image;
      if (row.name && file) imageByName.set(row.name, String(file).trim());
    }
    walk(form.sections ?? [], field => {
      // Keys of fields repeated in several groups are scoped: "<group>_<name>".
      const rowName = [...imageByName.keys()].find(n => field.key === n || field.key?.endsWith(`_${n}`));
      if (!rowName) return;
      const file = imageByName.get(rowName);
      const source = path.join(mediaDir, file);
      if (!fs.existsSync(source)) { report.missing.push(`form ${name}: image ${file}`); return; }
      fs.copyFileSync(source, path.join(imagesDir, file));
      field.image = `images/${file}`;
      report.images.push(`${name}.${field.key}: ${file}`);
    });
  }
  writeJson(formFile, form);
}

// Org unit levels (contact type of the same id) and person types (CHT "person" type).
const contactTypes = readJson(path.join(chtRoot, 'app_settings.json')).contact_types ?? [];
const typeIcon = id => contactTypes.find(c => c.id === id)?.icon;
for (const [folder, file, list, idOf] of [
  ['org_unit_levels', 'org_unit_level.json', report.levels, name => name],
  ['person_types', 'person_type.json', report.personTypes, () => 'person'],
]) {
  for (const name of fs.readdirSync(path.join(metadata, folder))) {
    const itemFile = path.join(metadata, folder, name, file);
    if (!fs.existsSync(itemFile)) continue;
    const item = readJson(itemFile);
    const icon = copyIcon(typeIcon(idOf(name)));
    if (!icon) { report.missing.push(`${folder} ${name}: icon '${typeIcon(idOf(name))}'`); continue; }
    item.icon = icon;
    writeJson(itemFile, item);
    list.push(`${name}: ${icon}`);
  }
}

console.log(JSON.stringify(report, null, 2));
