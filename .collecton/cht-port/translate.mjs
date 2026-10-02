// Applies dictionary.json (built by dictionary.py from the CHT sources) to the untranslated
// texts of the project, through the collecton-mcp translation operations. Writes the texts
// left untranslated to out/untranslated-<lang>.json.
import fs from 'node:fs';
const MCP = process.env.COLLECTON_MCP_DIST ?? '/Users/gilbertagbodamakou/Documents/shellit/Dev/collecton/collecton-mcp/dist';
const { loadProject } = await import(`${MCP}/project.js`);
const { assertProjectWritable } = await import(`${MCP}/guards.js`);
const tr = await import(`${MCP}/ops/translations.js`);

const ctx = loadProject(process.argv[2]);
assertProjectWritable(ctx);
const dictionary = JSON.parse(fs.readFileSync(new URL('./dictionary.json', import.meta.url), 'utf8'));
const normalize = (text) => String(text).replace(/\s+/g, ' ').trim().toLowerCase();
const byNormalized = new Map(Object.entries(dictionary).map(([fr, t]) => [normalize(fr), t]));

for (const lang of ['en', 'bm']) {
  const { untranslated } = tr.listUntranslated(ctx, lang);
  const found = {};
  const missing = [];
  for (const text of untranslated) {
    const entry = dictionary[text] ?? byNormalized.get(normalize(text));
    if (entry?.[lang]) found[text] = entry[lang];
    else missing.push(text);
  }
  // Texts without a source translation: technical keys of hidden fields and Collecton
  // permission labels keep their text; manual-en.json holds the few EN texts written by hand.
  const manual = lang === 'en' ? JSON.parse(fs.readFileSync(new URL('./manual-en.json', import.meta.url), 'utf8')) : {};
  const keepAsIs = (t) => /^[A-Za-z0-9_]+$/.test(t)
    || /^(Allows|Create|Edit|Delete|View) |^(Automatic|Manual) Sync$|^Deferred Collection$|^(The value is invalid|This value is required|Organisation Unit Name|Organization unit|Person)$/.test(t);
  for (const text of [...missing]) {
    if (keepAsIs(text)) found[text] = text;
    else if (manual[text]) found[text] = manual[text];
    else continue;
    missing.splice(missing.indexOf(text), 1);
  }
  if (Object.keys(found).length) tr.setTranslations(ctx, lang, found);
  fs.writeFileSync(new URL(`./out/untranslated-${lang}.json`, import.meta.url), JSON.stringify(missing, null, 1));
  console.log(`${lang}: ${untranslated.length} untranslated, ${Object.keys(found).length} translated, ${missing.length} left`);
}
