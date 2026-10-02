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
  if (Object.keys(found).length) tr.setTranslations(ctx, lang, found);
  fs.writeFileSync(new URL(`./out/untranslated-${lang}.json`, import.meta.url), JSON.stringify(missing, null, 1));
  console.log(`${lang}: ${untranslated.length} untranslated, ${Object.keys(found).length} translated from the sources, ${missing.length} left`);
}
