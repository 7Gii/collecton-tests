// Binds the registration fields of the 7 staff person types of muso through the MCP operations.
const MCP = process.env.COLLECTON_MCP_DIST ?? '/Users/gilbertagbodamakou/Documents/shellit/Dev/collecton/collecton-mcp/dist';
const { loadProject } = await import(`${MCP}/project.js`);
const { assertProjectWritable } = await import(`${MCP}/guards.js`);
const structure = await import(`${MCP}/ops/structure.js`);

const ctx = loadProject('/Users/gilbertagbodamakou/Desktop/collecton-muso');
assertProjectWritable(ctx);

const PREFIXED = new Set(['chw', 'chw_supervisor', 'tb_focal_point', 'stock_manager']);
const MUSO_ID = new Set(['chw', 'chw_supervisor']);
const TYPES = ['chw', 'chw_supervisor', 'health_area_supervisor', 'tb_focal_point', 'stock_manager', 'site_supervisor', 'chw_manager'];
const hidden = '(app) => false';
const problems = [];
const check = (where, r) => { if (!r.validation.valid) problems.push(`${where}: ${r.validation.errors.join('; ')}`); };

for (const type of TYPES) {
  const up = (key, patch) => check(`${type}.${key}`, structure.updateAttributeField(ctx, 'person_types', type, key, patch));
  const add = (dataElement, opts = {}) => check(`${type}.${dataElement}`, structure.addAttributeField(ctx, 'person_types', type, { dataElement, ...opts }));

  up('person_name', PREFIXED.has(type) ? { label: 'Nom', validationExpression: 'nameHasPrefix' } : { label: 'Nom' });
  up('person_gender', { label: 'Sexe', required: true });
  up('birthdate', { label: 'Date de naissance', calculationExpression: 'birthdate' });
  up('person_is_deceased', { label: 'Décédé' });

  add('language', { defaultValue: 'fr' });
  add('date_of_birth_method');
  add('dob_calendar', { required: true, condition: 'dobIsKnown', validation: 'dobNotInFuture' });
  add('age_years', { required: true, condition: 'dobIsApprox', validation: 'ageYearsInRange' });
  add('age_months', { required: true, condition: 'dobIsApprox', validation: 'ageMonthsInRange', defaultValue: '0' });
  add('phone', { validation: 'phoneValid' });
  add('alternate_phone', { validation: 'alternatePhoneValid' });
  if (MUSO_ID.has(type)) add('muso_id', { required: true, validation: 'musoIdValid' });
  if (type === 'tb_focal_point') {
    const group = 'Point Focal TB';
    add('s_site', { group });
    add('s_ha', { group, optionFilter: 'healthZonesOfSite' });
    add('s_cscom_area', { group, required: true, optionFilter: 'cscomsOfHealthZone' });
    add('cscom_area', { group, calculation: 'cscomArea' });
  }
  const main = 'Identification Details';
  add('is_active', { group: main, condition: hidden, defaultValue: 'true' });
  add('stm_is_active', type === 'stock_manager' ? { group: main, defaultValue: 'false' } : { group: main, condition: hidden, defaultValue: 'false' });
  add('role', { group: main, condition: hidden, defaultValue: type });
  add('geolocation', { group: main });
}

console.log(problems.length ? `PROBLEMS:\n${problems.join('\n')}` : 'ALL VALID');
