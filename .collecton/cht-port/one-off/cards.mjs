// Summary cards of the muso person types and org unit levels, through the MCP operation.
const MCP = process.env.COLLECTON_MCP_DIST ?? '/Users/gilbertagbodamakou/Documents/shellit/Dev/collecton/collecton-mcp/dist';
const { loadProject } = await import(`${MCP}/project.js`);
const { assertProjectWritable } = await import(`${MCP}/guards.js`);
const structure = await import(`${MCP}/ops/structure.js`);
const ctx = loadProject('/Users/gilbertagbodamakou/Desktop/collecton-muso');
assertProjectWritable(ctx);

const profile = (extra = []) => ({
  key: 'profile', title: 'Profil', fields: [
    { key: 'age', label: 'Age', value: 'ageLabel' },
    { key: 'phone', label: 'Téléphone', value: 'phoneValue' },
    { key: 'sex', label: 'Sexe', value: 'sexLabel' },
    { key: 'external_id', label: 'ID du Patient', value: 'externalIdValue' },
    { key: 'parent', label: 'Appartient à', value: 'parentLineage' },
    { key: 'role', label: 'Rôle', value: 'roleLabel' },
    ...extra,
  ],
});
const death = {
  key: 'death', title: 'Décès', displayExpression: 'isDeceased', fields: [
    { key: 'date', label: 'Date de décès', value: 'deathUnknown' },
    { key: 'place', label: 'Lieu du décès', value: 'deathUnknown' },
  ],
};
const parent = { key: 'parent', label: 'Appartient à', value: 'parentLineage' };

const plan = [
  ...['chw', 'chw_supervisor', 'health_area_supervisor', 'stock_manager', 'site_supervisor', 'chw_manager']
    .map(name => ['person_types', name, [profile(), death]]),
  ['person_types', 'tb_focal_point', [profile([{ key: 'cscom', label: 'CSCOM', value: 'cscomValue' }]), death]],
  ['org_unit_levels', 'c20_health_area', [{ key: 'profile', title: 'Profil', fields: [parent] }]],
  ['org_unit_levels', 'c30_supervisor_area', [{ key: 'profile', title: 'Profil', fields: [parent] }]],
  ['org_unit_levels', 'c40_chw_area', [{ key: 'profile', title: 'Profil', fields: [{ key: 'cscom', label: 'CSCOM', value: 'aireDeSante' }, parent] }]],
  ['org_unit_levels', 'c50_family', [{ key: 'profile', title: 'Profil', fields: [
    parent,
    { key: 'last_visit', label: 'Date de dernière visite', value: 'lastVisitDate' },
    { key: 'visits_this_month', label: 'Nombre de visites de ce mois-ci', value: 'visitsThisMonth' },
  ] }]],
];
const problems = [];
for (const [kind, name, cards] of plan) {
  const { validation } = structure.setSummaryCards(ctx, kind, name, cards);
  if (!validation.valid) problems.push(`${name}: ${validation.errors.join('; ')}`);
}
console.log(problems.length ? problems.join('\n') : `ALL VALID (${plan.length})`);
