/**
 * Logic for the 'menage' Org Unit Level.
 * Exports the value functions of the summary cards ('dashboard'). The summary of an
 * org unit only sees its own reports through app.getReports: the women and their
 * reports are read with app.find (raw rows, payload stored as a JSON string).
 */

async function femmesDuMenage(app) {
  const [typeFemme] = await app.find('person_types', [{ field: 'name', value: 'femme' }]);
  if (!typeFemme) return [];
  return app.find('persons', [
    { field: 'org_unit_id', value: app.selectedItem.id },
    { field: 'person_type_id', value: typeFemme.id },
    { field: 'is_active', value: true },
  ]);
}

// Rapports actifs d'une personne pour un formulaire, du plus récent au plus ancien.
async function rapports(app, formulaire, personneId) {
  const [form] = await app.find('form_definitions', [{ field: 'name', value: formulaire }]);
  if (!form) return [];
  const lignes = await app.find('data_submissions', [
    { field: 'target_id', value: personneId },
    { field: 'form_id', value: form.id },
    { field: 'is_active', value: true },
  ]);
  return lignes
    .map((r) => ({ ...r, payload: typeof r.payload === 'string' ? JSON.parse(r.payload || '{}') : r.payload || {} }))
    .sort((a, b) => b.created_at - a.created_at);
}

// Même règle que grossesseEnCours (indicateurs, fiche femme) : dernier diagnostic
// positif, sans accouchement enregistré depuis.
async function estEnceinte(app, femmeId) {
  const [diagnostic] = await rapports(app, 'diagnostic_de_grossesse', femmeId);
  if (!diagnostic || diagnostic.payload.resultat_du_test_de_grossesse !== 'positif') return false;
  const accouchements = await rapports(app, 'accouchement', femmeId);
  return !accouchements.some((a) => a.created_at >= diagnostic.created_at);
}

// La carte ne s'affiche que pour les ménages qui comptent au moins une femme.
export async function aDesFemmes(app) {
  return (await femmesDuMenage(app)).length > 0;
}

export async function nombreFemmes(app) {
  return (await femmesDuMenage(app)).length;
}

export async function nombreFemmesEnceintes(app) {
  let enceintes = 0;
  for (const femme of await femmesDuMenage(app)) {
    if (await estEnceinte(app, femme.id)) enceintes++;
  }
  return enceintes;
}
