/**
 * Proportion de femmes enceintes = femmes actuellement enceintes / femmes actives.
 */

// Grossesse en cours : dernier diagnostic de grossesse positif, sans accouchement
// enregistré depuis. Règle identique dans les indicateurs et la fiche (person_type.js).
function grossesseEnCours(app, femmeId) {
  const [diagnostic] = app.getReports({ form: 'diagnostic_de_grossesse', targetId: femmeId });
  if (!diagnostic || diagnostic.payload.resultat_du_test_de_grossesse !== 'positif') return false;
  return app.getReports({ form: 'accouchement', targetId: femmeId, from: diagnostic.created_at }).length === 0;
}

export function showIndicator(app) {
  return true;
}

export async function calculateValue(app) {
  const [typeFemme] = await app.find('person_types', [{ field: 'name', value: 'femme' }]);
  if (!typeFemme) return { numerator: 0, denominator: 0 };
  const femmes = await app.find('persons', [
    { field: 'person_type_id', value: typeFemme.id },
    { field: 'is_active', value: true },
  ]);
  const enceintes = femmes.filter((femme) => grossesseEnCours(app, femme.id)).length;
  return { numerator: enceintes, denominator: femmes.length };
}
