/**
 * Grossesse en cours (mini-indicateur de ligne, vue 'person') : badge d'alerte rouge
 * sur la ligne et la fiche des femmes actuellement enceintes, rien sinon.
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

export function calculateValue(app) {
  return grossesseEnCours(app, app.selectedItem.id) ? { icon: 'alert', color: 'red' } : null;
}
