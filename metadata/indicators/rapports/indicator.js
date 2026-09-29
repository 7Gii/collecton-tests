/**
 * Rapports (mini-indicateur de ligne, vue 'person') : nombre de rapports de la personne.
 */

export function showIndicator(app) {
  return true;
}

export function calculateValue(app) {
  return app.getReports({ targetId: app.selectedItem.id }).length;
}
