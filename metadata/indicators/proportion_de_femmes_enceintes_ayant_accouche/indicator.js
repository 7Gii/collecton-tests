/**
 * Proportion de femmes enceintes ayant accouché = femmes ayant un formulaire Accouchement
 * / femmes ayant eu au moins un diagnostic de grossesse positif.
 */

export function showIndicator(app) {
  return true;
}

export function calculateValue(app) {
  const diagnostiquees = new Set(
    app.getReports({ form: 'diagnostic_de_grossesse' })
      .filter((r) => r.payload.resultat_du_test_de_grossesse === 'positif')
      .map((r) => r.target_id),
  );
  const accouchees = new Set(
    app.getReports({ form: 'accouchement' })
      .filter((r) => diagnostiquees.has(r.target_id))
      .map((r) => r.target_id),
  );
  return { numerator: accouchees.size, denominator: diagnostiquees.size };
}
