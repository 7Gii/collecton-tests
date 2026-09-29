/**
 * Logic for the 'femme' Person Type.
 * Also exports the value functions of the summary cards ('dashboard').
 */

// Grossesse en cours : dernier diagnostic de grossesse positif, sans accouchement
// enregistré depuis. Règle identique dans les indicateurs et la fiche (person_type.js).
function grossesseEnCours(app, femmeId) {
  const [diagnostic] = app.getReports({ form: 'diagnostic_de_grossesse', targetId: femmeId });
  if (!diagnostic || diagnostic.payload.resultat_du_test_de_grossesse !== 'positif') return false;
  return app.getReports({ form: 'accouchement', targetId: femmeId, from: diagnostic.created_at }).length === 0;
}

export function statutGrossesse(app) {
  const [diagnostic] = app.getReports({ form: 'diagnostic_de_grossesse', targetId: app.selectedItem.id });
  if (!diagnostic) return 'Aucun diagnostic';
  return grossesseEnCours(app, app.selectedItem.id) ? 'Enceinte' : 'Non enceinte';
}

// Date affichée en jj/mm/aaaa (les champs date sont saisis en aaaa-mm-jj).
function formatDate(value) {
  const [annee, mois, jour] = String(value || '').split('-');
  return annee && mois && jour ? `${jour}/${mois}/${annee}` : null;
}

// Diagnostic positif de la grossesse en cours, ou null si la femme n'est pas enceinte.
function diagnosticEnCours(app) {
  if (!grossesseEnCours(app, app.selectedItem.id)) return null;
  const [diagnostic] = app.getReports({ form: 'diagnostic_de_grossesse', targetId: app.selectedItem.id });
  return diagnostic;
}

// Date prévue d'accouchement de la grossesse en cours (masquée hors grossesse).
export function datePrevueAccouchement(app) {
  const diagnostic = diagnosticEnCours(app);
  return diagnostic ? formatDate(diagnostic.payload.date_prevue_daccouchement) : null;
}

// Consultations prénatales depuis le diagnostic de la grossesse en cours (masqué hors grossesse).
export function nombreCpn(app) {
  const diagnostic = diagnosticEnCours(app);
  if (!diagnostic) return null;
  return app.getReports({ form: 'consultation_prenatale_cpn', targetId: app.selectedItem.id, from: diagnostic.created_at }).length;
}
