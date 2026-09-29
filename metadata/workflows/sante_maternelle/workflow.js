/**
 * Logique du workflow 'Santé maternelle'.
 * Chaque fonction reçoit `app` (workflow, person, submission, getReports).
 */

const AGE_MIN = 15;
const AGE_MAX = 40;

function ageEnAnnees(dateNaissance) {
  if (!dateNaissance) return null;
  const naissance = new Date(dateNaissance);
  if (isNaN(naissance.getTime())) return null;
  return (Date.now() - naissance.getTime()) / (365.25 * 24 * 60 * 60 * 1000);
}

// Disponibilité : femmes vivantes âgées de 15 à 40 ans.
export function estEligible(app) {
  const attributs = app.person?.attributes || {};
  if (attributs.person_is_deceased === true) return false;
  const age = ageEnAnnees(attributs.birthdate);
  return age !== null && age >= AGE_MIN && age <= AGE_MAX;
}

// Décision après le diagnostic : test de grossesse positif.
export function grossesseConfirmee(app) {
  const payload = app.workflow?.context?.submissions?.diagnostic?.payload
    || app.submission?.payload || {};
  return payload.resultat_du_test_de_grossesse === 'positif';
}

// Décision après chaque CPN : l'accouchement a-t-il eu lieu ?
export function accouchementSurvenu(app) {
  return app.submission?.payload?.accouchement_survenu === 'oui';
}

// Clôture automatique : femme déclarée décédée (fiche ou visite postnatale).
export function doitEtreCloture(app) {
  if (app.person?.attributes?.person_is_deceased === true) return true;
  return app.submission?.payload?.etat_de_la_mere === 'decedee';
}
