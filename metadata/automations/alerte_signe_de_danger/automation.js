/**
 * Alerte signe de danger : déclenchée à la soumission d'une CPN ou d'une visite postnatale.
 * Note : Collecton n'expose pas encore d'API d'envoi d'alerte aux automatisations ;
 * l'action journalise l'alerte côté serveur en attendant.
 */

function signesDeDanger(payload) {
  const signes = [];
  const cpn = payload.signes_de_danger;
  const liste = Array.isArray(cpn) ? cpn : typeof cpn === 'string' && cpn ? cpn.split(' ') : [];
  for (const s of liste) if (s && s !== 'aucun') signes.push(s);
  if (payload.etat_de_la_mere === 'signes_danger') signes.push('mere_signes_danger');
  if (payload.etat_du_nouveaune === 'signes_danger') signes.push('nouveau_ne_signes_danger');
  return signes;
}

export function checkCondition(app) {
  return signesDeDanger(app.submission?.payload || {}).length > 0;
}

export function runAction(app) {
  const signes = signesDeDanger(app.submission?.payload || {});
  console.log(
    `[ALERTE SIGNE DE DANGER] formulaire=${app.form?.name} personne=${app.submission?.personId} ` +
      `unite=${app.submission?.orgUnitId} signes=${signes.join(', ')}`
  );
}
