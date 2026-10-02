# Portage CHT -> Collecton (muso)

Outils qui ont transcrit les formulaires CHT de `muso-mali` dans ce projet. Ce dossier est
sous `.collecton/`, exclu de l'archive de déploiement.

- `extract.py` : exporte les XLSForms en JSON (`json/`).
- `convert.mjs` : écrit un formulaire Collecton via les opérations de `collecton-mcp`
  (verrou de plan vérifié avec `--plan`), clés = noms XLSForm, expressions traduites en
  fonctions de `form.js` ; ce qui n'est pas traduit va dans `out/<form>.review.json`.
- `xpath.mjs` (analyseur XPath -> JS) et `runtime.js` (sémantique XPath, en tête de chaque
  `form.js`) ; `forms.json` : titre, `showForm` et corrections de bugs source par formulaire.
- `test.mjs` (cas XPath), `runcheck.mjs` (exécute un formulaire avec le `FormEngine` de la
  lib), `validate.mjs` (validation du projet).
- `tasks.mjs` : crée les 19 règles de tâches (`task.json` + `task.js`) sous leurs plans
  (`task-plans.json`) ; `taskcheck.mjs` exécute les règles avec le moteur de tâches de la lib
  sur un scénario (`scenario*.json` : patients, rapports, utilisateur).
- `indicators.mjs` : crée les 22 indicateurs (fonctions de `targets.extras.js` copiées telles
  quelles derrière un adaptateur) ; `indcheck.mjs` les calcule avec le moteur de la lib sur un
  mois simulé.
- `purge.js` : source des règles de purge (portage de `purge.js` CHT, règles 0 à 14), installée
  en `utils/purge.js` par l'opération `set_purge_logic`.
- `one-off/` : scripts des étapes structure (champs des agents, fiches contact).

```bash
python3 extract.py ~/Documents/muso/Dev/github/muso/config-muso/muso-mali/forms/app patient_assessment
COLLECTON_MCP_MODE=dev node convert.mjs --project ../.. --form patient_assessment --plan <planId>
node runcheck.mjs ../.. patient_assessment s_child_temperature_pre_chw=38.5
```

Le formulaire cible ne doit pas exister (`convert.mjs` le crée). `COLLECTON_MCP_DIST` et
`COLLECTON_LIB_DIST` pointent vers les builds de `collecton-mcp` et `collecton-lib`.
