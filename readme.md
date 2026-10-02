# Prompt : créer le projet Collecton « muso » à partir de la config CHT muso-mali

> À coller dans une session Claude Code où le serveur MCP `collecton` est chargé en mode dev (`COLLECTON_MCP_MODE=dev`).

---

## Rôle et objectif

Tu es chargé de **créer un nouveau projet Collecton nommé `muso`** qui reproduit **à l'identique**, autant que la plateforme le permet, une partie de la configuration CHT de Muso Mali.

- **Source (lecture seule, ne rien modifier)** : `/Users/gilbertagbodamakou/Documents/muso/Dev/github/muso/config-muso/muso-mali`
- **Destination** : `/Users/gilbertagbodamakou/Desktop/collecton-muso` (le projet est initialisé à la racine de ce dépôt, à côté de ce readme)
- **Langues** : `fr` (par défaut), `en`, `bm`
- **Période de collecte** : perpétuelle, mensuelle (du 1er au dernier jour du mois), voir l'étape 1

Périmètre, et rien d'autre :

1. La hiérarchie et les utilisateurs (rôles et permissions)
2. Les contacts (types de personnes, formulaires d'enregistrement, fiche contact / summary)
3. Les formulaires d'évaluation enfant (< 5 ans) et adulte (≥ 5 ans), avec toutes les tâches qui en découlent et leurs formulaires de suivi
4. Les indicateurs (targets) liés à ces évaluations
5. La purge liée à ces formulaires
6. La période de collecte du projet (perpétuelle, mensuelle)

---

## Règles de travail

1. **Aucun import CHT n'existe dans Collecton.** Lis chaque fichier source (XLSX avec `openpyxl`, JS, JSON) et retranscris-le outil MCP par outil MCP. N'invente aucun champ, libellé, condition ou délai : tout doit venir de la source.
2. **Respecte le verrou de plan.** Chaque `create_*` exige un `plan_id` confirmé :
   - fais `plan_request`, puis montre-moi le `summary` et les `openQuestions` ;
   - appelle `confirm_plan` uniquement avec mon accord explicite, cité mot pour mot dans `user_confirmation`.
   - Un plan `program` n'autorise qu'**un seul** `person_type`. Fais des plans `org_unit_level`, `person_type` et `role` séparés pour la structure.
3. **Garde les identifiants.** Conserve les noms techniques CHT (noms de champs, ids de formulaires, ids de targets, noms de tâches) en snake_case pour garder la traçabilité. Les libellés sont les libellés FR exacts de la source (`translations/messages-fr.properties`, colonnes `label::fr` des XLSX, `*.properties.json`).
4. **Ne fais pas d'approximation silencieuse.** Quand une fonctionnalité CHT n'a pas d'équivalent :
   - n'approxime pas en silence ;
   - note-la dans un **rapport d'écarts** (`ECARTS_CHT_COLLECTON.md` à la racine du projet) : élément CHT, fichier et ligne source, solution retenue ou « non reproductible » ;
   - pour un écart bloquant, propose 2 options et demande-moi de choisir.
5. **Valide au fil de l'eau.** Lance `validate_project` après chaque grande étape et corrige jusqu'à 0 erreur.
6. **Ne déploie pas.** N'appelle pas `deploy_project` sans ma demande explicite.

---

## Correspondance des concepts CHT → Collecton

| CHT | Collecton |
|---|---|
| `contact_types` de type lieu | `org_unit_level` (niveau 1 = haut) |
| `person` + champ `role` du contact | un `person_type` par catégorie, rattaché à un niveau |
| `roles` + `permissions` (app_settings) | `create_role` avec `org_unit_level` + permissions Collecton |
| Utilisateur rattaché à un lieu (`user.parent`) | membre affecté à une org unit ; sync = branche de l'affectation |
| contact-summary `fields` / `cards` | `set_summary_cards` (+ `logic_script` dans le même appel) |
| `context` du contact-summary, `inputs` XForm, `instance('contact-summary')` | `app.person.attributes`, `app.getReports(...)` dans `form.js` |
| `properties.json` `expression` | fonction `showForm(app)` dans `form.js` |
| XLSForm `relevant` / `constraint` / `calculation` / `required` | `displayExpression` / `validationExpression` / `calculationExpression` / `required` (`update_form_field`) |
| `begin group` | `add_form_group` (expressions de groupe non modifiables par MCP → reporter la condition sur chaque champ) |
| `choices` | `create_option_set` |
| tâche CHT (`appliesTo: reports`, `appliesToType`) | `create_task` avec `applies_to: report` et `forms` |
| `events {days, start, end}` | `events {id, title, due_days, start_days, end_days}` |
| `appliesIf` / `resolvedIf` / `priority` | fonctions dans `task.js` (`set_task_logic`) |
| `actions[].modifyContent` (pré-remplissage) | pas d'équivalent : le formulaire de suivi relit l'évaluation source via `app.getReports` |
| targets (`count`/`percent`, `passesIf`, `goal`, `context`, `date`) | `create_indicator` + `update_indicator` (`type: number|percentage`, `goal_expression`, filtre de rôle et de période codé en JS) |
| `purge` (fonction) | `update_form` → `duration` (rétention en jours, par formulaire) |
| `uhc.visit_count.month_start_date` (mois de rapportage) | `set_collection_period` (`type: perpetual`, `start_day`, `end_day`, `margin_days`) |

---

## Étape 0 : vérifications préalables

1. Confirme que le MCP `collecton` répond (`get_project_summary` doit échouer proprement, puisqu'aucun projet n'existe encore).
2. Lis intégralement :
   - `app_settings.json` (`contact_types`, `roles`, `permissions`, `replication_depth`, `uhc`) ;
   - `forms/contact/*.xlsx` ;
   - `contact-summary.templated.js`, `contact-summary-extras.js`, `common.js`, `nools-extras.js` ;
   - `forms/app/patient_assessment.xlsx` + `.properties.json` ;
   - `forms/app/patient_assessment_over_5.xlsx` + `.properties.json` ;
   - `tasks.js`, `tasks.extras.js`, `targets.js`, `targets.extras.js`, `purge.js` ;
   - `translations/messages-fr.properties` (et `-en`, `-bm` s'ils existent) ;
   - `SETUP.md`, `harness.defaults.json`.
3. Présente-moi un plan global (étapes 1 à 10 ci-dessous) avant la première création.

---

## Étape 1 : projet

`plan_request(kind: "project")`, puis `confirm_plan`, puis `init_project` avec :
- `title: "muso"`, `organisation: "Muso"` ;
- `languages: ["fr","en","bm"]`, `default_language: "fr"`.

Puis `set_collection_period` pour une **collecte perpétuelle** (une période par mois), à poser dès la question « Collection periods » du plan `project` :
- `type: "perpetual"`, `start_day: 1`, `end_day: "LAST"` : aligné sur `uhc.visit_count.month_start_date: 1` de `app_settings.json` (le mois de rapportage CHT commence le 1er) ;
- `margin_days` : la source CHT ne la définit pas, demande-moi la valeur (elle reste modifiable ensuite dans la carte « Rapportage » de l'espace admin) ;
- la période se verrouille seule à `fin + marge` ; un rapport reçu après le verrou est marqué hors délai et rattaché à la période ouverte.

Note dans les écarts que CHT n'a pas de verrou de période : le mois CHT ne sert qu'aux compteurs (`visit_count`, targets du mois en cours).

---

## Étape 2 : hiérarchie (org unit levels)

Reproduis uniquement la hiérarchie Muso. Ignore les types hérités `district_hospital > health_center > clinic` et signale-les dans les écarts.

| level | name | Libellé FR | Source CHT | Icône CHT |
|---|---|---|---|---|
| 1 | `c10_site` | Site | c10_site | hospital |
| 2 | `c20_health_area` | Zone de Santé | c20_health_area | branch |
| 3 | `c30_supervisor_area` | Zone du Superviseur | c30_supervisor_area | medic-district-hospital |
| 4 | `c40_chw_area` | Zone de l'ASC | c40_chw_area | medic-health-center |
| 5 | `c50_family` | Famille | c50_family (`count_visits: true`) | medic-clinic |

**Attributs des lieux** (`forms/contact/cXX-create.xlsx`) : nom généré, `external_id`, `notes`, `geolocation`, `is_active`, `stm_is_active` (c30), `aire_de_sante` (c40), `mam_intrant` (c40), `family_id` (c50, regex `^[0-9]{4}$`).
- Le MCP ne pose que `ou_name` sur un niveau.
- Liste les attributs manquants dans les écarts et propose une solution : formulaire d'édition + automation `updateAttributes`, ou édition manuelle du JSON si les règles du dépôt Collecton l'autorisent. Demande-moi avant d'appliquer.

Noms générés à reproduire :
- « Site de ${contact_name} »
- « Aire de Santé - ${contact_name} »
- « Aire du Superviseur - ${contact_name} »
- « Zone de l'ASC - ${contact_name} »
- « Concession - ${contact_name} »

---

## Étape 3 : types de personnes (contacts)

Dans CHT, il n'y a qu'un type `person`. La catégorie est portée par le champ `role` (calculé dans `person-create.xlsx`, groupe `person`). Crée **un `person_type` par valeur de `role`**, rattaché au niveau où ce rôle est créé :

| person_type | Rattachement | Origine CHT | Préfixe de nom imposé |
|---|---|---|---|
| `patient` | c50_family | role = `patient` (défaut) | – |
| `chw` (ASC) | c40_chw_area | role = `chw` | `^ASC-.{2,}` |
| `chw_supervisor` (Superviseur) | c30_supervisor_area | role = `chw_supervisor` | `^SUP-.{2,}` |
| `health_area_supervisor` | c20_health_area | select_role | – |
| `tb_focal_point` (Point Focal TB) | c20_health_area | select_role | `^TBF-.{2,}` |
| `stock_manager` (Gestionnaire de stock) | c20_health_area | select_role | `^STM-.{2,}` |
| `site_supervisor` | c10_site | rôle par défaut du site | – |
| `chw_manager` (Responsable des ASC) | c10_site | select_role (filtre c10_site) | – |

**Champs à reproduire** depuis `person-create.xlsx` / `c50_family-create.xlsx`, avec types, libellés FR, `required`, `relevant`, `constraint`, choix :
- `name`, `short_name`, `sex` (Femme/Homme), `date_of_birth` + méthode (`calendar`/`approx`, âge en années/mois) ;
- `phone` et `alternate_phone` (regex 8 chiffres), `mother_phone` (nouveau-né) ;
- `person_type` (newborn / hors_zone_adult) et `s_person_type` ;
- `relation_to_parent` (Epoux, Frere/Soeur, Fils/Fille, Petit-fils/petite-fille, Cousin/Cousine, Neuveu/Niece, Pere/Mere, other, none) + `specify_other` ;
- `patient_enroll_id`, `can_patient_show_id`, `external_id` (« ID d'étude de patient Muso », longueur 11), `concession`, `menage` ;
- `muso_id` (≤ 6 caractères, chw et chw_supervisor), `language` (fr/en/bm) ;
- `is_active`, `stm_is_active` ;
- groupe `tb_focal_point_info` : site → zone de santé → `s_cscom_area`, avec la cascade de choix exacte → attribut `cscom_area`.

Pour la famille, le chef de ménage est créé dans le même formulaire, avec role `patient`.

`create_person_type` ne pose que les 4 attributs natifs (`person_name`, `person_gender`, `birthdate`, `person_is_deceased`). Traite les attributs supplémentaires comme à l'étape 2 : écart + proposition, et attends ma décision.

---

## Étape 4 : rôles utilisateurs et permissions

Crée les rôles suivants (`plan_request kind: role`, puis `create_role` avec `org_unit_level` et `permissions`). `muso_admin`, `gateway` et `muso-sih` sont des comptes console/intégration : note-les dans les écarts au lieu de créer des rôles mobiles, sauf avis contraire de ma part.

| Rôle CHT | Libellé FR | Niveau d'affectation | Offline | replication_depth CHT |
|---|---|---|---|---|
| `chw_uhc` | ASC Couverture de Santé Universelle | c40_chw_area | oui | – |
| `supervisor` | Superviseur | c30_supervisor_area | oui | 1 |
| `tb_focal_point` | Point Focal TB | c20_health_area | oui | 1 |
| `stock_manager` | Gestionnaire de stock | c20_health_area | oui | 1 |
| `chw_manager` | Responsable des ASCs | c10_site | oui | 2 |

**Correspondance des permissions CHT → Collecton** (à appliquer rôle par rôle, à partir de l'inversion de `app_settings.permissions`) :

| CHT | Collecton |
|---|---|
| `can_view_tasks` / `can_view_tasks_tab` | `permission_view_tasks` |
| `can_view_contacts` / `_tab` | `permission_view_people` |
| `can_view_reports` / `_tab` | `permission_view_reports` |
| `can_view_analytics` / `_tab` | `permission_view_indicators` |
| `can_update_reports` / `can_edit` | `permission_edit_report` |
| `can_delete_reports` | `permission_delete_report` |
| `can_create_people` | `permission_create_<person_type>` (types créables à partir du niveau du rôle et des niveaux inférieurs) |
| `can_create_places` | `permission_create_<org_unit_level>` (niveaux inférieurs) |
| `can_update_people` / `can_update_places` | `permission_edit_<…>` |
| `can_delete_contacts` | `permission_delete_<…>` |
| rôle offline | `permission_automatic_sync` + `permission_manual_sync` |
| `permission_deferred_collection` | ajoutée automatiquement par le template |

Permissions CHT sans équivalent (exports, messages, gateway, users, `can_skip_password_change`, `can_view_last_visited_date`…) : liste-les dans les écarts. N'utilise `add_permission` que si une logique (tâche, indicateur, formulaire) en a besoin.

**Important : le rôle métier `user.role`.**
- Dans les tâches et targets CHT, `user.role` est le `role` du contact de l'utilisateur (`chw`, `chw_supervisor`, `tb_focal_point`, `stock_manager`, `chw_manager`), et non le rôle CHT.
- Dans Collecton, utilise `app.currentUser.role.name` avec la table de correspondance ci-dessus, et écris-la en commentaire dans chaque script concerné.

---

## Étape 5 : fiche contact (summary cards)

Source : `contact-summary.templated.js` (l.96-181 et cartes), `contact-summary-extras.js`. Utilise `set_summary_cards`.

**`person` → sur chaque person_type concerné** :
- Âge (« Age ») ;
- Téléphone ;
- Sexe (« Sexe ») ;
- ID du Patient (`external_id`) ;
- Appartient à (lignée) ;
- Rôle (`formatRole` de `common.js` l.306-333 : TB FP, ASC, SUPERVISOR, SITE SUPERVISOR, HEALTH ZONE SUPERVISOR, STOCK MANAGER, CHW MANAGER) ;
- CSCOM (`cscom_area`, tb_focal_point seulement) ;
- carte « Décès » (date, lieu), si décédé.

**`c50_family`** :
- Appartient à ;
- « Date de dernière visite » (`oldestHomeVisitTimestamp`, sinon « jamais ») ;
- « Nombre de visites de ce mois-ci » (`numberOfDaysWithHomeVisits`). Une visite à domicile = une évaluation avec `visited_contact_uuid` égal à la famille.

**`c40_chw_area`** : CSCOM (`aire_de_sante`) + Appartient à. **`c30`, `c20`** : Appartient à.

Les cartes qui dépendent de formulaires hors périmètre ne sont pas créées, mais sont listées dans les écarts :
- Vaccination / HPV ;
- Fiche de stock ;
- cartes manager (visite de supervision, réunion individuelle, congés, réunion de groupe).

Les clés de `context` utilisées par les évaluations (`muted`, `alive`) deviennent des fonctions lues dans `showForm` :
- `alive` = `!person_is_deceased` ;
- `muted` : il n'existe pas d'équivalent natif, à noter dans les écarts.

---

## Étape 6 : formulaires d'évaluation

Un plan `program` « Prise en charge ICCM » pour le person_type `patient`, qui couvre :
- les 2 évaluations ;
- leurs formulaires de suivi (cible des tâches) ;
- les tâches ;
- les indicateurs.

### 6a. `patient_assessment` : « Evaluation des enfants jusqu'à 5 ans » (684 lignes)
### 6b. `patient_assessment_over_5` : « Evaluation des enfants à partir de 5 ans et des adultes » (453 lignes)

**Méthode de transcription des XLSX :**
- Parcours la feuille `survey` ligne par ligne avec `openpyxl`, puis :
  - chaque `select_*` → `create_option_set` (depuis `choices`) ;
  - chaque champ → `create_data_element` puis `add_form_field` ;
  - chaque `begin group` → `add_form_group`.
- Recopie à l'identique :
  - les noms de champs ;
  - les libellés FR (et EN/BM via `set_translations`) ;
  - `required`, `relevant` (→ `displayExpression`), `constraint` (→ `validationExpression` + `validationMessage`), `calculation` (→ `calculationExpression`).
- Les champs `calculate` restent des champs calculés avec les mêmes noms. Les tâches et indicateurs lisent notamment :
  - `referral`, `accompany_to_cscom`, `refer_to_cscom` ;
  - `treat_malaria`, `treat_diarrhea`, `treat_ari`, `treat_cough`, `fast_breathing`, `observe` ;
  - `assessment_date`, `visited_date`, `visited_contact_uuid` ;
  - `tb_referral`, `needs_signoff`, `cscom_area` ;
  - `treat_MAM`, `treat_SAM_without_complication`, `has_MAM`, `has_SAM_without_complication` ;
  - `has_danger_sign`, `tdr_done`, `tdr_result` ;
  - `patient_name`, `patient_sex`, `patient_phone`, `muso_id`, `chw_name`, `chw_muso_id`, `tb_diagnosis_fr/en/bm`.
- Les `inputs` (patient, famille, zone ASC, superviseur, user) se recalculent depuis `app.person`, la lignée de `app.orgUnit` et `app.currentUser`.
- `note` → champ en lecture seule ou `alertType`. Minuteur TDR (900 s) → écart.
- `prescription_summary` (db-doc, création d'un second document) → écart, avec une proposition d'automation `FORM_SUBMITTED`.

**`showForm`** (depuis `properties.json`, `expression`) :
- `patient_assessment` : person_type `patient`, vivant, utilisateur ASC affecté à un `c40_chw_area`, et (pas de date de naissance ou âge < 5 ans).
- `patient_assessment_over_5` : mêmes conditions, avec âge ≥ 5 ans.

**Formulaires de suivi à créer aussi** (cibles des tâches, à transcrire depuis `forms/app/<id>.xlsx`) :
- `referral_followup`, `referral_followup_under_5` ;
- `treatment_followup`, `treatment_followup_over_5` ;
- `moderate_malnutrition_followup`, `severe_malnutrition_followup` ;
- `notification_to_chw`, `redo_tdr` ;
- `tb_test_result_fp`, `tb_test_result_chw`.

Pour chacun : `showForm` renvoie `false` (le formulaire ne s'ouvre que par une tâche). Les champs `t_*` que CHT remplissait par `modifyContent` sont recalculés en lisant le rapport source avec `app.getReports({ form, targetId })`.

---

## Étape 7 : tâches liées aux évaluations

Source : `tasks.js` l.598-833 et les gabarits de `tasks.extras.js` (`defaultedTaskTemplate` l.166-199, `specificTreatmentTemplate` l.237-264, malnutrition l.266-361, superviseur l.439-563).

**Règles communes à reproduire dans chaque `task.js` :**
- `appliesIf` exige : patient vivant, non muet, et rôle de l'utilisateur dans `displayForUserAt` (par défaut ASC / c40 ; superviseur / c30 ; point focal / c20).
- Date de base de l'échéance = `visited_date` si `visited_contact_uuid` est renseigné, sinon `s_reported.s_reported_date`, sinon la date du rapport. Utilise `due_date` (fonction) quand c'est nécessaire.
- Si la date `assessment_date + jours` du dernier événement est antérieure à la date du rapport, l'échéance devient la date du rapport.
- `resolvedIf` par défaut : un rapport de l'`action_form` existe dans `[due - start, due + end + 1]`.

**Tâches à créer :**

| # | name | Formulaire source → action | Titre FR | Condition | Événements (jours / start / end) | Résolution supplémentaire |
|---|---|---|---|---|---|---|
| 1 | `patient_assessment_over_5_yields_referral_followup` | over_5 → `referral_followup` | Suivi référence | `referral` ou `accompany_to_cscom` ou `refer_to_cscom` = 'true' | 1, 2, 3 / 0 / 0 ; priorité haute | TB : `tb_test_result_chw` soumis ; dernier suivi `close_out` ou `referral` = 'true' ; nouvelle évaluation over_5 |
| 2 | `patient_assessment_yields_referral_followup_under_5` | patient_assessment → `referral_followup_under_5` | Suivi référence (moins de 5 ans) | idem | idem | idem |
| 3 | `referral_followup_yields_self` / `referral_followup_under_5_yields_self` | suivi → lui-même | idem | `referral='true'` dans le suivi | 1, 2, 3 | close_out, referral, nouvelle évaluation |
| 4 | traitement unique (malaria / ari / diarrhea) | patient_assessment → `treatment_followup` | Suivi traitement | ni accompagné ni référé ; un seul `treat_*` vrai | malaria 1, 2, 3 (0/0) ; ari 1, 2, 3, 5 (J5 : 1/3) ; diarrhée 5 (1/3) | close_out, nouvelle évaluation, referral |
| 5 | `patient_assessment_yields_treatment_followup` (multiple) | patient_assessment → `treatment_followup` | Suivi traitement | plus d'un traitement | 1, 2, 3, 5 (J5 : 1/3) | idem |
| 6a | Suivi MAM | (person) → `moderate_malnutrition_followup` | Suivi MAM | dernière évaluation avec `treat_MAM='true'`, âgée de moins de 97 jours | 7, 14, 21, 28, 42, 56, 70, 84 ; start 2 au premier puis 0 ; end = écart au suivant − 1, 14 au dernier ; priorité moyenne | `final_cat` ∈ {refer_with_no_followup, abandoned} ou `patient_deceased` |
| 6b | Suivi MAS sans complication | (person) → `severe_malnutrition_followup` | Suivi MAS Sans Complication | `treat_SAM_without_complication='true'` | 7, 14, 21, 28, 35, 42, 49, 56, 63, 70 | idem + `recovered` |
| 7 | traitement over_5 | over_5 → `treatment_followup_over_5` | Suivi traitement | `referral != 'true'` et `treat_malaria='true'` | 1, 2, 3 | close_out, referral, nouvelle évaluation over_5 |
| 8 | alerte superviseur (×2 formulaires) | évaluation → `notification_to_chw` | Alerte erreur TDR | dernière évaluation avec `verified === false` ; superviseur c30 | 0 / 0 / 3 | – |
| 9 | TDR à refaire (×2) | évaluation → `redo_tdr` | TDR à refaire | `verified === false` et `needs_signoff === 'true'` | 0 / 0 / 3 | – |
| 10 | Résultat TB point focal (×2) | évaluation → `tb_test_result_fp` | Résultat TB (Point Focal) | `tb_referral='true'`, `cscom_area` = celle de l'utilisateur, rôle tb_focal_point | 0 / 0 / 30 ; priorité haute | – |
| 11 | Résultat TB ASC (×2) | évaluation → `tb_test_result_chw` | Résultat TB | `tb_referral='true'`, rôle chw | 4 / 0 / 180 ; priorité haute | – |

Pour les tâches 1, 2, 3, 4, 5 et 7, vérifie les noms, icônes et conditions exacts dans `tasks.js` avant de les créer. Le tableau est un résumé, la source fait foi.

**Tâches de type « contacts » (6a, 6b, 8) :** utilise `applies_to: person`. Retrouve l'évaluation de référence avec `app.getReports`. Les libellés dynamiques (`Suivi N - x/total`, `chw_name (chw_muso_id)`) vont en `description` ou en `title` d'événement.

**Champ `verified` :** il est posé par la revue superviseur CHT. Si Collecton n'a pas d'équivalent, mets-le dans les écarts et demande-moi comment le traiter.

---

## Étape 8 : indicateurs liés aux évaluations

Source : `targets.js`, `targets.extras.js` (`basicTemplate` l.88-109).

**Règles communes :**
- visibles pour le rôle ASC uniquement (`context: user.role === "chw"`) ;
- patient vivant et non muet ;
- date du rapport = date dynamique (comme pour les tâches) ;
- période = mois en cours : CHT n'a pas de `date_range` explicite, donc le mois en cours s'applique par défaut. Code le filtre `from` / `to` avec `app.now`, sur les mêmes bornes que la période de collecte (du 1er au dernier jour du mois).

**Indicateurs d'erreurs de protocole :** tous en `type: percentage`, objectif 0. `numerator` = évaluations où l'erreur existe, `denominator` = évaluations des formulaires concernés. Ids et libellés FR exacts depuis `messages-fr.properties` l.1553-1635 :
- `percent-assessments-without-protocol-errors` (OU de toutes les erreurs) ;
- `danger-signs-without-referral-chw-error` ;
- `fever-present-malaria-rdt-not-performed-chw-error` ;
- `uncomplicated-malaria-without-act-chw-error` ;
- `incorrect-alu-dose-chw-error` (règles d'âge → dose ALU) ;
- `uncomplicated-diarrhea-without-zinc-chw-error`, `incorrect-zinc-dose-chw-error` ;
- `pneumonia-without-amoxicillin-chw-error`, `incorrect-amoxicillin-dose-chw-error` ;
- `incorrect-vitamin-a-dose-chw-error`, `incorrect-albendazole-dose-chw-error`, `incorrect-iron-folic-acid-dose-chw-error` ;
- `mam-no-complete-treatment-first-day-chw-error`, `sam-no-complete-treatment-first-day-chw-error` ;
- `incorrect-paracetamol-dose-chw-error` ;
- `no-24-hour-follow-up-chw-error`, `no-48-hour-follow-up-chw-error`, `no-72-hour-follow-up-chw-error`, `no-5-day-follow-up-chw-error` ;
- `condition-worsened-without-referral-chw-error`.

Chaque `passesIf` est une fonction de `targets.extras.js`. Traduis-la ligne à ligne en JS Collecton dans `indicator.js`, en conservant les mêmes noms de fonctions.

**Autres indicateurs :**
- `home_visits` (« Visite à domicile ») : `number`, compte des évaluations avec `visited_contact_uuid`, dédoublonnées par famille et par jour.
- `2-home-visits-per-fam` (« % de ménages ayant reçu au moins 2 VAD ») : `percentage`, objectif 100, ménages avec au moins 2 visites.

**Chemins suspects dans la source :** demande-moi s'il faut les reproduire tels quels (« identique ») ou les corriger, puis note la décision dans les écarts.
- `isFeverWithoutTdrChw` lit `s_malnutrition_mRDT.malnutrition_tdr_not_done` et `s_malaria.malaria_tdr_not_done` (sans `s_`), et `child_temperature` pour le formulaire over_5.
- `didNotAccompanyPatientToCscom` et `aggravatedWithoutReferral` lisent `group_diagnosis.s_accompany_refer_under5` / `over5`. Ces champs n'existent pas, le formulaire a `s_accompany_refer_CSCOM`.
- `masWithoutComplication…` lit `group_diagnosis.not_give_plumpy_nut`.

---

## Étape 9 : purge (rétention)

Source : `purge.js` (identique à `app_settings.purge`, exécutée chaque samedi à 22 h). Dans Collecton, seule la rétention par formulaire existe : `update_form` → `duration` en jours.

| Formulaire | duration (jours) |
|---|---|
| `patient_assessment` | 100 (cas MAM/MAS conservés pour les suivis ; les autres seraient purgés à 30 j, voir ci-dessous) |
| `patient_assessment_over_5` | 30 |
| `referral_followup`, `referral_followup_under_5`, `treatment_followup`, `treatment_followup_over_5` | 10 |
| `redo_tdr` | 30 |
| `moderate_malnutrition_followup`, `severe_malnutrition_followup` | 100 |
| `notification_to_chw`, `tb_test_result_fp`, `tb_test_result_chw` | non listés dans les règles 1 à 9, donc purgés à 300 j (règle 0) |

Écarts à documenter :
- **Purge conditionnelle de `patient_assessment` à 30 j** si ni `has_MAM` ni `has_SAM_without_complication` (règle 4.1). Elle n'est pas reproductible avec `duration` seul. Propose une option (par exemple 100 j pour toutes les évaluations) et demande-moi.
- **Purges par rôle** (règles 10 à 13 : chw_uhc, tb_focal_point > 30 j et formulaires autorisés, stock_manager, chw_manager) : pas d'équivalent. La sync par branche limite déjà ce qui descend sur l'appareil. Documente ce qui change.
- **`purge_tasks: [{event_name: 'self_assessment_'}]`** : les tâches Collecton ne sont pas stockées, donc rien à faire. Note-le.

---

## Étape 10 : traductions, validation, livrables

1. `sync_translations`, puis `list_untranslated` pour `en` et `bm`, puis `set_translations` à partir de `messages-en.properties`, `messages-bm.properties` et des colonnes `label::en` / `label::bm` des XLSX.
2. Relance `validate_project` jusqu'à 0 erreur.
3. Remets-moi :
   - le récapitulatif de ce qui a été créé (période de collecte, niveaux, person types, rôles, formulaires avec leur nombre de champs, tâches, indicateurs, rétentions) ;
   - un tableau de **contrôle de parité** : élément CHT → élément Collecton → statut (identique / adapté / non reproductible) ;
   - `ECARTS_CHT_COLLECTON.md` complet ;
   - la liste des décisions que j'ai prises pendant la session.

Ne déploie pas. Je lancerai `deploy_project` moi-même ou je te le demanderai.
