# Écarts CHT -> Collecton (projet muso)

Source : `config-muso/muso-mali` (CHT). Chaque ligne indique l'élément CHT, son
emplacement, la solution retenue et son statut : **identique**, **adapté** ou
**non reproductible**.

## Synthèse (02/10/2026)

**Créé dans le projet** (validation : 0 erreur ; avertissements = fenêtres de tâches de 0 jour, comme CHT) :

| Élément | Nombre | Détail |
|---|---|---|
| Période de collecte | 1 | perpétuelle, du 1er au dernier jour du mois, marge 5 jours |
| Niveaux | 5 | `c10_site` à `c50_family`, avec leurs champs (cascade Site > Zone > CSCOM) |
| Types de personnes | 8 | `patient`, `chw`, `chw_supervisor`, `health_area_supervisor`, `tb_focal_point`, `stock_manager`, `site_supervisor`, `chw_manager` ; fiches contact |
| Rôles | 5 | `chw_uhc`, `supervisor`, `tb_focal_point`, `stock_manager`, `chw_manager` |
| Formulaires | 12 | 2 évaluations et 10 suivis, 2 155 champs transcrits, rétention posée |
| Tâches | 19 | références, traitements, MAM / MAS, TB |
| Indicateurs | 22 | visites à domicile et 20 erreurs de protocole |
| Purge | 1 | `utils/purge.js`, règles 0 à 14 de CHT |
| Traductions | fr, en, bm | EN complet ; BM : 38 textes sans bambara dans les sources (liste plus bas) |

**Parité** (lignes des tableaux ci-dessous) : 44 identiques, 37 adaptées, 21 non reproductibles.

**Ajouté à Collecton pour ce portage** (lib, mobile, MCP) : champs supplémentaires des types
(`add_attribute_field`, `update_attribute_field`), nom technique à la création (`name`), filtres
d'options (`optionFilter`, propriétés des options), valeur par défaut des attributs, champs
d'affichage sans data element (titres colorés `headingLevel` / `color`, minuteur `timer`), clé de
champ et options de groupe (`key`, `condition`, `full_page`), `app.lineage` et `app.personsAt`,
valeurs passées par les tâches (`prefill` -> `app.taskInputs`), règles de purge par projet
(`utils/purge.js`, `set_purge_logic`).

**Outils du portage** : `.collecton/cht-port/` (exclu du déploiement), voir son `README.md`.

**Reste à faire hors de ce dépôt** : reconstruire l'app mobile (`npm run build:console`),
déployer (`deploy_project`, sur demande), tester sur téléphone, compléter les textes bambara.

**Textes sans bambara dans les sources** (l'application affiche le français) :

- 4 chiffres
- Cette dose n'est pas applicable à cet enfant.
- Couleur de la bande de shakir passée
- Différence entre date réelle de prise en charge et date du jour
- Est-ce que le fréquence respiratoire est rapide?
- Est-ce que l’enfant est vacciné contre la rougeole?
- Est-ce que l’enfant était vacciné lors dernier suivi?
- Géolocalisation
- L'ASC a choisi le CAT que l'application
- Le patient a été accompagné/reféré au CSCOM
- Le patient devrait continuer le suivi
- Le patient doit-il être referré?
- Nombre de bandes rouges
- Nombre de fois declaré vaccinée contre la rougeole
- Nombre de présence de l’enfant
- PB du dernier suivi
- Pas de farine enrichie données
- Patient has been accompany/referred to the CSCOM
- Pourquoi la bande de Shakir n’a pas été utilisé
- Présence de l’enfant durant le suivi
- Périmètre branchial (PB)
- Quantité de Plumpy Nut donnée cette semaine
- Quantité de Plumpy Nut donnée la semaine dernière
- Quelle est la suite à donner au suivi?
- Raison pour laquelle ne TDR n’a pas été fait
- Responsable des ASC
- Sachet de farine enrichie donnée: {app.currentForm["enriched_flour_dosage_label"] ?? ''} unité(s)
- Si le patient a les oedèmes bilatéraux
- Spécifier :
- Superviseur de site
- Superviseur zone de santé
- Traiter la Malnutrition Aigue Moderée
- Traiter la Malnutrition Aigue Moderée avec la farine enrichie
- Traiter la Malnutrition Aigue Moderée avec le plumpy sup
- Traiter la Malnutrition Aigue Sévère sans complication
- Utilisations des stocks
- date réelle de prise en charge
- real reported date

## Projet

| Élément CHT | Source | Solution Collecton | Statut |
|---|---|---|---|
| Mois de rapportage (`uhc.visit_count.month_start_date: 1`) | `app_settings.json` | Période de collecte perpétuelle, du 1er au dernier jour du mois, marge 5 jours (verrou le 6 à 00:00 UTC). CHT n'a pas de verrou : le mois ne sert qu'aux compteurs. | adapté |

## Hiérarchie (org unit levels)

| Élément CHT | Source | Solution Collecton | Statut |
|---|---|---|---|
| `c10_site` > `c20_health_area` > `c30_supervisor_area` > `c40_chw_area` > `c50_family` | `app_settings.json` `contact_types` | 5 niveaux, mêmes noms techniques, niveaux 1 à 5 | identique |
| Types hérités `district_hospital` > `health_center` > `clinic` | `app_settings.json` `contact_types` | Non repris (hors périmètre) | non reproductible |
| Icônes CHT des niveaux et des personnes | `app_settings.json`, `resources/` | Mêmes icônes copiées dans `assets/icons` (`.collecton/cht-port/media.mjs`) | identique |
| Nom généré du lieu (« Site de ${contact_name} »...) à partir du contact primaire créé dans le même formulaire | `forms/contact/cXX-create.xlsx` (`generated_name`, `is_name_generated`) | Nom saisi à la main (`ou_name`) : la fiche d'un lieu ne crée pas de personne | adapté |
| Contact primaire (`contact`, `create_new_person`, `select_person`) | `forms/contact/cXX-create.xlsx` | Non repris : le lieu et ses personnes sont créés séparément | non reproductible |
| Formulaires de création et d'édition distincts | `forms/contact/cXX-create.xlsx` / `cXX-edit.xlsx` | Une seule fiche par niveau (attributs) ; champs des deux formulaires réunis | adapté |
| `geolocation` calculée depuis le GPS du téléphone | `cXX-create.xlsx` (`inputs/meta/location`) | Champ `geolocation` de type geopoint, capturé par l'utilisateur | adapté |
| `is_active` (caché, défaut `true`) | `c20/c30/c40-create.xlsx` | Retiré (décision du 02/10/2026) : `is_active` est une colonne réservée des enregistrements Collecton (entité supprimée) ; un lieu ou une personne existe tant qu'il n'est pas supprimé | non reproductible |
| `stm_is_active`, `mam_intrant`, `is_in_proccm`, `disable_malaria_vaccine` visibles des comptes `pm_*` / `medic` seulement | `c30/c40-edit.xlsx` (`user_is_pm`) | Toujours affichés (décision du 02/10/2026) ; valeurs par défaut CHT conservées (`false`, `plumpy`, `true`) | adapté |
| Cascade Site > Zone de santé > CSCOM (`choice_filter`) | `c30/c40-*.xlsx`, `person-*.xlsx` | Filtres d'options ajoutés à Collecton (`optionFilter` + `properties` des options) ; listes `site`, `ha`, `cscom_area` | identique |
| `s_cscom_area` en choix multiple sur c30 | `c30_supervisor_area-edit.xlsx` | Clé `s_cscom_areas` (un data element porte un seul type ; c40 et le point focal TB gardent `s_cscom_area`) | adapté |
| Liste `yes_no` (valeurs `true` / `false`) | choix des XLSX | Liste `cht_yes_no` (la liste native `yes_no` vaut `yes` / `no`) ; valeurs identiques | identique |
| `family_id` affiché seulement si la langue de l'utilisateur est `fr` ; obligatoire en édition, facultatif en création | `c50_family-create/edit.xlsx` | Toujours affiché et obligatoire (une seule fiche ; langue non exposée aux scripts) ; regex `^[0-9]{4}$` conservée | adapté |
| Déplacement d'une famille vers une autre zone d'ASC (`move_to/next_chw_area`) | `c50_family-edit.xlsx` | Non repris | non reproductible |
| Champs d'audit `created_by*`, `last_edited_by*` | `cXX-*.xlsx` | Non repris : Collecton trace l'auteur côté serveur | adapté |

## Types de personnes

| Élément CHT | Source | Solution Collecton | Statut |
|---|---|---|---|
| Un seul type `person`, catégorie portée par `role` | `person-create.xlsx` (`role`, `select_role`) | 8 person types, un par valeur de `role` (`patient`, `chw`, `chw_supervisor`, `health_area_supervisor`, `tb_focal_point`, `stock_manager`, `site_supervisor`, `chw_manager`) ; attribut caché `role` à la valeur du type | adapté |
| Rôle par défaut `patient` à `c20_health_area` | `person-create.xlsx` (`role`) | Non repris : patients seulement dans les familles (décision du 02/10/2026) | non reproductible |
| `sex` (`female` / `male`) | `person-create.xlsx` | Champ natif `person_gender` (`F` / `M`) ; les scripts traduisent les valeurs là où CHT lit `female` / `male` | adapté |
| Clés `name`, `date_of_birth` | `person-create.xlsx` | Champs natifs `person_name`, `birthdate` | adapté |
| `short_name` | contact primaire des `cXX-create.xlsx` (absent de `person-create.xlsx`) | Non repris (le contact primaire n'est pas repris) | non reproductible |
| Formulaires de création et d'édition distincts | `person-create.xlsx` / `person-edit.xlsx` | Une seule fiche par type ; `muso_id` obligatoire (règle de l'édition) | adapté |
| Champs `fr` seulement (`s_person_type`, `patient_enroll_id`, `can_patient_show_id`) | `person-*.xlsx` (`place_language`) | Toujours affichés (langue non exposée aux scripts) | adapté |
| `is_active` (création cachée, édition par les comptes `pm`) | `person-*.xlsx` | Retiré (décision du 02/10/2026) : colonne réservée de Collecton | non reproductible |

## Rôles et permissions

| Élément CHT | Source | Solution Collecton | Statut |
|---|---|---|---|
| Rôles mobiles `chw_uhc`, `supervisor`, `tb_focal_point`, `stock_manager`, `chw_manager` | `app_settings.json` `roles` | 5 rôles, mêmes noms, affectés à c40, c30, c20, c20, c10 | identique |
| `muso_admin`, `gateway`, `muso-sih` | `app_settings.json` `roles` | Non créés : comptes console / intégration (administration dans l'espace admin) | non reproductible |
| `replication_depth` (1 ou 2) | `app_settings.json` | Un membre synchronise toute la branche de son affectation (pas de profondeur) | adapté |
| `can_view_*` / `can_view_*_tab` | `app_settings.json` `permissions` | `permission_view_tasks` / `_people` / `_reports` / `_indicators` | identique |
| `can_edit`, `can_update_reports` | `app_settings.json` `permissions` | `permission_edit_report`, plus `permission_edit_<type>` pour ce que le rôle crée (décision du 02/10/2026) | adapté |
| `can_create_people` / `can_create_places` | `app_settings.json` `permissions` | `permission_create_<type>` des types et niveaux du niveau du rôle et en dessous (chw_uhc : chw, patient, c50_family ; tb_focal_point : c30 à c50 et les 6 types de c20 et en dessous) | adapté |
| Rôle offline | `app_settings.json` `roles.*.offline` | `permission_automatic_sync` + `permission_manual_sync` ; `permission_deferred_collection` ajoutée par le template | identique |
| `can_export_*`, `can_update_messages`, `can_view_data_records`, `can_view_unallocated_data_records`, `can_skip_password_change`, `can_view_last_visited_date` | `app_settings.json` `permissions` | Pas d'équivalent Collecton | non reproductible |
| `user.role` (rôle du contact de l'utilisateur) dans tâches et targets | `tasks.js`, `targets.js` | `app.currentUser.role.name` : chw -> `chw_uhc`, chw_supervisor -> `supervisor`, tb_focal_point, stock_manager, chw_manager (même nom) | adapté |

## Fiche contact (summary cards)

| Élément CHT | Source | Solution Collecton | Statut |
|---|---|---|---|
| `fields` person : Age, Téléphone, Sexe, ID du Patient, Appartient à, Rôle (`formatRole`) | `contact-summary.templated.js` l.112-126, `common.js` l.306-333 | Carte « Profil » sur les 8 types ; fonctions du script du type (`ageLabel`, `parentLineage` via `app.find`, `roleLabel` = table `formatRole`) | identique |
| CSCOM du point focal TB | `contact-summary.templated.js` (`contact.role === 'tb_focal_point'`) | Champ CSCOM sur la carte du type `tb_focal_point` | identique |
| Carte « Décès » (date, lieu depuis `death_report`) | `contact-summary.templated.js` | Affichée si « Décédé » ; date et lieu = « Inconnu » (pas de date de décès ni de `death_report` dans le périmètre) | adapté |
| `c50_family` : Appartient à, Date de dernière visite (sinon « jamais »), Nombre de visites de ce mois-ci | `contact-summary.templated.js`, `contact-summary-extras.js` l.25-38 | Carte « Profil » ; visite = rapport dont `visited_contact_uuid` est la famille, date = `visited_date` sinon date du rapport ; jours distincts du mois en cours (mois Muso = du 1er au dernier jour) | identique |
| `oldestHomeVisitTimestamp` | `contact-summary-extras.js` l.25 | Nom trompeur côté CHT (il prend le maximum) : la date la plus récente est reprise | identique |
| `c40_chw_area` : CSCOM + Appartient à ; `c30`, `c20` : Appartient à | `contact-summary.templated.js` | Carte « Profil » | identique |
| Cartes vaccination / HPV, fiche de stock, cartes manager (supervision, réunion individuelle, congés, réunion de groupe) | `contact-summary.templated.js`, `contact-summary-extras.js` | Non reprises : formulaires hors périmètre | non reproductible |
| `context.alive` | `contact-summary.templated.js` | `!person_is_deceased`, lu dans `showForm` des évaluations | identique |
| `context.muted` | `contact-summary.templated.js` | Pas d'équivalent natif (pas de mise en sourdine) | non reproductible |
| Libellés « Age », « 3 ans »... traduits par CHT | `messages-*.properties` | Valeurs calculées en français dans les scripts (langue par défaut du projet) | adapté |

## Formulaires (évaluations et suivis)

Transcription automatique des 12 XLSX (`forms/app/*.xlsx`) : 2 155 champs, clés = noms XLSForm, chaque expression traduite en fonction de `form.js` (runtime XPath : types, comparaisons, dates en jours). Aucune expression non traduite.

| Élément CHT | Source | Solution Collecton | Statut |
|---|---|---|---|
| `relevant` / `constraint` / `calculation` / `required` / `default` / `choice_filter` | `survey` des XLSX | Fonctions de `form.js` (`rel_*`, `val_*`, `calc_*`, `req_*`, `def_*`, `filt_*`) ; une contrainte n'est pas testée sur une valeur vide, comme ODK | identique |
| Calcul d'un champ non pertinent (vide en ODK) | XLSForm | Les calculs renvoient `''` quand le champ ou un groupe parent est masqué | identique |
| Valeur d'un champ saisi puis masqué (effacée en ODK) | XLSForm | La valeur reste dans le rapport | adapté |
| Groupes imbriqués | `begin group` | Groupe de 1er niveau = page (mode wizard) ; groupes imbriqués aplatis, leur condition reportée sur leurs champs, leur libellé en titre de niveau 3 | adapté |
| Noms en double dans un formulaire (64) | XLSX | Clé préfixée par le groupe parent (`<groupe>_<nom>`) | adapté |
| Noms qui ne diffèrent que par la casse (`c_vitaminA_dosage` / `c_vitamina_dosage`, `CAT` / `cat`...) | XLSX | Clé de champ inchangée ; le data element ou la liste prend un nom préfixé par le formulaire (`<formulaire>__<nom>`), un nom étant un nom de fichier | adapté |
| Groupe `inputs` (contact, lignée, ASC, utilisateur) | XLSX | Lu à l'exécution : `app.person`, `app.lineage` (ajouté à Collecton), ASC de la zone via `app.personsAt` (ajouté), `app.currentUser` ; non stocké dans le rapport | adapté |
| `inputs/user/language` | XLSX | Toujours `fr` (langue non exposée aux scripts) | adapté |
| `inputs/user/is_in_sih` | XLSX | Toujours `false` | adapté |
| Entrées `t_*` passées par les tâches (`modifyContent`) | `tasks.js` | Fonction `prefill` des tâches, lue par le formulaire en `app.taskInputs` et stockée en `inputs_<nom>` | identique |
| `note` avec `h1` / `h2` / `h3` et couleur | `appearance` | Titres colorés ajoutés à Collecton (`headingLevel`, `color`) | identique |
| Minuteur `countdown-timer` (TDR 900 s, respiration 60 s) | `appearance`, `default` | Champ `timer` ajouté à Collecton (`timerSeconds`, bip et vibration) | identique |
| Icônes des formulaires | `*.properties.json` `icon`, `resources/` | Mêmes icônes copiées dans `assets/icons`, `Form.icon` | identique |
| Images des notes (TDR, résultats) | colonne `media::image`, `*-media/images` | Copiées dans `assets/images`, `FormField.image` (affichées hors-ligne) | identique |
| Icônes Font Awesome dans les libellés | `label::fr` | Retirées | adapté |
| Widgets `db:person` / `db:health_center` | XLSX | Non repris : la cible du formulaire est la personne | non reproductible |
| `instance::db-doc` (`prescription_summary`) | `patient_assessment.xlsx` | Non repris (pas de second document) | non reproductible |
| Visibilité (menu « + ») | `*.properties.json` | `showForm` : évaluations pour `chw_uhc`, patient vivant, selon l'âge ; « Résultat TB (Point Focal) » pour `tb_focal_point` (la source l'offre au menu, le README disait tâche seule) ; les 9 autres par leur tâche | identique |
| `summary.muted` dans `showForm` | `*.properties.json` | Pas d'équivalent | non reproductible |
| `next_visit_fr` : `if()` à 4 arguments (bug) | `moderate_malnutrition_followup.xlsx` | Corrigé comme `next_visit` (décision du 02/10/2026) | adapté |
| `=TRUE()` / `=FALSE()` (préfixe Excel) | `notification_to_chw`, `treatment_followup_over_5` | `=` retiré | identique |

## Tâches

19 règles (`metadata/tasks/`, générées par `.collecton/cht-port/tasks.mjs`), noms CHT `<source>_yields_<action>[_<sugar>]`.

| Élément CHT | Source | Solution Collecton | Statut |
|---|---|---|---|
| Références depuis les 2 évaluations (+1, +2, +3, priorité haute) | `tasks.js` l.594-678 | 2 règles ; résolution TB / close_out / referral / nouvelle évaluation / fenêtre | identique |
| `referralTemplate` depuis `treatment_followup(_over_5)`, `moderate_/severe_malnutrition_followup` | `tasks.js`, `tasks.extras.js` l.201 | 4 règles (absentes du tableau du README, présentes dans la source) | identique |
| `referralTemplate` depuis ANC, PNC, `simplecare_newborn_followup`, `tb_case_eval` | `tasks.js` | Non repris : formulaires hors périmètre | non reproductible |
| `selfSpawningTemplate` des 2 suivis de référence | `tasks.extras.js` l.363 | 2 règles ; comme CHT, la `resolvedIf` des options est ignorée (seuls les suivis postérieurs comptent) | identique |
| Traitements (paludisme, IRA, diarrhée, multiple, adulte) | `tasks.js`, `tasks.extras.js` l.237 | 5 règles, mêmes événements et fenêtres | identique |
| MAM / MAS sans complication (`appliesTo: contacts`, `this.patientAssessment`) | `tasks.extras.js` l.266 | Règles sur le patient ; l'évaluation de référence est recalculée dans chaque fonction ; suivis reliés par `inputs_source_id` | adapté |
| Résultat TB point focal / ASC | `tasks.extras.js` l.497-564 | 4 règles | identique |
| `user.cscom_area` du point focal | `tasks.extras.js` l.502 | CSCOM de la personne « Point Focal TB » de la zone de santé de l'utilisateur (celle au nom du compte s'il y en a plusieurs) | adapté |
| `modifyContent` (`t_*`, `source_id`) | `tasks.extras.js` | Fonction `prefill` ajoutée à Collecton : valeurs lues par le formulaire en `app.taskInputs`, stockées en `inputs_<nom>` | identique |
| `t_follow_up_type` lu dans `report.fields` par `referralTemplate` | `tasks.extras.js` l.217 | Repris tel quel (champ absent : valeur vide), comme CHT | identique |
| Date « dynamique » (`visited_date`, `s_reported.s_reported_date`, date du rapport) | `common.js` l.201 | Même règle (`dynamicDate`) ; `s_reported_date` n'existe pas dans ces formulaires | identique |
| `displayForUserAt` / `user.role` | `tasks.extras.js` | Rôle Collecton : `chw_uhc` (c40), `tb_focal_point` (c20) | adapté |
| Priorités `high` / `medium` | `tasks.extras.js` | Priorité Collecton 1 / 2 (décision du 02/10/2026) | identique |
| Étiquettes de priorité (« Suivi N - x/total », nom de l'ASC), `contactLabel` | `tasks.extras.js` | Non reprises | non reproductible |
| `muted` dans `appliesIf` / `resolvedIf` | `tasks.extras.js` | Pas d'équivalent | non reproductible |
| Tâches 8 (alerte erreur TDR) et 9 (TDR à refaire) | `tasks.js` | Non créées (décision du 01/10/2026) | non reproductible |

## Indicateurs (targets)

22 indicateurs (`metadata/indicators/`, générés par `.collecton/cht-port/indicators.mjs`), ids CHT conservés.

| Élément CHT | Source | Solution Collecton | Statut |
|---|---|---|---|
| Fonctions des erreurs de protocole (`isDangerSignNotReferred`... `aggravatedWithoutReferral`) | `targets.extras.js` | Copiées telles quelles, derrière un adaptateur (`getField` sur le payload à plat, dates façon Luxon) | identique |
| Chemins suspects (`malnutrition_tdr_not_done`, `child_temperature` en over_5, `s_accompany_refer_under5/over5`, `not_give_plumpy_nut`) et précédence de `aggravatedWithoutReferral` | `targets.extras.js` | Reproduits tels quels (décision du 01/10/2026) : ces champs n'existent pas, les tests valent faux comme dans CHT | identique |
| `percent-assessments-without-protocol-errors` : le titre dit « sans erreur » mais le calcul compte les évaluations avec au moins une erreur (objectif 0) | `targets.js` | Repris tel quel | identique |
| Pourcentage CHT (`passesIf` / instances, objectif 0 ou 100) | `targets.js` | Type `percentage` (`numerator` / `denominator`), `goal` en fonction | identique |
| Période : mois en cours (`month_start_date: 1`) | `app_settings.json` | Mois Muso du 1er au dernier jour sur l'horloge de l'app, comme la période de collecte ; date = date dynamique du rapport | identique |
| `context: user.role === "chw"` | `targets.extras.js` | Affiché pour le rôle `chw_uhc` | adapté |
| Patient vivant (`isAlive`) | `targets.extras.js` | Vérifié via `app.find('persons')` | identique |
| `muted` | `targets.extras.js` | Pas d'équivalent | non reproductible |
| `getNumberOfDaySinceDate` sur `new Date()` | `common.js` | Horloge de l'app (`app.now`, date choisie en collecte différée) | adapté |
| `home_visits` : `idType` famille~jour | `targets.js` | Paires (famille = `visited_contact_uuid`, jour) distinctes du mois | identique |
| `2-home-visits-per-fam` : `groupBy` famille, `passesIfGroupCount >= 2`, instance en échec par famille | `targets.js` | Familles avec au moins 2 jours de visite / toutes les familles (`c50_family`) de l'appareil ; famille = `visited_contact_uuid` plutôt que le parent du patient (même valeur) | adapté |
| Cibles manager, HPV, grossesse, `hoursFromIllnessToAssessment` | `targets.js`, `targets.extras.js` | Hors périmètre | non reproductible |
| Icônes CHT des cibles | `targets.js` | Icône par défaut | adapté |

## Purge (rétention)

| Élément CHT | Source | Solution Collecton | Statut |
|---|---|---|---|
| Fonction de purge (règles 0 à 14) | `purge.js` | Règles de purge ajoutées à Collecton : `utils/purge.js` exporte `shouldPurge(report, context)`, appelée sur l'appareil pour chaque soumission synchronisée ; mêmes listes et fenêtres | identique |
| Règle 4.1 : `patient_assessment` purgée à 30 jours sauf MAM / MAS sans complication | `purge.js` | Reprise telle quelle (`has_MAM`, `has_SAM_without_complication`) | identique |
| Règles 10 à 13 par rôle (`chw_uhc`, `tb_focal_point`, `stock_manager`, `chw_manager`) | `purge.js` | Rôle Collecton de l'utilisateur (mêmes noms) | identique |
| Exécution chaque samedi à 22 h (`cron`, `run_every_days: 7`) | `purge.js` | Après chaque synchronisation | adapté |
| `reported_date` | `purge.js` | Date de création sur l'appareil | identique |
| Une soumission non encore envoyée | CHT | Jamais purgée avant envoi | adapté |
| Durée par formulaire | (README étape 9) | `duration` = rétention maximale, qui borne aussi le premier téléchargement : `patient_assessment` 100, `patient_assessment_over_5` 30, suivis de référence et de traitement 10, `redo_tdr` 30, suivis MAM / MAS 100, `notification_to_chw`, `tb_test_result_fp`, `tb_test_result_chw` 300 | identique |
| `purge_tasks: [{ event_name: 'self_assessment_' }]` | `app_settings.json` | Rien à faire : les tâches Collecton ne sont pas stockées | non reproductible |
| La console garde tout | CHT (serveur) | Idem : la purge ne touche que l'appareil | identique |

## Décisions prises pendant la session

- 01/10/2026 : marge de la période perpétuelle = 5 jours.
- 01/10/2026 : tâches 8 (alerte erreur TDR) et 9 (TDR à refaire) non créées (pas d'équivalent du champ `verified`) ; leurs formulaires seront créés.
- 01/10/2026 : chemins suspects des targets reproduits à l'identique.
- 01/10/2026 : purge conditionnelle reproduite telle quelle ; Collecton étendu (script de purge par projet, à livrer).
- 01/10/2026 : champs supplémentaires des lieux et des personnes = data elements dans les définitions (attributs) des types.
- 02/10/2026 : filtres de choix ajoutés à Collecton ; géolocalisation = champ geopoint.
- 02/10/2026 : clé `s_cscom_areas` pour c30 ; numéro de ménage et champs « pm » toujours affichés.
- 02/10/2026 : 8 person types validés avec leurs champs ; sexe = champ natif M / F ; pas de patient à c20.
- 02/10/2026 : 5 rôles validés ; les rôles qui créent des personnes ou des lieux peuvent aussi les modifier.
- 02/10/2026 : minuteur et titres colorés ajoutés à Collecton ; plan « Evaluations et suivis des patients » (12 formulaires) confirmé ; `next_visit_fr` corrigé.
- 02/10/2026 : 19 tâches confirmées ; CSCOM du point focal = celui du « Point Focal TB » de sa zone ; priorités en fonctions (haute 1, moyenne 2) sans étiquette.
- 02/10/2026 : 22 indicateurs confirmés.
- 02/10/2026 : règles de purge par projet ajoutées à Collecton (`utils/purge.js`) ; `purge.js` CHT porté tel quel, durées par formulaire posées.
- 02/10/2026 : `is_active` retiré des personnes et des lieux.
