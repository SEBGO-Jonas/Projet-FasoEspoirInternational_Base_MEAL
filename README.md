# Dashboard MEAL — Faso Espoir International

Rapport Power BI multipage destiné à la direction de **Faso Espoir International (FEI)**, ONG humanitaire internationale intervenant au Burkina Faso, pour le suivi programmatique, financier, opérationnel et de redevabilité (**MEAL** — Monitoring, Évaluation, Redevabilité et Apprentissage).

> Projet Power BI au format PBIP (fichiers texte TMDL/PBIR), conçu pour être versionné avec Git.

## Sommaire

- [Aperçu](#aperçu)
- [Prérequis](#prérequis)
- [Démarrage rapide](#démarrage-rapide)
- [Structure du dépôt](#structure-du-dépôt)
- [Source de données](#source-de-données)
- [Modèle sémantique](#modèle-sémantique)
- [Rapport — 6 pages](#rapport--6-pages)
- [Limites connues](#limites-connues)
- [Maintenance](#maintenance)
- [Confidentialité des données](#confidentialité-des-données)

## Aperçu

| | |
|---|---|
| **Organisation** | Faso Espoir International (FEI) |
| **Pays d'intervention** | Burkina Faso |
| **Période couverte** | 2021 – 2025 |
| **Devise de référence** | USD |
| **Pages du rapport** | 6 (Accueil, Performance Programmatique, Suivi Financier, Opérations & Approvisionnement, Redevabilité & Qualité, Annexe) |
| **Mesures DAX** | 31, réparties en 4 dossiers (Programme, Finance, Opérations, Redevabilité) |

## Prérequis

- [Power BI Desktop](https://www.microsoft.com/store/productId/9NTXR16HNW1T) (version récente — le projet utilise le format PBIP/PBIR).
- Accès en lecture au fichier `FasoEspoirInternational_Base_MEAL.xlsx` (source de données).

## Démarrage rapide

```bash
git clone <url-du-depot>
```

Ouvrir uniquement le fichier **`Base_MEAL.pbip`** avec Power BI Desktop — il charge automatiquement le modèle sémantique et le rapport associés.

## Structure du dépôt

```text
Base_MEAL.pbip                            # fichier à ouvrir dans Power BI Desktop
Base_MEAL.SemanticModel/                  # modèle sémantique (format TMDL)
Base_MEAL.Report/                         # rapport (format PBIR)
FasoEspoirInternational_Base_MEAL.xlsx    # source de données du modèle
CongoTopFashion_Base_Donnees.xlsx         # non utilisé par ce modèle (jeu de données distinct)
README.md
```

## Source de données

Toutes les tables sont importées depuis `FasoEspoirInternational_Base_MEAL.xlsx` (25 feuilles). Le chemin du classeur est codé en dur dans les requêtes Power Query de chaque table — si le fichier est déplacé ou cloné dans un autre dossier, mettre à jour le chemin via **Transformer les données > Paramètres de la source de données** dans Power BI Desktop.

## Modèle sémantique

### Corrections apportées au modèle d'origine

- **En-têtes de colonnes réparés** sur les tables `Bailleurs`, `Dictionnaire_Donnees` et `Relations` (elles chargeaient des colonnes génériques `Column1..N`).
- **Colonnes de dates** (~30 colonnes sur 14 tables) retypées en véritables dates (elles étaient importées comme nombres entiers).
- **Table dupliquée `Table 1 (Beneficiaires)`** supprimée.
- **Table `Calendrier`** ajoutée (2021-2026) pour permettre l'intelligence temporelle (comparaisons annuelles, cumuls). Reliée à `Activites` et `Approvisionnements` ; `Decaissements`, `Livraisons` et `Plaintes` en héritent via leur relation à `Activites`.
- **Colonne calculée `Nom_Cible`** sur `Objectifs_Cadre_Logique` pour afficher un libellé lisible (nom de base ou de secteur) au lieu d'un code.
- **Relations ajoutées** : Interventions → Bailleurs, Interventions → Sous_Secteurs, Campagnes_Sensibilisation → Bases (et Activites → Personnel, inactive, activée via `USERELATIONSHIP` dans la mesure *Activités par Agent*).
- **Confidentialité** : les champs nominatifs et de contact des bénéficiaires (Nom, Prénom, Téléphone, Commentaires) ainsi que toutes les clés techniques `ID_*` sont masqués du volet de champs.

### Table `_Mesures`

31 mesures DAX regroupées en 4 dossiers d'affichage :

| Dossier | Exemples de mesures |
|---|---|
| **Programme** | Bénéficiaires Actifs, Taux de Réalisation Cadre Logique, Score de Vulnérabilité Moyen |
| **Finance** | Total Décaissé (USD), Taux de Décaissement, Coût par Bénéficiaire Actif (USD) |
| **Opérations** | Total Activités, Délai Moyen de Livraison, Bases en Rupture de Stock |
| **Redevabilité** | Nombre de Plaintes, Taux de Résolution des Plaintes, Score de Performance Moyen des Partenaires |

## Rapport — 6 pages

1. **Accueil** — synthèse direction : cartes KPI, tendance annuelle des décaissements, bénéficiaires actifs par région, cadre logique.
2. **Performance Programmatique** — profil des bénéficiaires (âge, vulnérabilité), détail cadre logique par cible et par année.
3. **Suivi Financier** — décaissements mensuels, valeur nette par secteur, coûts unitaires par intervention.
4. **Opérations & Approvisionnement** — activités par statut, délais de livraison, niveaux de stock.
5. **Redevabilité & Qualité** — plaintes par type/statut, performance des partenaires de mise en œuvre.
6. **Annexe** — paramètres organisationnels, note de confidentialité, dictionnaire de données.

Chaque page (sauf l'Annexe) porte 3 filtres synchronisés : **Année**, **Région**, **Secteur**. La navigation se fait via les onglets de page standard en bas de l'écran.

Thème visuel : `FEI_Humanitarian` (palette accessible daltonisme, ton institutionnel/humanitaire), défini dans `Base_MEAL.Report/StaticResources/SharedResources/BaseThemes/FEI_Humanitarian.json`.

## Limites connues

- Le filtre **Secteur** ne s'applique qu'aux visuels liés aux Interventions (pages Opérations/Finance) ; il ne filtre pas Bénéficiaires/Activités/Décaissements car Secteurs n'est pas connecté à ce cluster du modèle sans introduire un cycle de relations.
- Pas de bouton de navigation personnalisé — la navigation utilise les onglets de page natifs de Power BI.

## Maintenance

- **Actualiser les données** : bouton *Actualiser* dans Power BI Desktop (relit le classeur Excel source).
- **Ajouter une mesure** : l'ajouter dans la table `_Mesures` avec un `displayFolder` cohérent.
- **Étendre la période couverte** : modifier les bornes `CALENDAR(DATE(2021,1,1), DATE(2026,12,31))` dans la table `Calendrier`.

## Confidentialité des données

Les champs nominatifs et de contact des bénéficiaires (Nom, Prénom, Téléphone, Commentaires) sont masqués dans le modèle et ne sont utilisés dans aucun visuel. Seules les catégories agrégées (âge, sexe, région, vulnérabilité) sont exposées. Le classeur Excel source contient toutefois les données brutes : le traiter avec la même confidentialité que toute donnée sensible concernant des bénéficiaires humanitaires avant de le committer dans un dépôt, surtout si celui-ci est public.
