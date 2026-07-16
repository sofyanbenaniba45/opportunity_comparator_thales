# Comparateur d'Opportunités Commerciales

Application Streamlit locale pour comparer deux fichiers Excel d'opportunités entre deux trimestres.

**Auteur :** Sofyan BENANIBA

## Prérequis

- Python 3.10 ou supérieur
- pip

## Installation

```bash
cd opportunity-comparator
pip install -r requirements.txt
```

## Lancement

```bash
streamlit run app.py
```

L'application s'ouvre automatiquement dans votre navigateur à l'adresse `http://localhost:8501`.

## Utilisation

1. **Importer les fichiers** : déposez le fichier N-1 et le fichier actuel dans les zones d'upload.
2. **Comparer** : les KPIs et le tableau de détail s'affichent automatiquement.
3. **Filtrer** : utilisez les filtres par statut (Nouveau / Modifié / Supprimé / Inchangé) et par GBU/BL.
4. **Exporter** :
   - Cliquez **Générer le rapport Excel** puis téléchargez le `.xlsx`.
   - Cliquez **Générer le rapport Word** puis téléchargez le `.docx`.

## Format des fichiers

- Format : `.xlsx` ou `.xls`
- Les deux fichiers doivent avoir **exactement les mêmes colonnes**.
- La colonne **`Opportunity Name`** est obligatoire (clé de comparaison).
- La colonne contenant `GBU` ou `BL` dans son nom est détectée automatiquement pour le regroupement.

## Logique de comparaison

| Statut | Définition | Couleur |
|--------|-----------|---------|
| **Nouveau** | Présent dans actuel, absent dans N-1 | Jaune fluo |
| **Supprimé** | Présent dans N-1, absent dans actuel | Texte barré (app uniquement) |
| **Modifié** | Présent dans les deux, avec au moins une cellule différente | Cellule(s) changée(s) en orange |
| **Inchangé** | Identique dans les deux fichiers | Pas de mise en forme |

Les lignes **supprimées** sont visibles dans l'application mais **n'apparaissent pas** dans les exports Excel et Word.

## Structure du projet

```
opportunity-comparator/
├── app.py            # Interface Streamlit
├── comparison.py     # Logique de comparaison
├── excel_export.py   # Génération Excel
├── word_export.py    # Génération Word
├── requirements.txt  # Dépendances
└── README.md
```
