# DPGF Comparator

Outil d'analyse comparative automatisée des DPGF dans le cadre des appels d'offres.  
Développé pour les économistes de la construction.

---

## Ce que ça fait

Vous recevez une archive avec les dossiers entreprises. Vous lancez l'outil.  
Vous récupérez un Excel avec les prix de toutes les entreprises côte à côte, les alertes d'écart, les postes manquants.

**Pipeline en 5 étapes :**

1. **Extraction & détection** — Parcourt l'arborescence, identifie les dossiers entreprises, trouve les DPGF (Excel prioritaire sur PDF), renomme proprement tout ça dans `_REGROUPEMENT/`
2. **Parsing** — Excel (openpyxl), PDF natif (pdfplumber), PDF scanné (Tesseract OCR + fallback Claude Vision)
3. **Référence** — Charge votre trame vierge, ou reconstruit une trame depuis les offres si vous n'en avez pas
4. **Matching sémantique** — Claude API (batch de 20 postes) pour faire correspondre les libellés différents. Fallback matching direct par similarité textuelle si pas de clé API
5. **Export Excel** — Multi-feuilles, 1 feuille par lot, prix entreprises en colonnes, statistiques, alertes rouge > 20%, postes absents grisés, synthèse globale

---

## Installation

```bash
# Cloner ou extraire le projet
cd dpgf-comparator

# Installer les dépendances
pip install -r requirements.txt

# Configurer la clé API (requis pour le matching sémantique)
cp .env.example .env
# éditer .env et renseigner ANTHROPIC_API_KEY
```

**Prérequis système pour l'OCR :**
```bash
# Ubuntu/Debian
sudo apt-get install tesseract-ocr tesseract-ocr-fra poppler-utils

# macOS
brew install tesseract tesseract-lang poppler
```

---

## Utilisation

### Interface Streamlit (recommandé)

```bash
streamlit run app.py
```

Ouvre une interface dans le navigateur avec drag & drop, progression temps réel, prévisualisation et téléchargement.

### Ligne de commande

```bash
# Pipeline complet
python main.py \
  --entree /chemin/AO_NomOperation \
  --reference DPGF_Reference_vierge.xlsx \
  --sortie Comparatif_DPGF.xlsx

# Sans référence (reconstruction depuis les offres)
python main.py --entree /chemin/AO_NomOperation

# Tester uniquement l'étape 1 (extraction + détection)
python main.py --entree /chemin/AO_NomOperation --etape 1

# Jusqu'au parsing uniquement
python main.py --entree /chemin/AO_NomOperation --etape 2
```

**Options :**

| Option | Description | Défaut |
|--------|-------------|--------|
| `--entree` | Dossier AO ou archive .zip/.rar | — (obligatoire) |
| `--reference` | DPGF de référence vierge (Excel) | Mode reconstruction |
| `--sortie` | Chemin fichier Excel de sortie | `_dossier_AO_/Comparatif_DPGF.xlsx` |
| `--etape` | Arrêter après l'étape N (1 à 5) | 5 (pipeline complet) |

---

## Structure attendue du dossier AO

```
AO_NomOperation/
├── Entreprise_Dupont/
│   ├── Candidature/          ← ignoré
│   ├── Memoire_Technique/    ← ignoré
│   └── DPGF/
│       ├── 48291047.xlsx     ← détecté par contenu, pas par nom
│       └── scan_dpgf.pdf     ← ignoré si Excel présent
├── Entreprise_Martin/
│   └── DPGF/
│       └── offre_finale.pdf  ← PDF natif ou scanné selon contenu
└── Entreprise_Vinci/
    └── DPGF/
        └── reponse.xlsx
```

**Règles de détection :**
- Le nom de l'entreprise = nom du dossier parent (jamais le nom du fichier)
- Excel toujours prioritaire sur PDF
- Le DPGF est identifié par son contenu (colonnes prix/quantités/libellés), pas par son nom

---

## Fichier Excel de sortie

- **Feuille SYNTHÈSE** (première) : totaux par lot et par entreprise, min/max, écart %
- **1 feuille par lot** : tous les postes avec les prix de chaque entreprise en colonnes
- **Colonnes** : N° | Désignation | Unité | Quantité | Réf. PU | [Entreprises...] | Min | Max | Moyenne | Écart-type | ALERTE
- **Codes couleur** :
  - 🔴 Rouge : écart > 20% de la moyenne
  - 🟠 Orange : score de confiance MOYEN (à vérifier)
  - ⬜ Grisé : poste ABSENT dans l'offre
  - 🟡 Jaune pâle : poste REGROUPÉ avec un autre

---

## Tests

```bash
# Générer les données de test (3 entreprises GO, 20 postes)
python tests/generer_donnees_test.py

# Lancer tous les tests unitaires
python tests/test_pipeline.py
```

Les données de test génèrent 3 entreprises avec des cas réalistes :
- **Dupont** : postes standards, libellés proches de la référence
- **Martin** : libellés reformulés, 2 postes manquants, 1 regroupement
- **Vinci** : ordre des postes mélangé, 1 prix avec écart > 30%

---

## Configuration avancée

Tous les paramètres sont dans `config.py` :

```python
SEUIL_ALERTE_ECART = 0.20    # Seuil d'alerte rouge (20%)
BATCH_SIZE = 20               # Postes par appel API Claude
OCR_SCORE_MIN = 60            # Score confiance OCR minimum
TESSERACT_LANG = "fra"        # Langue OCR
```

---

## Variables d'environnement

Créer un fichier `.env` à la racine :

```env
ANTHROPIC_API_KEY=sk-ant-api03-...
```

Sans clé API, l'outil fonctionne en mode **matching direct** (similarité textuelle).  
Le matching sémantique Claude est nettement plus précis, surtout quand les libellés sont très différents.

---

## Architecture

```
dpgf-comparator/
├── main.py                    # CLI
├── config.py                  # Paramètres
├── app.py                     # Interface Streamlit
├── pipeline/
│   ├── extractor.py           # Extraction archive + détection entreprises
│   ├── file_detector.py       # Identification DPGF par contenu
│   ├── renamer.py             # _REGROUPEMENT + renommage
│   └── report.py              # Rapport de préparation JSON + console
├── parsers/
│   ├── excel_parser.py        # Parser Excel (en-têtes décalées, multi-feuilles)
│   ├── pdf_parser.py          # Parser PDF natif (pdfplumber)
│   └── ocr_parser.py          # OCR Tesseract + fallback Claude Vision
├── matching/
│   ├── semantic_matcher.py    # Matching Claude API (batch + cache + retry)
│   ├── direct_matcher.py      # Matching direct sans API (fallback)
│   ├── batch_manager.py       # Cache disque + estimation coût
│   └── prompts.py             # Prompts centralisés
├── output/
│   └── excel_exporter.py      # Excel multi-feuilles, couleurs, alertes
└── tests/
    ├── generer_donnees_test.py # Génération données de test réalistes
    └── test_pipeline.py       # Tests unitaires par étape
```
