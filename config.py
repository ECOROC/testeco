"""
config.py — Configuration globale de l'outil DPGF Comparator
"""

import os
from dotenv import load_dotenv

load_dotenv()

# ─── API Claude ───────────────────────────────────────────────────────────────
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL = "claude-sonnet-4-20250514"
CLAUDE_MAX_TOKENS = 2000

# ─── Paramètres de matching ───────────────────────────────────────────────────
BATCH_SIZE = 20                  # Nombre de postes par appel API
SEUIL_ALERTE_ECART = 0.20        # 20% d'écart → alerte rouge
SEUIL_CONFIANCE_MOYEN = "MOYEN"  # Score → orange

# ─── Paramètres OCR ───────────────────────────────────────────────────────────
OCR_SCORE_MIN = 60               # Score de confiance OCR minimum (%)
TESSERACT_LANG = "fra"           # Langue Tesseract

# ─── Extensions de fichiers ───────────────────────────────────────────────────
EXTENSIONS_EXCEL = {".xlsx", ".xls", ".xlsm"}
EXTENSIONS_PDF = {".pdf"}
EXTENSIONS_ARCHIVES = {".rar", ".zip", ".7z", ".tar", ".gz"}

# ─── Dossiers à ignorer lors du parcours ─────────────────────────────────────
DOSSIERS_IGNORES = {
    "_REGROUPEMENT", "__pycache__", ".git", "node_modules",
    "Candidature", "Memoire_Technique", "Mémoire_Technique",
    "memoire_technique", "candidature"
}

# ─── Mots-clés pour identifier un DPGF ───────────────────────────────────────
MOTS_CLES_DPGF = [
    "prix unitaire", "prix total", "quantite", "quantité",
    "unité", "unite", "libellé", "libelle", "montant",
    "désignation", "designation", "lot", "poste", "total ht",
    "p.u.", "montant ht", "qté", "prix en", "total en",
    "descriptif", "ouvrages", "fournitures"
]

# ─── Noms de colonnes typiques DPGF ──────────────────────────────────────────
COLONNES_NUMERO = [
    "n°", "num", "numéro", "repère", "art", "article", "ref", "poste", "rep"
]
COLONNES_LIBELLE = [
    "libellé", "libelle", "désignation", "designation", "description",
    "nature", "descriptif", "ouvrages", "désignation des ouvrages",
    "fournitures", "travaux"
]
COLONNES_UNITE = ["unité", "unite", "u", "unit"]
COLONNES_QUANTITE = [
    "quantité", "quantite", "qté", "qte", "qt",
    "quantité lacorps", "quantité entreprises", "quantite entreprise",
    "quantité entreprise", "qte entreprise", "q"
]
COLONNES_PU = [
    "prix unitaire", "p.u.", "pu", "prix u.", "px unitaire", "tarif",
    "prix en €", "prix en eur", "prix unit", "pu ht", "p.u ht",
    "prix €", "prix ht"
]
COLONNES_TOTAL = [
    "prix total", "montant", "total ht", "total", "montant ht", "pt",
    "total en €", "total en eur", "montant €", "total €",
    "montant total", "total ttc"
]

# ─── Logging ─────────────────────────────────────────────────────────────────
LOG_FILE = "dpgf_comparator.log"
LOG_LEVEL = "INFO"

# ─── Cache ────────────────────────────────────────────────────────────────────
CACHE_DIR = ".cache_matching"
