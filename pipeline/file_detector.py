"""
pipeline/file_detector.py — Identification du fichier DPGF dans un dossier entreprise
"""

import logging
from pathlib import Path
from typing import Optional
import re

from config import EXTENSIONS_EXCEL, EXTENSIONS_PDF, MOTS_CLES_DPGF

logger = logging.getLogger(__name__)


def detecter_fichier_dpgf(chemin_dossier: str) -> dict:
    """
    Analyse un dossier et identifie le fichier DPGF.
    Priorité : Excel > PDF

    Retourne :
    {
        "fichier": Path | None,
        "type": "excel" | "pdf_natif" | "pdf_scan" | None,
        "confiance": float,  # 0.0 à 1.0
        "message": str
    }
    """
    dossier = Path(chemin_dossier)
    if not dossier.exists():
        return {"fichier": None, "type": None, "confiance": 0.0, "message": "Dossier introuvable"}

    # Collecter tous les fichiers
    fichiers_excel = []
    fichiers_pdf = []

    for f in dossier.rglob("*"):
        if not f.is_file():
            continue
        ext = f.suffix.lower()
        if ext in EXTENSIONS_EXCEL:
            fichiers_excel.append(f)
        elif ext in EXTENSIONS_PDF:
            fichiers_pdf.append(f)

    logger.debug(f"  {len(fichiers_excel)} Excel, {len(fichiers_pdf)} PDF dans {dossier.name}")

    # Priorité 1 : Excel
    if fichiers_excel:
        meilleur = _choisir_meilleur_excel(fichiers_excel)
        if meilleur:
            return {
                "fichier": meilleur,
                "type": "excel",
                "confiance": _scorer_fichier_excel(meilleur),
                "message": f"Excel identifié : {meilleur.name}"
            }

    # Priorité 2 : PDF
    if fichiers_pdf:
        meilleur = _choisir_meilleur_pdf(fichiers_pdf)
        if meilleur:
            type_pdf = _detecter_type_pdf(meilleur)
            return {
                "fichier": meilleur,
                "type": type_pdf,
                "confiance": 0.7,
                "message": f"PDF identifié ({type_pdf}) : {meilleur.name}"
            }

    return {
        "fichier": None,
        "type": None,
        "confiance": 0.0,
        "message": "Aucun fichier DPGF détecté"
    }


def _choisir_meilleur_excel(fichiers: list[Path]) -> Optional[Path]:
    """
    Parmi plusieurs Excel, choisit le plus probable d'être un DPGF.
    Score basé sur le nom du fichier et le contenu (si possible).
    """
    if len(fichiers) == 1:
        return fichiers[0]

    scores = []
    for f in fichiers:
        score = _scorer_nom_fichier(f.name)
        scores.append((score, f))

    scores.sort(reverse=True)
    return scores[0][1]


def _scorer_nom_fichier(nom: str) -> float:
    """Score un nom de fichier selon les mots-clés DPGF."""
    nom_lower = nom.lower()
    score = 0.0

    mots_positifs = ["dpgf", "prix", "offre", "bpu", "dqe", "chiffrage", "devis", "reponse", "réponse"]
    mots_negatifs = ["candidature", "memoire", "mémoire", "administratif", "dc1", "dc2", "acte"]

    for mot in mots_positifs:
        if mot in nom_lower:
            score += 0.3

    for mot in mots_negatifs:
        if mot in nom_lower:
            score -= 0.5

    return score


def _scorer_fichier_excel(chemin: Path) -> float:
    """
    Tente d'ouvrir l'Excel et vérifie si le contenu ressemble à un DPGF.
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(chemin, read_only=True, data_only=True)
        score = 0.5  # Score de base

        for ws in wb.worksheets:
            contenu_feuille = ""
            lignes_lues = 0
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    if cell and isinstance(cell, str):
                        contenu_feuille += cell.lower() + " "
                lignes_lues += 1
                if lignes_lues > 30:
                    break

            mots_trouves = sum(1 for mot in MOTS_CLES_DPGF if mot in contenu_feuille)
            if mots_trouves >= 3:
                score = min(1.0, score + mots_trouves * 0.1)
                break

        wb.close()
        return score

    except Exception as e:
        logger.debug(f"Impossible d'analyser {chemin.name} : {e}")
        return 0.6  # Score par défaut si on ne peut pas lire


def _choisir_meilleur_pdf(fichiers: list[Path]) -> Optional[Path]:
    """Sélectionne le PDF le plus probable d'être un DPGF."""
    if len(fichiers) == 1:
        return fichiers[0]

    scores = []
    for f in fichiers:
        score = _scorer_nom_fichier(f.name)
        scores.append((score, f))

    scores.sort(reverse=True)
    return scores[0][1]


def _detecter_type_pdf(chemin: Path) -> str:
    """
    Détermine si un PDF est natif (texte extractible) ou scanné.
    Retourne "pdf_natif" ou "pdf_scan".
    """
    try:
        import pdfplumber
        with pdfplumber.open(chemin) as pdf:
            if not pdf.pages:
                return "pdf_scan"

            # Tester les 2 premières pages
            texte_total = ""
            for page in pdf.pages[:2]:
                texte = page.extract_text()
                if texte:
                    texte_total += texte

            # Si on a plus de 100 caractères, c'est un PDF natif
            if len(texte_total.strip()) > 100:
                return "pdf_natif"
            return "pdf_scan"

    except Exception:
        return "pdf_scan"
