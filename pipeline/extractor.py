"""
pipeline/extractor.py — Extraction d'archive et détection des dossiers entreprises
"""

import os
import shutil
import logging
import zipfile
from pathlib import Path
from typing import Optional

from config import EXTENSIONS_ARCHIVES, DOSSIERS_IGNORES

logger = logging.getLogger(__name__)


def extraire_archive(chemin_archive: str, dossier_destination: str) -> str:
    """
    Extrait une archive .rar, .zip, .7z dans le dossier de destination.
    Retourne le chemin du dossier extrait.
    """
    chemin = Path(chemin_archive)
    dest = Path(dossier_destination)
    dest.mkdir(parents=True, exist_ok=True)

    extension = chemin.suffix.lower()
    logger.info(f"Extraction de {chemin.name} vers {dest}")

    if extension == ".zip":
        with zipfile.ZipFile(chemin, 'r') as zf:
            zf.extractall(dest)
        logger.info("Archive ZIP extraite avec succès")

    elif extension == ".rar":
        try:
            import rarfile
            with rarfile.RarFile(chemin) as rf:
                rf.extractall(dest)
            logger.info("Archive RAR extraite avec succès")
        except ImportError:
            logger.warning("rarfile non disponible, tentative avec patool")
            try:
                import patoollib
                patoollib.extract_archive(str(chemin), outdir=str(dest))
            except Exception as e:
                raise RuntimeError(f"Impossible d'extraire le RAR : {e}")

    elif extension in {".7z", ".tar", ".gz"}:
        try:
            import patoollib
            patoollib.extract_archive(str(chemin), outdir=str(dest))
        except Exception as e:
            raise RuntimeError(f"Impossible d'extraire l'archive : {e}")

    else:
        raise ValueError(f"Format d'archive non supporté : {extension}")

    # Retourner le dossier racine extrait
    contenu = list(dest.iterdir())
    if len(contenu) == 1 and contenu[0].is_dir():
        return str(contenu[0])
    return str(dest)


def detecter_entreprises(dossier_racine: str) -> list[dict]:
    """
    Parcourt l'arborescence et détecte les dossiers entreprises.
    Un dossier entreprise est un sous-dossier de premier niveau contenant
    un sous-dossier DPGF (ou équivalent).

    Retourne une liste de dicts :
    {
        "nom": str,
        "chemin": str,
        "chemin_dpgf": str | None
    }
    """
    racine = Path(dossier_racine)
    entreprises = []

    logger.info(f"Analyse de l'arborescence : {racine}")

    for entree in sorted(racine.iterdir()):
        if not entree.is_dir():
            continue
        if entree.name in DOSSIERS_IGNORES or entree.name.startswith("."):
            continue
        if entree.name == "_REGROUPEMENT":
            continue

        nom_entreprise = entree.name
        chemin_dpgf = _trouver_dossier_dpgf(entree)

        entreprises.append({
            "nom": nom_entreprise,
            "chemin": str(entree),
            "chemin_dpgf": str(chemin_dpgf) if chemin_dpgf else None
        })
        logger.info(f"  Entreprise détectée : {nom_entreprise} | DPGF : {chemin_dpgf}")

    logger.info(f"{len(entreprises)} entreprise(s) détectée(s)")
    return entreprises


def _trouver_dossier_dpgf(dossier_entreprise: Path) -> Optional[Path]:
    """
    Recherche récursivement le sous-dossier DPGF dans un dossier entreprise.
    Cherche les dossiers nommés 'DPGF', 'Offre', 'Offre de prix', etc.
    """
    noms_dpgf = {"dpgf", "offre", "offre de prix", "prix", "chiffrage", "bpu", "dqe"}

    for entree in dossier_entreprise.rglob("*"):
        if entree.is_dir() and entree.name.lower() in noms_dpgf:
            return entree

    # Si pas de sous-dossier DPGF, prendre le dossier racine entreprise
    # (cas où les fichiers sont directement dans le dossier)
    return dossier_entreprise


def valider_structure(entreprises: list[dict]) -> dict:
    """
    Vérifie la cohérence de la structure détectée.
    Retourne un rapport de validation.
    """
    rapport = {
        "total": len(entreprises),
        "avec_dpgf": 0,
        "sans_dpgf": [],
        "anomalies": []
    }

    for e in entreprises:
        if e["chemin_dpgf"]:
            rapport["avec_dpgf"] += 1
        else:
            rapport["sans_dpgf"].append(e["nom"])
            rapport["anomalies"].append(f"Aucun dossier DPGF trouvé pour {e['nom']}")

    return rapport
