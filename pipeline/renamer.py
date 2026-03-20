"""
pipeline/renamer.py — Renommage des DPGF et création du dossier _REGROUPEMENT
"""

import logging
import shutil
from pathlib import Path

logger = logging.getLogger(__name__)


def creer_regroupement(dossier_racine: str, entreprises_avec_fichiers: list[dict]) -> dict:
    """
    Crée le dossier _REGROUPEMENT et y copie tous les DPGF renommés.

    entreprises_avec_fichiers : liste de dicts enrichis par file_detector :
    {
        "nom": str,
        "fichier_dpgf": Path,
        "type": str
    }

    Retourne un rapport de regroupement.
    """
    racine = Path(dossier_racine)
    dossier_regroupement = racine / "_REGROUPEMENT"
    dossier_regroupement.mkdir(exist_ok=True)

    logger.info(f"Création du dossier _REGROUPEMENT dans {racine}")

    rapport = {
        "dossier": str(dossier_regroupement),
        "fichiers_crees": [],
        "erreurs": []
    }

    for e in entreprises_avec_fichiers:
        if not e.get("fichier_dpgf"):
            rapport["erreurs"].append(f"{e['nom']} : aucun fichier DPGF à copier")
            continue

        fichier_src = Path(e["fichier_dpgf"])
        nom_propre = _nettoyer_nom_entreprise(e["nom"])
        extension = fichier_src.suffix.lower()
        nom_dest = f"DPGF_{nom_propre}{extension}"
        fichier_dest = dossier_regroupement / nom_dest

        try:
            shutil.copy2(fichier_src, fichier_dest)
            rapport["fichiers_crees"].append({
                "entreprise": e["nom"],
                "source": str(fichier_src),
                "destination": str(fichier_dest),
                "nom": nom_dest
            })
            logger.info(f"  Copié : {fichier_src.name} → {nom_dest}")

        except Exception as ex:
            msg = f"{e['nom']} : erreur lors de la copie — {ex}"
            rapport["erreurs"].append(msg)
            logger.error(f"  ERREUR : {msg}")

    logger.info(f"Regroupement terminé : {len(rapport['fichiers_crees'])} fichier(s) créé(s)")
    return rapport


def _nettoyer_nom_entreprise(nom: str) -> str:
    """
    Nettoie le nom d'une entreprise pour créer un nom de fichier valide.
    Exemple : "Entreprise_Dupont BTP & Fils" → "Dupont_BTP_Fils"
    """
    import re

    # Supprimer les préfixes génériques courants
    prefixes = [
        r"^entreprise[_\s]*",
        r"^groupe[_\s]*",
        r"^ste[_\s]*",
        r"^sarl[_\s]*",
        r"^sas[_\s]*",
        r"^sa[_\s]*",
        r"^eurl[_\s]*",
    ]
    nom_propre = nom
    for prefixe in prefixes:
        nom_propre = re.sub(prefixe, "", nom_propre, flags=re.IGNORECASE).strip()

    # Remplacer les caractères non alphanumériques par _
    nom_propre = re.sub(r"[^\w]", "_", nom_propre)
    # Supprimer les _ multiples
    nom_propre = re.sub(r"_+", "_", nom_propre)
    # Supprimer _ en début/fin
    nom_propre = nom_propre.strip("_")

    return nom_propre or nom
