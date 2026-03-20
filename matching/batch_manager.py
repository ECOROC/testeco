"""
matching/batch_manager.py — Gestion des batches d'appels API et cache disque
"""

import hashlib
import json
import logging
import os
from pathlib import Path

from config import BATCH_SIZE, CACHE_DIR

logger = logging.getLogger(__name__)


class CacheMatching:
    """Cache disque pour éviter de refaire les mêmes appels API."""

    def __init__(self, dossier_cache: str = CACHE_DIR):
        self.dossier = Path(dossier_cache)
        self.dossier.mkdir(exist_ok=True)

    def _cle_cache(self, prompt: str) -> str:
        """Génère une clé de cache depuis un prompt."""
        return hashlib.md5(prompt.encode("utf-8")).hexdigest()

    def get(self, prompt: str) -> dict | None:
        """Récupère un résultat depuis le cache."""
        cle = self._cle_cache(prompt)
        fichier = self.dossier / f"{cle}.json"
        if fichier.exists():
            try:
                with open(fichier, encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return None
        return None

    def set(self, prompt: str, resultat: dict | list) -> None:
        """Enregistre un résultat dans le cache."""
        cle = self._cle_cache(prompt)
        fichier = self.dossier / f"{cle}.json"
        try:
            with open(fichier, "w", encoding="utf-8") as f:
                json.dump(resultat, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.warning(f"Cache : impossible d'écrire {fichier.name} : {e}")

    def vider(self) -> None:
        """Vide le cache."""
        for f in self.dossier.glob("*.json"):
            f.unlink()
        logger.info("Cache vidé")


def decouper_en_batches(elements: list, taille: int = BATCH_SIZE) -> list[list]:
    """Découpe une liste en sous-listes de taille fixe."""
    return [elements[i:i + taille] for i in range(0, len(elements), taille)]


def estimer_cout_api(
    nb_postes_reference: int,
    nb_entreprises: int,
    tokens_par_appel: int = 1500
) -> dict:
    """
    Estime le coût API Claude pour le matching.
    Basé sur les tarifs Claude Sonnet 4 (approximatifs).
    """
    nb_batches = (nb_postes_reference + BATCH_SIZE - 1) // BATCH_SIZE
    nb_appels = nb_batches * nb_entreprises

    # Tarifs approximatifs Claude Sonnet (input + output)
    # ~3$ / 1M tokens input, ~15$ / 1M tokens output
    tokens_input = nb_appels * tokens_par_appel
    tokens_output = nb_appels * 500  # ~500 tokens de réponse JSON

    cout_input = (tokens_input / 1_000_000) * 3.0
    cout_output = (tokens_output / 1_000_000) * 15.0
    cout_total = cout_input + cout_output

    return {
        "nb_appels": nb_appels,
        "nb_batches_par_entreprise": nb_batches,
        "tokens_estimes": tokens_input + tokens_output,
        "cout_estime_eur": round(cout_total * 0.92, 4),  # Conversion USD → EUR approx
        "cout_estime_cts": round(cout_total * 0.92 * 100, 2)
    }
