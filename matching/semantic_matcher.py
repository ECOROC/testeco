"""
matching/semantic_matcher.py — Matching sémantique des postes DPGF via Claude API
"""

import json
import logging
import re
import time
from typing import Callable, Optional

import anthropic

from config import ANTHROPIC_API_KEY, CLAUDE_MODEL, CLAUDE_MAX_TOKENS, BATCH_SIZE
from matching.prompts import prompt_matching_batch
from matching.batch_manager import CacheMatching, decouper_en_batches

logger = logging.getLogger(__name__)


class SemanticMatcher:
    """
    Effectue le matching sémantique entre les postes de référence
    et les postes de chaque entreprise via l'API Claude.
    """

    def __init__(self, callback_progression: Optional[Callable] = None):
        if not ANTHROPIC_API_KEY:
            raise ValueError("Clé API Anthropic manquante (ANTHROPIC_API_KEY)")
        self.client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        self.cache = CacheMatching()
        self.callback = callback_progression  # fn(etape, total, message)

    def matcher_entreprise(
        self,
        postes_reference: list[dict],
        nom_entreprise: str,
        postes_entreprise: list[dict]
    ) -> list[dict]:
        """
        Mappe tous les postes de référence contre les postes d'une entreprise.

        Retourne une liste de résultats de matching :
        [
          {
            "ref_index": int,
            "poste_reference": dict,
            "indices_entreprise": [int, ...],
            "postes_entreprise": [dict, ...],
            "score_confiance": str,
            "type": str,
            "commentaire": str | None,
            "prix_unitaire_retenu": float | None,
            "prix_total_retenu": float | None
          }
        ]
        """
        logger.info(f"Matching {nom_entreprise} : {len(postes_reference)} postes de référence")

        resultats = []
        batches = decouper_en_batches(postes_reference, BATCH_SIZE)
        total_batches = len(batches)

        for i, batch in enumerate(batches):
            if self.callback:
                self.callback(i, total_batches, f"Batch {i+1}/{total_batches} — {nom_entreprise}")

            # Construire le prompt
            prompt = prompt_matching_batch(batch, nom_entreprise, postes_entreprise)

            # Vérifier le cache
            en_cache = self.cache.get(prompt)
            if en_cache:
                logger.debug(f"  Batch {i+1} : résultat depuis le cache")
                matchings_batch = en_cache
            else:
                matchings_batch = self._appeler_api(prompt, nom_entreprise, i+1)
                if matchings_batch:
                    self.cache.set(prompt, matchings_batch)

            if not matchings_batch:
                # Fallback : tout marquer ABSENT
                matchings_batch = [
                    {
                        "ref_index": j,
                        "indices_entreprise": [],
                        "score_confiance": "ABSENT",
                        "type": "absent",
                        "commentaire": "Erreur API — statut non déterminé"
                    }
                    for j in range(len(batch))
                ]

            # Reconstruire les résultats enrichis
            for matching in matchings_batch:
                ref_idx_local = matching.get("ref_index", 0)
                if ref_idx_local >= len(batch):
                    continue

                poste_ref = batch[ref_idx_local]
                indices_ent = matching.get("indices_entreprise", [])
                postes_ent_matchés = [
                    postes_entreprise[idx]
                    for idx in indices_ent
                    if idx < len(postes_entreprise)
                ]

                prix_pu, prix_total = self._consolider_prix(postes_ent_matchés, matching)

                resultats.append({
                    "ref_index": postes_reference.index(poste_ref) if poste_ref in postes_reference else -1,
                    "poste_reference": poste_ref,
                    "indices_entreprise": indices_ent,
                    "postes_entreprise": postes_ent_matchés,
                    "score_confiance": matching.get("score_confiance", "ABSENT"),
                    "type": matching.get("type", "absent"),
                    "commentaire": matching.get("commentaire"),
                    "prix_unitaire_retenu": prix_pu,
                    "prix_total_retenu": prix_total
                })

            # Délai entre les appels pour respecter les rate limits
            if i < total_batches - 1:
                time.sleep(0.5)

        logger.info(f"Matching {nom_entreprise} terminé : {len(resultats)} résultats")
        return resultats

    def _appeler_api(self, prompt: str, nom_entreprise: str, num_batch: int) -> list | None:
        """Appel à l'API Claude avec gestion des erreurs et retry."""
        max_tentatives = 3
        for tentative in range(max_tentatives):
            try:
                response = self.client.messages.create(
                    model=CLAUDE_MODEL,
                    max_tokens=CLAUDE_MAX_TOKENS,
                    messages=[{"role": "user", "content": prompt}]
                )

                texte = response.content[0].text.strip()
                texte = re.sub(r"```json|```", "", texte).strip()
                resultat = json.loads(texte)

                if isinstance(resultat, list):
                    return resultat
                logger.warning(f"Réponse API inattendue (non-liste) : {type(resultat)}")
                return None

            except json.JSONDecodeError as e:
                logger.error(f"JSON invalide — {nom_entreprise} batch {num_batch} tentative {tentative+1} : {e}")
                if tentative < max_tentatives - 1:
                    time.sleep(2)

            except anthropic.RateLimitError:
                logger.warning(f"Rate limit atteint, attente 60s...")
                time.sleep(60)

            except anthropic.APIError as e:
                logger.error(f"Erreur API Claude : {e}")
                if tentative < max_tentatives - 1:
                    time.sleep(5)

        return None

    def _consolider_prix(self, postes_matchés: list[dict], matching: dict) -> tuple:
        """
        Détermine le prix unitaire et total à retenir pour un matching.
        En cas de regroupement, somme les totaux.
        """
        if not postes_matchés:
            return None, None

        type_match = matching.get("type", "exact")

        if type_match == "regroupe":
            # Sommer les totaux si regroupement
            total = sum(
                p.get("prix_total") or 0
                for p in postes_matchés
                if p.get("prix_total") is not None
            )
            return None, total if total > 0 else None
        else:
            # Cas normal : prendre le premier poste matché
            p = postes_matchés[0]
            return p.get("prix_unitaire"), p.get("prix_total")


def matcher_toutes_entreprises(
    postes_reference: list[dict],
    entreprises_postes: dict[str, list[dict]],
    callback: Optional[Callable] = None
) -> dict[str, list[dict]]:
    """
    Lance le matching pour toutes les entreprises.

    entreprises_postes : {"Dupont": [...postes...], "Martin": [...postes...]}

    Retourne : {"Dupont": [...matchings...], "Martin": [...matchings...]}
    """
    matcher = SemanticMatcher(callback_progression=callback)
    resultats = {}

    for nom_entreprise, postes in entreprises_postes.items():
        logger.info(f"\n{'='*50}")
        logger.info(f"Démarrage matching : {nom_entreprise}")
        resultats[nom_entreprise] = matcher.matcher_entreprise(
            postes_reference, nom_entreprise, postes
        )

    return resultats
