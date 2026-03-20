"""
matching/direct_matcher.py — Matching direct sans API Claude
Utilisé quand pas de clé API disponible, ou pour les tests.

Stratégie : comparaison par similarité de texte (Levenshtein simplifié
+ correspondance de mots-clés). Moins précis que le matching sémantique
Claude, mais fonctionnel pour valider le pipeline.
"""

import re
import logging
from typing import Optional

logger = logging.getLogger(__name__)


SYNONYMES_GO = [
    # Chaque groupe = libellés équivalents
    {"terrassement", "fouille", "tranchée", "rigole", "puits", "excavation"},
    {"semelle", "fondation", "filante", "isolée", "continue", "ponctuelle"},
    {"voile", "mur", "refend", "vertical", "paroi"},
    {"béton", "beton", "b25", "b30", "b35", "dosé", "dose"},
    {"plancher", "dalle", "hourdis", "entrevous"},
    {"remblai", "remblaiment", "compactage", "compacté"},
    {"maçonnerie", "maconnerie", "bloc", "brique", "agglo"},
    {"enduit", "crépi", "façade", "monocouche", "gratté"},
    {"chape", "mortier", "sol", "flottante"},
    {"installation", "chantier", "repliement", "dépose"},
    {"longrine", "chevêtre", "poutre", "linteau"},
    {"poteau", "poteaux", "pilier", "colonne"},
    {"joint", "fractionnement", "dilatation", "profilé"},
    {"propreté", "proprete", "maigre", "150", "dosé"},
    {"dallage", "dallé", "sol", "plancher", "rez"},
]


def _bonus_synonymes(mots_a: set, mots_b: set) -> float:
    """Bonus si les mots appartiennent au même groupe sémantique GO."""
    bonus = 0.0
    for groupe in SYNONYMES_GO:
        overlap_a = mots_a & groupe
        overlap_b = mots_b & groupe
        if overlap_a and overlap_b:
            bonus += 0.15
    return min(bonus, 0.30)


def normaliser(texte: str) -> str:
    """Normalise un libellé pour la comparaison."""
    if not texte:
        return ""
    t = texte.lower()
    # Supprimer la ponctuation
    t = re.sub(r"[^\w\s]", " ", t)
    # Normaliser les espaces
    t = re.sub(r"\s+", " ", t).strip()
    # Supprimer les mots vides courants
    stops = {"de", "du", "des", "le", "la", "les", "en", "au", "aux", "et",
              "à", "a", "un", "une", "sur", "sous", "par", "pour", "avec",
              "dans", "d", "l", "se", "es", "est"}
    mots = [m for m in t.split() if m not in stops]
    return " ".join(mots)


def score_similarite(libelle_a: str, libelle_b: str) -> float:
    """
    Calcule un score de similarité entre deux libellés.
    Retourne un float entre 0.0 et 1.0.
    """
    a = normaliser(libelle_a)
    b = normaliser(libelle_b)

    if not a or not b:
        return 0.0

    mots_a = set(a.split())
    mots_b = set(b.split())

    if not mots_a or not mots_b:
        return 0.0

    # Score Jaccard sur les mots
    intersection = len(mots_a & mots_b)
    union = len(mots_a | mots_b)
    jaccard = intersection / union if union > 0 else 0.0

    # Bonus synonymes métier GO
    jaccard = min(1.0, jaccard + _bonus_synonymes(mots_a, mots_b))

    # Bonus si les mots techniques clés matchent
    mots_techniques = {
        "béton", "beton", "armé", "arme", "semelle", "voile", "poteau",
        "poutre", "dalle", "plancher", "dallage", "fondation", "fouille",
        "remblai", "terrassement", "maçonnerie", "maconnerie", "enduit",
        "chape", "hourdis", "longrine", "décapage", "decapage", "propreté",
        "proprete", "clôture", "cloture", "installation", "joint"
    }
    tech_a = mots_a & mots_techniques
    tech_b = mots_b & mots_techniques
    if tech_a and tech_b:
        tech_overlap = len(tech_a & tech_b) / max(len(tech_a), len(tech_b))
        jaccard = jaccard * 0.6 + tech_overlap * 0.4

    return min(1.0, jaccard)


def score_to_confiance(score: float) -> str:
    """Convertit un score numérique en niveau de confiance."""
    if score >= 0.55:
        return "ÉLEVÉ"
    elif score >= 0.30:
        return "MOYEN"
    elif score >= 0.10:
        return "FAIBLE"
    return "ABSENT"


def matcher_entreprise_direct(
    postes_reference: list[dict],
    nom_entreprise: str,
    postes_entreprise: list[dict]
) -> list[dict]:
    """
    Matching direct par similarité textuelle pour une entreprise.
    Retourne la même structure que le SemanticMatcher.
    """
    logger.info(f"Matching direct (sans API) — {nom_entreprise}")
    resultats = []

    for i, poste_ref in enumerate(postes_reference):
        libelle_ref = poste_ref.get("libelle", "")
        unite_ref = poste_ref.get("unite", "")

        meilleur_score = 0.0
        meilleur_idx = -1
        meilleur_poste = None

        for j, poste_ent in enumerate(postes_entreprise):
            libelle_ent = poste_ent.get("libelle", "")
            unite_ent = poste_ent.get("unite", "")

            score = score_similarite(libelle_ref, libelle_ent)

            # Pénalité si les unités sont incompatibles
            if unite_ref and unite_ent:
                u_ref = normaliser(unite_ref)
                u_ent = normaliser(unite_ent)
                if u_ref and u_ent and u_ref != u_ent:
                    # Unités différentes mais pas forcément incompatibles
                    compatibles = _unites_compatibles(u_ref, u_ent)
                    if not compatibles:
                        score *= 0.5

            if score > meilleur_score:
                meilleur_score = score
                meilleur_idx = j
                meilleur_poste = poste_ent

        confiance = score_to_confiance(meilleur_score)
        indices = [meilleur_idx] if meilleur_idx >= 0 and confiance != "ABSENT" else []
        postes_matchés = [meilleur_poste] if meilleur_poste and confiance != "ABSENT" else []

        pu = meilleur_poste.get("prix_unitaire") if meilleur_poste and confiance != "ABSENT" else None

        # Toujours recalculer le prix total avec les quantités de RÉFÉRENCE × PU entreprise
        # Ne jamais utiliser le prix_total brut du poste entreprise (les qts peuvent différer)
        pt = None
        if pu is not None:
            qt = poste_ref.get("quantite")
            if qt:
                pt = round(pu * qt, 2)

        resultats.append({
            "ref_index": i,
            "poste_reference": poste_ref,
            "indices_entreprise": indices,
            "postes_entreprise": postes_matchés,
            "score_confiance": confiance,
            "type": "exact" if confiance in ("ÉLEVÉ", "MOYEN") else ("partiel" if confiance == "FAIBLE" else "absent"),
            "commentaire": f"Score similarité : {meilleur_score:.2f}" if confiance != "ABSENT" else "Aucune correspondance trouvée",
            "prix_unitaire_retenu": pu,
            "prix_total_retenu": pt
        })

    absents = sum(1 for r in resultats if r["score_confiance"] == "ABSENT")
    logger.info(f"  {len(resultats)} matchings | {absents} absents | {len(resultats)-absents} trouvés")
    return resultats


def _unites_compatibles(u1: str, u2: str) -> bool:
    """Vérifie si deux unités de mesure sont compatibles."""
    groupes = [
        {"m2", "m²", "m 2"},
        {"m3", "m³", "m 3"},
        {"ml", "ml", "m", "lm"},
        {"u", "unite", "unité", "pce", "piece", "pièce"},
        {"forfait", "fft", "ens", "ensemble", "ft"},
        {"kg", "t", "tonne"},
    ]
    for groupe in groupes:
        if u1 in groupe and u2 in groupe:
            return True
    return u1 == u2


def matcher_toutes_entreprises_direct(
    postes_reference: list[dict],
    entreprises_postes: dict
) -> dict:
    """
    Matching direct pour toutes les entreprises, lot par lot.
    Garantit que les postes d'un lot de référence ne sont matchés
    qu'avec les postes du même lot chez l'entreprise.
    """
    resultats = {}
    lots = sorted(set(p['lot'] for p in postes_reference))

    for nom, postes_ent in entreprises_postes.items():
        matchings_entreprise = []

        for lot in lots:
            ref_lot = [p for p in postes_reference if p['lot'] == lot]
            ent_lot = [p for p in postes_ent if p['lot'] == lot]

            if not ent_lot:
                # Lot absent chez cette entreprise → tout marquer ABSENT
                for poste_ref in ref_lot:
                    matchings_entreprise.append({
                        "ref_index": postes_reference.index(poste_ref),
                        "poste_reference": poste_ref,
                        "indices_entreprise": [],
                        "postes_entreprise": [],
                        "score_confiance": "ABSENT",
                        "type": "absent",
                        "commentaire": f"Lot {lot} absent dans l'offre",
                        "prix_unitaire_retenu": None,
                        "prix_total_retenu": None
                    })
                logger.warning(f"  {nom} — Lot {lot} : aucun poste trouvé")
                continue

            matchings_lot = matcher_entreprise_direct(ref_lot, nom, ent_lot)
            matchings_entreprise.extend(matchings_lot)
            logger.info(f"  {nom} — Lot {lot} : {len(matchings_lot)} matchings")

        resultats[nom] = matchings_entreprise

    return resultats
