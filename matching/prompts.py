"""
matching/prompts.py — Prompts Claude centralisés pour le matching sémantique
"""


def prompt_matching_poste(
    libelle_reference: str,
    unite_reference: str,
    lot: str,
    nom_entreprise: str,
    postes_entreprise: list[dict]
) -> str:
    """
    Construit le prompt de matching pour un poste de référence
    contre la liste des postes d'une entreprise.
    """
    liste_postes = "\n".join([
        f"{i}. [{p.get('numero', '')}] {p.get('libelle', '')} "
        f"({p.get('unite', '')} — PU: {p.get('prix_unitaire', 'N/A')} €)"
        for i, p in enumerate(postes_entreprise)
    ])

    return f"""Tu es un expert en économie de la construction française.

Poste de référence :
- Libellé : {libelle_reference}
- Unité : {unite_reference}
- Lot : {lot}

Postes proposés par l'entreprise {nom_entreprise} :
{liste_postes}

Mission :
1. Identifier le ou les postes qui correspondent au poste de référence
2. Signaler si plusieurs postes sont regroupés en un seul chez l'entreprise
3. Signaler si le poste est absent
4. Attribuer un score : ÉLEVÉ / MOYEN / FAIBLE / ABSENT

Règles d'identification :
- Un poste peut avoir un libellé différent mais la même nature de travaux
- L'unité doit être cohérente (m² ≠ m³, mais m² ≈ m2)
- Si le poste est fragmenté en plusieurs sous-postes chez l'entreprise, liste-les tous
- Si un poste regroupe plusieurs postes de référence, indique type "regroupe"

Réponds UNIQUEMENT en JSON valide, sans texte avant ni après :
{{
  "indices_postes": [<entiers, indices dans la liste ci-dessus>],
  "score_confiance": "ÉLEVÉ|MOYEN|FAIBLE|ABSENT",
  "type": "exact|regroupe|partiel|absent",
  "commentaire": "<string explicatif ou null>"
}}"""


def prompt_matching_batch(
    postes_reference: list[dict],
    nom_entreprise: str,
    postes_entreprise: list[dict]
) -> str:
    """
    Prompt pour matcher un batch de postes de référence en une seule fois.
    Plus efficace en termes d'appels API.
    """
    ref_formatee = "\n".join([
        f"REF_{i}: [{p.get('lot', '')}] {p.get('libelle', '')} ({p.get('unite', '')})"
        for i, p in enumerate(postes_reference)
    ])

    ent_formatee = "\n".join([
        f"ENT_{i}: [{p.get('numero', '')}] {p.get('libelle', '')} "
        f"({p.get('unite', '')} — PU: {p.get('prix_unitaire', 'N/A')} €)"
        for i, p in enumerate(postes_entreprise)
    ])

    return f"""Tu es un expert en économie de la construction française.

POSTES DE RÉFÉRENCE :
{ref_formatee}

POSTES DE L'ENTREPRISE {nom_entreprise} :
{ent_formatee}

Pour chaque poste de référence (REF_0, REF_1, ...), identifie le ou les postes
correspondants dans la liste entreprise (ENT_0, ENT_1, ...).

Règles :
- Libellés différents peuvent désigner les mêmes travaux
- Un poste entreprise peut regrouper plusieurs postes de référence
- Indique ABSENT si le poste n'existe pas dans l'offre entreprise
- Score : ÉLEVÉ (correspondance certaine), MOYEN (probable), FAIBLE (douteux), ABSENT

Retourne UNIQUEMENT un tableau JSON valide :
[
  {{
    "ref_index": 0,
    "indices_entreprise": [<indices ENT correspondants>],
    "score_confiance": "ÉLEVÉ|MOYEN|FAIBLE|ABSENT",
    "type": "exact|regroupe|partiel|absent",
    "commentaire": null
  }},
  ...
]

Le tableau doit avoir exactement {len(postes_reference)} éléments (un par poste de référence)."""
