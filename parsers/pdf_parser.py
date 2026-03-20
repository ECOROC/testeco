"""
parsers/pdf_parser.py — Parsing des PDF natifs (texte extractible)
"""

import logging
import re
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)


def parser_pdf_natif(chemin: str) -> list[dict]:
    """
    Parse un PDF natif (texte extractible) et retourne les postes DPGF.
    """
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("pdfplumber est requis : pip install pdfplumber")

    chemin = Path(chemin)
    logger.info(f"Parsing PDF natif : {chemin.name}")

    postes = []
    lot_courant = "00"

    with pdfplumber.open(chemin) as pdf:
        logger.info(f"  {len(pdf.pages)} page(s) à traiter")

        for num_page, page in enumerate(pdf.pages, 1):
            logger.debug(f"  Page {num_page}")

            # Tenter l'extraction de tableaux en premier
            tableaux = page.extract_tables()
            if tableaux:
                for tableau in tableaux:
                    postes_tableau, lot_courant = _parser_tableau(
                        tableau, lot_courant, num_page
                    )
                    postes.extend(postes_tableau)
            else:
                # Fallback : extraction texte brut
                texte = page.extract_text()
                if texte:
                    postes_texte, lot_courant = _parser_texte_brut(
                        texte, lot_courant, num_page
                    )
                    postes.extend(postes_texte)

    logger.info(f"Total : {len(postes)} postes extraits")
    return postes


def _parser_tableau(tableau: list, lot_courant: str, num_page: int) -> tuple[list[dict], str]:
    """
    Parse un tableau pdfplumber (liste de listes) et extrait les postes.
    """
    if not tableau or len(tableau) < 2:
        return [], lot_courant

    postes = []

    # Identifier les colonnes depuis la première ligne non-vide
    mapping = {}
    header_idx = 0

    for i, row in enumerate(tableau[:5]):
        mapping = _detecter_colonnes_tableau(row)
        if mapping:
            header_idx = i
            break

    if not mapping:
        return [], lot_courant

    for row in tableau[header_idx + 1:]:
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue

        # Détecter titre de lot
        nouveau_lot = _detecter_lot_dans_ligne(row)
        if nouveau_lot:
            lot_courant = nouveau_lot
            continue

        poste = _extraire_poste_tableau(row, mapping, lot_courant)
        if poste:
            postes.append(poste)

    return postes, lot_courant


def _detecter_colonnes_tableau(row: list) -> dict:
    """Détecte le mapping des colonnes depuis une ligne d'en-têtes."""
    from config import (COLONNES_NUMERO, COLONNES_LIBELLE, COLONNES_UNITE,
                        COLONNES_QUANTITE, COLONNES_PU, COLONNES_TOTAL)
    mapping = {}

    for idx, cell in enumerate(row):
        if not cell:
            continue
        cell_lower = str(cell).lower().strip()

        if any(m in cell_lower for m in COLONNES_LIBELLE) and "libelle" not in mapping:
            mapping["libelle"] = idx
        elif any(m in cell_lower for m in COLONNES_NUMERO) and "numero" not in mapping:
            mapping["numero"] = idx
        elif any(m in cell_lower for m in COLONNES_UNITE) and "unite" not in mapping:
            mapping["unite"] = idx
        elif any(m in cell_lower for m in COLONNES_QUANTITE) and "quantite" not in mapping:
            mapping["quantite"] = idx
        elif any(m in cell_lower for m in COLONNES_PU) and "prix_unitaire" not in mapping:
            mapping["prix_unitaire"] = idx
        elif any(m in cell_lower for m in COLONNES_TOTAL) and "prix_total" not in mapping:
            mapping["prix_total"] = idx

    return mapping


def _extraire_poste_tableau(row: list, mapping: dict, lot: str) -> Optional[dict]:
    """Extrait un poste depuis une ligne de tableau."""
    def get(cle):
        idx = mapping.get(cle)
        if idx is None or idx >= len(row):
            return None
        val = row[idx]
        return str(val).strip() if val is not None else None

    libelle = get("libelle")
    if not libelle or len(libelle) < 3:
        return None

    pu = _to_float(get("prix_unitaire"))
    total = _to_float(get("prix_total"))
    qt = _to_float(get("quantite"))

    if pu is None and total is None and qt is None:
        return None

    return {
        "lot": lot,
        "numero": get("numero") or "",
        "libelle": libelle,
        "unite": get("unite") or "",
        "quantite": qt,
        "prix_unitaire": pu,
        "prix_total": total or ((pu or 0) * (qt or 0) if pu and qt else None)
    }


def _parser_texte_brut(texte: str, lot_courant: str, num_page: int) -> tuple[list[dict], str]:
    """
    Tente d'extraire des postes depuis du texte brut PDF (fallback).
    Heuristique : chercher des lignes avec des valeurs numériques en fin.
    """
    postes = []
    lignes = texte.split("\n")

    pattern_poste = re.compile(
        r"^(\d[\d\.]*)\s+(.{10,80}?)\s+([a-zA-Zé²³]+)\s+([\d\s,\.]+)\s+([\d\s,\.]+)\s+([\d\s,\.]+)",
        re.MULTILINE
    )

    for match in pattern_poste.finditer(texte):
        try:
            postes.append({
                "lot": lot_courant,
                "numero": match.group(1).strip(),
                "libelle": match.group(2).strip(),
                "unite": match.group(3).strip(),
                "quantite": _to_float(match.group(4)),
                "prix_unitaire": _to_float(match.group(5)),
                "prix_total": _to_float(match.group(6))
            })
        except Exception:
            continue

    return postes, lot_courant


def _detecter_lot_dans_ligne(row: list) -> Optional[str]:
    """Cherche un titre de lot dans une ligne de tableau."""
    for cell in row:
        if cell and isinstance(cell, str):
            match = re.search(r"^lot\s*n?°?\s*(\d+|[A-Z]{1,3})\b", cell.strip(), re.IGNORECASE)
            if match:
                return match.group(1).zfill(2)
    return None


def _to_float(val) -> Optional[float]:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    s = re.sub(r"[^\d,.\-]", "", s).replace(",", ".")
    try:
        return float(s) if s else None
    except ValueError:
        return None
