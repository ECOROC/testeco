"""
parsers/ocr_parser.py — Parsing des PDF scannés (OCR Tesseract + fallback Claude Vision)
"""

import base64
import json
import logging
import re
from pathlib import Path
from typing import Optional

from config import OCR_SCORE_MIN, TESSERACT_LANG, ANTHROPIC_API_KEY, CLAUDE_MODEL

logger = logging.getLogger(__name__)


def parser_pdf_scan(chemin: str) -> list[dict]:
    """
    Parse un PDF scanné via OCR.
    Stratégie :
    1. OCR avec Tesseract
    2. Si score de confiance < seuil → fallback Claude Vision
    """
    chemin = Path(chemin)
    logger.info(f"Parsing PDF scanné : {chemin.name}")

    # Convertir PDF en images
    images = _pdf_vers_images(chemin)
    if not images:
        logger.error("Impossible de convertir le PDF en images")
        return []

    logger.info(f"  {len(images)} page(s) à traiter en OCR")

    tous_postes = []
    pages_basse_confiance = []

    for num_page, image in enumerate(images, 1):
        texte, score = _ocr_tesseract(image)
        logger.info(f"  Page {num_page} — score OCR : {score:.0f}%")

        if score >= OCR_SCORE_MIN and texte:
            postes = _parser_texte_ocr(texte, num_page)
            tous_postes.extend(postes)
        else:
            logger.warning(f"  Page {num_page} : score OCR insuffisant ({score:.0f}%), mise en file pour Claude Vision")
            pages_basse_confiance.append((num_page, image))

    # Fallback Claude Vision pour les pages à faible confiance
    if pages_basse_confiance and ANTHROPIC_API_KEY:
        logger.info(f"  Fallback Claude Vision pour {len(pages_basse_confiance)} page(s)")
        for num_page, image in pages_basse_confiance:
            postes = _claude_vision_page(image, num_page)
            tous_postes.extend(postes)
    elif pages_basse_confiance:
        logger.warning("  Claude Vision non disponible (pas de clé API). Pages ignorées.")

    logger.info(f"Total : {len(tous_postes)} postes extraits")
    return tous_postes


def _pdf_vers_images(chemin: Path) -> list:
    """Convertit un PDF en liste d'images PIL."""
    try:
        from pdf2image import convert_from_path
        images = convert_from_path(str(chemin), dpi=300)
        return images
    except Exception as e:
        logger.error(f"Erreur conversion PDF→images : {e}")
        # Fallback avec PyMuPDF si disponible
        try:
            import fitz
            doc = fitz.open(str(chemin))
            images = []
            for page in doc:
                mat = fitz.Matrix(3.0, 3.0)  # 300 DPI équivalent
                pix = page.get_pixmap(matrix=mat)
                from PIL import Image
                import io
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                images.append(img)
            return images
        except Exception as e2:
            logger.error(f"Erreur fallback PyMuPDF : {e2}")
            return []


def _ocr_tesseract(image) -> tuple[str, float]:
    """
    Applique OCR Tesseract sur une image PIL.
    Retourne (texte, score_confiance).
    """
    try:
        import pytesseract
        from pytesseract import Output

        # OCR avec données de confiance
        data = pytesseract.image_to_data(
            image,
            lang=TESSERACT_LANG,
            output_type=Output.DICT,
            config="--psm 6"  # Mode : bloc uniforme de texte
        )

        texte = pytesseract.image_to_string(image, lang=TESSERACT_LANG, config="--psm 6")

        # Calculer le score de confiance moyen (ignorer les -1)
        confiances = [c for c in data["conf"] if c != -1]
        score = sum(confiances) / len(confiances) if confiances else 0.0

        return texte, score

    except Exception as e:
        logger.error(f"Erreur Tesseract : {e}")
        return "", 0.0


def _parser_texte_ocr(texte: str, num_page: int) -> list[dict]:
    """
    Parse le texte brut issu d'un OCR pour en extraire des postes DPGF.
    """
    postes = []
    lot_courant = "00"

    lignes = texte.split("\n")

    # Pattern pour détecter une ligne de poste :
    # N° | libellé | unité | qté | PU | total
    pattern = re.compile(
        r"(\d[\d\.]*)\s{2,}(.{10,}?)\s{2,}([a-zA-Zé²³/]+)\s+([\d\s,\.]+)\s+([\d\s,\.]+)\s+([\d\s,\.]+)"
    )

    for ligne in lignes:
        ligne = ligne.strip()
        if not ligne:
            continue

        # Détecter lot
        match_lot = re.search(r"^lot\s*n?°?\s*(\d+)", ligne, re.IGNORECASE)
        if match_lot:
            lot_courant = match_lot.group(1).zfill(2)
            continue

        match = pattern.match(ligne)
        if match:
            poste = {
                "lot": lot_courant,
                "numero": match.group(1),
                "libelle": match.group(2).strip(),
                "unite": match.group(3),
                "quantite": _to_float(match.group(4)),
                "prix_unitaire": _to_float(match.group(5)),
                "prix_total": _to_float(match.group(6))
            }
            if poste["libelle"] and len(poste["libelle"]) > 3:
                postes.append(poste)

    return postes


def _claude_vision_page(image, num_page: int) -> list[dict]:
    """
    Utilise Claude Vision pour extraire les postes d'une page scannée de mauvaise qualité.
    """
    try:
        import anthropic
        from PIL import Image
        import io

        # Convertir image en base64
        buffer = io.BytesIO()
        image.save(buffer, format="JPEG", quality=85)
        image_b64 = base64.b64encode(buffer.getvalue()).decode()

        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

        prompt = """Tu es un expert en économie de la construction française.
Cette image est une page d'un DPGF (Décomposition du Prix Global et Forfaitaire) scanné.

Extrait tous les postes de travaux visibles dans cette image.
Pour chaque poste, retourne un objet JSON avec ces champs :
- lot : numéro de lot (ex: "01")
- numero : numéro du poste (ex: "1.2.3")
- libelle : désignation complète du poste
- unite : unité de mesure (m², ml, m³, u, ens, forfait, etc.)
- quantite : quantité (nombre décimal ou null)
- prix_unitaire : prix unitaire HT (nombre décimal ou null)
- prix_total : montant total HT (nombre décimal ou null)

Retourne UNIQUEMENT un tableau JSON valide, sans texte avant ni après.
Si une valeur est illisible, mets null.
Ne retourne que les lignes de postes réels, pas les sous-totaux ni les en-têtes."""

        response = client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=2000,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/jpeg",
                            "data": image_b64
                        }
                    },
                    {"type": "text", "text": prompt}
                ]
            }]
        )

        texte_reponse = response.content[0].text.strip()
        # Nettoyer les éventuels blocs markdown
        texte_reponse = re.sub(r"```json|```", "", texte_reponse).strip()

        postes = json.loads(texte_reponse)
        logger.info(f"  Page {num_page} — Claude Vision : {len(postes)} poste(s) extraits")
        return postes

    except json.JSONDecodeError as e:
        logger.error(f"Claude Vision : réponse JSON invalide page {num_page} : {e}")
        return []
    except Exception as e:
        logger.error(f"Claude Vision erreur page {num_page} : {e}")
        return []


def _to_float(val) -> Optional[float]:
    if val is None:
        return None
    s = str(val).strip().replace(" ", "").replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        return float(s) if s else None
    except ValueError:
        return None
