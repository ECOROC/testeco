"""
main.py — Point d'entrée CLI du DPGF Comparator
"""

import argparse
import logging
import sys
from pathlib import Path

from config import LOG_FILE, LOG_LEVEL


def configurer_logging():
    level = getattr(logging, LOG_LEVEL.upper(), logging.INFO)
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s : %(message)s",
        datefmt="%H:%M:%S",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(LOG_FILE, encoding="utf-8")
        ]
    )


def run_pipeline(args):
    """Lance le pipeline complet d'analyse DPGF."""
    from pipeline.extractor import detecter_entreprises, extraire_archive
    from pipeline.file_detector import detecter_fichier_dpgf
    from pipeline.renamer import creer_regroupement
    from pipeline.report import generer_rapport_preparation, afficher_rapport_console
    from parsers.excel_parser import parser_excel
    from parsers.pdf_parser import parser_pdf_natif
    from parsers.ocr_parser import parser_pdf_scan

    logger = logging.getLogger("main")

    # ─── Étape 1 : Extraction et détection ────────────────────────────────────
    print("\n🔍 ÉTAPE 1 — Extraction et détection des entreprises")

    chemin_entree = Path(args.entree)

    # Si c'est une archive, extraire d'abord
    if chemin_entree.is_file():
        print(f"  Extraction de l'archive : {chemin_entree.name}")
        dossier_travail = chemin_entree.parent / (chemin_entree.stem + "_extrait")
        dossier_racine = extraire_archive(str(chemin_entree), str(dossier_travail))
    else:
        dossier_racine = str(chemin_entree)
        print(f"  Dossier source : {dossier_racine}")

    entreprises = detecter_entreprises(dossier_racine)

    if not entreprises:
        print("❌ Aucune entreprise détectée. Vérifiez la structure du dossier.")
        sys.exit(1)

    # Détecter les fichiers DPGF
    print(f"\n  {len(entreprises)} entreprise(s) détectée(s). Identification des DPGF...")
    for e in entreprises:
        if e["chemin_dpgf"]:
            resultat = detecter_fichier_dpgf(e["chemin_dpgf"])
            e.update({
                "fichier_dpgf": str(resultat["fichier"]) if resultat["fichier"] else None,
                "type": resultat["type"],
                "confiance": resultat["confiance"],
                "message": resultat["message"]
            })
        else:
            e.update({"fichier_dpgf": None, "type": None, "confiance": 0.0, "message": "Pas de dossier DPGF"})

    # Créer le regroupement
    regroupement = creer_regroupement(dossier_racine, entreprises)

    # Rapport de préparation
    rapport = generer_rapport_preparation(dossier_racine, entreprises, regroupement)
    afficher_rapport_console(rapport)

    if args.etape == 1:
        print("✅ Étape 1 terminée. Rapport disponible dans _REGROUPEMENT/rapport_preparation.json")
        return

    # ─── Étape 2 : Parsing ────────────────────────────────────────────────────
    print("\n📄 ÉTAPE 2 — Parsing des fichiers DPGF")

    postes_par_entreprise = {}

    for e in entreprises:
        if not e.get("fichier_dpgf"):
            print(f"  ⚠️  {e['nom']} : pas de fichier à parser")
            postes_par_entreprise[e["nom"]] = []
            continue

        print(f"  Parsing {e['nom']} ({e['type']})...")
        fichier = e["fichier_dpgf"]

        try:
            if e["type"] == "excel":
                postes = parser_excel(fichier)
            elif e["type"] == "pdf_natif":
                postes = parser_pdf_natif(fichier)
            elif e["type"] == "pdf_scan":
                postes = parser_pdf_scan(fichier)
            else:
                postes = []

            postes_par_entreprise[e["nom"]] = postes
            print(f"  ✓ {e['nom']} : {len(postes)} poste(s) extraits")

        except Exception as ex:
            logger.error(f"Erreur parsing {e['nom']} : {ex}")
            print(f"  ✗ {e['nom']} : erreur — {ex}")
            postes_par_entreprise[e["nom"]] = []

    if args.etape == 2:
        print("\n✅ Étape 2 terminée.")
        for nom, postes in postes_par_entreprise.items():
            print(f"   {nom} : {len(postes)} postes")
        return

    # ─── Étape 3 : DPGF de référence ─────────────────────────────────────────
    print("\n📋 ÉTAPE 3 — Chargement du DPGF de référence")

    if args.reference:
        print(f"  Chargement de la référence : {args.reference}")
        postes_reference = parser_excel(args.reference)
        print(f"  ✓ {len(postes_reference)} poste(s) de référence chargés")
    else:
        print("  Mode 'depuis zéro' : reconstruction de la trame depuis les offres")
        postes_reference = _reconstruire_reference(postes_par_entreprise)
        print(f"  ✓ {len(postes_reference)} poste(s) reconstruits")

    if not postes_reference:
        print("❌ Aucun poste de référence. Impossible de continuer.")
        sys.exit(1)

    if args.etape == 3:
        print("\n✅ Étape 3 terminée.")
        return

    # ─── Étape 4 : Matching sémantique ───────────────────────────────────────
    print("\n🤖 ÉTAPE 4 — Matching sémantique via Claude API")

    from config import ANTHROPIC_API_KEY
    from matching.batch_manager import estimer_cout_api

    estimation = estimer_cout_api(len(postes_reference), len(postes_par_entreprise))
    print(f"  Estimation : {estimation['nb_appels']} appel(s) API")
    print(f"  Coût estimé : {estimation['cout_estime_cts']:.1f} centimes EUR")

    if ANTHROPIC_API_KEY:
        from matching.semantic_matcher import matcher_toutes_entreprises
        print("  Mode : Matching sémantique Claude API")

        def callback_progression(etape, total, message):
            pct = int((etape / max(total, 1)) * 100)
            print(f"  [{pct:3d}%] {message}")

        matchings = matcher_toutes_entreprises(
            postes_reference,
            postes_par_entreprise,
            callback=callback_progression
        )
    else:
        from matching.direct_matcher import matcher_toutes_entreprises_direct
        print("  ⚠️  Pas de clé API — Mode matching direct (similarité textuelle)")
        print("     → Pour activer le matching sémantique : export ANTHROPIC_API_KEY=sk-...")
        matchings = matcher_toutes_entreprises_direct(
            postes_reference,
            postes_par_entreprise
        )

    if args.etape == 4:
        print("\n✅ Étape 4 terminée.")
        return

    # ─── Étape 5 : Export Excel ───────────────────────────────────────────────
    print("\n📊 ÉTAPE 5 — Génération du fichier Excel de comparaison")

    from output.excel_exporter import generer_excel_comparatif

    nom_sortie = args.sortie or str(Path(dossier_racine) / "Comparatif_DPGF.xlsx")
    nom_operation = Path(dossier_racine).name

    chemin_excel = generer_excel_comparatif(
        postes_reference,
        matchings,
        nom_sortie,
        nom_operation
    )

    print(f"\n✅ Fichier Excel généré : {chemin_excel}")
    print(f"   {len(postes_reference)} postes de référence")
    print(f"   {len(entreprises)} entreprises comparées")


def _reconstruire_reference(postes_par_entreprise: dict) -> list[dict]:
    """
    Reconstruit une trame de référence depuis les postes des entreprises.
    Utilise l'entreprise ayant le plus de postes comme base.
    """
    if not postes_par_entreprise:
        return []

    meilleure = max(postes_par_entreprise, key=lambda k: len(postes_par_entreprise[k]))
    return postes_par_entreprise[meilleure]


def main():
    configurer_logging()

    parser = argparse.ArgumentParser(
        description="DPGF Comparator — Analyse comparative des offres d'appel d'offres",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemples :
  # Lancer le pipeline complet
  python main.py --entree /chemin/AO_Operation --reference DPGF_Reference.xlsx

  # Tester uniquement l'étape 1 (extraction et détection)
  python main.py --entree /chemin/AO_Operation --etape 1

  # Interface Streamlit
  streamlit run app.py
        """
    )

    parser.add_argument(
        "--entree", required=True,
        help="Dossier AO extrait ou archive (.zip/.rar)"
    )
    parser.add_argument(
        "--reference", default=None,
        help="DPGF de référence vierge (Excel, optionnel)"
    )
    parser.add_argument(
        "--sortie", default=None,
        help="Chemin du fichier Excel de sortie"
    )
    parser.add_argument(
        "--etape", type=int, default=5, choices=[1, 2, 3, 4, 5],
        help="Arrêter après cette étape (défaut : 5 = pipeline complet)"
    )

    args = parser.parse_args()
    run_pipeline(args)


if __name__ == "__main__":
    main()
