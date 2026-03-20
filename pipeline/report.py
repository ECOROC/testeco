"""
pipeline/report.py — Génération du rapport de préparation
"""

import json
import logging
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)


def generer_rapport_preparation(
    dossier_racine: str,
    entreprises: list[dict],
    regroupement: dict
) -> dict:
    """
    Génère un rapport complet de la phase de préparation.
    Retourne le rapport sous forme de dict et l'écrit en JSON.
    """
    rapport = {
        "date_analyse": datetime.now().isoformat(),
        "dossier_racine": dossier_racine,
        "resume": {
            "total_entreprises": len(entreprises),
            "avec_excel": sum(1 for e in entreprises if e.get("type") == "excel"),
            "avec_pdf_natif": sum(1 for e in entreprises if e.get("type") == "pdf_natif"),
            "avec_pdf_scan": sum(1 for e in entreprises if e.get("type") == "pdf_scan"),
            "sans_fichier": sum(1 for e in entreprises if not e.get("fichier_dpgf")),
            "fichiers_regroupes": len(regroupement.get("fichiers_crees", [])),
            "erreurs": len(regroupement.get("erreurs", []))
        },
        "entreprises": [],
        "anomalies": regroupement.get("erreurs", []),
        "dossier_regroupement": regroupement.get("dossier", "")
    }

    for e in entreprises:
        rapport["entreprises"].append({
            "nom": e["nom"],
            "fichier_detecte": Path(e["fichier_dpgf"]).name if e.get("fichier_dpgf") else None,
            "type": e.get("type"),
            "confiance": e.get("confiance", 0.0),
            "statut": "OK" if e.get("fichier_dpgf") else "MANQUANT",
            "message": e.get("message", "")
        })

    # Écrire le rapport JSON
    chemin_rapport = Path(dossier_racine) / "_REGROUPEMENT" / "rapport_preparation.json"
    try:
        chemin_rapport.parent.mkdir(parents=True, exist_ok=True)
        with open(chemin_rapport, "w", encoding="utf-8") as f:
            json.dump(rapport, f, ensure_ascii=False, indent=2)
        logger.info(f"Rapport écrit : {chemin_rapport}")
    except Exception as e:
        logger.warning(f"Impossible d'écrire le rapport JSON : {e}")

    return rapport


def afficher_rapport_console(rapport: dict) -> None:
    """Affiche un résumé du rapport dans la console."""
    r = rapport["resume"]
    print("\n" + "=" * 60)
    print("  RAPPORT DE PRÉPARATION — DPGF COMPARATOR")
    print("=" * 60)
    print(f"  Entreprises détectées   : {r['total_entreprises']}")
    print(f"  Fichiers Excel          : {r['avec_excel']}")
    print(f"  Fichiers PDF natifs     : {r['avec_pdf_natif']}")
    print(f"  Fichiers PDF scannés    : {r['avec_pdf_scan']}")
    print(f"  Sans fichier détecté    : {r['sans_fichier']}")
    print(f"  Fichiers regroupés      : {r['fichiers_regroupes']}")

    if r["erreurs"] > 0:
        print(f"\n  ⚠️  ANOMALIES ({r['erreurs']}) :")
        for anomalie in rapport["anomalies"]:
            print(f"     - {anomalie}")

    print("\n  DÉTAIL PAR ENTREPRISE :")
    print(f"  {'Entreprise':<35} {'Type':<12} {'Statut':<10} {'Confiance'}")
    print("  " + "-" * 65)
    for e in rapport["entreprises"]:
        confiance_pct = f"{e['confiance'] * 100:.0f}%"
        statut_icon = "✓" if e["statut"] == "OK" else "✗"
        print(f"  {e['nom']:<35} {(e['type'] or 'N/A'):<12} {statut_icon} {e['statut']:<8} {confiance_pct}")

    print("=" * 60 + "\n")
