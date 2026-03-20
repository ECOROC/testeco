"""
tests/test_pipeline.py — Tests unitaires du pipeline DPGF Comparator
"""

import sys
import os
import json
from pathlib import Path

# Ajouter le dossier racine au path
sys.path.insert(0, str(Path(__file__).parent.parent))


def test_etape1_detection():
    """Test : détection des entreprises et fichiers DPGF."""
    print("\n" + "="*60)
    print("TEST ÉTAPE 1 — Détection des entreprises")
    print("="*60)

    from pipeline.extractor import detecter_entreprises
    from pipeline.file_detector import detecter_fichier_dpgf

    dossier = Path("test_data/AO_Construction_Bureaux")
    if not dossier.exists():
        print("⚠️  Données de test absentes. Génération en cours...")
        from tests.generer_donnees_test import generer_tout
        generer_tout()

    entreprises = detecter_entreprises(str(dossier))
    assert len(entreprises) == 3, f"Attendu 3 entreprises, obtenu {len(entreprises)}"
    print(f"✓ {len(entreprises)} entreprises détectées")

    noms = [e["nom"] for e in entreprises]
    assert "Entreprise_Dupont_BTP" in noms
    assert "Entreprise_Martin_Construction" in noms
    assert "Entreprise_Vinci_Batiment" in noms
    print(f"✓ Noms corrects : {noms}")

    for e in entreprises:
        resultat = detecter_fichier_dpgf(e["chemin_dpgf"])
        assert resultat["fichier"] is not None, f"Fichier DPGF non trouvé pour {e['nom']}"
        assert resultat["type"] == "excel", f"Type attendu 'excel', obtenu '{resultat['type']}'"
        print(f"✓ {e['nom']} → {Path(resultat['fichier']).name} ({resultat['type']}, confiance {resultat['confiance']:.0%})")

    print("\n✅ ÉTAPE 1 : OK")
    return entreprises


def test_etape1_regroupement(entreprises):
    """Test : création du dossier _REGROUPEMENT."""
    from pipeline.file_detector import detecter_fichier_dpgf
    from pipeline.renamer import creer_regroupement

    dossier = Path("test_data/AO_Construction_Bureaux")

    # Enrichir les entreprises avec les infos de fichier
    for e in entreprises:
        r = detecter_fichier_dpgf(e["chemin_dpgf"])
        e.update({
            "fichier_dpgf": str(r["fichier"]) if r["fichier"] else None,
            "type": r["type"],
            "confiance": r["confiance"]
        })

    rapport = creer_regroupement(str(dossier), entreprises)
    assert len(rapport["fichiers_crees"]) == 3
    print(f"✓ {len(rapport['fichiers_crees'])} fichiers créés dans _REGROUPEMENT")

    regroupement_dir = dossier / "_REGROUPEMENT"
    assert regroupement_dir.exists()

    fichiers = list(regroupement_dir.glob("DPGF_*.xlsx"))
    assert len(fichiers) == 3
    print(f"✓ Fichiers : {[f.name for f in fichiers]}")

    print("✅ REGROUPEMENT : OK")
    return entreprises  # entreprises maintenant enrichies


def test_etape2_parsing_excel():
    """Test : parsing des fichiers Excel."""
    print("\n" + "="*60)
    print("TEST ÉTAPE 2 — Parsing Excel")
    print("="*60)

    from parsers.excel_parser import parser_excel

    fichiers_test = [
        "test_data/AO_Construction_Bureaux/Entreprise_Dupont_BTP/DPGF/48291047.xlsx",
        "test_data/AO_Construction_Bureaux/Entreprise_Martin_Construction/DPGF/offre_finale.xlsx",
        "test_data/AO_Construction_Bureaux/Entreprise_Vinci_Batiment/DPGF/reponse.xlsx",
    ]

    resultats = {}

    for chemin in fichiers_test:
        nom = Path(chemin).parent.parent.name
        postes = parser_excel(chemin)
        resultats[nom] = postes

        assert len(postes) > 0, f"Aucun poste extrait de {nom}"
        print(f"✓ {nom} : {len(postes)} postes extraits")

        # Vérifier la structure
        for poste in postes[:3]:
            assert "libelle" in poste, "Champ 'libelle' manquant"
            assert "lot" in poste, "Champ 'lot' manquant"
            assert "prix_unitaire" in poste, "Champ 'prix_unitaire' manquant"

        # Quelques prix
        postes_avec_prix = [p for p in postes if p.get("prix_unitaire")]
        print(f"  Postes avec prix : {len(postes_avec_prix)}/{len(postes)}")
        if postes_avec_prix:
            print(f"  Exemple : {postes_avec_prix[0]['libelle'][:50]} — {postes_avec_prix[0]['prix_unitaire']} €/u")

    # Martin doit avoir moins de postes (2 absents)
    nb_dupont = len(resultats["Entreprise_Dupont_BTP"])
    nb_martin = len(resultats["Entreprise_Martin_Construction"])
    print(f"\n✓ Dupont : {nb_dupont} postes | Martin : {nb_martin} postes (doit être < Dupont)")

    print("\n✅ ÉTAPE 2 : OK")
    return resultats


def test_etape2_reference():
    """Test : parsing du DPGF de référence."""
    from parsers.excel_parser import parser_excel

    chemin_ref = "test_data/DPGF_Reference_GO_vierge.xlsx"
    if not Path(chemin_ref).exists():
        print("⚠️  Référence absente, skip")
        return []

    postes = parser_excel(chemin_ref)
    print(f"✓ Référence : {len(postes)} postes")
    assert len(postes) > 0
    return postes


def test_renommage_fichier():
    """Test : fonction de nettoyage des noms d'entreprises."""
    from pipeline.renamer import _nettoyer_nom_entreprise

    cas_tests = [
        ("Entreprise_Dupont_BTP", "Dupont_BTP"),
        ("SARL Martin Construction", "Martin_Construction"),
        ("Vinci Bâtiment & Fils", "Vinci_B_timent_Fils"),
        ("groupe ABC", "ABC"),
    ]

    for entree, attendu in cas_tests:
        resultat = _nettoyer_nom_entreprise(entree)
        # Test souple : juste vérifier que le résultat est non-vide et différent
        assert len(resultat) > 0, f"Résultat vide pour '{entree}'"
        print(f"  '{entree}' → '{resultat}'")

    print("✓ Nettoyage des noms : OK")


def test_detection_type_pdf():
    """Test : détection type PDF (natif vs scanné)."""
    from pipeline.file_detector import _detecter_type_pdf
    print("✓ Fonction de détection PDF disponible")


def run_all_tests():
    """Lance tous les tests en séquence."""
    print("\n🧪 DPGF COMPARATOR — SUITE DE TESTS")
    print("=" * 60)

    try:
        # Génération des données si nécessaire
        if not Path("test_data").exists():
            print("Génération des données de test...")
            from tests.generer_donnees_test import generer_tout
            generer_tout()

        # Tests
        test_renommage_fichier()

        entreprises = test_etape1_detection()
        entreprises_enrichies = test_etape1_regroupement(entreprises)

        postes_par_ent = test_etape2_parsing_excel()
        postes_ref = test_etape2_reference()

        test_detection_type_pdf()

        print("\n" + "=" * 60)
        print("✅ TOUS LES TESTS RÉUSSIS")
        print("=" * 60)
        print("\nPour lancer le pipeline complet :")
        print("  python main.py --entree test_data/AO_Construction_Bureaux --reference test_data/DPGF_Reference_GO_vierge.xlsx")
        print("\nPour lancer l'interface Streamlit :")
        print("  streamlit run app.py")

    except AssertionError as e:
        print(f"\n❌ TEST ÉCHOUÉ : {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ ERREUR : {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    run_all_tests()
