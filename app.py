"""
app.py — DPGF Comparator — Interface simplifiée lot par lot
"""

import os
import re
import json
import tempfile
from pathlib import Path

import streamlit as st
import pandas as pd

# ─── Config page ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DPGF Comparator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.header {
    background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
    color: white; padding: 1.8rem 2.5rem; border-radius: 12px;
    margin-bottom: 2rem;
}
.header h1 { color: white; margin: 0; font-size: 1.7rem; font-weight: 700; }
.header p  { color: #BDD7EE; margin: 0.3rem 0 0; font-size: 0.9rem; }

.step-box {
    background: white; border: 1px solid #E2E8F0;
    border-radius: 10px; padding: 1.5rem 2rem; margin-bottom: 1.5rem;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
.step-number {
    display: inline-block; background: #2E75B6; color: white;
    border-radius: 50%; width: 28px; height: 28px; text-align: center;
    line-height: 28px; font-weight: 700; font-size: 0.85rem;
    margin-right: 0.6rem;
}
.step-title { font-weight: 600; font-size: 1.05rem; color: #1F4E79; }

.lot-chip {
    display: inline-block; background: #EBF5FB; color: #1F4E79;
    border: 1px solid #BDD7EE; border-radius: 20px;
    padding: 0.25rem 0.75rem; margin: 0.2rem;
    font-size: 0.85rem; font-weight: 500;
}
.lot-chip.active {
    background: #1F4E79; color: white; border-color: #1F4E79;
}

.file-card {
    background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px;
    padding: 0.7rem 1rem; margin: 0.3rem 0;
    display: flex; align-items: center; gap: 0.5rem;
}
.file-ok   { border-left: 3px solid #2E7D32; }
.file-warn { border-left: 3px solid #F59E0B; }
.file-err  { border-left: 3px solid #CC0000; }

.metric-row { display: flex; gap: 1rem; margin: 1rem 0; }
.metric-box {
    flex: 1; background: #F8FAFC; border: 1px solid #E2E8F0;
    border-radius: 8px; padding: 1rem; text-align: center;
}
.metric-val { font-size: 1.8rem; font-weight: 700; color: #1F4E79; }
.metric-lbl { font-size: 0.8rem; color: #64748B; margin-top: 0.2rem; }

.info-box {
    background: #EBF8FF; border: 1px solid #BEE3F8;
    border-radius: 8px; padding: 0.8rem 1.2rem;
    color: #2C5282; font-size: 0.88rem; margin: 0.8rem 0;
}
.warn-box {
    background: #FFFBEB; border: 1px solid #FCD34D;
    border-radius: 8px; padding: 0.8rem 1.2rem;
    color: #92400E; font-size: 0.88rem; margin: 0.8rem 0;
}
</style>
""", unsafe_allow_html=True)

# ─── En-tête ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header">
  <h1>📊 DPGF Comparator</h1>
  <p>Analyse comparative des offres — Économie de la construction</p>
</div>
""", unsafe_allow_html=True)

# ─── Session state ────────────────────────────────────────────────────────────
if "lots"         not in st.session_state: st.session_state.lots = []
if "lot_actif"    not in st.session_state: st.session_state.lot_actif = None
if "cle_api"      not in st.session_state: st.session_state.cle_api = os.getenv("ANTHROPIC_API_KEY","")

# ═══════════════════════════════════════════════════════════════════════════════
# ÉTAPE 1 — DÉCLARER LES LOTS
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="step-box">
  <span class="step-number">1</span>
  <span class="step-title">Déclarer les lots de l'appel d'offres</span>
</div>
""", unsafe_allow_html=True)

col_num, col_nom, col_btn = st.columns([1, 3, 1])
with col_num:
    new_num = st.text_input("N° de lot", placeholder="01", key="inp_num",
                             label_visibility="visible")
with col_nom:
    new_nom = st.text_input("Intitulé du lot", placeholder="Gros Oeuvre",
                             key="inp_nom", label_visibility="visible")
with col_btn:
    st.write("")
    st.write("")
    if st.button("➕ Ajouter", use_container_width=True):
        num = new_num.strip().zfill(2)
        nom = new_nom.strip()
        if num and nom:
            # Éviter les doublons
            if not any(l["num"] == num for l in st.session_state.lots):
                st.session_state.lots.append({"num": num, "nom": nom})
                st.rerun()
            else:
                st.warning(f"Le lot {num} existe déjà.")
        else:
            st.warning("Renseignez le numéro et l'intitulé du lot.")

# Affichage des lots déclarés
if st.session_state.lots:
    cols_lots = st.columns(min(len(st.session_state.lots), 6))
    for i, lot in enumerate(sorted(st.session_state.lots, key=lambda x: x["num"])):
        with cols_lots[i % 6]:
            actif = st.session_state.lot_actif == lot["num"]
            label = f"LOT {lot['num']}\n{lot['nom']}"
            if st.button(label, key=f"btn_lot_{lot['num']}",
                         use_container_width=True,
                         type="primary" if actif else "secondary"):
                st.session_state.lot_actif = lot["num"]
                st.rerun()
    # Bouton supprimer
    with st.expander("Gérer les lots"):
        lot_a_sup = st.selectbox(
            "Supprimer un lot",
            ["—"] + [f"LOT {l['num']} — {l['nom']}" for l in st.session_state.lots]
        )
        if st.button("🗑 Supprimer", type="secondary"):
            if lot_a_sup != "—":
                num_sup = lot_a_sup.split(" ")[1]
                st.session_state.lots = [l for l in st.session_state.lots
                                          if l["num"] != num_sup]
                if st.session_state.lot_actif == num_sup:
                    st.session_state.lot_actif = None
                st.rerun()
else:
    st.markdown("""
    <div class="info-box">
    💡 Commencez par déclarer vos lots. Exemple : N° <b>01</b> — Intitulé <b>Gros Oeuvre</b>
    </div>
    """, unsafe_allow_html=True)

st.divider()

# ═══════════════════════════════════════════════════════════════════════════════
# ÉTAPE 2 — ANALYSER UN LOT
# ═══════════════════════════════════════════════════════════════════════════════

if not st.session_state.lots:
    st.markdown("""
    <div class="step-box">
      <span class="step-number">2</span>
      <span class="step-title">Analyser un lot</span>
      <p style="color:#94A3B8;margin-top:0.5rem;font-size:0.9rem;">
        → Déclarez d'abord vos lots ci-dessus, puis cliquez sur un lot pour l'analyser.
      </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

if not st.session_state.lot_actif:
    st.markdown("""
    <div class="step-box">
      <span class="step-number">2</span>
      <span class="step-title">Analyser un lot</span>
      <p style="color:#64748B;margin-top:0.5rem;font-size:0.9rem;">
        ☝️ Cliquez sur un lot ci-dessus pour démarrer l'analyse.
      </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Lot sélectionné
lot_sel  = next(l for l in st.session_state.lots
                if l["num"] == st.session_state.lot_actif)
lot_label = f"LOT {lot_sel['num']} — {lot_sel['nom'].upper()}"

st.markdown(f"""
<div class="step-box">
  <span class="step-number">2</span>
  <span class="step-title">Analyser : {lot_label}</span>
</div>
""", unsafe_allow_html=True)

# ─── Règle de nommage ─────────────────────────────────────────────────────────
st.markdown("""
<div class="info-box">
<b>📌 Règle de nommage des fichiers</b><br>
Renommez chaque fichier DPGF avec le nom de l'entreprise avant de l'importer.<br>
Exemples : <code>Dupont_BTP.xlsx</code> &nbsp;|&nbsp; <code>Martin_Construction.pdf</code>
&nbsp;|&nbsp; <code>Vinci.xlsx</code><br>
Le nom du fichier (sans extension) devient automatiquement le nom de l'entreprise dans le comparatif.
</div>
""", unsafe_allow_html=True)

# ─── Upload fichiers ──────────────────────────────────────────────────────────
col_up1, col_up2 = st.columns([3, 2])

with col_up1:
    st.subheader("📂 DPGF des entreprises")
    fichiers_ent = st.file_uploader(
        "Déposez les DPGF (Excel ou PDF) — un fichier par entreprise",
        type=["xlsx", "xls", "pdf"],
        accept_multiple_files=True,
        key=f"upload_ent_{lot_sel['num']}",
        label_visibility="collapsed"
    )

with col_up2:
    st.subheader("📋 DPGF de référence")
    fichier_ref = st.file_uploader(
        "Votre trame vierge (optionnel)",
        type=["xlsx", "xls"],
        key=f"upload_ref_{lot_sel['num']}",
        label_visibility="collapsed"
    )
    st.caption("Si absent : l'analyse se base sur les offres entreprises")

    st.subheader("🔑 Clé API Anthropic")
    cle = st.text_input(
        "Clé API",
        value=st.session_state.cle_api,
        type="password",
        label_visibility="collapsed",
        placeholder="sk-ant-... (optionnel — améliore le matching)"
    )
    if cle: st.session_state.cle_api = cle
    st.caption("Sans clé : matching textuel automatique")

# ─── Aperçu des fichiers déposés ──────────────────────────────────────────────
if fichiers_ent:
    st.markdown("**Fichiers détectés :**")
    entreprises_detectees = []
    problemes = []

    for f in fichiers_ent:
        nom_entreprise = Path(f.name).stem.replace("_", " ").strip()
        ext = Path(f.name).suffix.lower()
        taille_ko = len(f.getbuffer()) // 1024

        # Vérifier que le nom semble être un nom d'entreprise (pas une suite de chiffres)
        est_valide = bool(re.search(r'[a-zA-ZÀ-ÿ]{2,}', nom_entreprise))

        if est_valide:
            icone = "✅" if ext in (".xlsx", ".xls") else "📄"
            type_f = "Excel" if ext in (".xlsx", ".xls") else "PDF"
            st.markdown(f"""
            <div class="file-card file-ok">
              {icone} <b>{nom_entreprise}</b>
              <span style="color:#64748B;font-size:0.85rem;">
                — {type_f} — {taille_ko} Ko
              </span>
            </div>
            """, unsafe_allow_html=True)
            entreprises_detectees.append({
                "nom": nom_entreprise, "fichier": f,
                "type": type_f, "ext": ext
            })
        else:
            st.markdown(f"""
            <div class="file-card file-warn">
              ⚠️ <b>{f.name}</b>
              <span style="color:#92400E;font-size:0.85rem;">
                — Nom invalide. Renommez ce fichier avec le nom de l'entreprise.
              </span>
            </div>
            """, unsafe_allow_html=True)
            problemes.append(f.name)

    if problemes:
        st.markdown(f"""
        <div class="warn-box">
        ⚠️ {len(problemes)} fichier(s) avec un nom invalide (suite de chiffres, etc.).<br>
        Renommez-les avec le nom de l'entreprise et redéposez-les.
        </div>
        """, unsafe_allow_html=True)

# ─── Bouton Lancer l'analyse ──────────────────────────────────────────────────
st.write("")
lancer = st.button(
    f"🚀 Lancer l'analyse — {lot_label}",
    type="primary",
    use_container_width=True,
    disabled=not fichiers_ent
)

if lancer and fichiers_ent:
    entreprises_ok = [e for e in entreprises_detectees]
    if not entreprises_ok:
        st.error("Aucun fichier valide. Vérifiez les noms de fichiers.")
        st.stop()

    # ── Sauvegarder les fichiers dans un dossier temporaire ───────────────────
    tmp_dir = Path(tempfile.mkdtemp())

    with st.status(f"Analyse en cours — {lot_label}...", expanded=True) as status:

        # 1. Parsing des fichiers entreprises
        st.write("📄 Extraction des postes...")
        from parsers.excel_parser import parser_excel
        from parsers.pdf_parser import parser_pdf_natif
        from parsers.ocr_parser import parser_pdf_scan
        from pipeline.file_detector import _detecter_type_pdf

        postes_par_entreprise = {}
        for ent in entreprises_ok:
            fichier = ent["fichier"]
            chemin = tmp_dir / fichier.name
            with open(chemin, "wb") as fh:
                fh.write(fichier.getbuffer())

            try:
                if ent["ext"] in (".xlsx", ".xls"):
                    postes = parser_excel(str(chemin))
                else:
                    type_pdf = _detecter_type_pdf(chemin)
                    if type_pdf == "pdf_natif":
                        postes = parser_pdf_natif(str(chemin))
                    else:
                        postes = parser_pdf_scan(str(chemin))

                # Forcer le numéro de lot sur tous les postes
                for p in postes:
                    p["lot"] = lot_sel["num"]

                postes_par_entreprise[ent["nom"]] = postes
                st.write(f"  ✓ {ent['nom']} — {len(postes)} poste(s) extrait(s)")

            except Exception as e:
                st.warning(f"  ⚠ {ent['nom']} — Erreur : {e}")
                postes_par_entreprise[ent["nom"]] = []

        # 2. DPGF de référence
        st.write("📋 Chargement de la référence...")
        if fichier_ref:
            chemin_ref = tmp_dir / fichier_ref.name
            with open(chemin_ref, "wb") as fh:
                fh.write(fichier_ref.getbuffer())
            postes_ref = parser_excel(str(chemin_ref))
            for p in postes_ref:
                p["lot"] = lot_sel["num"]
            st.write(f"  ✓ Référence : {len(postes_ref)} poste(s)")
        else:
            # Reconstruire depuis l'entreprise avec le plus de postes
            meilleure = max(postes_par_entreprise,
                            key=lambda k: len(postes_par_entreprise[k]))
            postes_ref = [
                {**p, "prix_unitaire": None, "prix_total": None}
                for p in postes_par_entreprise[meilleure]
            ]
            st.write(f"  ✓ Référence reconstruite depuis {meilleure} — {len(postes_ref)} poste(s)")

        # Propager les quantités depuis les entreprises si manquantes dans la référence
        for p_ref in postes_ref:
            if not p_ref.get("quantite"):
                for nom_ent, postes_ent in postes_par_entreprise.items():
                    for p_ent in postes_ent:
                        if _libelles_proches(p_ref.get("libelle",""), p_ent.get("libelle","")):
                            p_ref["quantite"] = p_ent.get("quantite")
                            break

        # 3. Matching
        st.write("🤖 Matching des postes...")
        if st.session_state.cle_api:
            os.environ["ANTHROPIC_API_KEY"] = st.session_state.cle_api
            # Recharger config avec la nouvelle clé
            import importlib, config
            importlib.reload(config)
            try:
                from matching.semantic_matcher import matcher_toutes_entreprises
                matchings = matcher_toutes_entreprises(postes_ref, postes_par_entreprise)
                st.write(f"  ✓ Matching sémantique Claude API")
            except Exception as e:
                st.warning(f"  ⚠ Matching API échoué ({e}), fallback textuel")
                from matching.direct_matcher import matcher_toutes_entreprises_direct
                matchings = matcher_toutes_entreprises_direct(postes_ref, postes_par_entreprise)
        else:
            from matching.direct_matcher import matcher_toutes_entreprises_direct
            matchings = matcher_toutes_entreprises_direct(postes_ref, postes_par_entreprise)
            st.write(f"  ✓ Matching textuel (sans clé API)")

        # 4. Export Excel
        st.write("📊 Génération du fichier Excel...")
        from output.excel_exporter import generer_excel_comparatif

        nom_fichier = f"DPGF_LOT{lot_sel['num']}_{lot_sel['nom'].replace(' ','_')}.xlsx"
        chemin_out  = tmp_dir / nom_fichier

        generer_excel_comparatif(
            postes_ref, matchings, str(chemin_out),
            f"LOT {lot_sel['num']} — {lot_sel['nom']}"
        )

        status.update(label="✅ Analyse terminée !", state="complete")

    # ── Métriques de résultat ─────────────────────────────────────────────────
    nb_ent     = len([e for e in postes_par_entreprise if postes_par_entreprise[e]])
    nb_postes  = len(postes_ref)
    nb_absents = sum(
        sum(1 for m in ml if m.get("score_confiance") == "ABSENT")
        for ml in matchings.values()
    )
    nb_alertes = 0
    import numpy as np
    for i, poste in enumerate(postes_ref):
        prix = []
        for ml in matchings.values():
            m = next((x for x in ml
                      if x["poste_reference"].get("libelle") == poste.get("libelle")), None)
            if m and m.get("prix_unitaire_retenu"):
                prix.append(m["prix_unitaire_retenu"])
        if len(prix) >= 2:
            moy = sum(prix)/len(prix)
            if moy > 0 and (max(prix)-min(prix))/moy > 0.20:
                nb_alertes += 1

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Entreprises", nb_ent)
    c2.metric("Postes analysés", nb_postes)
    c3.metric("⚠ Alertes écart", nb_alertes)
    c4.metric("Postes absents", nb_absents)

    # ── Aperçu tableau ────────────────────────────────────────────────────────
    st.subheader("Aperçu des prix")
    lignes = []
    for poste in postes_ref[:30]:
        ligne = {
            "N°": poste.get("numero",""),
            "Désignation": poste.get("libelle","")[:55],
            "Unité": poste.get("unite",""),
        }
        for nom_ent, ml in matchings.items():
            m = next((x for x in ml
                      if x["poste_reference"].get("libelle") == poste.get("libelle")), None)
            sc = m.get("score_confiance","?") if m else "?"
            pu = m.get("prix_unitaire_retenu") if m else None
            if sc == "ABSENT":
                ligne[nom_ent] = "ABSENT"
            elif pu is not None:
                ligne[nom_ent] = f"{pu:.2f} €"
            else:
                ligne[nom_ent] = f"({sc})"
        lignes.append(ligne)

    df = pd.DataFrame(lignes)

    def colorier(val):
        if val == "ABSENT": return "background-color: #DDDDDD; color: #666"
        if "FAIBLE" in str(val): return "background-color: #FFF0E0"
        return ""

    # Appliquer style uniquement sur les colonnes entreprises
    cols_ent = [c for c in df.columns if c not in ("N°","Désignation","Unité")]
    st.dataframe(
        df.style.applymap(colorier, subset=cols_ent),
        use_container_width=True, height=420
    )

    # ── Téléchargement ────────────────────────────────────────────────────────
    st.write("")
    with open(chemin_out, "rb") as fh:
        st.download_button(
            label=f"⬇️ Télécharger — {nom_fichier}",
            data=fh.read(),
            file_name=nom_fichier,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

# ─── Utilitaire ───────────────────────────────────────────────────────────────
def _libelles_proches(a: str, b: str) -> bool:
    """Vérifie si deux libellés sont proches pour l'appariement des quantités."""
    mots_a = set(re.sub(r"[^\w]"," ",a.lower()).split())
    mots_b = set(re.sub(r"[^\w]"," ",b.lower()).split())
    if not mots_a or not mots_b: return False
    return len(mots_a & mots_b) / max(len(mots_a), len(mots_b)) > 0.5
