import pandas as pd
import streamlit as st
from itertools import product
from datetime import datetime, time as dt_time
import time
from io import BytesIO

st.set_page_config(page_title="Outils Commerciaux", layout="wide")

# ─────────────────────────────────────────────
# CSS personnalisé — tableaux & UI
# ─────────────────────────────────────────────
st.markdown("""
<style>
    /* ── Palette & typographie ── */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* ── En-têtes de page ── */
    h1 { font-size: 1.8rem !important; font-weight: 700 !important; color: #1a202c !important; }
    h2 { font-size: 1.3rem !important; font-weight: 600 !important; color: #2d3748 !important; }
    h3 { font-size: 1.1rem !important; font-weight: 600 !important; color: #4a5568 !important; }

    /* ── Tableaux Streamlit (st.dataframe) ── */
    [data-testid="stDataFrame"] table {
        border-collapse: collapse;
        width: 100%;
        font-size: 0.875rem;
    }
    [data-testid="stDataFrame"] thead tr th {
        background-color: #1e3a5f !important;
        color: #ffffff !important;
        font-weight: 600 !important;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        padding: 10px 14px !important;
        border: none !important;
        white-space: nowrap;
    }
    [data-testid="stDataFrame"] tbody tr:nth-child(even) td {
        background-color: #f0f4f8 !important;
    }
    [data-testid="stDataFrame"] tbody tr:last-child td {
        background-color: #dbeafe !important;
        font-weight: 700 !important;
        border-top: 2px solid #1e3a5f !important;
    }
    [data-testid="stDataFrame"] tbody tr:hover td {
        background-color: #e0ecff !important;
    }
    [data-testid="stDataFrame"] tbody td {
        padding: 9px 14px !important;
        border-bottom: 1px solid #e2e8f0 !important;
        color: #1a202c;
    }

    /* ── Métriques ── */
    [data-testid="stMetric"] {
        background: linear-gradient(135deg, #f8fafc 0%, #eef2ff 100%);
        border: 1px solid #c7d2fe;
        border-radius: 10px;
        padding: 14px 18px !important;
    }
    [data-testid="stMetricLabel"] { color: #4338ca !important; font-weight: 600 !important; font-size: 0.78rem !important; text-transform: uppercase; letter-spacing: 0.05em; }
    [data-testid="stMetricValue"] { color: #1e1b4b !important; font-weight: 700 !important; font-size: 1.4rem !important; }

    /* ── Boutons ── */
    .stButton > button, .stDownloadButton > button {
        background-color: #1e3a5f !important;
        color: white !important;
        border: none !important;
        border-radius: 7px !important;
        font-weight: 600 !important;
        padding: 9px 20px !important;
        transition: background 0.2s ease, transform 0.1s ease;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background-color: #2d5282 !important;
        transform: translateY(-1px);
    }

    /* ── Séparateur section ── */
    .section-title {
        border-left: 4px solid #1e3a5f;
        padding-left: 10px;
        margin: 24px 0 12px 0;
        font-size: 1.05rem;
        font-weight: 600;
        color: #1e3a5f;
    }

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1e3a5f 0%, #2d5282 100%) !important;
    }
    [data-testid="stSidebar"] * { color: white !important; }
    [data-testid="stSidebar"] .stRadio label { font-size: 0.9rem !important; }

    /* ── Alertes ── */
    .stInfo    { background-color: #eff6ff !important; border-left: 4px solid #3b82f6 !important; border-radius: 6px; }
    .stSuccess { background-color: #f0fdf4 !important; border-left: 4px solid #22c55e !important; border-radius: 6px; }
    .stWarning { background-color: #fffbeb !important; border-left: 4px solid #f59e0b !important; border-radius: 6px; }
    .stError   { background-color: #fef2f2 !important; border-left: 4px solid #ef4444 !important; border-radius: 6px; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Utilitaires
# ─────────────────────────────────────────────
def to_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()


def normaliser_auteur(nom: str) -> str:
    """
    Normalise un nom d'auteur pour regrouper les variantes d'inversion prénom/nom.
    'Arthur PITAULT', 'Pitault Arthur', 'PITAULT arthur' → clé identique 'arthur pitault'
    """
    if pd.isna(nom) or str(nom).strip() == "":
        return "(sans commercial)"
    mots = str(nom).strip().lower().split()
    return " ".join(sorted(mots))


def formatter_auteur(cle: str) -> str:
    """Capitalise chaque mot de la clé normalisée pour l'affichage."""
    if cle == "(sans commercial)":
        return "(Sans commercial)"
    return " ".join(m.capitalize() for m in cle.split())


# ─────────────────────────────────────────────
# Mapping colonnes CSV produit
# ─────────────────────────────────────────────
COL_CODE        = "Produit - Code / Référence"
COL_PIM_PRODUIT = "Produit - pim_key"
COL_PIM_FAMILLE = "Famille Produit - pim_key"
COL_PIM_MARQUE  = "Marque Produit - pim_key"
COL_PIM_FOURN   = "Fournisseur produit - pim_key"
COL_PRIX_VENTE  = "OffreProduit - Prix de vente HT"
COL_PRIX_ACHAT  = "OffreProduit - Prix d'achat HT"
COL_OFFRE_ID    = "OffreProduit - Id"

# ─────────────────────────────────────────────
# Mapping colonnes CSV commande (format détail)
# ─────────────────────────────────────────────
COL_DETAIL_ACHAT = "Detail de commande - prixAchatHt"
COL_DETAIL_VENTE = "Detail de commande - prixFinalHt"
COL_DETAIL_QTE   = "Detail de commande - Quantité"


# ─────────────────────────────────────────────
# Calcul taux de marge depuis les détails
# ─────────────────────────────────────────────
def compute_taux_marge_from_detail(row) -> float | None:
    """
    Calcule le taux de marge à partir des colonnes de détail.

    - prixFinalHt  : prix de vente de la LIGNE en centimes (quantité incluse)
    - prixAchatHt  : prix d'achat UNITAIRE en centimes → × quantité
    - Quantité     : quantités par ligne
    CA réel = Prix produits (HT) − Remise (HT)
    Coût    = Σ(prixAchatHt × Quantité) / 100
    """
    try:
        achats = [int(x) for x in str(row[COL_DETAIL_ACHAT]).split("|")]
        qtes   = [int(x) for x in str(row[COL_DETAIL_QTE]).split("|")]
        if len(achats) != len(qtes):
            return None
        total_achat_eur = sum(a * q for a, q in zip(achats, qtes)) / 100.0
        prix_produits   = float(row["Prix produits (HT)"])
        remise          = float(row["Remise (HT)"]) if pd.notna(row.get("Remise (HT)")) else 0.0
        ca_reel         = prix_produits - remise
        if ca_reel > 0:
            return round((ca_reel - total_achat_eur) / ca_reel * 100, 2)
        return None
    except Exception:
        return None


def compute_total_achat(row) -> float | None:
    try:
        achats = [int(x) for x in str(row[COL_DETAIL_ACHAT]).split("|")]
        qtes   = [int(x) for x in str(row[COL_DETAIL_QTE]).split("|")]
        return round(sum(a * q for a, q in zip(achats, qtes)) / 100.0, 2)
    except Exception:
        return None


# ─────────────────────────────────────────────
# NAVIGATION
# ─────────────────────────────────────────────
st.sidebar.title("⚙️ Outils Commerciaux")
st.sidebar.markdown("---")
page = st.sidebar.radio(
    "Navigation",
    ["📦 Calculateur Prix Promo", "📊 Analyse CA par Commercial"],
    label_visibility="collapsed"
)


# ══════════════════════════════════════════════
# PAGE 1 — CALCULATEUR PRIX PROMO
# ══════════════════════════════════════════════
if page == "📦 Calculateur Prix Promo":

    if "log" not in st.session_state:
        st.session_state["log"] = []
    if "calcul_done" not in st.session_state:
        st.session_state["calcul_done"] = False

    log_container = st.empty()

    def update_status(message: str):
        st.session_state["log"].append(
            f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')} — {message}"
        )
        log_container.text_area(
            "Journal des actions",
            "\n".join(st.session_state["log"]),
            height=200,
            disabled=True
        )
        time.sleep(0.1)

    st.title("📦 Calculateur de Prix Promo")
    st.sidebar.header("Paramètres")

    st.markdown('<p class="section-title">Chargement des fichiers</p>', unsafe_allow_html=True)

    st.info(
        "**Champs à sélectionner dans l'interface d'export** (l'ordre des colonnes n'est pas important) :\n"
        "- `Code / Référence Produit`\n"
        "- `pim_key Produit`\n"
        "- `pim_key Famille Produit`\n"
        "- `pim_key Marque Produit`\n"
        "- `pim_key Fournisseur produit`\n"
        "- `Prix de vente HT OffreProduit`\n"
        "- `Prix d'achat HT OffreProduit`\n"
        "- `Id OffreProduit`"
    )

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        produit_file   = st.file_uploader("📄 Export produit (CSV)",    type=["csv"],  key="produit_csv")
    with col_f2:
        exclusion_file = st.file_uploader("🚫 Fichier exclusion (Excel)", type=["xlsx"], key="exclusion")
    with col_f3:
        remise_file    = st.file_uploader("💰 Fichier remise (Excel)",    type=["xlsx"], key="remise")

    st.markdown('<p class="section-title">Période promotionnelle</p>', unsafe_allow_html=True)
    col_d1, col_d2, col_d3, col_d4 = st.columns(4)
    with col_d1:
        start_date = st.date_input("Date de début",  value=datetime.now().date(), key="sd")
    with col_d2:
        start_time = st.time_input("Heure de début", value=dt_time(0, 0),        key="st")
    with col_d3:
        end_date   = st.date_input("Date de fin",    value=datetime.now().date(), key="ed")
    with col_d4:
        end_time   = st.time_input("Heure de fin",   value=dt_time(23, 59),      key="et")

    start_datetime = datetime.combine(start_date, start_time)
    end_datetime   = datetime.combine(end_date,   end_time)

    st.markdown("")
    if st.button("🚀 Démarrer le calcul"):
        st.session_state["log"] = []
        st.session_state["calcul_done"] = False

        try:
            if not (produit_file and exclusion_file and remise_file):
                st.error("Veuillez charger tous les fichiers requis.")
                update_status("Erreur : fichiers manquants.")
            elif not (start_datetime and end_datetime):
                st.error("Veuillez spécifier les dates et heures de début et de fin.")
                update_status("Erreur : dates ou heures manquantes.")
            else:
                # ── Chargement produits ───────────────────────────────────────
                update_status("Chargement des données produit (CSV)...")
                data = pd.read_csv(produit_file)
                update_status(f"Lignes chargées : {len(data):,}")

                colonnes_requises = [COL_CODE, COL_PIM_PRODUIT, COL_PIM_FAMILLE,
                                     COL_PIM_MARQUE, COL_PIM_FOURN,
                                     COL_PRIX_VENTE, COL_PRIX_ACHAT, COL_OFFRE_ID]
                colonnes_manquantes = [c for c in colonnes_requises if c not in data.columns]
                if colonnes_manquantes:
                    st.error(f"Colonnes manquantes dans le fichier produit : {colonnes_manquantes}")
                    update_status(f"Erreur : colonnes manquantes : {colonnes_manquantes}")
                    st.stop()

                # ── Éclatement multi-offres ───────────────────────────────────
                cols_a_eclater = [COL_PRIX_VENTE, COL_PRIX_ACHAT, COL_OFFRE_ID]
                for col in cols_a_eclater:
                    data[col] = data[col].astype(str).str.split('|')
                avant = len(data)
                data = data.explode(cols_a_eclater).reset_index(drop=True)
                if len(data) > avant:
                    update_status(f"Éclatement multi-offres : {avant:,} → {len(data):,} lignes.")
                update_status(f"Produits / offres à traiter : {len(data):,}")

                data[COL_PRIX_VENTE] = pd.to_numeric(
                    data[COL_PRIX_VENTE].astype(str).str.replace(",", "."), errors="coerce")
                data[COL_PRIX_ACHAT] = pd.to_numeric(
                    data[COL_PRIX_ACHAT].astype(str).str.replace(",", "."), errors="coerce")

                before = len(data)
                data[COL_OFFRE_ID] = data[COL_OFFRE_ID].replace("nan", pd.NA)
                data = data.dropna(subset=[COL_PRIX_VENTE, COL_PRIX_ACHAT, COL_OFFRE_ID])
                if (ignores := before - len(data)) > 0:
                    update_status(f"{ignores:,} ligne(s) ignorée(s) : offre sans prix ou ID.")

                # ── Exclusions ────────────────────────────────────────────────
                update_status("Chargement des exclusions...")
                exclusions_data = pd.ExcelFile(exclusion_file)

                excl_code_agz    = exclusions_data.parse('Code AGZ')['Code AGZ'].dropna().astype(str).tolist()
                excl_fournisseur = exclusions_data.parse('Founisseur ')['Identifiant fournisseur seul'].dropna().astype(str).tolist()
                excl_marque      = exclusions_data.parse('Marque')['Identifiant marque seul'].dropna().astype(str).tolist()
                excl_ff          = exclusions_data.parse('Fournisseur famille')[
                    ['Identifiant fournisseur', 'Identifiant famille']
                ].astype(str)

                all_fournisseurs    = excl_ff['Identifiant fournisseur'].unique()
                all_familles        = excl_ff['Identifiant famille'].unique()
                all_combinations_df = pd.DataFrame(
                    list(product(all_fournisseurs, all_familles)),
                    columns=['Identifiant fournisseur', 'Identifiant famille']
                )

                update_status("Application des exclusions...")
                for col in [COL_PIM_PRODUIT, COL_PIM_FOURN, COL_PIM_MARQUE, COL_PIM_FAMILLE]:
                    data[col] = data[col].astype(str)

                data['Exclusion Reason'] = None
                data.loc[data[COL_CODE].astype(str).isin(excl_code_agz),
                         'Exclusion Reason'] = 'Exclus — Code AGZ'
                data.loc[data[COL_PIM_FOURN].isin(excl_fournisseur),
                         'Exclusion Reason'] = 'Exclus — Fournisseur'
                data.loc[data[COL_PIM_MARQUE].isin(excl_marque),
                         'Exclusion Reason'] = 'Exclus — Marque'

                data_merged = data.merge(
                    all_combinations_df, how='left',
                    left_on=[COL_PIM_FOURN, COL_PIM_FAMILLE],
                    right_on=['Identifiant fournisseur', 'Identifiant famille'],
                    indicator=True
                )
                data_merged.loc[data_merged['_merge'] == 'both', 'Exclusion Reason'] = (
                    'Exclus — Fournisseur × Famille'
                )

                data_excluded  = data_merged[data_merged['Exclusion Reason'].notna()].copy()
                data_processed = data_merged[data_merged['Exclusion Reason'].isna()].copy()
                for df_ in [data_processed, data_excluded]:
                    df_.drop(columns=['Identifiant fournisseur', 'Identifiant famille', '_merge'],
                             inplace=True)

                update_status(f"Produits exclus : {len(data_excluded):,}")
                update_status(f"Produits à traiter : {len(data_processed):,}")

                # ── Remises ───────────────────────────────────────────────────
                update_status("Chargement des remises...")
                remises = pd.read_excel(remise_file)

                # ── Calcul des prix promo ─────────────────────────────────────
                update_status("Calcul des prix promo...")
                result, margin_issues, exclusion_reasons_from_calc = [], [], []

                for _, row in data_processed.iterrows():
                    pv = row[COL_PRIX_VENTE]
                    pa = row[COL_PRIX_ACHAT]
                    if pd.isna(pv) or pd.isna(pa) or pv <= 0:
                        continue

                    marge = round((pv - pa) / pv * 100, 2)
                    remise_appliquee, remise_raison = 0, ""
                    for _, r in remises.iterrows():
                        if r['Marge minimale'] <= marge <= r['Marge maximale']:
                            remise_appliquee = r['Remise'] / 100
                            remise_raison    = (
                                f"Remise {r['Remise']}% "
                                f"(marge entre {r['Marge minimale']}% et {r['Marge maximale']}%)"
                            )
                            break

                    prix_promo       = round(pv * (1 - remise_appliquee), 2)
                    prix_promo_cents = int(round(prix_promo * 100))
                    taux_marge_promo = round((prix_promo - pa) / prix_promo * 100, 2) if prix_promo > 0 else None

                    if pv != prix_promo and pd.notna(taux_marge_promo):
                        result.append({
                            'Offre produit (cocher EST identifiant)': row[COL_OFFRE_ID],
                            'Type':   'promo',
                            'Prix':   prix_promo_cents,
                            "Date d'application":               start_datetime.strftime('%Y-%m-%d %H:%M:%S'),
                            'Date fin (pour promo uniquement)': end_datetime.strftime('%Y-%m-%d %H:%M:%S'),
                            'Prix (ne pas importer)':           f"{prix_promo:.2f}",
                        })
                        if taux_marge_promo < 5 or taux_marge_promo > 80:
                            margin_issues.append({
                                COL_CODE:                        row[COL_CODE],
                                COL_OFFRE_ID:                    row[COL_OFFRE_ID],
                                'Prix de vente HT':              pv,
                                "Prix d'achat HT":               pa,
                                'Prix promo calculé (HT)':       prix_promo,
                                'Prix promo calculé (centimes)': prix_promo_cents,
                                'Taux marge promo':              taux_marge_promo,
                            })
                    else:
                        exclusion_reasons_from_calc.append({
                            COL_CODE:               row[COL_CODE],
                            COL_OFFRE_ID:           row[COL_OFFRE_ID],
                            'Raison exclusion':     'Prix promo ≥ prix de vente',
                            'Prix de vente HT':     pv,
                            "Prix d'achat HT":      pa,
                            'Remise appliquée (%)': remise_appliquee * 100,
                            'Raison de la remise':  remise_raison,
                        })

                # ── Construction fichier exclus final ─────────────────────────
                if not data_excluded.empty:
                    excluded_from_exclus = data_excluded[[
                        COL_CODE, COL_OFFRE_ID, COL_PRIX_VENTE, COL_PRIX_ACHAT, 'Exclusion Reason'
                    ]].copy()
                    excluded_from_exclus.rename(columns={
                        COL_PRIX_VENTE:     'Prix de vente HT',
                        COL_PRIX_ACHAT:     "Prix d'achat HT",
                        'Exclusion Reason': 'Raison exclusion'
                    }, inplace=True)
                    excluded_from_exclus['Remise appliquée (%)'] = ""
                    excluded_from_exclus['Raison de la remise']  = ""
                else:
                    excluded_from_exclus = pd.DataFrame(columns=[
                        COL_CODE, COL_OFFRE_ID, 'Prix de vente HT', "Prix d'achat HT",
                        'Raison exclusion', 'Remise appliquée (%)', 'Raison de la remise'
                    ])

                exclusion_final_df = pd.concat(
                    [excluded_from_exclus, pd.DataFrame(exclusion_reasons_from_calc)],
                    ignore_index=True
                )

                st.session_state["result_df"]            = pd.DataFrame(result)
                st.session_state["margin_issues_df"]     = pd.DataFrame(margin_issues)
                st.session_state["exclusion_reasons_df"] = exclusion_final_df
                st.session_state["calcul_done"]          = True
                update_status(f"✅ Calcul terminé — {len(result):,} offres promo générées.")

        except Exception as e:
            st.error(f"Une erreur est survenue : {e}")
            update_status(f"Erreur : {e}")

    if st.session_state.get("calcul_done"):
        st.success(f"✅ **{len(st.session_state['result_df']):,} offres promo** prêtes à l'export.")
        st.markdown('<p class="section-title">Téléchargements</p>', unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button(
                "⬇️ Résultats (CSV)",
                data=st.session_state["result_df"].to_csv(index=False, sep=';', encoding="utf-8"),
                file_name="prix_promo_output.csv",
                mime="text/csv"
            )
        with col2:
            st.download_button(
                "⬇️ Problèmes de marge (Excel)",
                data=to_excel(st.session_state["margin_issues_df"]),
                file_name="produits_problemes_marge.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col3:
            st.download_button(
                "⬇️ Produits exclus (Excel)",
                data=to_excel(st.session_state["exclusion_reasons_df"]),
                file_name="produits_exclus.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# ══════════════════════════════════════════════
# PAGE 2 — ANALYSE CA PAR COMMERCIAL
# ══════════════════════════════════════════════
elif page == "📊 Analyse CA par Commercial":

    st.title("📊 Analyse CA par Commercial")

    with st.expander("ℹ️ Formats de fichier acceptés", expanded=False):
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown(
                "**Format A — colonne `taux_marge` pré-calculée**\n\n"
                "- `Commande - Reference`\n"
                "- `Commande - Auteur`\n"
                "- `Commande - Etat`\n"
                "- `Commande - Prix produits (HT)`\n"
                "- `Commande - Prix final (HT)`\n"
                "- `Commande - taux_marge`"
            )
        with col_b:
            st.markdown(
                "**Format B — détails de commande**\n\n"
                "- `Commande - Reference`\n"
                "- `Commande - Auteur`\n"
                "- `Commande - Etat`\n"
                "- `Commande - Prix produits (HT)`\n"
                "- `Commande - Prix final (HT)`\n"
                "- `Commande - Remise (HT)`\n"
                "- `Detail de commande - prixAchatHt` *(centimes, séparés par |)*\n"
                "- `Detail de commande - prixFinalHt` *(centimes, séparés par |)*\n"
                "- `Detail de commande - Quantité` *(séparés par |)*"
            )

    csv_file = st.file_uploader("📄 Charger le fichier export commandes (CSV)", type=["csv"], key="ca_csv")

    if csv_file is not None:
        df = pd.read_csv(csv_file)
        df.columns = [c.replace("Commande - ", "").strip() for c in df.columns]

        df["Prix produits (HT)"] = pd.to_numeric(df["Prix produits (HT)"], errors="coerce")
        df["Prix final (HT)"]    = pd.to_numeric(df["Prix final (HT)"],    errors="coerce")
        df["Remise (HT)"]        = pd.to_numeric(df.get("Remise (HT)"),    errors="coerce").fillna(0)

        col_detail_achat_s = COL_DETAIL_ACHAT.replace("Commande - ", "").strip()
        col_detail_vente_s = COL_DETAIL_VENTE.replace("Commande - ", "").strip()
        col_detail_qte_s   = COL_DETAIL_QTE.replace("Commande - ", "").strip()

        has_detail_cols = all(c in df.columns for c in [col_detail_achat_s, col_detail_vente_s, col_detail_qte_s])
        has_taux_marge  = "taux_marge" in df.columns

        if has_detail_cols:
            st.info("📋 **Format B détecté** — taux de marge calculé depuis les détails de commande.")
            df = df.rename(columns={
                col_detail_achat_s: COL_DETAIL_ACHAT,
                col_detail_vente_s: COL_DETAIL_VENTE,
                col_detail_qte_s:   COL_DETAIL_QTE,
            })
            df["taux_marge"]     = df.apply(compute_taux_marge_from_detail, axis=1)
            df["total_achat_HT"] = df.apply(compute_total_achat, axis=1)

            if (nb_sans_marge := df["taux_marge"].isna().sum()) > 0:
                st.warning(f"⚠️ {nb_sans_marge:,} commande(s) sans taux de marge calculable.")

        elif has_taux_marge:
            st.info("📋 **Format A détecté** — taux de marge lu depuis la colonne `taux_marge`.")
            df["taux_marge"]     = pd.to_numeric(df["taux_marge"], errors="coerce")
            df["total_achat_HT"] = None

        else:
            st.error(
                "❌ Format non reconnu. Le fichier doit contenir soit `taux_marge`, "
                "soit les trois colonnes de détail."
            )
            st.stop()

        # ── Normalisation des auteurs ─────────────────────────────────────────
        # Regroupe les variantes d'inversion prénom/nom (ex: "Arthur PITAULT" = "Pitault Arthur")
        df["Auteur"] = (
            df["Auteur"]
            .fillna("(Sans commercial)")
            .apply(normaliser_auteur)
            .apply(formatter_auteur)
        )
        df["Etat"] = df["Etat"].fillna("(Inconnu)").str.strip()

        # ── Filtres ───────────────────────────────────────────────────────────
        st.markdown('<p class="section-title">Filtres</p>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)

        auteurs_dispo = sorted(df["Auteur"].unique().tolist())
        etats_dispo   = sorted(df["Etat"].unique().tolist())

        with col1:
            auteurs_sel = st.multiselect(
                "👤 Commercial(aux)", options=auteurs_dispo, default=[],
                placeholder="Tous les commerciaux…"
            )
        with col2:
            etats_preselectes = [e for e in ["en_preparation", "expedie", "valide"] if e in etats_dispo]
            etats_sel = st.multiselect(
                "📌 État(s)", options=etats_dispo, default=etats_preselectes,
                placeholder="Tous les états…"
            )

        masque_auteur = df["Auteur"].isin(auteurs_sel) if auteurs_sel else pd.Series([True] * len(df), index=df.index)
        masque_etat   = df["Etat"].isin(etats_sel)    if etats_sel   else pd.Series([True] * len(df), index=df.index)
        df_filtre = df[masque_auteur & masque_etat].copy()

        st.markdown(f"**{len(df_filtre):,} commandes** correspondent aux filtres sélectionnés.")

        if df_filtre.empty:
            st.warning("Aucune commande ne correspond à la sélection.")
        else:
            # ── Calcul des valeurs ────────────────────────────────────────────
            if has_detail_cols:
                df_filtre["ca_reel"]      = df_filtre["Prix produits (HT)"] - df_filtre["Remise (HT)"].fillna(0)
                df_filtre["valeur_achat"] = df_filtre["total_achat_HT"]
                df_filtre["valeur_marge"] = df_filtre["ca_reel"] - df_filtre["total_achat_HT"]
            else:
                df_filtre["ca_reel"]      = df_filtre["Prix produits (HT)"]
                df_filtre["valeur_marge"] = df_filtre["Prix produits (HT)"] * df_filtre["taux_marge"] / 100
                df_filtre["valeur_achat"] = df_filtre["Prix produits (HT)"] - df_filtre["valeur_marge"]

            # ── Agrégation par commercial ─────────────────────────────────────
            agg = (
                df_filtre
                .groupby("Auteur", as_index=False)
                .agg(
                    Nb_commandes   =("Reference",          "count"),
                    CA_produits_HT =("Prix produits (HT)", "sum"),
                    CA_final_HT    =("Prix final (HT)",    "sum"),
                    _val_marge     =("valeur_marge",       "sum"),
                    _val_achat     =("valeur_achat",       "sum"),
                    Taux_marge_moy =("taux_marge",         "mean"),
                )
                .sort_values("CA_final_HT", ascending=False)
            )
            agg["Taux_marge_pondere"] = agg["_val_marge"] / agg["CA_produits_HT"] * 100
            agg.drop(columns=["_val_marge", "_val_achat"], inplace=True)

            # ── Ligne TOTAL ───────────────────────────────────────────────────
            total_ca_ht     = df_filtre["ca_reel"].sum()
            total_val_marge = df_filtre["valeur_marge"].sum()
            total = pd.DataFrame([{
                "Auteur":             "TOTAL",
                "Nb_commandes":       int(agg["Nb_commandes"].sum()),
                "CA_produits_HT":     agg["CA_produits_HT"].sum(),
                "CA_final_HT":        agg["CA_final_HT"].sum(),
                "Taux_marge_moy":     df_filtre["taux_marge"].mean(),
                "Taux_marge_pondere": total_val_marge / total_ca_ht * 100 if total_ca_ht else 0,
            }])
            agg_display = pd.concat([agg, total], ignore_index=True)

            # ── Formatage pour affichage ──────────────────────────────────────
            def fmt_eur(v):
                try:    return f"{v:,.2f} €".replace(",", " ").replace(".", ",")
                except: return v

            def fmt_pct(v):
                try:    return f"{v:.2f} %"
                except: return v

            def fmt_int(v):
                try:    return f"{int(v):,}".replace(",", " ")
                except: return v

            agg_display["CA_produits_HT"]     = agg_display["CA_produits_HT"].apply(fmt_eur)
            agg_display["CA_final_HT"]        = agg_display["CA_final_HT"].apply(fmt_eur)
            agg_display["Taux_marge_moy"]     = agg_display["Taux_marge_moy"].apply(fmt_pct)
            agg_display["Taux_marge_pondere"] = agg_display["Taux_marge_pondere"].apply(fmt_pct)
            agg_display["Nb_commandes"]       = agg_display["Nb_commandes"].apply(fmt_int)

            agg_display = agg_display.rename(columns={
                "Auteur":             "Commercial",
                "Nb_commandes":       "Nb commandes",
                "CA_produits_HT":     "CA Produits HT",
                "CA_final_HT":        "CA Final HT",
                "Taux_marge_moy":     "Taux marge moyen",
                "Taux_marge_pondere": "Taux marge pondéré",
            })

            st.markdown('<p class="section-title">Récapitulatif par commercial</p>', unsafe_allow_html=True)
            st.dataframe(agg_display, use_container_width=True, hide_index=True)

            # ── Indicateurs globaux ───────────────────────────────────────────
            st.markdown('<p class="section-title">Indicateurs globaux</p>', unsafe_allow_html=True)
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Nb commandes",       f"{df_filtre['Reference'].count():,}".replace(",", " "))
            m2.metric("CA Produits HT",     fmt_eur(total_ca_ht))
            m3.metric("CA Final HT",        fmt_eur(df_filtre["Prix final (HT)"].sum()))
            m4.metric("Taux marge moyen",   fmt_pct(df_filtre["taux_marge"].mean()))
            m5.metric("Taux marge pondéré", fmt_pct(total_val_marge / total_ca_ht * 100 if total_ca_ht else 0))

            # ── Exports ───────────────────────────────────────────────────────
            st.markdown('<p class="section-title">Téléchargements</p>', unsafe_allow_html=True)
            agg_export = agg.copy()
            col_dl1, col_dl2 = st.columns(2)

            with col_dl1:
                st.download_button(
                    "⬇️ Récapitulatif par commercial (Excel)",
                    data=to_excel(agg_export),
                    file_name="ca_par_commercial.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col_dl2:
                detail_cols = ["Reference", "Auteur", "Etat",
                               "Prix produits (HT)", "Prix final (HT)",
                               "taux_marge", "valeur_marge"]
                if has_detail_cols:
                    detail_cols += ["total_achat_HT", COL_DETAIL_ACHAT, COL_DETAIL_VENTE, COL_DETAIL_QTE]
                detail_export = df_filtre[[c for c in detail_cols if c in df_filtre.columns]].copy()

                st.download_button(
                    "⬇️ Détail des commandes (Excel)",
                    data=to_excel(detail_export),
                    file_name="detail_commandes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    else:
        st.info("👆 Chargez un fichier CSV pour démarrer l'analyse.")
