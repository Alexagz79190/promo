import pandas as pd
import streamlit as st
from itertools import product
from datetime import datetime, time as dt_time
import time
from io import BytesIO

st.set_page_config(page_title="Outils Commerciaux", layout="wide")

# --- Fonction pour convertir un DataFrame en fichier Excel (bytes) ---
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# ─────────────────────────────────────────────
# NAVIGATION
# ─────────────────────────────────────────────
st.sidebar.title("Navigation")
page = st.sidebar.radio(
    "Choisir une page",
    ["📦 Calculateur Prix Promo", "📊 Analyse CA par Commercial"],
    label_visibility="collapsed"
)

# ══════════════════════════════════════════════
# PAGE 1 — CALCULATEUR PRIX PROMO (code original)
# ══════════════════════════════════════════════
if page == "📦 Calculateur Prix Promo":

    # --- Initialisation du session_state ---
    if "log" not in st.session_state:
        st.session_state["log"] = []
    if "calcul_done" not in st.session_state:
        st.session_state["calcul_done"] = False

    log_container = st.empty()

    def update_status(message):
        st.session_state["log"].append(f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - {message}")
        st.session_state["log_display"] = "\n".join(st.session_state["log"])
        log_container.text_area("Journal des actions", st.session_state["log_display"], height=200, disabled=True)
        time.sleep(0.1)

    def load_file(file_type):
        return st.file_uploader(f"Charger le fichier {file_type} (format Excel)", type=["xlsx"], key=file_type)

    st.title("Calculateur de Prix Promo")
    st.sidebar.header("Paramètres")

    st.subheader("Chargement des fichiers")
    produit_file = load_file("export produit")
    st.info(
        "Les champs attendus pour le fichier **export produit** sont :\n"
        "- Identifiant produit\n"
        "- Fournisseur : identifiant\n"
        "- Famille : identifiant\n"
        "- Marque : identifiant\n"
        "- Code produit\n"
        "- Prix de vente en cours\n"
        "- Prix d'achat avec option\n"
        "- Prix de revient"
    )
    exclusion_file = load_file("exclusion produit")
    remise_file = load_file("remise")

    st.subheader("Sélection des dates")
    start_date = st.date_input("Date de début", value=datetime.now().date())
    start_time = st.time_input("Heure de début", value=dt_time(0, 0))
    end_date = st.date_input("Date de fin", value=datetime.now().date())
    end_time = st.time_input("Heure de fin", value=dt_time(23, 59))
    start_datetime = datetime.combine(start_date, start_time)
    end_datetime = datetime.combine(end_date, end_time)

    st.subheader("Options de calcul")
    price_option = st.radio("Choisissez les options de calcul :",
                            options=["Prix d'achat avec option", "Prix de revient"])

    if st.button("Démarrer le calcul"):
        try:
            if not (produit_file and exclusion_file and remise_file):
                st.error("Veuillez charger tous les fichiers requis.")
                update_status("Erreur : Fichiers manquants.")
            elif not (start_datetime and end_datetime):
                st.error("Veuillez spécifier les dates et heures de début et de fin.")
                update_status("Erreur : Dates ou heures manquantes.")
            else:
                update_status("Chargement des données produit...")
                data = pd.read_excel(produit_file, sheet_name='Worksheet')
                update_status(f"Nombre de produits chargés : {len(data)}")

                update_status("Chargement des exclusions depuis exclus.xlsx...")
                exclusions_data = pd.ExcelFile(exclusion_file)
                excl_code_agz = exclusions_data.parse('Code AGZ')['Code AGZ'].dropna().astype(str).tolist()
                excl_fournisseur = exclusions_data.parse('Founisseur ')['Identifiant fournisseur seul'].dropna().astype(int).tolist()
                excl_marque = exclusions_data.parse('Marque')['Identifiant marque seul'].dropna().astype(int).tolist()
                excl_fournisseur_famille = exclusions_data.parse('Fournisseur famille')[['Identifiant fournisseur', 'Identifiant famille']]

                all_fournisseurs = excl_fournisseur_famille['Identifiant fournisseur'].unique()
                all_familles = excl_fournisseur_famille['Identifiant famille'].unique()
                all_combinations = list(product(all_fournisseurs, all_familles))
                all_combinations_df = pd.DataFrame(all_combinations, columns=['Identifiant fournisseur', 'Identifiant famille'])

                update_status("Application des exclusions issues du fichier exclus.xlsx...")
                data['Exclusion Reason'] = None
                data.loc[data['Code produit'].astype(str).isin(excl_code_agz), 'Exclusion Reason'] = 'Exclus car présent dans Code AGZ fichier exclus'
                data.loc[data['Fournisseur : identifiant'].isin(excl_fournisseur), 'Exclusion Reason'] = 'Exclus car présent dans Fournisseur fichier exclus'
                data.loc[data['Marque : identifiant'].isin(excl_marque), 'Exclusion Reason'] = 'Exclus car présent dans Marque fichier exclus'

                data_merged = data.merge(
                    all_combinations_df,
                    how='left',
                    left_on=['Fournisseur : identifiant', 'Famille : identifiant'],
                    right_on=['Identifiant fournisseur', 'Identifiant famille'],
                    indicator=True
                )
                data_merged.loc[data_merged['_merge'] == 'both', 'Exclusion Reason'] = 'Exclus car présent dans Fournisseur famille du fichier exclus'

                data_excluded = data_merged[data_merged['Exclusion Reason'].notna()].copy()
                data_processed = data_merged[data_merged['Exclusion Reason'].isna()].copy()
                data_processed = data_processed.drop(columns=['Identifiant fournisseur', 'Identifiant famille', '_merge'])
                data_excluded = data_excluded.drop(columns=['Identifiant fournisseur', 'Identifiant famille', '_merge'])

                update_status(f"Produits exclus via exclus.xlsx : {len(data_excluded)}")
                update_status(f"Produits restants après exclusions : {len(data_processed)}")

                update_status("Chargement des remises...")
                remises = pd.read_excel(remise_file)
                price_column = "Prix d'achat avec option" if price_option == "Prix d'achat avec option" else "Prix de revient"

                update_status("Calcul des prix promo...")
                result = []
                margin_issues = []
                exclusion_reasons_from_calc = []
                for _, row in data_processed.iterrows():
                    prix_vente = row['Prix de vente en cours']
                    prix_base = row[price_column]
                    marge = (prix_vente - prix_base) / prix_vente * 100
                    marge = round(marge, 2)
                    remise_appliquee = 0
                    remise_raison = ""
                    for _, remise_row in remises.iterrows():
                        if remise_row['Marge minimale'] <= marge <= remise_row['Marge maximale']:
                            remise_appliquee = remise_row['Remise'] / 100
                            remise_raison = f"Remise appliquée : {remise_row['Remise']}% (Marge entre {remise_row['Marge minimale']}% et {remise_row['Marge maximale']}%)"
                            break
                    prix_promo = round(prix_vente * (1 - remise_appliquee), 2)
                    prix_base_for_margin = row["Prix d'achat avec option"]
                    taux_marge_promo = round((prix_promo - prix_base_for_margin) / prix_promo * 100, 2)
                    if prix_vente != prix_promo and pd.notna(taux_marge_promo):
                        result.append({
                            'Identifiant produit': row['Identifiant produit'],
                            'Prix promo HT': str(prix_promo).replace('.', ','),
                            'Date de début prix promo': start_datetime.strftime('%d/%m/%Y %H:%M:%S'),
                            'Date de fin prix promo': end_datetime.strftime('%d/%m/%Y %H:%M:%S'),
                            'Taux marge prix promo': str(taux_marge_promo).replace('.', ',')
                        })
                        if taux_marge_promo < 5 or taux_marge_promo > 80:
                            margin_issues.append({
                                'Code produit': row['Code produit'],
                                'Prix de vente en cours': prix_vente,
                                "Prix d'achat avec option": row["Prix d'achat avec option"],
                                'Prix de revient': row['Prix de revient'],
                                'Prix promo calculé': prix_promo
                            })
                    else:
                        exclusion_reasons_from_calc.append({
                            'Code produit': row['Code produit'],
                            'Raison exclusion': 'Exclus car le prix promo est supérieur ou égal au prix de vente',
                            'Prix de vente en cours': prix_vente,
                            "Prix d'achat avec option": row["Prix d'achat avec option"],
                            'Prix de revient': row['Prix de revient'],
                            'Remise appliquée': remise_appliquee * 100,
                            'Raison de la remise': remise_raison
                        })

                if not data_excluded.empty:
                    excluded_from_exclus = data_excluded[['Code produit', 'Prix de vente en cours',
                                                          "Prix d'achat avec option", "Prix de revient", "Exclusion Reason"]].copy()
                    excluded_from_exclus.rename(columns={"Exclusion Reason": "Raison exclusion"}, inplace=True)
                    excluded_from_exclus["Remise appliquée"] = ""
                    excluded_from_exclus["Raison de la remise"] = ""
                else:
                    excluded_from_exclus = pd.DataFrame(columns=["Code produit", "Raison exclusion", "Prix de vente en cours",
                                                                 "Prix d'achat avec option", "Prix de revient", "Remise appliquée", "Raison de la remise"])

                exclusion_from_calc_df = pd.DataFrame(exclusion_reasons_from_calc)
                exclusion_final_df = pd.concat([excluded_from_exclus, exclusion_from_calc_df], ignore_index=True)

                st.session_state["result_df"] = pd.DataFrame(result)
                st.session_state["margin_issues_df"] = pd.DataFrame(margin_issues)
                st.session_state["exclusion_reasons_df"] = exclusion_final_df
                st.session_state["calcul_done"] = True

                update_status("Calcul terminé. Les fichiers de résultats sont prêts au téléchargement.")
        except Exception as e:
            st.error(f"Une erreur est survenue : {e}")
            update_status(f"Erreur : {e}")

    if st.session_state.get("calcul_done"):
        st.download_button("Télécharger les résultats",
                           data=st.session_state["result_df"].to_csv(index=False, sep=';', encoding="utf-8"),
                           file_name="prix_promo_output.csv")
        st.download_button("Télécharger les produits avec problèmes de marge",
                           data=to_excel(st.session_state["margin_issues_df"]),
                           file_name="produits_avec_problemes_de_marge.xlsx")
        st.download_button("Télécharger les produits exclus",
                           data=to_excel(st.session_state["exclusion_reasons_df"]),
                           file_name="produits_exclus.xlsx")


# ══════════════════════════════════════════════
# PAGE 2 — ANALYSE CA PAR COMMERCIAL
# ══════════════════════════════════════════════
elif page == "📊 Analyse CA par Commercial":

    st.title("📊 Analyse CA par Commercial")

    csv_file = st.file_uploader("Charger le fichier export commandes (CSV)", type=["csv"], key="ca_csv")

    if csv_file is not None:
        # Chargement du CSV
        df = pd.read_csv(csv_file)

        # Nettoyage : colonnes numériques
        df["Prix produits (HT)"] = pd.to_numeric(df["Prix produits (HT)"], errors="coerce")
        df["Prix final (HT)"]    = pd.to_numeric(df["Prix final (HT)"],    errors="coerce")
        df["taux_marge"]         = pd.to_numeric(df["taux_marge"],         errors="coerce")

        # Remplacement des Auteur vides par une étiquette lisible
        df["Auteur"] = df["Auteur"].fillna("(Sans commercial)").str.strip()
        df["Etat"]   = df["Etat"].fillna("(Inconnu)").str.strip()

        # ── Filtres ──────────────────────────────────
        st.subheader("Filtres")
        col1, col2 = st.columns(2)

        auteurs_dispo = sorted(df["Auteur"].unique().tolist())
        etats_dispo   = sorted(df["Etat"].unique().tolist())

        with col1:
            auteurs_sel = st.multiselect(
                "Auteur(s)",
                options=auteurs_dispo,
                default=auteurs_dispo,
                placeholder="Sélectionner des commerciaux…"
            )
        with col2:
            etats_sel = st.multiselect(
                "État(s)",
                options=etats_dispo,
                default=etats_dispo,
                placeholder="Sélectionner des états…"
            )

        # ── Filtrage ─────────────────────────────────
        df_filtre = df[
            df["Auteur"].isin(auteurs_sel) &
            df["Etat"].isin(etats_sel)
        ].copy()

        st.markdown(f"**{len(df_filtre):,} commandes** correspondent aux filtres sélectionnés.")

        if df_filtre.empty:
            st.warning("Aucune commande ne correspond à la sélection.")
        else:
            # ── Agrégation par Auteur ────────────────
            agg = (
                df_filtre
                .groupby("Auteur", as_index=False)
                .agg(
                    Nb_commandes        = ("Reference",         "count"),
                    CA_produits_HT      = ("Prix produits (HT)", "sum"),
                    CA_final_HT         = ("Prix final (HT)",    "sum"),
                    Taux_marge_moyen    = ("taux_marge",         "mean"),
                )
                .sort_values("CA_final_HT", ascending=False)
            )

            # ── Ligne TOTAL ──────────────────────────
            total = pd.DataFrame([{
                "Auteur":            "**TOTAL**",
                "Nb_commandes":      agg["Nb_commandes"].sum(),
                "CA_produits_HT":    agg["CA_produits_HT"].sum(),
                "CA_final_HT":       agg["CA_final_HT"].sum(),
                "Taux_marge_moyen":  df_filtre["taux_marge"].mean(),
            }])
            agg_display = pd.concat([agg, total], ignore_index=True)

            # ── Mise en forme ────────────────────────
            def fmt_eur(v):
                try:
                    return f"{v:,.2f} €".replace(",", " ").replace(".", ",")
                except Exception:
                    return v

            def fmt_pct(v):
                try:
                    return f"{v:.2f} %"
                except Exception:
                    return v

            agg_display["CA_produits_HT"]   = agg_display["CA_produits_HT"].apply(fmt_eur)
            agg_display["CA_final_HT"]       = agg_display["CA_final_HT"].apply(fmt_eur)
            agg_display["Taux_marge_moyen"]  = agg_display["Taux_marge_moyen"].apply(fmt_pct)

            agg_display = agg_display.rename(columns={
                "Auteur":           "Commercial",
                "Nb_commandes":     "Nb commandes",
                "CA_produits_HT":   "CA Produits HT",
                "CA_final_HT":      "CA Final HT",
                "Taux_marge_moyen": "Taux marge moyen",
            })

            # ── Affichage ────────────────────────────
            st.subheader("Récapitulatif par commercial")
            st.dataframe(agg_display, use_container_width=True, hide_index=True)

            # ── Métriques rapides ────────────────────
            st.subheader("Indicateurs globaux")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Nb commandes",      f"{df_filtre['Reference'].count():,}".replace(",", " "))
            m2.metric("CA Produits HT",    fmt_eur(df_filtre["Prix produits (HT)"].sum()))
            m3.metric("CA Final HT",       fmt_eur(df_filtre["Prix final (HT)"].sum()))
            m4.metric("Taux marge moyen",  fmt_pct(df_filtre["taux_marge"].mean()))

            # ── Export ───────────────────────────────
            # On re-calcule un df propre pour l'export (sans formatage)
            agg_export = (
                df_filtre
                .groupby("Auteur", as_index=False)
                .agg(
                    Nb_commandes        = ("Reference",          "count"),
                    CA_produits_HT      = ("Prix produits (HT)", "sum"),
                    CA_final_HT         = ("Prix final (HT)",    "sum"),
                    Taux_marge_moyen    = ("taux_marge",         "mean"),
                )
                .sort_values("CA_final_HT", ascending=False)
            )
            st.download_button(
                "⬇️ Télécharger le récapitulatif (Excel)",
                data=to_excel(agg_export),
                file_name="ca_par_commercial.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("👆 Chargez un fichier CSV pour démarrer l'analyse.")
        st.markdown(
            "**Colonnes attendues dans le fichier :**  \n"
            "`Reference`, `Auteur`, `Etat`, `Prix produits (HT)`, `Prix final (HT)`, `taux_marge`"
        )
