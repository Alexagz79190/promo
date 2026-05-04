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
# Mapping des colonnes du CSV produit
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
# NAVIGATION
# ─────────────────────────────────────────────
st.sidebar.title("Navigation")
page = st.sidebar.radio(
    "Choisir une page",
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

    def update_status(message):
        st.session_state["log"].append(f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - {message}")
        st.session_state["log_display"] = "\n".join(st.session_state["log"])
        log_container.text_area("Journal des actions", st.session_state["log_display"], height=200, disabled=True)
        time.sleep(0.1)

    st.title("Calculateur de Prix Promo")
    st.sidebar.header("Paramètres")

    st.subheader("Chargement des fichiers")

    produit_file = st.file_uploader("Charger le fichier export produit (format CSV)", type=["csv"], key="produit_csv")
    st.info(
        "Les champs attendus pour le fichier **export produit** (CSV) — l'ordre des colonnes n'est pas important :\n"
        f"- `{COL_CODE}`\n"
        f"- `{COL_PIM_PRODUIT}`\n"
        f"- `{COL_PIM_FAMILLE}`\n"
        f"- `{COL_PIM_MARQUE}`\n"
        f"- `{COL_PIM_FOURN}`\n"
        f"- `{COL_PRIX_VENTE}`\n"
        f"- `{COL_PRIX_ACHAT}`\n"
        f"- `{COL_OFFRE_ID}`"
    )

    exclusion_file = st.file_uploader("Charger le fichier exclusion produit (format Excel)", type=["xlsx"], key="exclusion")
    remise_file    = st.file_uploader("Charger le fichier remise (format Excel)",            type=["xlsx"], key="remise")

    st.subheader("Sélection des dates")
    start_date = st.date_input("Date de début",  value=datetime.now().date(), key="sd")
    start_time = st.time_input("Heure de début", value=dt_time(0, 0),        key="st")
    end_date   = st.date_input("Date de fin",    value=datetime.now().date(), key="ed")
    end_time   = st.time_input("Heure de fin",   value=dt_time(23, 59),      key="et")
    start_datetime = datetime.combine(start_date, start_time)
    end_datetime   = datetime.combine(end_date,   end_time)

    if st.button("Démarrer le calcul"):
        st.session_state["log"] = []
        st.session_state["calcul_done"] = False

        try:
            if not (produit_file and exclusion_file and remise_file):
                st.error("Veuillez charger tous les fichiers requis.")
                update_status("Erreur : Fichiers manquants.")
            elif not (start_datetime and end_datetime):
                st.error("Veuillez spécifier les dates et heures de début et de fin.")
                update_status("Erreur : Dates ou heures manquantes.")
            else:
                # ── Chargement produits ───────────────────────────────────────
                update_status("Chargement des données produit (CSV)...")
                data = pd.read_csv(produit_file)
                update_status(f"Nombre de lignes chargées : {len(data)}")

                # Vérification que toutes les colonnes attendues sont présentes
                colonnes_requises = [COL_CODE, COL_PIM_PRODUIT, COL_PIM_FAMILLE,
                                     COL_PIM_MARQUE, COL_PIM_FOURN,
                                     COL_PRIX_VENTE, COL_PRIX_ACHAT, COL_OFFRE_ID]
                colonnes_manquantes = [c for c in colonnes_requises if c not in data.columns]
                if colonnes_manquantes:
                    st.error(f"Colonnes manquantes dans le fichier produit : {colonnes_manquantes}")
                    update_status(f"Erreur : colonnes manquantes : {colonnes_manquantes}")
                    st.stop()

                # ── Dédoublonnage multi-offres (valeurs séparées par |) ───────
                # Certains produits ont plusieurs offres sur une seule ligne,
                # ex : Prix de vente = "1.96|1.93", Id = "uuid1|uuid2"
                # On éclate chaque colonne multi-valeurs en autant de lignes.
                cols_a_eclater = [COL_PRIX_VENTE, COL_PRIX_ACHAT, COL_OFFRE_ID]
                for col in cols_a_eclater:
                    data[col] = data[col].astype(str).str.split('|')
                avant_eclatement = len(data)
                data = data.explode(cols_a_eclater).reset_index(drop=True)
                if len(data) > avant_eclatement:
                    update_status(
                        f"Dédoublonnage multi-offres : {avant_eclatement} ligne(s) → "
                        f"{len(data)} ligne(s) après éclatement."
                    )
                update_status(f"Nombre de produits/offres à traiter : {len(data)}")

                # Nettoyage colonnes numériques (gestion virgule/point)
                data[COL_PRIX_VENTE] = pd.to_numeric(
                    data[COL_PRIX_VENTE].astype(str).str.replace(",", "."), errors="coerce")
                data[COL_PRIX_ACHAT] = pd.to_numeric(
                    data[COL_PRIX_ACHAT].astype(str).str.replace(",", "."), errors="coerce")

                # Suppression des lignes sans prix valides
                before = len(data)
                data = data.dropna(subset=[COL_PRIX_VENTE, COL_PRIX_ACHAT])
                if len(data) < before:
                    update_status(f"{before - len(data)} ligne(s) ignorée(s) : prix manquants ou non numériques.")

                # ── Chargement exclusions ─────────────────────────────────────
                update_status("Chargement des exclusions...")
                exclusions_data = pd.ExcelFile(exclusion_file)

                excl_code_agz = (exclusions_data.parse('Code AGZ')['Code AGZ']
                                 .dropna().astype(str).tolist())
                excl_fournisseur = (exclusions_data.parse('Founisseur ')['Identifiant fournisseur seul']
                                    .dropna().astype(str).tolist())
                excl_marque = (exclusions_data.parse('Marque')['Identifiant marque seul']
                               .dropna().astype(str).tolist())
                excl_ff = exclusions_data.parse('Fournisseur famille')[
                    ['Identifiant fournisseur', 'Identifiant famille']
                ].astype(str)

                all_fournisseurs    = excl_ff['Identifiant fournisseur'].unique()
                all_familles        = excl_ff['Identifiant famille'].unique()
                all_combinations_df = pd.DataFrame(
                    list(product(all_fournisseurs, all_familles)),
                    columns=['Identifiant fournisseur', 'Identifiant famille']
                )

                # ── Application des exclusions ────────────────────────────────
                update_status("Application des exclusions...")

                data[COL_PIM_PRODUIT] = data[COL_PIM_PRODUIT].astype(str)
                data[COL_PIM_FOURN]   = data[COL_PIM_FOURN].astype(str)
                data[COL_PIM_MARQUE]  = data[COL_PIM_MARQUE].astype(str)
                data[COL_PIM_FAMILLE] = data[COL_PIM_FAMILLE].astype(str)

                data['Exclusion Reason'] = None
                data.loc[data[COL_CODE].astype(str).isin(excl_code_agz),
                         'Exclusion Reason'] = 'Exclus car présent dans Code AGZ fichier exclus'
                data.loc[data[COL_PIM_FOURN].isin(excl_fournisseur),
                         'Exclusion Reason'] = 'Exclus car présent dans Fournisseur fichier exclus'
                data.loc[data[COL_PIM_MARQUE].isin(excl_marque),
                         'Exclusion Reason'] = 'Exclus car présent dans Marque fichier exclus'

                data_merged = data.merge(
                    all_combinations_df,
                    how='left',
                    left_on=[COL_PIM_FOURN, COL_PIM_FAMILLE],
                    right_on=['Identifiant fournisseur', 'Identifiant famille'],
                    indicator=True
                )
                data_merged.loc[data_merged['_merge'] == 'both', 'Exclusion Reason'] = (
                    'Exclus car présent dans Fournisseur famille du fichier exclus'
                )

                data_excluded  = data_merged[data_merged['Exclusion Reason'].notna()].copy()
                data_processed = data_merged[data_merged['Exclusion Reason'].isna()].copy()
                data_processed = data_processed.drop(columns=['Identifiant fournisseur', 'Identifiant famille', '_merge'])
                data_excluded  = data_excluded.drop(columns=['Identifiant fournisseur', 'Identifiant famille', '_merge'])

                update_status(f"Produits exclus via fichier exclus : {len(data_excluded)}")
                update_status(f"Produits restants après exclusions : {len(data_processed)}")

                # ── Chargement remises ────────────────────────────────────────
                update_status("Chargement des remises...")
                remises = pd.read_excel(remise_file)

                # ── Calcul des prix promo ─────────────────────────────────────
                update_status("Calcul des prix promo...")
                result                      = []
                margin_issues               = []
                exclusion_reasons_from_calc = []

                for _, row in data_processed.iterrows():
                    prix_vente = row[COL_PRIX_VENTE]
                    prix_achat = row[COL_PRIX_ACHAT]

                    if pd.isna(prix_vente) or pd.isna(prix_achat) or prix_vente <= 0:
                        continue

                    marge = round((prix_vente - prix_achat) / prix_vente * 100, 2)

                    remise_appliquee = 0
                    remise_raison    = ""
                    for _, remise_row in remises.iterrows():
                        if remise_row['Marge minimale'] <= marge <= remise_row['Marge maximale']:
                            remise_appliquee = remise_row['Remise'] / 100
                            remise_raison    = (
                                f"Remise appliquée : {remise_row['Remise']}% "
                                f"(Marge entre {remise_row['Marge minimale']}% et {remise_row['Marge maximale']}%)"
                            )
                            break

                    prix_promo       = round(prix_vente * (1 - remise_appliquee), 2)
                    prix_promo_cents = int(round(prix_promo * 100))
                    taux_marge_promo = round((prix_promo - prix_achat) / prix_promo * 100, 2) if prix_promo > 0 else None

                    if prix_vente != prix_promo and pd.notna(taux_marge_promo):
                        result.append({
                            COL_OFFRE_ID:            row[COL_OFFRE_ID],
                            'Type de prix':          'promo',
                            'Prix promo (centimes)': prix_promo_cents,
                            'Date de début':         start_datetime.strftime('%d/%m/%Y %H:%M:%S'),
                            'Date de fin':           end_datetime.strftime('%d/%m/%Y %H:%M:%S'),
                        })
                        if taux_marge_promo < 5 or taux_marge_promo > 80:
                            margin_issues.append({
                                COL_CODE:                        row[COL_CODE],
                                COL_OFFRE_ID:                    row[COL_OFFRE_ID],
                                'Prix de vente HT':              prix_vente,
                                "Prix d'achat HT":               prix_achat,
                                'Prix promo calculé (HT)':       prix_promo,
                                'Prix promo calculé (centimes)': prix_promo_cents,
                                'Taux marge promo':              taux_marge_promo,
                            })
                    else:
                        exclusion_reasons_from_calc.append({
                            COL_CODE:               row[COL_CODE],
                            COL_OFFRE_ID:           row[COL_OFFRE_ID],
                            'Raison exclusion':     'Exclus car le prix promo est supérieur ou égal au prix de vente',
                            'Prix de vente HT':     prix_vente,
                            "Prix d'achat HT":      prix_achat,
                            'Remise appliquée (%)': remise_appliquee * 100,
                            'Raison de la remise':  remise_raison,
                        })

                # ── Construction du fichier exclus final ──────────────────────
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
                        COL_CODE, COL_OFFRE_ID,
                        'Prix de vente HT', "Prix d'achat HT",
                        'Raison exclusion', 'Remise appliquée (%)', 'Raison de la remise'
                    ])

                exclusion_from_calc_df = pd.DataFrame(exclusion_reasons_from_calc)
                exclusion_final_df     = pd.concat([excluded_from_exclus, exclusion_from_calc_df], ignore_index=True)

                st.session_state["result_df"]            = pd.DataFrame(result)
                st.session_state["margin_issues_df"]     = pd.DataFrame(margin_issues)
                st.session_state["exclusion_reasons_df"] = exclusion_final_df
                st.session_state["calcul_done"]          = True

                update_status(f"Calcul terminé — {len(result)} offres promo générées.")

        except Exception as e:
            st.error(f"Une erreur est survenue : {e}")
            update_status(f"Erreur : {e}")

    if st.session_state.get("calcul_done"):
        st.success(f"✅ {len(st.session_state['result_df'])} offres promo prêtes.")

        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button(
                "⬇️ Télécharger les résultats (CSV)",
                data=st.session_state["result_df"].to_csv(index=False, sep=';', encoding="utf-8"),
                file_name="prix_promo_output.csv",
                mime="text/csv"
            )
        with col2:
            st.download_button(
                "⬇️ Produits avec problèmes de marge (Excel)",
                data=to_excel(st.session_state["margin_issues_df"]),
                file_name="produits_avec_problemes_de_marge.xlsx",
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
# PAGE 2 — ANALYSE CA PAR COMMERCIAL (inchangée)
# ══════════════════════════════════════════════
elif page == "📊 Analyse CA par Commercial":

    st.title("📊 Analyse CA par Commercial")

    csv_file = st.file_uploader("Charger le fichier export commandes (CSV)", type=["csv"], key="ca_csv")

    if csv_file is not None:
        df = pd.read_csv(csv_file)
        df.columns = [c.replace("Commande - ", "").strip() for c in df.columns]

        df["Prix produits (HT)"] = pd.to_numeric(df["Prix produits (HT)"], errors="coerce")
        df["Prix final (HT)"]    = pd.to_numeric(df["Prix final (HT)"],    errors="coerce")
        df["taux_marge"]         = pd.to_numeric(df["taux_marge"],         errors="coerce")

        df["Auteur"] = df["Auteur"].fillna("(Sans commercial)").str.strip()
        df["Etat"]   = df["Etat"].fillna("(Inconnu)").str.strip()

        st.subheader("Filtres")
        col1, col2 = st.columns(2)

        auteurs_dispo = sorted(df["Auteur"].unique().tolist())
        etats_dispo   = sorted(df["Etat"].unique().tolist())

        with col1:
            auteurs_sel = st.multiselect("Auteur(s)", options=auteurs_dispo, default=[],
                                         placeholder="Sélectionner des commerciaux…")
        with col2:
            etats_preselectes = [e for e in ["en_preparation", "expedie", "valide"] if e in etats_dispo]
            etats_sel = st.multiselect("État(s)", options=etats_dispo, default=etats_preselectes,
                                       placeholder="Sélectionner des états…")

        masque_auteur = df["Auteur"].isin(auteurs_sel) if auteurs_sel else pd.Series([True] * len(df), index=df.index)
        masque_etat   = df["Etat"].isin(etats_sel)    if etats_sel   else pd.Series([True] * len(df), index=df.index)
        df_filtre = df[masque_auteur & masque_etat].copy()

        st.markdown(f"**{len(df_filtre):,} commandes** correspondent aux filtres sélectionnés.")

        if df_filtre.empty:
            st.warning("Aucune commande ne correspond à la sélection.")
        else:
            df_filtre["valeur_marge"] = df_filtre["Prix produits (HT)"] * df_filtre["taux_marge"] / 100

            agg = (
                df_filtre
                .groupby("Auteur", as_index=False)
                .agg(
                    Nb_commandes      =("Reference",          "count"),
                    CA_produits_HT    =("Prix produits (HT)", "sum"),
                    CA_final_HT       =("Prix final (HT)",    "sum"),
                    _val_marge        =("valeur_marge",       "sum"),
                    Taux_marge_simple =("taux_marge",         "mean"),
                )
                .sort_values("CA_final_HT", ascending=False)
            )
            agg["Taux_marge_pondere"] = agg["_val_marge"] / agg["CA_produits_HT"] * 100
            agg = agg.drop(columns=["_val_marge"])

            total_ca_ht     = df_filtre["Prix produits (HT)"].sum()
            total_val_marge = df_filtre["valeur_marge"].sum()
            total = pd.DataFrame([{
                "Auteur":             "**TOTAL**",
                "Nb_commandes":       agg["Nb_commandes"].sum(),
                "CA_produits_HT":     agg["CA_produits_HT"].sum(),
                "CA_final_HT":        agg["CA_final_HT"].sum(),
                "Taux_marge_simple":  df_filtre["taux_marge"].mean(),
                "Taux_marge_pondere": total_val_marge / total_ca_ht * 100 if total_ca_ht else 0,
            }])
            agg_display = pd.concat([agg, total], ignore_index=True)

            def fmt_eur(v):
                try:    return f"{v:,.2f} €".replace(",", " ").replace(".", ",")
                except: return v

            def fmt_pct(v):
                try:    return f"{v:.2f} %"
                except: return v

            agg_display["CA_produits_HT"]     = agg_display["CA_produits_HT"].apply(fmt_eur)
            agg_display["CA_final_HT"]        = agg_display["CA_final_HT"].apply(fmt_eur)
            agg_display["Taux_marge_simple"]  = agg_display["Taux_marge_simple"].apply(fmt_pct)
            agg_display["Taux_marge_pondere"] = agg_display["Taux_marge_pondere"].apply(fmt_pct)

            agg_display = agg_display.rename(columns={
                "Auteur":             "Commercial",
                "Nb_commandes":       "Nb commandes",
                "CA_produits_HT":     "CA Produits HT",
                "CA_final_HT":        "CA Final HT",
                "Taux_marge_simple":  "Taux marge moyen",
                "Taux_marge_pondere": "Taux marge pondéré",
            })

            st.subheader("Récapitulatif par commercial")
            st.dataframe(agg_display, use_container_width=True, hide_index=True)

            st.subheader("Indicateurs globaux")
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Nb commandes",       f"{df_filtre['Reference'].count():,}".replace(",", " "))
            m2.metric("CA Produits HT",     fmt_eur(total_ca_ht))
            m3.metric("CA Final HT",        fmt_eur(df_filtre["Prix final (HT)"].sum()))
            m4.metric("Taux marge moyen",   fmt_pct(df_filtre["taux_marge"].mean()))
            m5.metric("Taux marge pondéré", fmt_pct(total_val_marge / total_ca_ht * 100 if total_ca_ht else 0))

            agg_export = agg.copy()
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                st.download_button(
                    "⬇️ Télécharger le récapitulatif (Excel)",
                    data=to_excel(agg_export),
                    file_name="ca_par_commercial.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col_dl2:
                detail_export = df_filtre[["Reference", "Auteur", "Etat",
                                           "Prix produits (HT)", "Prix final (HT)",
                                           "taux_marge", "valeur_marge"]].copy()
                st.download_button(
                    "⬇️ Télécharger le détail des commandes (Excel)",
                    data=to_excel(detail_export),
                    file_name="detail_commandes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.info("👆 Chargez un fichier CSV pour démarrer l'analyse.")
        st.markdown(
            "**Colonnes attendues dans le fichier :**  \n"
            "`Reference`, `Auteur`, `Etat`, `Prix produits (HT)`, `Prix final (HT)`, `taux_marge`"
        )
