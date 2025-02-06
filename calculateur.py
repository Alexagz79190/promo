import pandas as pd
import streamlit as st
from itertools import product
from datetime import datetime, time as dt_time  # on renomme pour éviter le conflit avec le module time
import time  # pour time.sleep

# --- Initialisation du session_state ---
if "log" not in st.session_state:
    st.session_state["log"] = []
if "calcul_done" not in st.session_state:
    st.session_state["calcul_done"] = False

# --- Conteneur pour afficher le journal en direct ---
log_container = st.empty()

def update_status(message):
    """Ajoute un message au journal et met à jour l'affichage."""
    st.session_state["log"].append(f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - {message}")
    st.session_state["log_display"] = "\n".join(st.session_state["log"])
    log_container.text_area("Journal des actions", st.session_state["log_display"], height=200, disabled=True)
    time.sleep(0.1)

def load_file(file_type):
    """Charge un fichier Excel et retourne l'objet uploader."""
    return st.file_uploader(f"Charger le fichier {file_type} (format Excel)", type=["xlsx"], key=file_type)

# --- Interface principale ---
st.title("Calculateur de Prix Promo")
st.sidebar.header("Paramètres")

# --- Section de chargement des fichiers ---
st.subheader("Chargement des fichiers")

# Upload du fichier export produit
produit_file = load_file("export produit")

# Affichage des champs attendus (en dur) pour le fichier export produit
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

# Dès que le fichier export produit est chargé, proposer le bouton "Exporter les champs nécessaires"
if produit_file is not None:
    st.write("Fichier **export produit** chargé.")
    if st.button("Exporter les champs nécessaires"):
        # Champs en dur
        fields = [
            "Identifiant produit",
            "Fournisseur : identifiant",
            "Famille : identifiant",
            "Marque : identifiant",
            "Code produit",
            "Prix de vente en cours",
            "Prix d'achat avec option",
            "Prix de revient"
        ]
        st.session_state["export_fields"] = "\n".join(fields).encode("utf-8")
        update_status("Champs nécessaires exportés avec succès.")
    if "export_fields" in st.session_state:
        st.download_button("Télécharger les champs nécessaires",
                           data=st.session_state["export_fields"],
                           file_name="champs_export_produit.txt",
                           encoding="utf-8-sig")

# Upload des autres fichiers
exclusion_file = load_file("exclusion produit")
remise_file = load_file("remise")

# --- Sélection des dates et heures ---
st.subheader("Sélection des dates")
start_date = st.date_input("Date de début", value=datetime.now().date())
start_time = st.time_input("Heure de début", value=dt_time(0, 0))
end_date = st.date_input("Date de fin", value=datetime.now().date())
end_time = st.time_input("Heure de fin", value=dt_time(23, 59))
start_datetime = datetime.combine(start_date, start_time)
end_datetime = datetime.combine(end_date, end_time)

# --- Options de calcul ---
st.subheader("Options de calcul")
price_option = st.radio("Choisissez les options de calcul :",
                        options=["Prix d'achat avec option", "Prix de revient"])

# --- Bouton pour démarrer le calcul ---
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
            
            update_status("Application des exclusions issues du fichier exclus.xlsx...")
            data['Exclusion Reason'] = None
            data.loc[data['Code produit'].astype(str).isin(excl_code_agz), 'Exclusion Reason'] = 'Exclus car présent dans Code AGZ fichier exclus'
            data.loc[data['Fournisseur : identifiant'].isin(excl_fournisseur), 'Exclusion Reason'] = 'Exclus car présent dans Fournisseur fichier exclus'
            data.loc[data['Marque : identifiant'].isin(excl_marque), 'Exclusion Reason'] = 'Exclus car présent dans Marque fichier exclus'
            
            # Vérification par la combinaison fournisseur-famille
            data_merged = data.merge(
                excl_fournisseur_famille,
                how='left',
                left_on=['Fournisseur : identifiant', 'Famille : identifiant'],
                right_on=['Identifiant fournisseur', 'Identifiant famille'],
                indicator=True
            )
            data_merged.loc[data_merged['_merge'] == 'both', 'Exclusion Reason'] = 'Exclus car présent dans Fournisseur famille du fichier exclus'
            
            # Séparation des produits exclus via le fichier exclus.xlsx et des produits à traiter
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
                taux_marge_promo = round((prix_promo - prix_base) / prix_promo * 100, 2)
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
            
            # Préparation du DataFrame des produits exclus
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
            
            # Stockage des résultats dans le session_state
            st.session_state["result_df"] = pd.DataFrame(result)
            st.session_state["margin_issues_df"] = pd.DataFrame(margin_issues)
            st.session_state["exclusion_reasons_df"] = exclusion_final_df
            st.session_state["calcul_done"] = True
            
            update_status("Calcul terminé. Les fichiers de résultats sont prêts au téléchargement.")
    except Exception as e:
        st.error(f"Une erreur est survenue : {e}")
        update_status(f"Erreur : {e}")

# --- Boutons de téléchargement (après calcul) ---
if st.session_state.get("calcul_done"):
    st.download_button("Télécharger les résultats",
                       data=st.session_state["result_df"].to_csv(index=False, sep=';', encoding="utf-8-sig"),
                       file_name="prix_promo_output.csv")
    st.download_button("Télécharger les produits avec problèmes de marge",
                       data=st.session_state["margin_issues_df"].to_csv(index=False, sep=';', encoding="utf-8-sig"),
                       file_name="produits_avec_problemes_de_marge.csv")
    st.download_button("Télécharger les produits exclus",
                       data=st.session_state["exclusion_reasons_df"].to_csv(index=False, sep=';', encoding="utf-8-sig"),
                       file_name="produits_exclus.csv")
