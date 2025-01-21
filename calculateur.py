import pandas as pd
import os
import streamlit as st
from itertools import product

# Helper function to load files
def load_file(file_type):
    uploaded_file = st.file_uploader(f"Charger le fichier {file_type} (format Excel)", type=["xlsx"])
    return uploaded_file

# Helper function to update status
def update_status(message):
    st.session_state["log"].append(message)
    st.text_area("Journal des actions", value="\n".join(st.session_state["log"]), height=200, disabled=True)

# Initialize session state for logs
if "log" not in st.session_state:
    st.session_state["log"] = []

# Main Streamlit App
st.title("Calculateur de Prix Promo")
st.sidebar.header("Paramètres")

# File upload section
st.subheader("Chargement des fichiers")
produit_file = load_file("export produit")
exclusion_file = load_file("exclusion produit")
remise_file = load_file("remise")

# Date selection
st.subheader("Sélection des dates")
start_date = st.text_input("Date de début (dd/mm/yyyy hh:mm:ss)")
end_date = st.text_input("Date de fin (dd/mm/yyyy hh:mm:ss)")

# Price type selection
st.subheader("Options de calcul")
price_option = st.radio("Choisissez les options de calcul :", 
                        options=["Prix d'achat avec option", "Prix de revient"])

# Export necessary fields
if st.button("Exporter les champs nécessaires"):
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
    st.download_button("Télécharger les champs nécessaires", 
                       data="\n".join(fields).encode("utf-8"), 
                       file_name="champs_export_produit.txt")
    update_status("Champs nécessaires exportés avec succès.")

# Process data
if st.button("Démarrer le calcul"):
    try:
        if not (produit_file and exclusion_file and remise_file):
            st.error("Veuillez charger tous les fichiers requis.")
            update_status("Erreur : Fichiers manquants.")
        elif not start_date or not end_date:
            st.error("Veuillez spécifier les dates de début et de fin.")
            update_status("Erreur : Dates manquantes.")
        else:
            # Load data
            update_status("Chargement des données produit...")
            data = pd.read_excel(produit_file, sheet_name='Worksheet')
            update_status(f"Nombre de produits chargés : {len(data)}")

            update_status("Chargement des exclusions...")
            exclusions_data = pd.ExcelFile(exclusion_file)
            excl_code_agz = exclusions_data.parse('Code AGZ')['Code AGZ'].dropna().astype(str).tolist()
            excl_fournisseur = exclusions_data.parse('Founisseur ')['Identifiant fournisseur seul'].dropna().astype(int).tolist()
            excl_marque = exclusions_data.parse('Marque')['Identifiant marque seul'].dropna().astype(int).tolist()
            excl_fournisseur_famille = exclusions_data.parse('Fournisseur famille')[['Identifiant fournisseur', 'Identifiant famille']]

            update_status("Génération des combinaisons fournisseur-famille...")
            fournisseur_famille_combinations = pd.DataFrame(
                product(
                    excl_fournisseur_famille['Identifiant fournisseur'].unique(),
                    excl_fournisseur_famille['Identifiant famille'].unique()
                ),
                columns=['Identifiant fournisseur', 'Identifiant famille']
            )

            # Apply exclusions
            update_status("Application des exclusions...")
            data = data[~data['Code produit'].astype(str).isin(excl_code_agz)]
            data = data[~data['Fournisseur : identifiant'].isin(excl_fournisseur)]
            data = data[~data['Marque : identifiant'].isin(excl_marque)]

            data = data.merge(
                fournisseur_famille_combinations,
                how='left',
                left_on=['Fournisseur : identifiant', 'Famille : identifiant'],
                right_on=['Identifiant fournisseur', 'Identifiant famille'],
                indicator=True
            )
            data = data[data['_merge'] == 'left_only']
            data = data.drop(columns=['Identifiant fournisseur', 'Identifiant famille', '_merge'])

            update_status(f"Produits restants après exclusions : {len(data)}")

            # Load discounts
            update_status("Chargement des remises...")
            remises = pd.read_excel(remise_file)

            # Choose price type
            price_column = 'Prix d\'achat avec option' if price_option == "Prix d'achat avec option" else 'Prix de revient'

            # Calculate promo prices
            update_status("Calcul des prix promo...")
            result = []
            for _, row in data.iterrows():
                prix_vente = row['Prix de vente en cours']
                prix_base = row[price_column]
                marge = (prix_vente - prix_base) / prix_vente * 100
                remise = 0
                for _, remise_row in remises.iterrows():
                    if remise_row['Marge minimale'] <= marge <= remise_row['Marge maximale']:
                        remise = remise_row['Remise'] / 100
                        break
                prix_promo = round(prix_vente * (1 - remise), 2)
                taux_marge_promo = round((prix_promo - prix_base) / prix_promo * 100, 2)
                if pd.notna(taux_marge_promo):
                    result.append({
                        'Identifiant produit': row['Identifiant produit'],
                        'Prix promo HT': str(prix_promo).replace('.', ','),
                        'Date de début prix promo': start_date,
                        'Date de fin prix promo': end_date,
                        'Taux marge prix promo': str(taux_marge_promo).replace('.', ',')
                    })

            result_df = pd.DataFrame(result)
            st.download_button("Télécharger les résultats", 
                               data=result_df.to_csv(index=False, sep=';', encoding='utf-8'), 
                               file_name="prix_promo_output.csv")
            update_status("Calcul terminé. Fichier exporté avec succès.")

    except Exception as e:
        st.error(f"Une erreur est survenue : {e}")
        update_status(f"Erreur : {e}")
