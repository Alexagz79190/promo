import pandas as pd
import streamlit as st
from itertools import product
from datetime import datetime, time
import time  # pour insérer des délais

# --- Initialisation du session_state ---
if "log" not in st.session_state:
    st.session_state["log"] = []
if "log_display" not in st.session_state:
    st.session_state["log_display"] = ""
if "calcul_done" not in st.session_state:
    st.session_state["calcul_done"] = False

# --- Création d'un conteneur pour le journal d'événements ---
log_container = st.empty()

def update_status(message):
    """Ajoute un message au journal et met à jour le conteneur."""
    st.session_state["log"].append(f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - {message}")
    st.session_state["log_display"] = "\n".join(st.session_state["log"])
    # Mise à jour immédiate du journal dans le conteneur :
    log_container.text_area("Journal des actions", st.session_state["log_display"], height=200, disabled=True)
    # Délai pour forcer l'actualisation de l'interface
    time.sleep(0.1)

def load_file(file_type):
    uploaded_file = st.file_uploader(f"Charger le fichier {file_type} (format Excel)", type=["xlsx"], key=file_type)
    return uploaded_file

# --- Titre et paramètres ---
st.title("Calculateur de Prix Promo")
st.sidebar.header("Paramètres")

# --- Section de chargement des fichiers ---
st.subheader("Chargement des fichiers")
produit_file = load_file("export produit")
exclusion_file = load_file("exclusion produit")
remise_file = load_file("remise")

# --- Sélection des dates ---
st.subheader("Sélection des dates")
start_date = st.date_input("Date de début", value=datetime.now().date())
start_time = st.time_input("Heure de début", value=time(0, 0))
end_date = st.date_input("Date de fin", value=datetime.now().date())
end_time = st.time_input("Heure de fin", value=time(23, 59))
start_datetime = datetime.combine(start_date, start_time)
end_datetime = datetime.combine(end_date, end_time)

# --- Options de calcul ---
st.subheader("Options de calcul")
price_option = st.radio("Choisissez les options de calcul :",
                        options=["Prix d'achat avec option", "Prix de revient"])

# --- Export des champs nécessaires ---
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
    st.session_state["export_fields"] = "\n".join(fields).encode("utf-8")
    update_status("Champs nécessaires exportés avec succès.")

if "export_fields" in st.session_state:
    st.download_button("Télécharger les champs nécessaires",
                       data=st.session_state["export_fields"],
                       file_name="champs_export_produit.txt")

# --- Bouton pour lancer le calcul ---
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
            
            update_status("Application des exclusions...")
            exclusion_reasons = []
            data['Exclusion Reason'] = None
            data.loc[data['Code produit'].astype(str).isin(excl_code_agz), 'Exclusion Reason'] = 'Exclus car présent dans code AGZ fichier exclus'
            data.loc[data['Fournisseur : identifiant'].isin(excl_fournisseur), 'Exclusion Reason'] = 'Exclus car présent dans Fournisseur fichier exclus'
            data.loc[data['Marque : identifiant'].isin(excl_marque), 'Exclusion Reason'] = 'Exclus car présent dans Marque fichier exclus'
            data = data.merge(
                fournisseur_famille_combinations,
                how='left',
                left_on=['Fournisseur : identifiant', 'Famille : identifiant'],
                right_on=['Identifiant fournisseur', 'Identifiant famille'],
                indicator=True
            )
            data.loc[data['_merge'] == 'both', 'Exclusion Reason'] = 'Exclus car présent dans Fournisseur famille du fichier exclus'
            data = data[data['_merge'] == 'left_only']
            data = data.drop(columns=['Identifiant fournisseur', 'Identifiant famille', '_merge'])
            update_status(f"Produits restants après exclusions : {len(data)}")

            update_status("Chargement des remises...")
            remises = pd.read_excel(remise_file)
            price_column = 'Prix d\'achat avec option' if price_option == "Prix d'achat avec option" else 'Prix de revient'

            update_status("Calcul des prix promo...")
            result = []
            margin_issues = []
            for _, row in data.iterrows():
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
                            'Prix d\'achat avec option': row['Prix d\'achat avec option'],
                            'Prix de revient': row['Prix de revient'],
                            'Prix promo calculé': prix_promo
                        })
                else:
                    exclusion_reasons.append({
                        'Code produit': row['Code produit'],
                        'Raison exclusion': 'Exclus car le prix promo est supérieur ou égal au prix de vente',
                        'Prix de vente en cours': prix_vente,
                        'Prix d\'achat avec option': row['Prix d\'achat avec option'],
                        'Prix de revient': row['Prix de revient'],
                        'Remise appliquée': remise_appliquee * 100,
                        'Raison de la remise': remise_raison
                    })

            st.session_state["result_df"] = pd.DataFrame(result)
            st.session_state["margin_issues_df"] = pd.DataFrame(margin_issues)
            st.session_state["exclusion_reasons_df"] = pd.DataFrame(exclusion_reasons)
            st.session_state["calcul_done"] = True

            update_status("Calcul terminé. Les fichiers de résultats sont prêts au téléchargement.")
    except Exception as e:
        st.error(f"Une erreur est survenue : {e}")
        update_status(f"Erreur : {e}")

# --- Affichage des boutons de téléchargement (ils restent affichés) ---
if st.session_state.get("calcul_done"):
    st.download_button("Télécharger les résultats",
                       data=st.session_state["result_df"].to_csv(index=False, sep=';', encoding='utf-8'),
                       file_name="prix_promo_output.csv")
    st.download_button("Télécharger les produits avec problèmes de marge",
                       data=st.session_state["margin_issues_df"].to_csv(index=False, sep=';', encoding='utf-8'),
                       file_name="produits_avec_problemes_de_marge.csv")
    st.download_button("Télécharger les produits exclus",
                       data=st.session_state["exclusion_reasons_df"].to_csv(index=False, sep=';', encoding='utf-8'),
                       file_name="produits_exclus.csv")
