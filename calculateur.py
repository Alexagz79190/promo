import pandas as pd
import streamlit as st
from itertools import product
from datetime import datetime, time as dt_time  # alias pour éviter le conflit avec le module time
import time  # pour time.sleep
from io import BytesIO  # pour convertir les DataFrames en fichier Excel en mémoire

# --- Fonction pour convertir un DataFrame en fichier Excel (bytes) ---
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

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

# --- Interface principale ---
st.title("Calculateur de Prix Promo")
st.sidebar.header("Paramètres")

# --- Section de chargement des fichiers ---
st.subheader("Chargement des fichiers")
produit_file = st.file_uploader("Charger le fichier export produit (format Excel)", type=["xlsx"], key="produit")
exclusion_file = st.file_uploader("Charger le fichier exclusion produit (format Excel)", type=["xlsx"], key="exclusion")
remise_file = st.file_uploader("Charger le fichier remise (format Excel)", type=["xlsx"], key="remise")

# --- Bouton pour démarrer le calcul ---
if st.button("Démarrer le calcul"):
    try:
        if not (produit_file and exclusion_file and remise_file):
            st.error("Veuillez charger tous les fichiers requis.")
            update_status("Erreur : Fichiers manquants.")
        else:
            update_status("Chargement des exclusions depuis exclus.xlsx...")
            exclusions_data = pd.ExcelFile(exclusion_file)
            excl_fournisseur_famille = exclusions_data.parse('Fournisseur famille')[['Identifiant fournisseur', 'Identifiant famille']]
            
            # Création du produit cartésien fournisseur x famille
            fournisseurs_uniques = excl_fournisseur_famille['Identifiant fournisseur'].unique()
            familles_uniques = excl_fournisseur_famille['Identifiant famille'].unique()
            toutes_combinaisons = pd.DataFrame(list(product(fournisseurs_uniques, familles_uniques)),
                                               columns=['Identifiant fournisseur', 'Identifiant famille'])
            
            update_status("Chargement des données produit...")
            data = pd.read_excel(produit_file, sheet_name='Worksheet')
            
            # Vérification par le produit cartésien
            data_merged = data.merge(
                toutes_combinaisons,
                how='left',
                left_on=['Fournisseur : identifiant', 'Famille : identifiant'],
                right_on=['Identifiant fournisseur', 'Identifiant famille'],
                indicator=True
            )
            
            data_merged.loc[data_merged['_merge'] == 'both', 'Exclusion Reason'] = 'Exclus car présent dans toutes les combinaisons de Fournisseur x Famille'
            
            # Séparation des produits exclus
            data_excluded = data_merged[data_merged['Exclusion Reason'].notna()].copy()
            data_processed = data_merged[data_merged['Exclusion Reason'].isna()].copy()
            
            update_status(f"Produits exclus via produit cartésien : {len(data_excluded)}")
            update_status(f"Produits restants après exclusions : {len(data_processed)}")
            
            # Stockage des résultats dans le session_state
            st.session_state["exclusion_reasons_df"] = data_excluded[['Fournisseur : identifiant', 'Famille : identifiant', 'Exclusion Reason']]
            st.session_state["calcul_done"] = True
            
            update_status("Calcul terminé. Les fichiers de résultats sont prêts au téléchargement.")
    except Exception as e:
        st.error(f"Une erreur est survenue : {e}")
        update_status(f"Erreur : {e}")

# --- Bouton de téléchargement ---
if st.session_state.get("calcul_done"):
    st.download_button("Télécharger les produits exclus",
                       data=to_excel(st.session_state["exclusion_reasons_df"]),
                       file_name="produits_exclus.xlsx")
