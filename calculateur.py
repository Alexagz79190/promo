import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from itertools import product

def load_file(entry_widget, file_type):
    """Helper function to load a file using Tkinter file dialog."""
    file_path = filedialog.askopenfilename(title=f"Charger le fichier {file_type}", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)
    else:
        messagebox.showwarning("Avertissement", f"Aucun fichier {file_type} sélectionné.")

def update_status(message):
    """Update the status box with a message."""
    status.config(state="normal")
    status.insert(tk.END, f"{message}\n")
    status.config(state="disabled")
    status.see(tk.END)

def export_fields():
    """Export the required fields to a TXT file."""
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
    default_output_path = "champs_export_produit.txt"
    output_path = filedialog.asksaveasfilename(title="Enregistrer les champs nécessaires", initialfile=default_output_path, defaultextension=".txt", filetypes=[("Text files", "*.txt")])
    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(fields))
        update_status(f"Fichier des champs exporté avec succès à : {output_path}")
        messagebox.showinfo("Succès", f"Fichier des champs exporté avec succès à : {output_path}")
    else:
        update_status("Aucun emplacement de fichier spécifié pour les champs.")
        messagebox.showwarning("Avertissement", "Aucun emplacement de fichier spécifié pour les champs.")

def start_processing():
    """Process the product data with selected options."""
    try:
        # Get file paths from entries
        produit_path = produit_entry.get()
        exclusion_path = exclusion_entry.get()
        remise_path = remise_entry.get()

        if not os.path.exists(produit_path) or not os.path.exists(exclusion_path) or not os.path.exists(remise_path):
            messagebox.showerror("Erreur", "Veuillez vérifier que tous les fichiers sont sélectionnés.")
            return

        # Get promo dates
        start_date = start_date_entry.get()
        end_date = end_date_entry.get()
        if not start_date or not end_date:
            messagebox.showerror("Erreur", "Veuillez spécifier les dates de début et de fin.")
            return

        # Load data
        update_status("Chargement des données produit...")
        data = pd.read_excel(produit_path, sheet_name='Worksheet')
        update_status(f"Nombre de produits chargés : {len(data)}")

        update_status("Chargement des exclusions...")
        exclusions_data = pd.ExcelFile(exclusion_path)
        excl_code_agz = exclusions_data.parse('Code AGZ')['Code AGZ'].dropna().astype(str).tolist()
        excl_fournisseur = exclusions_data.parse('Founisseur ')['Identifiant fournisseur seul'].dropna().astype(int).tolist()
        excl_marque = exclusions_data.parse('Marque')['Identifiant marque seul'].dropna().astype(int).tolist()
        excl_fournisseur_famille = exclusions_data.parse('Fournisseur famille')[['Identifiant fournisseur', 'Identifiant famille']]

        # Générer toutes les combinaisons possibles de fournisseurs et familles
        update_status("Génération des combinaisons fournisseur-famille...")
        fournisseur_famille_combinations = pd.DataFrame(
            product(
                excl_fournisseur_famille['Identifiant fournisseur'].unique(),
                excl_fournisseur_famille['Identifiant famille'].unique()
            ),
            columns=['Identifiant fournisseur', 'Identifiant famille']
        )

        update_status("Application des exclusions...")
        data = data[~data['Code produit'].astype(str).isin(excl_code_agz)]
        data = data[~data['Fournisseur : identifiant'].isin(excl_fournisseur)]
        data = data[~data['Marque : identifiant'].isin(excl_marque)]

        # Exclure les combinaisons fournisseur-famille
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

        update_status("Chargement des remises...")
        remises = pd.read_excel(remise_path)

        # Choose price type
        price_type = price_option.get()
        if price_type == "achat":
            price_column = 'Prix d\'achat avec option'
        else:
            price_column = 'Prix de revient'

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
        default_output_path = "prix_promo_output.csv"
        output_path = filedialog.asksaveasfilename(title="Enregistrer le fichier de sortie", initialfile=default_output_path, defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if output_path:
            result_df.to_csv(output_path, sep=';', index=False, encoding='utf-8')
            update_status(f"Fichier exporté avec succès à : {output_path}")
            messagebox.showinfo("Succès", f"Fichier exporté avec succès à : {output_path}")
        else:
            update_status("Aucun emplacement de fichier spécifié.")
            messagebox.showwarning("Avertissement", "Aucun emplacement de fichier spécifié.")

    except Exception as e:
        update_status(f"Erreur : {e}")
        messagebox.showerror("Erreur", f"Une erreur est survenue : {e}")

# Create Tkinter window
root = tk.Tk()
root.title("Calculateur de prix promo")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack()

# Entries for file selection
tk.Label(frame, text="Fichier export produit :").grid(row=0, column=0, sticky="w")
produit_entry = tk.Entry(frame, width=50)
produit_entry.grid(row=0, column=1, padx=5)
tk.Button(frame, text="Parcourir", command=lambda: load_file(produit_entry, "export produit")).grid(row=0, column=2)

tk.Label(frame, text="Fichier exclusion produit :").grid(row=1, column=0, sticky="w")
exclusion_entry = tk.Entry(frame, width=50)
exclusion_entry.grid(row=1, column=1, padx=5)
tk.Button(frame, text="Parcourir", command=lambda: load_file(exclusion_entry, "exclusion produit")).grid(row=1, column=2)

tk.Label(frame, text="Fichier remise :").grid(row=2, column=0, sticky="w")
remise_entry = tk.Entry(frame, width=50)
remise_entry.grid(row=2, column=1, padx=5)
tk.Button(frame, text="Parcourir", command=lambda: load_file(remise_entry, "remise")).grid(row=2, column=2)

# Promo date selection
tk.Label(frame, text="Date de début (dd/mm/yyyy hh:mm:ss) :").grid(row=3, column=0, sticky="w")
start_date_entry = tk.Entry(frame, width=50)
start_date_entry.grid(row=3, column=1, padx=5)

tk.Label(frame, text="Date de fin (dd/mm/yyyy hh:mm:ss) :").grid(row=4, column=0, sticky="w")
end_date_entry = tk.Entry(frame, width=50)
end_date_entry.grid(row=4, column=1, padx=5)

# Price type selection
tk.Label(frame, text="Choisissez les options de calcul :").grid(row=5, column=0, columnspan=2, sticky="w")
price_option = tk.StringVar(value="achat")
rbtn_achat = tk.Radiobutton(frame, text="Prix d'achat avec option", variable=price_option, value="achat")
rbtn_achat.grid(row=6, column=0, sticky="w")
rbtn_revient = tk.Radiobutton(frame, text="Prix de revient", variable=price_option, value="revient")
rbtn_revient.grid(row=6, column=1, sticky="w")

# Buttons for processing
process_btn = tk.Button(frame, text="Démarrer le calcul", command=start_processing)
process_btn.grid(row=7, column=0, columnspan=3, pady=10)

export_fields_btn = tk.Button(frame, text="Exporter les champs nécessaires", command=export_fields)
export_fields_btn.grid(row=8, column=0, columnspan=3, pady=10)

# Status box
status = tk.Text(frame, height=10, state="disabled", wrap="word")
status.grid(row=9, column=0, columnspan=3, pady=10)

root.mainloop()
