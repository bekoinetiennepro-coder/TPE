import pandas as pd
from tkinter import *
from tkinter import ttk, filedialog, messagebox
import os
from config import get_connection

df_final = pd.DataFrame()

# ================== 🔍 Lecture automatique du bon header ===================
def lire_fichier_excel_auto(fichier):
    try:
        temp_df = pd.read_excel(fichier, header=None, nrows=10)
        header_row = None

        for i, row in temp_df.iterrows():
            ligne_str = [str(x).strip().lower() for x in row.fillna("")]
            if "opérateur" in ligne_str and "site" in ligne_str:
                header_row = i
                break

        if header_row is None:
            raise Exception("Impossible de détecter la ligne contenant 'Operateur' et 'Site'.")

        df = pd.read_excel(fichier, header=header_row)
        return df

    except Exception as e:
        raise Exception(f"Erreur lecture fichier : {e}")

def importer_fichier():
    global df_final
    try:
        fichier = filedialog.askopenfilename(
            title="Sélectionner un fichier Excel",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls")]
        )
        if not fichier:
            return

        # Lire le fichier sans header pour détecter la bonne ligne
        df_temp = pd.read_excel(fichier, header=None)

        # Trouver la ligne contenant les mots-clés
        header_row = None
        for i, row in df_temp.iterrows():
            line_str = " ".join(str(x).lower() for x in row.tolist())
            if "opérateur" in line_str and "site" in line_str:
                header_row = i
                break

        if header_row is None:
            messagebox.showerror("Erreur lecture fichier",
                                 "Impossible de détecter la ligne contenant 'Operateur' et 'Site'.")
            return

        # Relire le fichier à partir de la ligne détectée
        df = pd.read_excel(fichier, header=header_row)

        # Normaliser les noms de colonnes
        df.columns = [col.strip().lower().replace(" ", "_") for col in df.columns]

        colonnes_requises = ["opérateur", "site", "montant", "commission", "comptabiliser"]
        manquantes = [c for c in colonnes_requises if c not in df.columns]
        if manquantes:
            messagebox.showerror("Erreur", f"Colonnes manquantes : {manquantes}")
            return

        # Regrouper par site et opérateur
        df_group = df.groupby(["opérateur", "site"], as_index=False).sum()

        # Ajouter Debit et Crédit
        lignes = []
        for _, row in df_group.iterrows():
            site = row["site"]
            montant = row["montant"]
            commission = row["commission"]
            compta = row["comptabiliser"]
            operateur = row["opérateur"]

            lignes.append({
                "Operateur": operateur,
                "Site": site,
                "Type": "Débit",
                "Compte": "45800",
                "Montant": montant
            })
            lignes.append({
                "Operateur": operateur,
                "Site": site,
                "Type": "Crédit",
                "Compte": "2100",
                "Montant": commission
            })
            lignes.append({
                "Operateur": operateur,
                "Site": site,
                "Type": "Crédit",
                "Compte": "36000",
                "Montant": compta
            })

        df_final = pd.DataFrame(lignes)
        messagebox.showinfo("Succès", "Importation réussie ✅")

    except Exception as e:
        messagebox.showerror("Erreur importation", f"Erreur lecture fichier : {e}")
# ================== 📤 Export Excel ===================
def exporter_resultat():
    try:
        if df_final.empty:
            messagebox.showwarning("Avertissement", "Aucune donnée à exporter.")
            return

        fichier = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichier Excel", "*.xlsx")],
            title="Enregistrer le résultat"
        )
        if fichier:
            df_final.to_excel(fichier, index=False)
            messagebox.showinfo("Succès", f"Résultat exporté : {os.path.basename(fichier)}")
    except Exception as e:
        messagebox.showerror("Erreur", f"Échec exportation : {e}")

# ================== 🖥️ Interface graphique ===================
root = Tk()
root.title("Logiciel Comptable - Détection automatique + MySQL")
root.geometry("950x600")

frame = Frame(root)
frame.pack(pady=10)

btn_import = Button(frame, text="📂 Importer un fichier Excel", command=importer_fichier)
btn_import.grid(row=0, column=0, padx=10)

btn_export = Button(frame, text="💾 Exporter le résultat", command=exporter_resultat)
btn_export.grid(row=0, column=1, padx=10)

colonnes = ["Site", "Opérateur", "Type", "Compte", "Montant"]
tree = ttk.Treeview(root, columns=colonnes, show="headings", height=20)
for col in colonnes:
    tree.heading(col, text=col)
    tree.column(col, width=150, anchor=CENTER)
tree.pack(fill=BOTH, expand=True, padx=10, pady=10)

root.mainloop()

