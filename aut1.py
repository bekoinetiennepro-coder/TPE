import pandas as pd
import os
import shutil
import time
from datetime import datetime
from config import get_connection
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import re
# ----------------- CONFIGURATION -----------------
DOSSIER_ENTREE = r"C:\MonProjetCompta\input"
DOSSIER_ARCHIVE = r"C:\MonProjetCompta\archive"
DOSSIER_EXPORT = r"C:\MonProjetCompta\export"
INTERVALLE_CHECK = 10  # secondes pour la surveillance continue

# ----------------- NOM DU FICHIER D'EXPORT -----------------
def nom_fichier_export():
    date_str = datetime.now().strftime("%Y-%m-%d")
    return os.path.join(DOSSIER_EXPORT, f"resultat_{date_str}.xlsx")

# ----------------- DATAFRAME GLOBAL -----------------
df_final = pd.DataFrame()

# ----------------- DÉTECTION DE L'OPÉRATEUR -----------------
def detecter_operateur_depuis_nom(fichier):
    """Détecte le nom de l'opérateur à partir du nom du fichier."""
    nom_fichier = os.path.basename(fichier).upper()
    if "ORANGE" in nom_fichier:
        return "ORANGE"
    elif "MTN" in nom_fichier:
        return "MTN"
    elif "MOOV" in nom_fichier:
        return "MOOV"
    elif "WAVE" in nom_fichier:
        return "WAVE"
    else:
        return "INCONNU"




# ----------------- TRAITEMENT D'UN FICHIER -----------------
def traiter_fichier(fichier):
    global df_final
    try:
        print(f"\n[INFO] Traitement du fichier : {fichier}")
        operateur_fichier = detecter_operateur_depuis_nom(fichier)
        print(f"[INFO] Opérateur détecté : {operateur_fichier}")

        # Lecture du fichier pour identifier la ligne d'en-têtes
        df_temp = pd.read_excel(fichier, header=None)
        header_row = None

       
            # Détection automatique pour les autres opérateurs
        for i, row in df_temp.iterrows():
            line_str = " ".join(str(x).lower() for x in row.tolist())
            if all(keyword in line_str for keyword in ["montant", "commission", "site"]):
                header_row = i
                break
                
        if header_row is None:
            print(f"[ERREUR] Impossible de détecter la ligne d'en-têtes pour {fichier}")
            return

        # Lecture avec les bons en-têtes
        df = pd.read_excel(fichier, header=header_row)
        df.columns = [col.strip().lower().replace(" ", "_") for col in df.columns]

        # Ajout automatique de la colonne opérateur si absente
        if "opérateur" not in df.columns:
            df["opérateur"] = operateur_fichier
        else:
            df["opérateur"].fillna(operateur_fichier, inplace=True)

        # ----------------- Vérification des colonnes selon opérateur -----------------
                # ----------------- Vérification des colonnes selon opérateur -----------------
        manquantes = []

       
        if operateur_fichier == "MOOV":
            colonnes_requises = ["opérateur", "site", "montant", "commission", "montant_à_comptabiliser"]
  
        
        elif operateur_fichier == "MTN":
            colonnes_requises = ["opérateur", "site", "montant", "commission", "montant_à_comptabiliser"]
            
        
        elif operateur_fichier == "WAVE":
              # 🔍 Recherche intelligente des colonnes spécifiques à Orange
            colonnes_requises = ["opérateur", "site", "montant", "commission", "montant_à_comptabiliser"]

            # Trouver la colonne du montant (ex: "crédit" ou "credit")
            col_montant = next((c for c in df.columns if "montant_brut" in c or "montant_brut" in c), None)
            # Trouver la colonne de commission (ex: "compte:" ou "compte1")
                    
            if col_montant:
                df.rename(columns={col_montant: "montant"}, inplace=True)
            
            
            if "montant" not in df.columns or "commission" not in df.columns:
                print(f"[ERREUR] Colonnes non trouvées pour WAVE : montant={col_montant}, commission={col_commission}")
                print("[DEBUG] Colonnes détectées :", df.columns.tolist())
                return
         
        elif operateur_fichier == "ORANGE":
            # 🔍 Recherche intelligente des colonnes spécifiques à Orange
            colonnes_requises = ["opérateur", "site", "montant", "commission", "montant_à_comptabiliser"]

            # Trouver la colonne du montant (ex: "crédit" ou "credit")
            col_montant = next((c for c in df.columns if "crédit" in c or "credit" in c), None)
            # Trouver la colonne de commission (ex: "compte:" ou "compte1")
            col_commission = next((c for c in df.columns if "compte:_0767540922" in c), None)
             # Trouver la colonne de commission (ex: "compte:" ou "compte1")
            col_comptabiliser = next((c for c in df.columns if "net_à_comptabiliser" in c  ), None)
            
            if col_montant:
                df.rename(columns={col_montant: "montant"}, inplace=True)
            if col_commission:
                df.rename(columns={col_commission: "commission"}, inplace=True)

            if col_comptabiliser:
                df.rename(columns={col_comptabiliser: "montant_à_comptabiliser"}, inplace=True)
            
            
            if "montant" not in df.columns or "commission" not in df.columns:
                print(f"[ERREUR] Colonnes non trouvées pour ORANGE : montant={col_montant}, commission={col_commission}")
                print("[DEBUG] Colonnes détectées :", df.columns.tolist())
                return
            # 🧮 Normalisation des valeurs pour ORANGE
            if "commission" in df.columns:
                # convertir en numérique proprement
                df["commission"] = pd.to_numeric(df["commission"], errors="coerce").fillna(0)
                # remettre en positif (si valeurs négatives)
                df["commission"] = df["commission"].abs()

        else:
            colonnes_requises = ["opérateur", "site", "montant", "commission", "comptabiliser"]

        # ----------------- Vérification finale -----------------
        manquantes = [c for c in colonnes_requises if c not in df.columns]
        if manquantes:
            print(f"[ERREUR] Colonnes manquantes dans {fichier} : {manquantes}")
            print("[DEBUG] Colonnes présentes :", df.columns.tolist())
            return

        
        
        # ----------------- Regroupement -----------------
        df_group = df.groupby(["opérateur", "site"], as_index=False).sum(numeric_only=True)

        # ----------------- Génération des écritures comptables -----------------
        lignes = []
        for _, row in df_group.iterrows():
            site = row["site"]
            montant = row["montant"]
            commission = row["commission"]
            compta = row["montant_à_comptabiliser"]
            operateur = row["opérateur"].upper()
            Code="C52"
            
            
            
            if operateur == "MOOV":
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "C", "Compte": "558013", "Montant": montant})
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "C", "Compte": "558013", "Montant": montant,"TypeOp":"OD","Auxiliare":"TPE" + str(site)})
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "D", "Compte": "6327", "Montant": commission})
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "D", "Compte": "521650", "Montant": compta})
            elif operateur == "MTN":
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "C", "Compte": "558012", "Montant": montant})
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "D", "Compte": "6327", "Montant": commission})
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "D", "Compte": "531100", "Montant": compta})
            elif operateur == "ORANGE":
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "C", "Compte": "558011", "Montant": montant})
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "D", "Compte": "6327", "Montant": commission})
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "D", "Compte": "521650", "Montant": compta})
            elif operateur == "WAVE":
                # --- Nettoyage / Conversion des colonnes WAVE ---
            
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "C", "Compte": "558015", "Montant": montant})
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "C", "Compte": "558015", "Montant": montant,"TypeOp":"OD","Auxiliare":"TPE" + str(site)})
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "D", "Compte": "6327", "Montant": commission})
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "D", "Compte": "521450", "Montant": compta})
            else:
                print(f"[AVERTISSEMENT] Opérateur non reconnu pour {fichier}: {operateur}")

        # ----------------- Ajout au dataframe global -----------------
        if lignes:
            df_final = pd.concat([df_final, pd.DataFrame(lignes)], ignore_index=True)

        # ----------------- Déplacement du fichier -----------------
        shutil.move(fichier, os.path.join(DOSSIER_ARCHIVE, os.path.basename(fichier)))
        print(f"[OK] Fichier traité et archivé : {fichier}")

    except Exception as e:
        print(f"[ERREUR] Traitement du fichier {fichier} : {e}")

# ----------------- INSERTION DANS LA BASE -----------------
def inserer_donnees_bdd(df):
    try:
        if df.empty:
            print("[INFO] Aucune donnée à insérer dans la BDD.")
            return
        conn = get_connection()
        cursor = conn.cursor()
        sql = """
            INSERT INTO ecritures_comptables (site, operateur, type_ecriture, compte, montant)
            VALUES (%s, %s, %s, %s, %s)
        """
        valeurs = [(row["Site"], row["Operateur"], row["Type"], row["Compte"], float(row["Montant"]))
                   for _, row in df.iterrows()]
        cursor.executemany(sql, valeurs)
        conn.commit()
        cursor.close()
        conn.close()
        print(f"[BDD] {len(valeurs)} écritures enregistrées ✅")
    except Exception as e:
        print(f"[ERREUR BDD] {e}")

# ----------------- EXPORT DU JOUR -----------------
def exporter_resultat():
    if df_final.empty:
        print("[INFO] Aucune donnée à exporter aujourd'hui.")
        return
    fichier_export = nom_fichier_export()
    df_final.to_excel(fichier_export, index=False)
    print(f"[EXPORT] Résultat exporté : {fichier_export}")

# ----------------- TRAITEMENT DE TOUS LES FICHIERS -----------------
def automatiser_import():
    fichiers = [os.path.join(DOSSIER_ENTREE, f) for f in os.listdir(DOSSIER_ENTREE)
                if f.lower().endswith((".xlsx", ".xls",".xlsb"))]
    for fichier in fichiers:
        traiter_fichier(fichier)

    inserer_donnees_bdd(df_final)
    exporter_resultat()

# ----------------- SURVEILLANCE EN CONTINU -----------------
class SurveillanceHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith((".xlsx", ".xls",".xlsb")):
            traiter_fichier(event.src_path)
            inserer_donnees_bdd(df_final)
            exporter_resultat()

def surveillance_continue():
    event_handler = SurveillanceHandler()
    observer = Observer()
    observer.schedule(event_handler, path=DOSSIER_ENTREE, recursive=False)
    observer.start()
    print("[INFO] Surveillance continue activée...")
    try:
        while True:
            time.sleep(INTERVALLE_CHECK)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

# ----------------- MAIN -----------------
if __name__ == "__main__":
    automatiser_import()
    surveillance_continue()
