import pandas as pd
import os
import shutil
import time
from datetime import datetime
from config import get_connection
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import re
import csv
# ----------------- CONFIGURATION -----------------
DOSSIER_ENTREE = r"C:\MonProjetCompta\input"
DOSSIER_ARCHIVE = r"C:\MonProjetCompta\archive"
DOSSIER_EXPORT = r"C:\MonProjetCompta\export"
INTERVALLE_CHECK = 10  # secondes pour la surveillance continue


# ----------------- NOM DU FICHIER D'EXPORT -----------------
def nom_fichier_export():
    date_str = datetime.now().strftime("%Y-%m-%d")
    return os.path.join(DOSSIER_EXPORT, f"resultat_{date_str}.txt")

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

        # Cas spécial WAVE → le header est toujours à la ligne 2 (index = 1)
        if operateur_fichier == "WAVE":
            header_row = 0
        
            # Détection automatique pour les autres opérateurs
        for i, row in df_temp.iterrows():
            line_str = " ".join(str(x).lower() for x in row.tolist())
            if "site" in line_str or "site_" in line_str:  # MOOV / MTN / ORANGE
                header_row = i
                break

        if header_row is not None:
            print(f"[INFO] Ligne d'en-têtes détectée à l'index {header_row}")
        else:
            print(f"[ERREUR] Impossible de détecter la ligne d'en-têtes pour {fichier}")
            return
                
        # Lecture avec les bons en-têtes
        df = pd.read_excel(fichier, header=header_row)
        #df.columns = [col.strip().lower().replace(" ", "_") for col in df.columns]
        df.columns = (
            df.columns.str.strip()
                .str.lower()
                .str.replace(" ", "_")
                .str.replace("é", "e")
                .str.replace("è", "e")
                .str.replace("ê", "e")
                .str.replace("à", "a")
                .str.replace("ô", "o")
                .str.replace("î", "i")
                .str.replace("'", "")
        )

        # Ajout automatique de la colonne opérateur si absente
        if "opérateur" not in df.columns:
            df["opérateur"] = operateur_fichier
        else:
            df["opérateur"].fillna(operateur_fichier, inplace=True)

        # ----------------- Vérification des colonnes selon opérateur -----------------
                # ----------------- Vérification des colonnes selon opérateur -----------------
        manquantes = []
        colonnes_requises = []
       
        if operateur_fichier == "MOOV":
            colonnes_requises = ["opérateur", "site", "montant", "commission", "montant_a_comptabiliser", "date_operation"]
           
        elif operateur_fichier == "MTN":
            colonnes_requises = ["opérateur", "site", "montant", "commission", "montant_a_comptabiliser","date_operation"]
               
        # elif operateur_fichier == "WAVE":

        #     # ----------- 1) Mappage intelligent des colonnes -----------
        #     col_site = next((c for c in df.columns if "nom_marchand" in c), None)
        #     col_montant = next((c for c in df.columns if "montant_brut" in c and "net" not in c), None)
        #     col_commission = next((c for c in df.columns if "frais" in c), None)
        #     col_comptabiliser = next((c for c in df.columns if "montant_net" in c), None)

        #     # Renommage
        #     if col_site:
        #         df.rename(columns={col_site: "site"}, inplace=True)
        #         df["site"] = df["site"].astype(str).str.extract(r"(\d+)", expand=False)

        #     if col_montant:
        #         df.rename(columns={col_montant: "montant"}, inplace=True)

        #     if col_commission:
        #         df.rename(columns={col_commission: "commission"}, inplace=True)

        #     if col_comptabiliser:
        #         df.rename(columns={col_comptabiliser: "montant_à_comptabiliser"}, inplace=True)

        #     # ----------- 2) Conversion en numérique -----------
        #     for col in ["montant", "commission", "montant_à_comptabiliser"]:
        #         if col in df.columns:
        #             df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        #             #df["commission"] = df["commission"].abs()
        #     print("[DEBUG] Colonnes WAVE après nettoyage :", df.columns.tolist())

        #     # ----------- 3) Suppression des lignes cumulées WAVE -----------
        #     # (Celles qui créent les totaux 59830 / 480 / 59350)
        #     # 1️⃣ Supprimer les lignes où montant == montant_à_comptabiliser (typique cumul)
        #     df = df[df["montant"] != df["montant_à_comptabiliser"]]
        #     # 2️⃣ Supprimer les lignes avec commission = 0 (cumul du jour)
        #     df = df[~((df["commission"] == 0 ) & (df["montant"] > 0))]
            
        #     # 3️⃣ Supprimer doublons exacts
        #     if all(col in df.columns for col in ["montant", "commission", "montant_à_comptabiliser", "site"]):
        #         df = df.drop_duplicates(subset=["site", "montant", "commission", "montant_à_comptabiliser"])

        #     # Debug pour ton site 204
        #     if "site" in df.columns:
        #         df_204 = df[df["site"] == "204"]
        #         print(f"[DEBUG] Lignes finales conservées site 204 ({len(df_204)} lignes) :")
        #         print(df_204[["site", "montant", "commission", "montant_à_comptabiliser"]])

        #     # ----------- SUPPRIMER LE GRAND CUMUL FINAL -----------

        #     # 5️⃣ Supprimer la ligne où montant = somme totale (cumul final)
        #     montant_max = df["montant"].max()
        #     commission_max = df["commission"].max()

        #     # ligne cumulée = montant max ET commission = commission max
        #     df = df[~((df["montant"] == montant_max) & (df["commission"] == commission_max))]
            
            

        elif operateur_fichier == "ORANGE":
            # 🔍 Recherche intelligente des colonnes spécifiques à Orange
            colonnes_requises = ["opérateur", "site", "montant", "commission", "montant_a_comptabiliser","date_operation"]

            # Trouver la colonne du montant (ex: "crédit" ou "credit")
            col_montant = next((c for c in df.columns if "credit" in c or "crédit" in c), None)
            # Trouver la colonne de commission (ex: "compte:" ou "compte1")
            col_commission = next((c for c in df.columns if "compte:_0767540922" in c), None)
             # Trouver la colonne de commission (ex: "compte:" ou "compte1")
            col_comptabiliser = next((c for c in df.columns if "net_a_comptabiliser" in c  ), None)
            col_date = next((c for c in df.columns if "date" in c ),None)
            
            if col_montant:
                df.rename(columns={col_montant: "montant"}, inplace=True)
            if col_commission:
                df.rename(columns={col_commission: "commission"}, inplace=True)

            if col_comptabiliser:
                df.rename(columns={col_comptabiliser: "montant_a_comptabiliser"}, inplace=True)
            if col_date:
                df.rename(columns={col_date: "date_operation"},inplace=True)
            
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


        elif operateur_fichier == "WAVE":
            # 🔍 Recherche intelligente des colonnes spécifiques à Orange
            colonnes_requises = ["opérateur", "site", "montant", "commission", "montant_a_comptabiliser","date_operation"]

            # Trouver la colonne du montant (ex: "crédit" ou "credit")
            col_montant = next((c for c in df.columns if "montant_brut" in c ), None)
            # Trouver la colonne de commission (ex: "compte:" ou "compte1")
            col_commission = next((c for c in df.columns if "frais" in c), None)
             # Trouver la colonne de commission (ex: "compte:" ou "compte1")
            col_comptabiliser = next((c for c in df.columns if "montant_net" in c  ), None)
            col_date = next((c for c in df.columns if "date" in c ),None)
            
            if col_montant:
                df.rename(columns={col_montant: "montant"}, inplace=True)
            if col_commission:
                df.rename(columns={col_commission: "commission"}, inplace=True)

            if col_comptabiliser:
                df.rename(columns={col_comptabiliser: "montant_a_comptabiliser"}, inplace=True)
            if col_date:
                df.rename(columns={col_date: "date_operation"},inplace=True)
            
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
            colonnes_requises = ["opérateur", "site", "montant", "commission", "comptabiliser","terminal"]

        # ----------------- Vérification finale -----------------
        manquantes = [c for c in colonnes_requises if c not in df.columns]
        if manquantes:
            print(f"[ERREUR] Colonnes manquantes dans {fichier} : {manquantes}")
            print("[DEBUG] Colonnes présentes :", df.columns.tolist())
            return

       
        # ----------------- Regroupement final avec terminal, date_operation et periode -----------------

        # 1️⃣ Sauvegarder la colonne terminal (si présente)
        # if "terminal" in df.columns:
        #     df_terminal = df.groupby(["opérateur", "site"])["terminal"].first().reset_index()
        # else:
        #     df_terminal = None

        # 2️⃣ Sauvegarder la colonne date_operation (si présente)
        if "date_operation" in df.columns:
            # Conversion en datetime puis garder la date seulement
            df["date_operation"] = pd.to_datetime(df["date_operation"], errors="coerce").dt.date
            # Format jj/mm/aaaa
            df["date_operation"] = df["date_operation"].apply(lambda x: x.strftime("%d/%m/%Y") if pd.notnull(x) else "")
            # Sauvegarder la date par site avant groupby
            df_date = df.groupby(["opérateur", "site"])["date_operation"].first().reset_index()
        else:
            df_date = None

        # 3️⃣ Groupby pour sommer les colonnes numériques
        df_group = df.groupby(["opérateur", "site"], as_index=False).sum(numeric_only=True)

        # 4️⃣ Réattacher la colonne terminal
       
        # 5️⃣ Réattacher la colonne date_operation
        if df_date is not None:
            df_group = df_group.merge(df_date, on=["opérateur", "site"], how="left")

        # 6️⃣ Créer la colonne periode à partir de date_operation
        if "date_operation" in df_group.columns:
            # Convertir en datetime
            df_group["date_operation_dt"] = pd.to_datetime(df_group["date_operation"], format="%d/%m/%Y", errors="coerce")
            # Créer periode au format jj-mmm-aa (ex: 12-nov-25)
            df_group["periode"] = df_group["date_operation_dt"].dt.strftime("%d-%b-%y").str.lower()
            # Supprimer colonne temporaire
            df_group.drop(columns=["date_operation_dt"], inplace=True)

        # -----------------------------------------
#       EXPORT FINAL EN MODE APPEND
# -----------------------------------------

        # chemin_sortie = nom_fichier_export()

        # df_group.to_csv(
        #     chemin_sortie,
        #     sep=";",             
        #     index=False,
        #     header=False,        # pas d’en-tête
        #     encoding="utf-8-sig",
        #     mode="a",            # 🔥 ajouter au fichier existant
        # )

        # print("Lignes ajoutées dans :", chemin_sortie)



        # # 2️⃣ Regroupement normal (sum supprime les colonnes texte)
        # df_group = df.groupby(["opérateur", "site"], as_index=False).sum(numeric_only=True)

        # # 3️⃣ Réattacher la colonne terminal après le groupby
        # if df_terminal is not None:
        #     df_group = df_group.merge(df_terminal, on=["opérateur", "site"], how="left")

       
        # ----------------- Génération des écritures comptables -----------------
        lignes = []
        for _, row in df_group.iterrows():
            site = row["site"]
            montant = row["montant"]
            commission =  row["commission"]
            compta = row["montant_a_comptabiliser"]
            operation = row["date_operation"]
            operateur = row["opérateur"].upper()
            Code="C52"
            libelle55= str(site) + " Clollecte Globale TPE " + operateur +  " du " + str(operation)
            libelle52= str(site) + " Collecte Vente TPE du " +  str(operation)
            libelle63= str(site) + " Commission sur Collecte TPE du " + str(operation)
            comptmoov = 558013
            comptebrumoov_om = 521650
            comptemtn = 558012
            comptebrumtn = 531100
            compteom = 558011
            comptewa = 558015
            comptebruwav = 521450
            comptefr = 6327
            auxi = "TPE" + str(site)
         
            
            
            #Ligne Moov
            if operateur == "MOOV":
                
                lignes.append({"Journal":Code, "Date": operation,"Mouvement":"OD", "Compte": comptefr , "Sens Ecriture": "D","Type Ecriture":"", "Montant": commission,"Auxil/Analyst": "","Libelle":libelle63})
                lignes.append({"Journal":Code, "Date": operation,"Mouvement":"OD", "Compte": comptefr , "Sens Ecriture": "D", "Type Ecriture":"A","Montant": commission,"Auxil/Analyst": site,"Libelle":libelle63})
                
                
                lignes.append({"Journal":Code, "Date": operation,"Mouvement":"OD", "Compte": comptmoov, "Sens Ecriture": "C", "Type Ecriture":"X", "Montant": montant,"Auxil/Analyst": auxi ,"Libelle":libelle55 })
                
                lignes.append({"Journal":Code, "Date": operation,"Mouvement":"OD", "Compte": comptebrumoov_om, "Sens Ecriture": "D",  "Montant": compta,"Libelle":libelle52})
              
               
                
            #ligne Mtn    
            elif operateur == "MTN":
                
              lignes.append({"Journal":Code, "Date": operation,"Mouvement":"OD", "Compte": comptefr , "Sens Ecriture": "D","Type Ecriture":"", "Montant": commission,"Auxil/Analyst": "","Libelle":libelle63})
              lignes.append({"Journal":Code, "Date": operation,"Mouvement":"OD", "Compte": comptefr , "Sens Ecriture": "D", "Type Ecriture":"A","Montant": commission,"Auxil/Analyst": site,"Libelle":libelle63})
                
                
              lignes.append({"Journal":Code, "Date": operation,"Mouvement":"OD", "Compte": comptemtn, "Sens Ecriture": "C", "Type Ecriture":"X", "Montant": montant,"Auxil/Analyst": auxi ,"Libelle":libelle55 })
            
              lignes.append({"Journal":Code, "Date": operation,"Mouvement":"OD", "Compte": comptebrumtn, "Sens Ecriture": "D",  "Montant": compta,"Libelle":libelle52})
              
                
            #ligne orange   
            elif operateur == "ORANGE":
                
                lignes.append({"Journal":Code, "Date": operation,"Mouvement":"OD", "Compte": comptefr , "Sens Ecriture": "D","Type Ecriture":"", "Montant": commission,"Auxil/Analyst": "","Libelle":libelle63})
                lignes.append({"Journal":Code, "Date": operation,"Mouvement":"OD", "Compte": comptefr , "Sens Ecriture": "D", "Type Ecriture":"A","Montant": commission,"Auxil/Analyst": site,"Libelle":libelle63})
                
                
                lignes.append({"Journal":Code, "Date": operation,"Mouvement":"OD", "Compte": compteom, "Sens Ecriture": "C", "Type Ecriture":"X", "Montant": montant,"Auxil/Analyst": auxi ,"Libelle":libelle55 })
                
                lignes.append({"Journal":Code, "Date": operation,"Mouvement":"OD", "Compte": comptebrumoov_om, "Sens Ecriture": "D",  "Montant": compta,"Libelle":libelle52})
              
              
            #ligne wave  
            elif operateur == "WAVE":
                
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "C", "Compte": "558015", "Montant": montant})
                
                lignes.append({"code":Code,"Operateur": operateur, "Site": site, "Type": "D", "Compte": "6327", "Montant": commission})
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
   # df_final.to_excel(fichier_export, index=False)
   

    df_final.to_csv(
    fichier_export,
    sep=";",
    index=False,
    header=False,
    encoding="utf-8-sig",
    mode="a"
    )
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
