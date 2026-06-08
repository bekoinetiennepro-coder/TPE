import pandas as pd
import os
import shutil
import time
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

BASE_RESEAU = r"\\192.109.69.46\GenerationTPE"
# ----------------- CONFIGURATION -----------------
#DOSSIER_ENTREE = os.path.join(BASE_RESEAU, "input_Constation_des_versements_TPE")
DOSSIER_ENTREE = os.path.join(BASE_RESEAU, "Archives")

DOSSIER_ARCHIVE = os.path.join(BASE_RESEAU, "archive_Constation_des_versements_TPE")
DOSSIER_EXPORT = os.path.join(BASE_RESEAU,"Constation_des_versements_TPE/MAI")

#INTERVALLE_CHECK = 10  # secondes pour la surveillance continue

# ----------------- CONFIGURATION -----------------

# ----------------- NOM DU FICHIER D'EXPORT -----------------
def nom_fichier_export():
    date_str = datetime.now().strftime("%Y-%m-%d")
    return os.path.join(DOSSIER_EXPORT, f"Contre_Partie_solde.txt")


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
    elif "CARTE BANCAIRE" in nom_fichier:
        return "CARTE BANCAIRE"
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
        manquantes = []
        colonnes_requises = []
       
        if operateur_fichier == "MOOV":
            colonnes_requises = ["opérateur", "site", "montant", "commission", "montant_a_comptabiliser", "date_operation"]
           
        elif operateur_fichier == "MTN":
            colonnes_requises = ["opérateur", "site", "montant", "commission", "montant_a_comptabiliser","date_operation"]
               
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
            #  Normalisation des valeurs pour ORANGE
            if "commission" in df.columns:
                # convertir en numérique proprement
                df["commission"] = pd.to_numeric(df["commission"], errors="coerce").fillna(0)
                # remettre en positif (si valeurs négatives)
                df["commission"] = df["commission"].abs()


        elif operateur_fichier == "WAVE":
            #  Recherche intelligente des colonnes spécifiques à Orange
            colonnes_requises = ["opérateur", "site", "montant", "commission", "montant_a_comptabiliser","date_operation"]

            # Trouver la colonne du montant (ex: "crédit" ou "credit")
            col_montant = next((c for c in df.columns if "montant_brut" in c ), None)
            
            col_date = next((c for c in df.columns if "horodatage" in c ),None)
            
            if col_montant:
                df.rename(columns={col_montant: "montant"}, inplace=True)
            
            if col_date:
                df.rename(columns={col_date: "date_operation"},inplace=True)
                
            
            if "montant" not in df.columns or "commission" not in df.columns:
                print(f"[ERREUR] Colonnes non trouvées pour WAVE : montant={col_montant}, commission={col_commission}")
                print("[DEBUG] Colonnes détectées :", df.columns.tolist())
                return
            
        # ---------------- SPECIAL : TRAITEMENT CARTE BANCAIRE ----------------
        
        if operateur_fichier == "CARTE BANCAIRE":

            # Toujours utiliser Code Site
            df["site"] = df["code_site"]

            # Nettoyage
            df.drop(columns=[c for c in ["code_site"] if c in df.columns], inplace=True)

            df.rename(columns={
                "date": "date_operation",
                "carte_bancaire": "montant"
            }, inplace=True)

            df["commission"] = 0
            df["montant_a_comptabiliser"] = df["montant"]

        else:
            colonnes_requises = ["opérateur", "site", "montant", "commission", "montant_a_comptabiliser"]

        # ----------------- Vérification finale -----------------
        manquantes = [c for c in colonnes_requises if c not in df.columns]
        if manquantes:
            print(f"[ERREUR] Colonnes manquantes dans {fichier} : {manquantes}")
            print("[DEBUG] Colonnes présentes :", df.columns.tolist())
            return

        # Sauvegarder la colonne date_operation (si présente)
        if "date_operation" in df.columns:
            
            if operateur_fichier == "WAVE":
        # WAVE : format ISO 8601 avec timezone
                df["date_operation"] = pd.to_datetime(
                    df["date_operation"],
                    errors="coerce",
                    #utc=True  
                    dayfirst=True # WAVE fournit des timestamps en UTC, on peut les convertir en heure locale si besoin exemeple :2026-11-12T14:23:45Z → 11/12/2026
                ).dt.date
            else:
                # MOOV / MTN / ORANGE / CARTE BANCAIRE
                df["date_operation"] = pd.to_datetime(
                    df["date_operation"],
                    dayfirst=True,
                    errors="coerce"
                ).dt.date

            # Format jj/mm/aaaa
            df["date_operation"] = df["date_operation"].apply(lambda x: x.strftime("%d/%m/%Y") if pd.notnull(x) else "")
            # Sauvegarder la date par site avant groupby
            df_date = df.groupby(["opérateur", "site"])["date_operation"].first().reset_index()
        else:
            df_date = None
        for col in ["montant", "commission", "montant_a_comptabiliser"]:
                if col in df.columns:

                    df[col] = (
                        df[col]
                        .astype(str)
                        .str.replace("\xa0", "", regex=False)  # espace insécable
                        .str.replace(" ", "", regex=False)
                        .str.replace(",", ".", regex=False)
                        .str.strip()
                    )

                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)    

        # Groupby pour sommer les colonnes numériques
        df_group = df.groupby(["opérateur", "site"], as_index=False).sum(numeric_only=True)
       
       # Arrondir proprement après le groupby
        df_group["montant"] = df_group["montant"].round(2)
        df_group["commission"] = df_group["commission"].round(2)
        df_group["montant_a_comptabiliser"] = df_group["montant_a_comptabiliser"].round(2)
        
        # Réattacher la colonne date_operation
        if df_date is not None:
            df_group = df_group.merge(df_date, on=["opérateur", "site"], how="left")

        # 6️⃣ Créer la colonne periode à partir de date_operation
        if "date_operation" in df_group.columns:
            # Convertir en datetime
           # df_group["date_operation_dt"] = pd.to_datetime(df_group["date_operation"], format="%d/%m/%Y", errors="coerce")
            df_group["date_operation_dt"] = pd.to_datetime(
                df_group["date_operation"],
                dayfirst=True,
                errors="coerce"
            )
            # Créer periode au format jj-mmm-aa (ex: 12-nov-25)
            # df_group["periode"] = df_group["date_operation_dt"].dt.strftime("%d-%b-%y").str.lower()
            # # Supprimer colonne temporaire
            # df_group.drop(columns=["date_operation_dt"], inplace=True)
            df_group["periode"] = (
            df_group["date_operation_dt"]
            .dt.strftime("%d-%b-%y")
            .str.lower()
            .replace({
                "jan": "jan",
                "feb": "fév",
                "mar": "mars",
                "apr": "avr",
                "may": "mai",
                "jun": "juin",
                "jul": "juil",
                "aug": "août",
                "sep": "sept",
                "oct": "oct",
                "nov": "nov",
                "dec": "déc"
            }, regex=True)
            )

            # Supprimer colonne temporaire
            df_group.drop(columns=["date_operation_dt"], inplace=True)

        # ----------------- Génération des écritures comptables -----------------
       # ----------------- Génération des écritures comptables -----------------
        lignes = []

        ordre_operateurs = ["WAVE", "MTN", "MOOV", "ORANGE", "CARTE BANCAIRE"]

        comptes_operateurs = {
            "MOOV": 558013,
            "MTN": 558012,
            "ORANGE": 558011,
            "WAVE": 558015,
            "CARTE BANCAIRE": 558014
        }

        # 🔥 TRAITEMENT PAR SITE
        for site, df_site in df_group.groupby("site"):

            total_site = 0
            date_operation = df_site["date_operation"].iloc[0]
            periode = df_site["periode"].iloc[0]
            auxi = f"TPE{site}"
            Code = "C52"
               
            # 🔹 DEBITS : un par opérateur existant
            for operateur in ordre_operateurs:
                ligne_op = df_site[df_site["opérateur"].str.upper() == operateur]

                if ligne_op.empty:
                    continue

                montant = ligne_op["montant"].sum()
                total_site += montant

                compte = comptes_operateurs[operateur]
                libelle = f"{site} Collecte Vente TPE {operateur} du {date_operation}"
                if montant == 0 or pd.isna(montant):
                 continue    
                lignes.append({
                    "Journal": Code,
                    "Date": date_operation,
                    "Mouvement": "OD",
                    "Compte": compte,
                    "Sens Ecriture": "D",
                    "Type Ecriture": "",
                    "Montant": montant,
                    "Auxil/Analyst": "",
                    "Periode": periode,
                    "Libelle": libelle
                })

            # 🔹 CREDIT UNIQUE PAR SITE
            if total_site > 0:
                libelle_credit = f"{site} Collecte Globale TPE du {date_operation}"

                lignes.append({
                    "Journal": Code,
                    "Date": date_operation,
                    "Mouvement": "OD",
                    "Compte": 55801,
                    "Sens Ecriture": "C",
                    "Type Ecriture": "X",
                    "Montant": total_site,
                    "Auxil/Analyst": auxi,
                    "Periode": periode,
                    "Libelle": libelle_credit
                })

        # ----------------- Ajout au dataframe global -----------------
        if lignes:
            df_final = pd.concat([df_final, pd.DataFrame(lignes)], ignore_index=True)

        # ----------------- Déplacement du fichier -----------------
        shutil.move(fichier, os.path.join(DOSSIER_ARCHIVE, os.path.basename(fichier)))
        print(f"[OK] Fichier traité et archivé : {fichier}")

    except Exception as e:
        print(f"[ERREUR] Traitement du fichier {fichier} : {e}")



# ----------------- EXPORT DU JOUR -----------------
def exporter_resultat():
    if df_final.empty:
        print("[INFO] Aucune donnée à exporter aujourd'hui.")
        return

    if "Date" in df_final.columns:
        date_operation = (
            pd.to_datetime(
                df_final["Date"],
                dayfirst=True,      
                errors="coerce"
            )
            .dropna()
            .iloc[0]
            .strftime("%d-%m-%Y")
        )
    else:
        date_operation = datetime.now().strftime("%d-%m-%Y")

    fichier_export = f"{nom_fichier_export()}_{date_operation}.txt"

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

    exporter_resultat()

# ----------------- SURVEILLANCE EN CONTINU -----------------
# class SurveillanceHandler(FileSystemEventHandler):
#     def on_created(self, event):
#         if not event.is_directory and event.src_path.lower().endswith((".xlsx", ".xls",".xlsb")):
#             traiter_fichier(event.src_path)
#             exporter_resultat()

# def surveillance_continue():
#     event_handler = SurveillanceHandler()
#     observer = Observer()
#     observer.schedule(event_handler, path=DOSSIER_ENTREE, recursive=False)
#     observer.start()
#     print("[INFO] Surveillance continue activée...")
#     try:
#         while True:
#             time.sleep(INTERVALLE_CHECK)
#     except KeyboardInterrupt:
#         observer.stop()
#     observer.join()

# ----------------- MAIN -----------------
if __name__ == "__main__":
    automatiser_import()
   # surveillance_continue()
