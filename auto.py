import pandas as pd
import os
import shutil
import time
from datetime import datetime
from config import get_connection
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ----------------- CONFIGURATION -----------------
DOSSIER_ENTREE = r"C:\MonProjetCompta\input"
DOSSIER_ARCHIVE = r"C:\MonProjetCompta\archive"
DOSSIER_EXPORT = r"C:\MonProjetCompta\export"
INTERVALLE_CHECK = 10  # en secondes pour la surveillance continue

# Nom du fichier d'export du jour
def nom_fichier_export():
    date_str = datetime.now().strftime("%Y-%m-%d")
    return os.path.join(DOSSIER_EXPORT, f"resultat_{date_str}.xlsx")

# DataFrame global pour accumuler les écritures du jour
df_final = pd.DataFrame()

# ----------------- TRAITEMENT D'UN FICHIER -----------------
def traiter_fichier(fichier):
    global df_final
    try:
        print(f"[INFO] Traitement du fichier : {fichier}")
        df_temp = pd.read_excel(fichier, header=None)
        header_row = None
        for i, row in df_temp.iterrows():
            line_str = " ".join(str(x).lower() for x in row.tolist())
            if "site" in line_str or "Site" in line_str:
                
                header_row = i
                break
        if header_row is None:
            print(f"[ERREUR] Impossible de détecter la ligne d'en-têtes pour {fichier}")
            return

        df = pd.read_excel(fichier, header=header_row)
        df.columns = [col.strip().lower().replace(" ", "_") for col in df.columns]

        colonnes_requises = ["opérateur", "site", "montant", "commission", "comptabiliser"]
        manquantes = [c for c in colonnes_requises if c not in df.columns]
        if manquantes:
            print(f"[ERREUR] Colonnes manquantes dans {fichier} : {manquantes}")
            return

        # Regroupement par opérateur et site
        df_group = df.groupby(["opérateur", "site"], as_index=False).sum(numeric_only=True)

        # Génération des écritures comptables
        lignes = []
        for _, row in df_group.iterrows():
            site = row["site"]
            montant = row["montant"]
            commission = row["commission"]
            compta = row["comptabiliser"]
            operateur = row["opérateur"]

            if operateur.upper() == "MOOV":
                lignes.append({"Operateur": operateur, "Site": site, "Type": "C", "Compte": "55801", "Montant": montant})
                lignes.append({"Operateur": operateur, "Site": site, "Type": "D", "Compte": "6327", "Montant": commission})
                lignes.append({"Operateur": operateur, "Site": site, "Type": "D", "Compte": "521650", "Montant": compta})
            elif operateur.upper() == "MTN":
                lignes.append({"Operateur": operateur, "Site": site, "Type": "C", "Compte": "55801", "Montant": montant})
                lignes.append({"Operateur": operateur, "Site": site, "Type": "D", "Compte": "6327", "Montant": commission})
                lignes.append({"Operateur": operateur, "Site": site, "Type": "D", "Compte": "531100", "Montant": compta})

        df_final = pd.concat([df_final, pd.DataFrame(lignes)], ignore_index=True)

        # Déplacer le fichier traité dans l'archive
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

# ----------------- TRAITEMENT DE TOUS LES FICHIERS DU DOSSIER -----------------
def automatiser_import():
    fichiers = [os.path.join(DOSSIER_ENTREE, f) for f in os.listdir(DOSSIER_ENTREE)
                if f.lower().endswith((".xlsx", ".xls"))]
    for fichier in fichiers:
        traiter_fichier(fichier)

    inserer_donnees_bdd(df_final)
    exporter_resultat()

# ----------------- SURVEILLANCE EN CONTINU -----------------
class SurveillanceHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith((".xlsx", ".xls")):
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
    # 1️⃣ Traitement initial des fichiers déjà présents
    automatiser_import()
    # 2️⃣ Lancement de la surveillance continue
    surveillance_continue()
