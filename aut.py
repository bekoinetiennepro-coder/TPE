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
INTERVALLE_CHECK = 10  # secondes

# ----------------- NOM DU FICHIER D'EXPORT -----------------
def nom_fichier_export():
    date_str = datetime.now().strftime("%Y-%m-%d")
    return os.path.join(DOSSIER_EXPORT, f"resultat_{date_str}.xlsx")

# ----------------- DATAFRAME GLOBAL -----------------
df_final = pd.DataFrame()

# ----------------- DÉTECTION DE L'OPÉRATEUR -----------------
def detecter_operateur_depuis_nom(fichier):
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

# ----------------- DÉTECTION ROBUSTE DE LA LIGNE D'EN-TÊTE -----------------
def detecter_ligne_entete(df_temp):
    mots_cles = ["site", "montant", "commission", "transaction", "identifiant", "marchand"]
    for i, row in df_temp.iterrows():
        texte = " ".join(str(x).lower() for x in row.tolist())
        score = sum(1 for mot in mots_cles if mot in texte)
        if score >= 2:
            return i
    return None

# ----------------- DÉTECTION AUTOMATIQUE DES COLONNES -----------------
def normaliser_colonnes(df, operateur):
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]

    colonnes_cibles = {
        "site": ["site","solde"],
        "montant": ["montant_brut", "montant"],
        "commission": ["commission", "com","frais"],
        "montant_à_comptabiliser": ["montant_a_comptabiliser", "net_a_comptabiliser", "montant_net"]
    }

    for cle, variantes in colonnes_cibles.items():
        for var in variantes:
            col_trouvee = next((c for c in df.columns if var in c), None)
            if col_trouvee:
                df.rename(columns={col_trouvee: cle}, inplace=True)
                break

    # Ajouter colonne opérateur si absente
    if "opérateur" not in df.columns:
        df["opérateur"] = operateur
    else:
        df["opérateur"].fillna(operateur, inplace=True)

    # Normaliser la commission (positif)
    if "commission" in df.columns:
        df["commission"] = pd.to_numeric(df["commission"], errors="coerce").fillna(0).abs()

    return df

# ----------------- TRAITEMENT D'UN FICHIER -----------------
def traiter_fichier(fichier):
    global df_final
    try:
        print(f"\n[INFO] Traitement du fichier : {fichier}")
        operateur_fichier = detecter_operateur_depuis_nom(fichier)
        print(f"[INFO] Opérateur détecté : {operateur_fichier}")

        df_temp = pd.read_excel(fichier, header=None, nrows=20)
        header_row = detecter_ligne_entete(df_temp)
        if header_row is None:
            print(f"[ERREUR] Impossible de détecter la ligne d'en-têtes pour {fichier}")
            print("[DEBUG] Aperçu du fichier :", df_temp.head(10))
            return

        df = pd.read_excel(fichier, header=header_row)
        df = normaliser_colonnes(df, operateur_fichier)

        # Vérification colonnes importantes
        colonnes_requises = ["site", "montant", "commission", "montant_à_comptabiliser", "opérateur"]
        manquantes = [c for c in colonnes_requises if c not in df.columns]
        if manquantes:
            print(f"[ERREUR] Colonnes manquantes dans {fichier} : {manquantes}")
            print("[DEBUG] Colonnes détectées :", df.columns.tolist())
            return

        # Agrégation par opérateur / site
        df_group = df.groupby(["opérateur", "site"], as_index=False).sum(numeric_only=True)

        # Génération des écritures comptables
        lignes = []
        for _, row in df_group.iterrows():
            site = row["site"]
            montant = row["montant"]
            commission = row["commission"]
            compta = row["montant_à_comptabiliser"]
            operateur = row["opérateur"].upper()
            Code = "C52"

            if operateur == "MOOV":
                lignes.extend([
                    {"code":Code,"Operateur":operateur,"Site":site,"Type":"C","Compte":"558013","Montant":montant},
                    {"code":Code,"Operateur":operateur,"Site":site,"Type":"C","Compte":"558013","Montant":montant,"TypeOp":"OD","Auxiliare":"TPE"+str(site)},
                    {"code":Code,"Operateur":operateur,"Site":site,"Type":"D","Compte":"6327","Montant":commission},
                    {"code":Code,"Operateur":operateur,"Site":site,"Type":"D","Compte":"521650","Montant":compta}
                ])
            elif operateur == "MTN":
                lignes.extend([
                    {"code":Code,"Operateur":operateur,"Site":site,"Type":"C","Compte":"558012","Montant":montant},
                    {"code":Code,"Operateur":operateur,"Site":site,"Type":"D","Compte":"6327","Montant":commission},
                    {"code":Code,"Operateur":operateur,"Site":site,"Type":"D","Compte":"531100","Montant":compta}
                ])
            elif operateur == "ORANGE":
                lignes.extend([
                    {"code":Code,"Operateur":operateur,"Site":site,"Type":"C","Compte":"558011","Montant":montant},
                    {"code":Code,"Operateur":operateur,"Site":site,"Type":"D","Compte":"6327","Montant":commission},
                    {"code":Code,"Operateur":operateur,"Site":site,"Type":"D","Compte":"521650","Montant":compta}
                ])
            elif operateur == "WAVE":
                lignes.extend([
                    {"code":Code,"Operateur":operateur,"Site":site,"Type":"C","Compte":"558015","Montant":montant},
                    {"code":Code,"Operateur":operateur,"Site":site,"Type":"D","Compte":"6327","Montant":commission},
                    {"code":Code,"Operateur":operateur,"Site":site,"Type":"D","Compte":"521450","Montant":compta}
                ])
            else:
                print(f"[AVERTISSEMENT] Opérateur non reconnu pour {fichier}: {operateur}")

        if lignes:
            df_final = pd.concat([df_final, pd.DataFrame(lignes)], ignore_index=True)

        shutil.move(fichier, os.path.join(DOSSIER_ARCHIVE, os.path.basename(fichier)))
        print(f"[OK] Fichier traité et archivé : {fichier}")

    except Exception as e:
        print(f"[ERREUR] Traitement du fichier {fichier} : {e}")

# ----------------- INSERTION DANS LA BDD -----------------
def inserer_donnees_bdd(df):
    if df.empty:
        print("[INFO] Aucune donnée à insérer dans la BDD.")
        return
    try:
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

# ----------------- EXPORT -----------------
def exporter_resultat():
    if df_final.empty:
        print("[INFO] Aucune donnée à exporter aujourd'hui.")
        return
    fichier_export = nom_fichier_export()
    df_final.to_excel(fichier_export, index=False)
    print(f"[EXPORT] Résultat exporté : {fichier_export}")

# ----------------- TRAITEMENT DES FICHIERS -----------------
def automatiser_import():
    fichiers = [os.path.join(DOSSIER_ENTREE, f) for f in os.listdir(DOSSIER_ENTREE)
                if f.lower().endswith((".xlsx", ".xls", ".xlsb"))]
    for fichier in fichiers:
        traiter_fichier(fichier)
    inserer_donnees_bdd(df_final)
    exporter_resultat()

# ----------------- SURVEILLANCE EN CONTINU -----------------
class SurveillanceHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith((".xlsx", ".xls", ".xlsb")):
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
