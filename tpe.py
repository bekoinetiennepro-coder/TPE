import pandas as pd
import os
import shutil
import time
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ================= CONFIGURATION =================
# ----------------- CONFIGURATION -----------------

BASE_RESEAU = r"\\192.109.69.46\Compta_Reglement\TPE"
# ----------------- CONFIGURATION -----------------
DOSSIER_ENTREE = os.path.join(BASE_RESEAU, "brutTpe")
DOSSIER_ARCHIVE = os.path.join(BASE_RESEAU, "Archives")
# ================= DOSSIER EXPORT DYNAMIQUE =================
def obtenir_dossier_export(date_operation):

    try:
        # convertir la date
        date_dt = pd.to_datetime(
            date_operation,
            dayfirst=True,
            errors="coerce"
        )

        # nom du mois en français
        mois_fr = {
            1: "JANVIER",
            2: "FEVRIER",
            3: "MARS",
            4: "AVRIL",
            5: "MAI",
            6: "JUIN",
            7: "JUILLET",
            8: "AOUT",
            9: "SEPTEMBRE",
            10: "OCTOBRE",
            11: "NOVEMBRE",
            12: "DECEMBRE"
        }

        nom_mois = mois_fr[date_dt.month]

        # chemin dynamique
        dossier = os.path.join(
            BASE_RESEAU,
            "data_tpe",
            nom_mois
        )

        # création auto du dossier
        os.makedirs(dossier, exist_ok=True)

        return dossier

    except Exception as e:
        print(f"[ERREUR DOSSIER EXPORT] {e}")

        # fallback
        dossier = os.path.join(BASE_RESEAU, "data_tpe")
        os.makedirs(dossier, exist_ok=True)

        return dossier


INTERVALLE_CHECK = 10  # secondes pour la surveillance continue

# ----------------- NOM DU FICHIER D'EXPORT -----------------
def nom_fichier_export(date_operation):
    
    dossier_export = obtenir_dossier_export(date_operation)

    return os.path.join(
        dossier_export,
        "Contre_Partie.txt"
    )

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
            #  Recherche intelligente des colonnes spécifiques à Orange
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
                df["date_operation"] = pd.to_datetime(
                    df["date_operation"],
                    errors="coerce",
                    #utc=True  # WAVE fournit des timestamps en UTC, on peut les convertir en heure locale si besoin exemeple :2026-11-12T14:23:45Z → 11/12/2026
                    dayfirst=True
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

        # Groupby pour sommer les colonnes numériques
        df_group = df.groupby(["opérateur", "site"], as_index=False).sum(numeric_only=True)
       
       # Arrondir proprement après le groupby
        df_group["montant"] = df_group["montant"].round(2)
        df_group["commission"] = df_group["commission"].round(2)
        df_group["montant_a_comptabiliser"] = df_group["montant_a_comptabiliser"].round(2)
        
        # Réattacher la colonne date_operation
        if df_date is not None:
            df_group = df_group.merge(df_date, on=["opérateur", "site"], how="left")

        # 6️Créer la colonne periode à partir de date_operation
        if "date_operation" in df_group.columns:
            # Convertir en datetime
           # df_group["date_operation_dt"] = pd.to_datetime(df_group["date_operation"], format="%d/%m/%Y", errors="coerce")
            df_group["date_operation_dt"] = pd.to_datetime(
                df_group["date_operation"],
                dayfirst=True,
                errors="coerce"
            )
            # Créer periode au format jj-mmm-aa (ex: 12-nov-25)
            df_group["periode"] = df_group["date_operation_dt"].dt.strftime("%d-%b-%y").str.lower()
            # Supprimer colonne temporaire
            df_group.drop(columns=["date_operation_dt"], inplace=True)

        # ----------------- Génération des écritures comptables -----------------
       # ----------------- CUMUL PAR DATE -----------------
        df_cumul = df_group.groupby(
            ["date_operation", "opérateur"],
            as_index=False
        )["montant_a_comptabiliser"].sum()

        df_cumul.rename(columns={
            "montant_a_comptabiliser": "compta_cumule"
        }, inplace=True)

        # merge
        df_group = df_group.merge(
            df_cumul,
            on=["date_operation", "opérateur"],
            how="left"
        )

        # ----------------- GENERATION DES ECRITURES -----------------
        lignes = []
        deja_fait = set()

        for _, row in df_group.iterrows():
            site = row["site"]
            montant = row["montant"]
            commission = row["commission"]
            compta = row["compta_cumule"]
            operation = row["date_operation"]
            operateur = row["opérateur"].upper()
            periode = row["periode"]

            Code = "C52"

            libelle55 = f"{site} Collecte Globale TPE {operateur} du {operation}"
            libelle52 = f"{site} Collecte Vente TPE du {operation}"
            libelle63 = f"{site} Commission {operateur} sur Collecte TPE du {operation}"

            comptmoov = 5584
            comptebrumoov_om = 521650
            comptemtn = 5583
            comptebrutom=521651
            comptebrumtn = 531100
            compteom = 5582
            comptewa = 5586
            comptebruwav = 521450
            comptecartebancaire = 5584
            comptecartebancairebrut = 558014
            comptefr = 632700

            auxi = "TPE" + str(site)

            # ---------------- COMMISSIONS (PAR SITE) ----------------
            lignes.append({
                "Journal": Code, "Date": operation, "Mouvement": "OD",
                "Compte": comptefr, "Sens Ecriture": "D",
                "Type Ecriture": "", "Montant": commission,
                "Auxil/Analyst": "", "Periode": periode, "Libelle": libelle63
            })

            lignes.append({
                "Journal": Code, "Date": operation, "Mouvement": "OD",
                "Compte": comptefr, "Sens Ecriture": "D",
                "Type Ecriture": "A", "Montant": commission,
                "Auxil/Analyst": site, "Periode": periode, "Libelle": libelle63
            })

            # ---------------- DEBIT CUMULÉ (UNE FOIS PAR DATE) ----------------
            key = (operation, operateur)

            if key not in deja_fait:

                if operateur == "MOOV":
                    compte_debit = comptebrumoov_om
                elif operateur == "ORANGE":
                    compte_debit = comptebrutom  # ou un compte spécifique ORANGE si tu veux

                elif operateur == "MTN":
                    compte_debit = comptebrumtn
                elif operateur == "WAVE":
                    compte_debit = comptebruwav
                elif operateur == "CARTE BANCAIRE":
                    compte_debit = comptecartebancairebrut
                else:
                    compte_debit = None

                if compte_debit:
                    lignes.append({
                        "Journal": Code,
                        "Date": operation,
                        "Mouvement": "OD",
                        "Compte": compte_debit,
                        "Sens Ecriture": "D",
                        "Type Ecriture": "",
                        "Montant": compta,
                        "Auxil/Analyst": "",
                        "Periode": periode,
                        "Libelle": libelle52
                    })

                deja_fait.add(key)

            # ---------------- CREDIT (PAR SITE) ----------------
            if operateur == "MOOV":
                compte_credit = comptmoov
            elif operateur == "MTN":
                compte_credit = comptemtn
            elif operateur == "ORANGE":
                compte_credit = compteom
            elif operateur == "WAVE":
                compte_credit = comptewa
            elif operateur == "CARTE BANCAIRE":
                if montant == 0 or pd.isna(montant):
                    continue
                compte_credit = comptecartebancaire
            else:
                print(f"[AVERTISSEMENT] Opérateur non reconnu : {operateur}")
                continue

            lignes.append({
                "Journal": Code, "Date": operation, "Mouvement": "OD",
                "Compte": compte_credit, "Sens Ecriture": "C",
                "Type Ecriture": "X",
                "Montant": montant,
                "Auxil/Analyst": auxi,
                "Periode": periode,
                "Libelle": libelle55
            })
        # ----------------- Ajout au dataframe global -----------------
        if lignes:
            df_final = pd.concat([df_final, pd.DataFrame(lignes)], ignore_index=True)

        # ----------------- Déplacement du fichier -----------------
        shutil.move(fichier, os.path.join(DOSSIER_ARCHIVE, os.path.basename(fichier)))
        print(f"[OK] Fichier traité et archivé : {fichier}")

    except Exception as e:
        print(f"[ERREUR] Traitement du fichier {fichier} : {e}")


def cumuler_debit(df):
    comptes_cibles = [521650,521651, 521450, 531100,558014]  # comptes à cumuler

    df = df.copy()

    #  Sécurisation des types
    df["Compte"] = pd.to_numeric(df["Compte"], errors="coerce")
    df["Sens Ecriture"] = df["Sens Ecriture"].astype(str).str.strip()
    df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")

    #  Filtrer uniquement les lignes à cumuler (DEBIT)
    df_debit = df[
        (df["Compte"].isin(comptes_cibles)) &
        (df["Sens Ecriture"] == "D")
    ].copy()

    if df_debit.empty:
        print("[INFO] Rien à cumuler")
        return df

    #  DEBUG utile
    print("DEBUG df_debit :")
    print(df_debit[["Date", "Compte", "Montant"]].head(20))

    #  Garder les autres lignes
    df_autres = df[
        ~(
            (df["Compte"].isin(comptes_cibles)) &
            (df["Sens Ecriture"] == "D")
        )
    ]

    #  GROUPBY SIMPLIFIÉ (clé du problème)
    df_group = df_debit.groupby(
        ["Date", "Compte"],
        as_index=False
    )["Montant"].sum()

    #  Reconstruction des colonnes fixes
    df_group["Journal"] = "C52"
    df_group["Mouvement"] = "OD"
    df_group["Sens Ecriture"] = "D"
    df_group["Type Ecriture"] = ""

    # récupérer la période correcte
    df_periode = df_debit.groupby(["Date", "Compte"])["Periode"].first().reset_index()
    df_group = df_group.merge(df_periode, on=["Date", "Compte"], how="left")

    #  Mapping opérateur
    mapping = {
        521650: "MOOV",
        521651: "ORANGE",
        521450: "WAVE",
        531100: "MTN",
        558014: "CARTE BANCAIRE"
        
        
    }

    df_group["operateur"] = df_group["Compte"].map(mapping)

    #  Libellé global
    df_group["Libelle"] = (
        "Collecte Globale TPE " +
        df_group["operateur"] +
        " du " + df_group["Date"]
    )

    df_group["Auxil/Analyst"] = ""

    #  Réorganiser les colonnes comme ton export initial
    df_group = df_group[
        ["Journal", "Date", "Mouvement", "Compte", "Sens Ecriture",
         "Type Ecriture", "Montant", "Auxil/Analyst", "Periode", "Libelle"]
    ]

    #  Fusion finale
    df_final_new = pd.concat([df_autres, df_group], ignore_index=True)

    print("DEBUG comptes finaux :", df_final_new["Compte"].unique())

    return df_final_new

def exporter_resultat():
    if df_final.empty:
        print("[INFO] Aucune donnée à exporter aujourd'hui.")
        return

    if "Date" in df_final.columns:
        date_operation = (
            pd.to_datetime(
                df_final["Date"],
                dayfirst=True,      #  CORRECTION CRITIQUE
                errors="coerce"
            )
            .dropna()
            .iloc[0]
            .strftime("%d-%m-%Y")
        )
    else:
        date_operation = datetime.now().strftime("%d-%m-%Y")

    #fichier_export = f"{nom_fichier_export()}_{date_operation}.txt"
    fichier_export = f"{nom_fichier_export(date_operation)}_{date_operation}.txt"
    df_to_export = cumuler_debit(df_final)

    df_to_export.to_csv(
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


# ----------------- MAIN -----------------
if __name__ == "__main__":
    automatiser_import()
    #surveillance_continue()
