import pandas as pd       # outil pour travailler avec des tableaux de donnees (lire CSV, filtrer, grouper)
from pathlib import Path  # outil pour manipuler les chemins de fichiers
from datetime import datetime  # outil pour travailler avec les dates et heures

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# nom du fichier CSV qu'on va lire (le fichier brut genere par scan_metadata.py)
FICHIER_ENTREE  = "raw_file_metadata.csv"

# nom du fichier CSV qu'on va produire a la fin avec toutes les colonnes
# dans Excel : masquer les 9 premieres colonnes pour voir seulement les features
FICHIER_SORTIE  = "cleaned_file_metadata.csv"

# date du jour utilisee comme reference pour calculer les ages des fichiers
# on la stocke une seule fois pour que toutes les fonctions utilisent la meme date
DATE_REFERENCE = datetime.now()

# extensions a supprimer car elles correspondent a des fichiers systeme inutiles
# .tmp = temporaires | .log = journaux | .ini = configuration | .sys = systeme | .lnk = raccourcis
EXTENSIONS_A_SUPPRIMER = {".tmp", ".log", ".ini", ".sys", ".lnk"}

# correspondance entre extension et categorie de fichier
# permet de savoir si un fichier est un document, une image, une video, etc.
CARTE_CATEGORIES = {
    "document" : {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt", ".odt"},
    "image"    : {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".svg", ".webp", ".tiff", ".ico"},
    "video"    : {".mp4", ".avi", ".mkv", ".mov", ".wmv", ".flv", ".webm"},
    "audio"    : {".mp3", ".wav", ".aac", ".flac", ".ogg", ".wma"},
    "code"     : {".py", ".js", ".ts", ".java", ".c", ".cpp", ".cs", ".html", ".css", ".sql"},
    "archive"  : {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2"},
    "donnees"  : {".csv", ".json", ".xml", ".yaml", ".yml", ".db", ".sqlite"},
}

# ---------------------------------------------------------------------------
# Fonctions utilitaires — Nettoyage
# ---------------------------------------------------------------------------

# cette fonction retourne la categorie d'un fichier selon son extension
# Ex: ".pdf" -> "document" | ".jpg" -> "image" | ".mp4" -> "video"
# si aucune categorie ne correspond, elle retourne "autre"
def obtenir_categorie(extension):
    # on parcourt chaque categorie et ses extensions
    for categorie, extensions in CARTE_CATEGORIES.items():
        # si l'extension est dans cette categorie, on retourne la categorie
        if extension in extensions:
            return categorie
    # si aucune categorie ne correspond, on retourne "autre"
    return "autre"

# cette fonction retourne le repertoire racine a partir du chemin complet
# Ex: "C:/Users/AOmpa/Documents/rapport.pdf" -> "Users"
def obtenir_repertoire(chemin):
    try:
        # on decoupe le chemin en morceaux
        parties = Path(str(chemin)).parts
        # on retourne le deuxieme morceau (index 1) qui est le repertoire racine
        # si le chemin est trop court, on retourne "racine"
        return parties[1] if len(parties) > 1 else "racine"
    except Exception:
        # si le chemin est casse ou vide, on retourne "inconnu" au lieu de planter
        return "inconnu"

# ---------------------------------------------------------------------------
# Fonctions utilitaires — Feature Engineering
# ---------------------------------------------------------------------------

# cette fonction retourne le dossier parent du fichier (sans le nom du fichier)
# Ex: /home/user/backup/report/text.pdf -> /home/user/backup/report
def obtenir_dossier_parent(chemin):
    try:
        # Path(chemin).parent enleve le nom du fichier et retourne juste le dossier
        return str(Path(chemin).parent)
    except Exception:
        # si le chemin est casse ou vide, on retourne "inconnu" au lieu de planter
        return "inconnu"

# cette fonction retourne la profondeur du dossier (combien de niveaux il y a)
# Ex: /home/user/backup/report/text.pdf -> 4 (home, user, backup, report)
def calculer_profondeur(chemin):
    try:
        # .parts decoupe le chemin en morceaux
        # len() compte le nombre de morceaux
        # -1 pour ne pas compter la racine "/" qui n'est pas un vrai dossier
        return len(Path(chemin).parent.parts) - 1
    except Exception:
        # si le chemin est casse, on retourne 0 au lieu de planter
        return 0

# cette fonction classe le fichier selon sa taille
# Small : moins de 1 MB | Medium : entre 1 MB et 100 MB | Large : plus de 100 MB
def categoriser_taille(taille_octets):
    if taille_octets < 1_048_576:       # 1_048_576 = 1 MB en octets
        return "Small"
    elif taille_octets < 104_857_600:   # 104_857_600 = 100 MB en octets
        return "Medium"
    else:
        return "Large"

# cette fonction retourne la taille dans l'unite la plus adaptee (KB, MB ou GB)
# Ex: 800 octets -> "0.78 KB" | 1 500 000 octets -> "1.43 MB"
# :.2f signifie "arrondi a 2 decimales"
def formater_taille_lisible(taille_octets):
    if taille_octets >= 1_073_741_824:  # 1 GB en octets
        # on divise par 1 GB et on arrondit a 2 decimales
        return f"{taille_octets / 1_073_741_824:.2f} GB"
    elif taille_octets >= 1_048_576:    # 1 MB en octets
        # on divise par 1 MB et on arrondit a 2 decimales
        return f"{taille_octets / 1_048_576:.2f} MB"
    else:
        # on divise par 1024 pour convertir en KB
        return f"{taille_octets / 1_024:.2f} KB"

# cette fonction retourne le nombre de jours depuis la creation du fichier
# Ex: cree le 10 juin 2026 -> file_age = 2 (si aujourd'hui = 12 juin 2026)
def calculer_age_fichier(date_creation):
    try:
        # on soustrait la date de creation a la date d'aujourd'hui pour obtenir une duree
        # .replace(tzinfo=None) enleve le fuseau horaire pour eviter les erreurs
        delta = DATE_REFERENCE - date_creation.replace(tzinfo=None)
        # .days extrait le nombre de jours | max evite les nombres negatifs
        return max(delta.days, 0)
    except Exception:
        # si la date est invalide, on retourne -1
        return -1

# cette fonction retourne le nombre de jours depuis la derniere modification
# Ex: modifie le 12 mai 2026 -> days_since_modification = 31
def calculer_jours_depuis_modif(date_modification):
    try:
        # meme logique que calculer_age_fichier mais pour la date de modification
        delta = DATE_REFERENCE - date_modification.replace(tzinfo=None)
        return max(delta.days, 0)
    except Exception:
        # si la date est invalide, on retourne -1
        return -1

# ---------------------------------------------------------------------------
# Fonction principale
# Elle fait tout dans l'ordre : nettoyage, enrichissement, feature engineering, export
# ---------------------------------------------------------------------------

def nettoyer_donnees(chemin_entree, chemin_sortie):

    # on charge le fichier CSV brut genere par scan_metadata.py
    # dtype=str signifie qu'on lit tout comme du texte pour eviter les conversions automatiques
    df = pd.read_csv(chemin_entree, dtype=str)

    # on stocke le nombre total de lignes pour afficher le bilan a la fin
    total = len(df)
    print(f"Lignes chargees : {total}")

    # ------------------------------------------------------------------
    # NETTOYAGE
    # ------------------------------------------------------------------

    # Condition 1 — Suppression des lignes avec valeurs manquantes
    # si un fichier n'a pas de nom, de chemin ou de hash, il est inutilisable
    # on n'inclut pas "owner" car certains fichiers peuvent ne pas avoir de proprietaire
    df = df.dropna(subset=["nom_fichier", "chemin", "hash_md5", "taille_octets"])

    # Condition 2 — Suppression des fichiers de taille nulle ou negative
    # pd.to_numeric convertit la colonne en nombre pour pouvoir la comparer
    # errors="coerce" transforme les valeurs invalides en NaN au lieu de planter
    df = df.copy()
    df["taille_octets"] = pd.to_numeric(df["taille_octets"], errors="coerce")
    # on garde seulement les fichiers qui ont une taille superieure a 0
    df = df[df["taille_octets"] > 0]

    # Condition 3 — Suppression des lignes avec hash MD5 incorrect
    # un hash MD5 valide fait toujours exactement 32 caracteres
    # .str.strip() enleve les espaces | .str.len() compte les caracteres
    df = df[df["hash_md5"].str.strip().str.len() == 32]

    # Condition 4 — Suppression des lignes avec dates incoherentes
    # on convertit les dates en vrai format date que Python peut comparer
    # errors="coerce" transforme les dates invalides en NaT au lieu de planter
    df["date_creation"]     = pd.to_datetime(df["date_creation"],     errors="coerce")
    df["date_modification"] = pd.to_datetime(df["date_modification"], errors="coerce")
    # la date de modification ne peut pas etre anterieure a la date de creation
    # si c'est le cas, la donnee est erronee et on supprime la ligne
    df = df[df["date_modification"] >= df["date_creation"]]

    # Condition 5 — Normalisation des extensions en minuscule
    # .str.strip() enleve les espaces | .str.lower() met en minuscule
    # comme ca ".PDF" et ".pdf" sont traites de la meme facon
    df["extension"] = df["extension"].str.strip().str.lower()

    # Condition 6 — Suppression des extensions non pertinentes
    # .isin() verifie si l'extension est dans la liste EXTENSIONS_A_SUPPRIMER
    # "~" devant inverse le resultat : on garde ceux qui NE sont PAS dans la liste
    df = df[~df["extension"].isin(EXTENSIONS_A_SUPPRIMER)]

    # Condition 7 — Suppression des espaces dans les noms de fichiers
    # .str.strip() enleve les espaces en debut et fin de chaine
    # Ex: "  rapport.pdf  " devient "rapport.pdf"
    df["nom_fichier"] = df["nom_fichier"].str.strip()

    # Condition 8 — Suppression des chemins dupliques
    # un meme chemin ne peut pas apparaitre deux fois dans le tableau
    # .drop_duplicates() supprime les lignes en double sur la colonne "chemin"
    df = df.drop_duplicates(subset=["chemin"])

    # ------------------------------------------------------------------
    # ENRICHISSEMENT — colonnes calculees
    # ------------------------------------------------------------------

    # taille_ko : taille en kilo-octets (plus lisible que les octets bruts)
    # on divise la taille en octets par 1024 et on arrondit a 2 decimales
    df["taille_ko"]  = (df["taille_octets"] / 1024).round(2)

    # categorie : categorie du fichier selon son extension
    # on appelle la fonction obtenir_categorie pour chaque ligne de la colonne "extension"
    df["categorie"]  = df["extension"].apply(obtenir_categorie)

    # repertoire : repertoire racine du fichier
    # on appelle la fonction obtenir_repertoire pour chaque ligne de la colonne "chemin"
    df["repertoire"] = df["chemin"].apply(obtenir_repertoire)

    # ------------------------------------------------------------------
    # FEATURE ENGINEERING
    # ------------------------------------------------------------------

    # Feature 1 — folder_path : dossier parent du fichier
    # on prend la colonne "chemin" et on applique la fonction obtenir_dossier_parent
    # Ex: /home/user/backup/report/text.pdf -> /home/user/backup/report
    df["folder_path"] = df["chemin"].apply(obtenir_dossier_parent)

    # Feature 2 — folder_depth : profondeur du dossier parent
    # on prend la colonne "chemin" et on applique la fonction calculer_profondeur
    # Ex: /home/user/backup/report/text.pdf -> 4
    df["folder_depth"] = df["chemin"].apply(calculer_profondeur)

    # Feature 3 — size_category : categorie de taille du fichier
    # on prend la colonne "taille_octets" et on applique la fonction categoriser_taille
    # Small : < 1 MB | Medium : 1 MB - 100 MB | Large : > 100 MB
    df["size_category"] = df["taille_octets"].apply(categoriser_taille)

    # Feature 4 — file_age : nombre de jours depuis la creation du fichier
    # on prend la colonne "date_creation" et on applique la fonction calculer_age_fichier
    # Ex: cree le 10 juin 2026 -> file_age = 2
    df["file_age"] = df["date_creation"].apply(calculer_age_fichier)

    # Feature 5 — days_since_modification : jours depuis la derniere modification
    # on prend la colonne "date_modification" et on applique calculer_jours_depuis_modif
    # Ex: modifie le 12 mai 2026 -> days_since_modification = 31
    df["days_since_modification"] = df["date_modification"].apply(calculer_jours_depuis_modif)

    # Feature 6 — file_size_readable : taille lisible en KB / MB / GB
    # on prend la colonne "taille_octets" et on applique la fonction formater_taille_lisible
    # Ex: 800 octets -> "0.78 KB" | 1 500 000 octets -> "1.43 MB"
    df["file_size_readable"] = df["taille_octets"].apply(formater_taille_lisible)

    # ------------------------------------------------------------------
    # Reformatage final des dates en texte lisible
    # on force la reconversion en datetime avant le strftime pour eviter les erreurs
    # "%Y-%m-%d %H:%M:%S" donne par exemple "2026-06-13 10:35:22"
    # ------------------------------------------------------------------
    df["date_creation"]     = pd.to_datetime(df["date_creation"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M:%S")
    df["date_modification"] = pd.to_datetime(df["date_modification"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M:%S")

    # ------------------------------------------------------------------
    # ORDRE DES COLONNES
    # on met d'abord les 9 colonnes a masquer dans Excel (colonnes brutes du scan)
    # ensuite les colonnes nettoyees + features (colonnes a garder visibles)
    # dans Excel : selectionner colonnes A a I -> clic droit -> Masquer
    # ------------------------------------------------------------------
    colonnes_ordonnees = [
        # --- 9 colonnes a masquer dans Excel ---
        "identifiant",        # numero unique du fichier
        "nom_fichier",        # nom du fichier
        "extension",          # extension du fichier
        "chemin",             # chemin complet (necessaire pour exact_duplicates.py)
        "taille_octets",      # taille en octets (necessaire pour exact_duplicates.py)
        "date_creation",      # date de creation
        "date_modification",  # date de modification
        "hash_md5",           # empreinte digitale (necessaire pour exact_duplicates.py)
        "owner",              # proprietaire du fichier

        # --- colonnes a garder visibles dans Excel ---
        "taille_ko",                # taille en kilo-octets
        "categorie",                # categorie du fichier (document, image, video, etc.)
        "repertoire",               # repertoire racine du fichier
        "folder_path",              # dossier parent du fichier
        "folder_depth",             # profondeur du dossier
        "size_category",            # Small / Medium / Large
        "file_age",                 # nombre de jours depuis la creation
        "days_since_modification",  # nombre de jours depuis la derniere modification
        "file_size_readable",       # taille lisible en KB / MB / GB
    ]

    # on reordonne le tableau selon l'ordre defini ci-dessus
    df = df[colonnes_ordonnees]

    # ------------------------------------------------------------------
    # EXPORT — un seul fichier avec toutes les colonnes
    # index=False : on n'ecrit pas les numeros de lignes dans le CSV
    # encoding="utf-8" : format d'encodage pour que les accents s'affichent correctement
    # ------------------------------------------------------------------
    df.to_csv(chemin_sortie, index=False, encoding="utf-8")

    # bilan final
    print(f"Fichier exporte    : {chemin_sortie}")
    print(f"Lignes totales     : {total}")
    print(f"Lignes conservees  : {len(df)}")
    print(f"Lignes supprimees  : {total - len(df)}")
    print(f"Colonnes totales   : {len(df.columns)}")
    print(f"\nDans Excel : masquer colonnes A a I pour voir seulement les features")

# ---------------------------------------------------------------------------
# Point d'entree
# ---------------------------------------------------------------------------

def principal():
    # on lance la fonction principale avec le fichier d'entree et de sortie
    nettoyer_donnees(FICHIER_ENTREE, FICHIER_SORTIE)

if __name__ == "__main__":
    # cette ligne dit "si on lance directement ce fichier, appelle la fonction principal"
    # c'est la porte d'entree du programme
    principal()