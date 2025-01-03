from flask import Flask, render_template
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font
from datetime import datetime
import os

app = Flask(__name__)

# Définir les chemins NAS
NAS_BASE_PATH = r"\\nas23\MRO_GlobalSupplyChain\LMPA\Plan de progrès LMP\Gestion de projets\2024 - PO Right First Time\Pré-mesure"
FILE_PATHS = {
    "clients": os.path.join(NAS_BASE_PATH, "Clients.txt"),
    "createurs": os.path.join(NAS_BASE_PATH, "Createur-OS.txt"),
    "sous_traitants": os.path.join(NAS_BASE_PATH, "Sous-Traitants.txt"),
    "categories": os.path.join(NAS_BASE_PATH, "list_po_issues.txt"),
}

# Lire un fichier en gérant les encodages
def read_file_from_nas(files_path):
    try:
        with open(files_path, 'r', encoding='utf-8') as file:
            return [line.strip() for line in file if line.strip()]
    except UnicodeDecodeError:
        try:
            with open(files_path, 'r', encoding='iso-8859-1') as file:
                return [line.strip() for line in file if line.strip()]
        except Exception as ee:
            print(f"Erreur lors de la lecture du fichier : {files_path}. Détails : {ee}")
            return []
    except FileNotFoundError:
        print(f"Erreur : Fichier introuvable -> {files_path}")
        return []  # Retourne une liste vide si le fichier est introuvable

# Vérifier si le NAS général est accessible et gérer les fichiers manquants
def get_data_from_nas():
    if not os.path.exists(NAS_BASE_PATH):
        return {
            "error_message": "Erreur : Le NAS principal n'est pas accessible. Veuillez vérifier votre connexion réseau.",
            "nas_link": NAS_BASE_PATH,
            "file_errors": [],
            "clients": [],
            "createurs": [],
            "flux": [],
            "sous_traitants": [],
            "categories": []
        }

    # Si le NAS est accessible, essayer de lire les fichiers
    file_errors = []
    data = {
        "clients": read_file_from_nas(FILE_PATHS["clients"]),
        "createurs": read_file_from_nas(FILE_PATHS["createurs"]),
        "flux": ['UNS', 'SERV'],
        "sous_traitants": read_file_from_nas(FILE_PATHS["sous_traitants"]),
        "categories": read_file_from_nas(FILE_PATHS["categories"]),
    }

    # Vérifier les fichiers manquants ou vides
    for key, file_path in FILE_PATHS.items():
        if not data[key]:
            file_errors.append(f"Le fichier {os.path.basename(file_path)} est introuvable ou vide.")

    return {
        "error_message": None,
        "nas_link": NAS_BASE_PATH,
        "file_errors": file_errors,
        **data,
    }


def enregistrer_donnees(numero_os, createur, num_po, client, flux, sous_traitant, categorie, site="WRS",
                        fichier="RESULTAT A TELECHARGER EN LOCAL.xlsx"):
    # Déterminer le chemin du fichier dans le répertoire local
    chemin_fichier = os.path.join(os.getcwd(), fichier)

    # Extraire la cause et la description de la catégorie
    if "-" in categorie:
        cause, description = map(str.strip, categorie.split("-", 1))
    else:
        cause, description = categorie, ""

    # Obtenir la date et l'utilisateur
    date_actuelle = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    utilisateur = os.getlogin()

    # Vérifier si le fichier existe, sinon le créer
    if not os.path.exists(chemin_fichier):
        wb = Workbook()
        ws = wb.active
        # Ajouter les en-têtes
        ws.append([
            "Date d'ouverture", "QUI", "OS", "CREE PAR", "PO", "CLIENTS",
            "FLUX", "Sous-traitant", "Cause", "description anomalie", "Site"
        ])
        # Mettre en forme les en-têtes
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
        wb.save(chemin_fichier)

    # Charger le fichier existant
    wb = load_workbook(chemin_fichier)
    ws = wb.active

    # Trouver la première ligne vide
    ligne_vide = ws.max_row + 1

    # Insérer les données
    ws.append([
        date_actuelle, utilisateur, numero_os, createur, num_po, client,
        flux, sous_traitant, cause, description, site
    ])

    # Ajuster la largeur des colonnes si nécessaire
    for col in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col)
        ws.column_dimensions[col_letter].width = max(len(str(ws.cell(row=1, column=col).value)) + 5, 15)

    # Sauvegarder le fichier
    wb.save(chemin_fichier)
    print(f"Données enregistrées dans '{fichier}' avec succès.")

@app.route('/')
def home():
    datas = get_data_from_nas()

    # Priorité : vérifier si le NAS est inaccessible
    if datas["error_message"]:
        return render_template(
            'nas_error.html',
            error_message=datas["error_message"],
            nas_link=datas["nas_link"]
        )

    # Vérifier s'il y a des erreurs liées aux fichiers
    if datas["file_errors"]:
        return render_template(
            'file_error.html',
            file_errors=datas["file_errors"],
            nas_contact="raphael.carabeuf@safrangroup.com"
        )

    # Sinon, afficher la page principale
    return render_template(
        'index.html',
        error_message=None,
        nas_link=datas["nas_link"],
        file_errors=datas["file_errors"],
        clients=datas["clients"],
        createurs=datas["createurs"],
        flux=datas["flux"],
        sous_traitants=datas["sous_traitants"],
        categories=datas["categories"]
    )


if __name__ == '__main__':
    app.run(debug=True)
