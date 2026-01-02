#On import les modules openpyxl et os

#On va l'utiliser pour ouvrir les fichier xlsx et ainsi les utiliser
import openpyxl

#il va servir a donner les chemins d'accès pour aller aux différents fichiers que nous allons utiliser
import os


#CHEMIN DES FICHIERS EXCEL
#Chemin pour accéder au fichier excel des coefs
fichier_de_ref = '/workspaces/SAE105_EXCEL_-_VALIDATION_UE/tests/test/Coef.xlsx'

#CHEMIN POUR ACCEDER AU DOSSIER QUI CONTIENT LES FICHIERS EXCEL DES NOTES DE TOUS LES ETUDIANTS
dossier_notes = '/workspaces/SAE105_EXCEL_-_VALIDATION_UE/tests/test/notes_S1'


#INITIALISATION DES TABLEAUX
# on créait des tableaux vide pour pouvoir mettre les données que nous voulons utilisées

#On créait un tableau_coef pour y mettre les données qui sont présent dans le fichiers Coef.xlsx pour que ca soit plus simple a utiliser
tableau_coef = []

#On créait un tableau_note pour y mettres les données qui sont présent dans les fichiers xlsx qui sont dans le dossier notes_S1 pour pouvoir les utiliser
note_matiere = []

#On créait un Gros_tableau_Notes qui va contenir les notes de chaque matière de chaque élèves que nous allons utiliser
Gros_Tableau_Notes = []

#C'est le dictionnaire que nous allons utiliser plus tard 
notes = {}



#FONCTION POUR LIRE LE FICHIER EXCEL DES COEFS
def lecture_du_fichier_coef_dico(fichier_coef):
    tableau_coef = []
    Fichier_coef= openpyxl.load_workbook(fichier_de_ref, data_only = True)

    Onglets_coef = Fichier_coef.sheetnames
    feuille_active = Fichier_coef[Onglets_coef[0]]
    entetes = [cell.value for cell in feuille_active[1]]
        

    for ligne in feuille_active.iter_rows(
        min_row=2, max_row=feuille_active.max_row,
        values_only=True):
        tableau_coef.append(dict(zip(entetes, ligne)))

    return tableau_coef



#   FONCTION POUR LIRE LES FICHIERS EXCEL DES NOTES DES ETUDIANTS   
def lire_fichier_excel(fichier , Dossier_semestre):
    note_matiere = []
    Fichier_Matière = os.path.basename(fichier)
    Dossier_semestre_basename = os.path.basename(Dossier_semestre)
    Fichier_Excel = openpyxl.load_workbook(os.path.join(Dossier_semestre, fichier), data_only = True)
    Onglets = Fichier_Excel.sheetnames
    feuille_active = Fichier_Excel[Onglets[0]]
    entetes = [cell.value for cell in feuille_active[1]][1:4]

        
    nblignes= feuille_active['A2'].value
    for ligne in feuille_active.iter_rows(
        min_row=2, max_row=nblignes, min_col=2, max_col=4,
        values_only=True):
        note = dict(zip(entetes, ligne)) | {
            "Fichier_Matière": Fichier_Matière,
            "Dossier_semestre": Dossier_semestre_basename       
        }
        note_matiere.append(note)   
    return note_matiere
   


#Début du programme principal
# Lecture du fichier des coefs  
tableau_coef = lecture_du_fichier_coef_dico(fichier_de_ref)


# Lecture des fichiers des notes des étudiants
for fichier in os.listdir(dossier_notes):
    Gros_Tableau_Notes = Gros_Tableau_Notes + lire_fichier_excel(fichier , dossier_notes)


# liste des UE


liste_ue = list({item["Unité_d_Enseignement"] for item in tableau_coef})

liste_semestre = list({item["Semestre"] for item in tableau_coef})



for ue in liste_ue:
    print("================================== " , ue)
    for matière in tableau_coef:

        matiere_ue = matière["Unité_d_Enseignement"]
        matiere_fichier = matière["Fichier"]
        matiere_coef = float(matière["Coefficient"])
        #SI la matière existe dans l'UE en cours
        if  matiere_ue == ue :
            for eleve in Gros_Tableau_Notes:
                cle = (eleve["Nom"], eleve["Prénom"],ue)  # identifiant unique
                # on ne traite que la matière en cours 
                #print('>>>>>>',eleve["Fichier_Matière"])
                if eleve["Fichier_Matière"] == matière["Fichier"] :
                    #print("***",eleve["Fichier_Matière"],eleve["Note"],matiere_coef)
                    # Creation du tableau de notes final avec la pondération
                    if cle not in notes:
                        notes[cle] = 0
                    notes[cle] += float(eleve["Note"]) * matiere_coef  / 100


    # Affichage
for (nom, prenom,ue), total in notes.items():
    print(nom, prenom ,  ue , " → note :", total)




# Début du code HTML
html_content = """
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tableau de Notes Stylisé</title>
    <style>
        /* Configuration générale */
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #eef2f5;
            display: flex;
            justify-content: center;
            padding: 40px;
        }

        /* Conteneur pour le tableau avec défilement horizontal */
        .table-container {
            width: 100%;
            max-width: 1200px;
            background-color: #fff;
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
            overflow-x: auto; /* Permet le scroll sur petit écran */
            padding: 20px;
        }

        /* Style de base du tableau */
        .styled-table {
            width: 100%;
            border-collapse: collapse;
            margin: 0;
            font-size: 0.95em;
            min-width: 1000px; /* Force une largeur minimale pour éviter que les colonnes ne s'écrasent */
        }

        /* En-tête du tableau */
        .styled-table thead th {
            background-color: #0056b3; /* Bleu professionnel */
            color: #ffffff;
            text-align: center; /* Centrer les titres */
            font-weight: 600;
            padding: 15px 10px;
            border-bottom: 2px solid #004494;
            white-space: nowrap; /* Empêche le texte de passer à la ligne */
        }

        /* Cellules du corps */
        .styled-table td {
            padding: 12px 10px;
            border-bottom: 1px solid #dee2e6;
            text-align: center; /* Centrer le contenu des cellules */
            color: #333;
        }

        /* Alignement à gauche pour les noms et prénoms */
        .styled-table td:nth-child(1),
        .styled-table td:nth-child(2) {
            text-align: left;
            font-weight: 500;
        }

        /* Effet zébré (une ligne sur deux) */
        .styled-table tbody tr:nth-child(even) {
            background-color: #f8f9fa;
        }

        /* Effet de survol (hover) */
        .styled-table tbody tr:hover {
            background-color: #e2e6ea;
            transition: background-color 0.2s ease;
        }

        /* Cellules vides */
        .styled-table td:empty::after {
            content: "-";
            color: #aaa;
        }

        /* Styles spécifiques pour les statuts */
        .status-val {
            color: #28a745; /* Vert pour VAL */
            font-weight: bold;
        }
        .status-att {
            color: #ffc107; /* Jaune/Orange pour ATT */
            font-weight: bold;
        }

    </style>
</head>
<body>

    <div class="table-container">
        <table class="styled-table">
            <thead>
                <tr>
                    <th>Nom</th>
                    <th>Prénom</th>
                    <th>UE1.1</th>
                    <th>UE1.2</th>
                    <th>UE1.3</th>
                    <th>Etat UE1</th>
                    <th>UE2.1</th>
                    <th>UE2.2</th>
                    <th>UE2.3</th>
                    <th>Etat UE2</th>
                    <th>Etat BUT1</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td><span class="status-val">VAL</span></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                </tr>
                <tr>
                    <td>Nom2</td>
                    <td>Prénom2</td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td><span class="status-att">ATT</span></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                </tr>
                <tr>
                    <td>Nom3</td>
                    <td>Prénom3</td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                </tr>
                <tr>
                    <td>Nom4</td>
                    <td>Prénom4</td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                </tr>
            </tbody>
        </table>
    </div>

</body>
</html>
"""

# Boucle Python pour ajouter les lignes (TR) et cellules (TD)
for (nom, prenom, ue), total in notes.items():
    html_content += f"""
    <tr>
        <td>{nom}</td>
        <td>{prenom}</td>
        <td>{ue}</td>
        <td>{total}</td>
    </tr>
    """

# Fin du code HTML
html_content += """
</table>
</body>
</html>
"""

# Écriture dans un fichier
with open("mes_notes.html", "w", encoding="utf-8") as file:
    file.write(html_content)

print("Le fichier 'mes_notes.html' a été créé !")





