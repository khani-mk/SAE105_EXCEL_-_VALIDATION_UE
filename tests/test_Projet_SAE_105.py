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
                    print("***",eleve["Fichier_Matière"],eleve["Note"],matiere_coef)
                    # Creation du tableau de notes final avec la pondération
                    if cle not in notes:
                        notes[cle] = 0
                    notes[cle] += float(eleve["Note"]) * matiere_coef/100


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
        <title>Tableau de Notes BUT.1</title>
        <style>
                body{
                    font-family: Arial, sans-serif;
                    margin: 20px;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                }

                table{
                    border:solid;
                    border-collapse: collapse;
                }
                td{
                    border: solid;  
                }
                #nv{
                    background-color: red;
                }
                
                #v{
                    background-color: green
                }
                #a{
                    background-color: yellow
                }
                .excellent { background-color: green; font-weight: bold; }
                .average { background-color: #e6b800; font-weight: bold; } /* Darker yellow for readability */
                .fail { background-color: red; font-weight: bold; }
        </style>
    </head>
    <body>
    <table>
    <tr>
        <td>Nom</td>
        <td>Prénom</td>
        <td>UE</td>
        <td>Note Totale</td>
        <td id="a">EN ATTENTE DE VALIDATION</td>
        <td id="v">VALIDÉ</td>
        <td id="nv">NON VALIDÉ</td>
    </tr>
"""

# Boucle pour ajouter les lignes (TR) et cellules (TD)
for eleve in Gros_Tableau_Notes:
    grade = eleve["Note"]
    name = eleve["Nom"] + eleve["Prénom"]
    css_class = ""
    status = ""
    
    if grade > 10:
        css_class = "excellent" # VERT
        status = "Validé"
    elif 8 <= grade <= 10:
        css_class = "average"   # JAUNE
        status = "En attente de validation"
    else: # >8
        css_class = "fail"      # ROUGE
        status = "Non validé"
for (nom, prenom, ue), total in notes.items():
    html_content += f"""
    <tr>
        <td>{nom}</td>
        <td>{prenom}</td>
        <td>{ue}</td>
        <td>{total}</td>
        <td class="{css_class}">{status}</td>
    </tr>
    """

# Fin du code HTML
html_content += """
        
    <table>
    </body>
    </html>
"""

# Écriture dans un fichier
with open("mes_notes.html", "w", encoding="utf-8") as file:
    file.write(html_content)

print("Le fichier 'mes_notes.html' a été créé !")