#On import les modules openpyxl et os

#On va l'utiliser pour ouvrir les fichier xlsx et ainsi les utiliser
import openpyxl

#il va servir a donner les chemins d'accès pour aller aux différents fichiers que nous allons utiliser
import os


#CHEMIN DES FICHIERS EXCEL
#Chemin pour accéder au fichier excel des coefs
fichier_de_ref = '/workspaces/SAE105_EXCEL_-_VALIDATION_UE/tests/test/Coef.xlsx'

#CHEMIN POUR ACCEDER AU DOSSIER QUI CONTIENT LES FICHIERS EXCEL DES NOTES DE TOUS LES ETUDIANTS
 

dossiers_notes = []

dossiers_notes.append('/workspaces/SAE105_EXCEL_-_VALIDATION_UE/tests/test/notes_S1')
dossiers_notes.append('/workspaces/SAE105_EXCEL_-_VALIDATION_UE/tests/test/notes_S2')

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
    feuille_active = Fichier_coef[Onglets_coef[0]] # la feuille active est le premier onglet du fichier
    
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
for dossier_notes in dossiers_notes:    
    for fichier in os.listdir(dossier_notes):
        Gros_Tableau_Notes = Gros_Tableau_Notes + lire_fichier_excel(fichier , dossier_notes)


# liste des UE


liste_ue = list({item["Unité_d_Enseignement"] for item in tableau_coef})

liste_semestre = list({item["Semestre"] for item in tableau_coef})



for ue in liste_ue:
    #print("================================== " , ue)
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
                    notes[cle] += float(eleve["Note"]) * matiere_coef/ 20


# =================================================================
# 1. RESTRUCTURATION DES DONNÉES
# =================================================================

# On liste toutes les UEs uniques pour créer les colonnes du tableau
toutes_les_ues = sorted(list({k[2] for k in notes.keys()}))

# On regroupe les notes par élève : dictionnaire { (Nom, Prenom): {UE1: Note, UE2: Note} }
eleves_dict = {}

for (nom, prenom, ue), note in notes.items():
    cle_eleve = (nom, prenom)
    if cle_eleve not in eleves_dict:
        eleves_dict[cle_eleve] = {}
    eleves_dict[cle_eleve][ue] = note

# =================================================================
# 2. GÉNÉRATION DU HTML
# =================================================================

# Début du code HTML
html_content = """
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Tableau Récapitulatif</title>
    <style>
        body{ font-family: Arial, sans-serif; padding: 20px; }
        table{ border-collapse: collapse; width: 100%; }
        td, th{ border: 1px solid black; padding: 8px; text-align: center; }
        th { background-color: #f2f2f2; }
        .excellent { background-color: #90ee90; } /* Vert clair */
        .average { background-color: #ffd700; }   /* Or/Jaune */
        .fail { background-color: #ffcccb; }      /* Rouge clair */
    </style>
</head>
<body>
    <h2>Récapitulatif des Notes par Élève  - (Vert=Validé/Jaune=En cours/Rouge=Non Validé)</h2>
    <table>
    <tr>
        <th>Nom</th>
        <th>Prénom</th>
"""

# Ajout dynamique des colonnes pour chaque UE
for ue in toutes_les_ues:
    html_content += f"<th>{ue}</th>"

html_content += "</tr>"

# Boucle pour chaque élève (une ligne par élève)
for (nom, prenom), notes_ues in eleves_dict.items():
    html_content += f"<tr><td>{nom}</td><td>{prenom}</td>"
    
    # Pour cet élève, on regarde chaque UE (pour bien aligner les colonnes)
    for ue in toutes_les_ues:
        # Si l'élève a une note pour cette UE, on la récupère, sinon on met "N/A"
        if ue in notes_ues:
            note_finale = float(notes_ues[ue])
            
            # Logique des couleurs
            if note_finale > 10:
                css_class = "excellent"
            elif 8 <= note_finale <= 10:
                css_class = "average"
            else:
                css_class = "fail"
            
            # On ajoute la case avec la couleur
            html_content += f'<td class="{css_class}">{round(note_finale, 2)}</td>'
        else:
            # Pas de note pour cette UE
            html_content += "<td>-</td>"

    html_content += "</tr>"

# Fin du HTML
html_content += """
    </table>
</body>
</html>
"""

# Écriture dans le fichier
with open("mes_notes.html", "w", encoding="utf-8") as file:
    file.write(html_content)

print("Le fichier 'mes_notes.html' a été généré avec succès (1 ligne par élève) !")
# Écriture dans un fichier
with open("mes_notes.html", "w", encoding="utf-8") as file:
    file.write(html_content)

print("Le fichier 'mes_notes.html' a été créé !")