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
                    notes[cle] += float(eleve["Note"]) * matiere_coef / 100

    


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
    
    




def calculer_decision_annee(notes_ues):
    """
    Applique les règles de passage du BUT basées sur l'image :
    - Avoir au moins 2 UE >= 10
    - Avoir toutes les UE >= 8 (aucune < 8)
    """
    valeurs_notes = list(notes_ues.values())
    
    # Si l'étudiant n'a pas toutes les notes (ex: absent), on considère incomplet
    # On suppose ici qu'il y a 3 UE (UE1, UE2, UE3) comme dans l'exemple, 
    # mais le code s'adapte au nombre d'UE présentes.
    if not valeurs_notes:
        return "Incomplet", "fail"

    # Règle 1 : Vérifier si une UE est < 8 (Éliminatoire)
    for note in valeurs_notes:
        if note < 8:
            return "REFUSÉ (UE < 8)", "fail"

    # Règle 2 : Compter le nombre d'UE >= 10
    nb_ue_sup_10 = sum(1 for note in valeurs_notes if note >= 10)

    # Il faut au moins 2 UE >= 10 pour valider
    if nb_ue_sup_10 >= 2:
        return "ADMIS (Passage BUT2)", "excellent"
    else:
        # Cas où l'étudiant a tout > 8 mais n'a pas deux notes > 10 (ex: 9, 9, 9)
        return "REFUSÉ (Pas assez d'UE > 10)", "fail"

    
# =================================================================
# 2. GÉNÉRATION DU HTML (Mise à jour avec Règles BUT)
# =================================================================

html_content = """
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Tableau Récapitulatif BUT</title>
    <style>
        body{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; }
        table{ border-collapse: collapse; width: 100%; box-shadow: 0 0 20px rgba(0,0,0,0.1); }
        td, th{ border: 1px solid #ddd; padding: 12px; text-align: center; }
        th { background-color: #009879; color: white; }
        tr:nth-child(even){background-color: #f2f2f2;}
        
        /* Couleurs des notes individuelles */
        .ue-validee { background-color: #c6efce; color: #006100; } /* Vert */
        .ue-compensee { background-color: #ffeb9c; color: #9c5700; } /* Jaune/Orange */
        .ue-echec { background-color: #ffc7ce; color: #9c0006; }    /* Rouge */

        /* Couleurs de la décision finale */
        .excellent { background-color: #28a745; color: white; font-weight: bold; }
        .fail { background-color: #dc3545; color: white; font-weight: bold; }
    </style>
</head>
<body>
    <h2>BUT.1-R&T</h2>
    <p>Règle : Aucune UE < 8 <b>ET</b> au moins 2 UE >= 10.</p>
    <table>
    <tr>
        <th>Nom</th>
        <th>Prénom</th>
"""

# Ajout dynamique des colonnes UE
for ue in toutes_les_ues:
    html_content += f"<th>{ue}</th>"

# Ajout de la colonne Décision
html_content += "<th>Etat BUT1</th></tr>"

# Boucle pour chaque élève
for (nom, prenom), notes_ues in eleves_dict.items():
    
    # 1. On calcule d'abord la décision globale pour l'année
    decision_texte, decision_class = calculer_decision_annee(notes_ues)

    html_content += f"<tr><td>{nom}</td><td>{prenom}</td>"
    
    # 2. Affichage des notes par UE
    for ue in toutes_les_ues:
        if ue in notes_ues:
            note_finale = float(notes_ues[ue])
            
            # Logique d'affichage des couleurs par cellule (UE)
            # Si note > 10 : Vert
            # Si 8 <= note < 10 : Jaune (Potentiellement compensé)
            # Si note < 8 : Rouge
            
            if note_finale >= 10:
                css_class = "ue-validee"
            elif note_finale >= 8:
                css_class = "ue-compensee"
            else:
                css_class = "ue-echec"
            
            html_content += f'<td class="{css_class}">{round(note_finale, 2)}</td>'
        else:
            html_content += "<td>-</td>"

    # 3. Ajout de la colonne Décision Finale
    html_content += f'<td class="{decision_class}">{decision_texte}</td>'
    html_content += "</tr>"

html_content += """
    </table>
</body>
</html>
"""

# Écriture dans le fichier
with open("mes_notes.html", "w", encoding="utf-8") as file:
    file.write(html_content)

print("Le fichier 'mes_notes.html' a été mis à jour avec les règles de passage BUT !")