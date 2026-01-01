import openpyxl
import os


#CHEMIN DES FICHIERS EXCEL

#Chemin pour accéder au fichier excel des coefs
fichier_de_ref = '/workspaces/SAE105_EXCEL_-_VALIDATION_UE/tests/test/Coef.xlsx'

#CHEMIN POUR ACCEDER AU DOSSIER QUI CONTIENT LES FICHIERS EXCEL DES NOTES DE TOUS LES ETUDIANTS
dossier_notes = '/workspaces/SAE105_EXCEL_-_VALIDATION_UE/tests/test/notes_S1'

#INITIALISATION DES TABLEAUX    
tableau_coef = []
Gros_Tableau_Notes = []
note_matiere = []



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
#for fichier in os.listdir(dossier_notes):
#    Gros_Tableau_Notes = Gros_Tableau_Notes + lire_fichier_excel(fichier , dossier_notes)

Gros_Tableau_Notes = Gros_Tableau_Notes + lire_fichier_excel('/workspaces/SAE105_EXCEL_-_VALIDATION_UE/tests/test/notes_S1/Architecture_des_systemes_numeriques_et_informatiques.xlsx', dossier_notes)    
Gros_Tableau_Notes = Gros_Tableau_Notes + lire_fichier_excel('/workspaces/SAE105_EXCEL_-_VALIDATION_UE/tests/test/notes_S1/Anglais_de_communication_technique.xlsx', dossier_notes)    

print(Gros_Tableau_Notes)
#print(tableau_coef)
# liste des UE


liste_ue = list({item["Unité_d_Enseignement"] for item in tableau_coef})

liste_semestre = list({item["Semestre"] for item in tableau_coef})

noteUE = {}

for ue in liste_ue:
    print("================================== " , ue)
    for matière in tableau_coef:
        #on VERIFIE SI la matière existe dans l'UE en cours
        if matière["Unité_d_Enseignement"] == ue :
            print("--------- " , matière["Fichier"])
            # on récuère le coef pour la matière pour l'ue en cours
            coef = float(matière["Coefficient"])
            # on calcul la note de l'élève pondéré par le coef
            for eleve in Gros_Tableau_Notes:
                cle = (eleve["Nom"], eleve["Prénom"],ue)  # identifiant unique
                # on ne traite que la matière en cours 
                if eleve["Fichier_Matière"] == matière["Fichier"] :
                    print(eleve["Fichier_Matière"])
                    if cle not in noteUE:
                        noteUE[cle] = 0
                        noteUE[cle] += eleve["Note"] 
                break    
#for ue in liste_ue:
    #for matière in tableau_coef:
        #if matière["Unité_d_Enseignement"] == ue:
            

                    

    # Affichage
for (nom, prenom,ue), total in noteUE.items():
    print(nom, prenom ,  ue , "→ total des notes :", total)





