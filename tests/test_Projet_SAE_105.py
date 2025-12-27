import openpyxl
import os

tableau_coef= []

Fichier_coef_S1= openpyxl.load_workbook('/workspaces/SAE105_EXCEL_-_VALIDATION_UE/tests/test/Coef.xlsx', data_only = True)
Onglets_coef = Fichier_coef_S1.sheetnames
feuille_active = Fichier_coef_S1[Onglets_coef[0]] 
#On va lire les données qui sont présente dans le fichier pour pouvoir applique les coef sur les fichiers après
for ligne in feuille_active.iter_rows(
    min_row=2, max_row=22, min_col=2, max_col=5,
    values_only=True):
    tableau_coef.append(list(ligne))

print(tableau_coef)

notes = []

Fichier = openpyxl.load_workbook('/workspaces/SAE105_EXCEL_-_VALIDATION_UE/tests/test/notes_S1/Anglais_technique_1.xlsx', data_only = True)
Onglets = Fichier.sheetnames
feuille_active = Fichier[Onglets[0]] 
# c'est le premier onglet dans la feuille Excel
#parcours de chaque ligne de l'onglet et affichage de la valeur
#de chaque cellule

#Nombre de ligne 
nblignes= feuille_active['A2'].value

#on va lire les colonnes B C D 
for ligne in feuille_active.iter_rows(
    min_row=2, max_row=nblignes, min_col=2, max_col=4, # on lis les colonne B C D 
                                                       #( qui corresponde à 2 3 4 ( prénom , nom , note ))
    values_only=True):
    notes.append(list(ligne))

for i in range(len(notes)):
    notes[i].extend(["Semestre", "UE" , "Matière", "Coef"]) 

print(notes)

#ferme le fichier excel
Fichier.close()

























# J'AI MIS LE CODE EN COMMENTAIRE POUR TESTES LE MIEN

#import ast
#import openpyxl

#workbook = openpyxl.load_workbook(r'C:\Users\admin\Downloads\notes_S1\notes.xlsx', data_only = True)
#titres_onglets = workbook.sheetnames
#toutes_les_donnees = {} 

#for nom_onglet in workbook.sheetnames:
    
    # On ouvre l'onglet actuel grâce à son nom
    #feuille_actuelle = workbook[nom_onglet]
    
    # On récupère toutes les lignes de cet onglet sous forme de liste
    # La fonction list() transforme le générateur 'values' en une vraie liste manipulable
    #lignes_de_l_onglet = list(feuille_actuelle.values)
    
    # 4. Stockage
    # On range les lignes dans notre dictionnaire, sous le nom de l'onglet
    #toutes_les_donnees[nom_onglet] = lignes_de_l_onglet
    
    #print(f"Onglet '{nom_onglet}' traité : {len(lignes_de_l_onglet)} lignes récupérées.")

#workbook.close()

# --- EXEMPLE D'UTILISATION DES DONNÉES ---

# Maintenant, si tu veux accéder aux données du premier onglet (ex: 'UE 1.1') :
#nom_premier_onglet = workbook.sheetnames[0] # On récupère le nom
#data_onglet_1 = toutes_les_donnees[nom_premier_onglet] # On ouvre le tiroir

# Afficher la valeur de la ligne 2, colonne 3 de cet onglet
# Attention : en Python les index commencent à 0, donc ligne 2 est à l'index 1
#print(data_onglet_1[4][5])
# #FIN DU CODE EN COMMENTAIRE



#Géneration de la page html
def genere_page_web( nom_fichier, titre, corps):    
    with open(nom_fichier, 'w', encoding='utf-8') as f:
        f.write(corps)
    print(f"le fichier {nom_fichier} a été généré avec succès !")
    def main():
        corps = """
                <style>
                table{
                    border:solid;
                    border-collapse: collapse;
                }
                td{
                    border: solid;  
                }
                #ue1{
                    background-color: red;
                }
                
                #ue2{
                    background-color: green
                }
                   
                </style>
                <body>
                    <table>
                    <tr>
                        <td> Nom  </td> 
                        <td> Prénom</td>
                        <td>UE 1.1</td>
                        <td>UE 1.2</td>
                        <td>UE 1.3</td>
                        <td>UE 1 </td>   
                        <td>Etat UE 1 </td>
                        <td>UE 2.1 </td>
                        <td>UE 2.2 </td>
                        <td>UE 2.3 </td>
                        <td>UE 2 </td>
                    </tr>
                    <tr>
                        <td> Nom  </td> 
                        <td> Prénom</td>
                        <td></td>
                        <td>*</td>
                        <td>*</td>>
                        <td id="ue1"*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td id="ue2">*</td>
                    </tr>
                     <tr>
                        <td> Nom  </td> 
                        <td> Prénom</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>    
                        <td id="ue1">*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td id="ue2">*</td>
                    </tr>
                     <tr>
                        <td> Nom  </td> 
                        <td> Prénom</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>    
                        <td id="ue1">*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td id="ue2">*</td>
                    </tr>
                
                     <tr>
                        <td> Nom  </td> 
                        <td> Prénom</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>    
                        <td id="ue1">*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td id="ue2">*</td>
                    </tr>
                     <tr>
                        <td> Nom  </td> 
                        <td> Prénom</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>    
                        <td id="ue1">*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td id="ue2">*</td>
                    </tr>
                
                     <tr>
                        <td> Nom  </td> 
                        <td> Prénom</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>    
                        <td id="ue1">*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td id="ue2">*</td>
                    </tr>
                     <tr>
                        <td> Nom  </td> 
                        <td> Prénom</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>    
                        <td id="ue1">*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td id="ue2">*</td>
                    </tr>
                
                     <tr>
                        <td> Nom  </td> 
                        <td> Prénom</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>    
                        <td id="ue1">*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td>*</td>
                        <td id="ue2">*</td>
                    </tr>
                
                    </table>
                </body>
                </html>
            
            """
    genere_page_web("index.html", "mon_titre", corps) 

 
#if __name__ == "__main__":
#    main() # type: ignore
