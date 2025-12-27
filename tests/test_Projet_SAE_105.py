import ast
import openpyxl

workbook = openpyxl.load_workbook(r'C:\Users\admin\Downloads\notes_S1\notes.xlsx', data_only = True)
titres_onglets = workbook.sheetnames
toutes_les_donnees = {} 

for nom_onglet in workbook.sheetnames:
    
    # On ouvre l'onglet actuel grâce à son nom
    feuille_actuelle = workbook[nom_onglet]
    
    # On récupère toutes les lignes de cet onglet sous forme de liste
    # La fonction list() transforme le générateur 'values' en une vraie liste manipulable
    lignes_de_l_onglet = list(feuille_actuelle.values)
    
    # 4. Stockage
    # On range les lignes dans notre dictionnaire, sous le nom de l'onglet
    toutes_les_donnees[nom_onglet] = lignes_de_l_onglet
    
    print(f"Onglet '{nom_onglet}' traité : {len(lignes_de_l_onglet)} lignes récupérées.")

workbook.close()

# --- EXEMPLE D'UTILISATION DES DONNÉES ---

# Maintenant, si tu veux accéder aux données du premier onglet (ex: 'UE 1.1') :
nom_premier_onglet = workbook.sheetnames[0] # On récupère le nom
data_onglet_1 = toutes_les_donnees[nom_premier_onglet] # On ouvre le tiroir

# Afficher la valeur de la ligne 2, colonne 3 de cet onglet
# Attention : en Python les index commencent à 0, donc ligne 2 est à l'index 1
print(data_onglet_1[4][5])


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

 
if __name__ == "__main__":
    main() # type: ignore
