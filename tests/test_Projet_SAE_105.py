import openpyxl
import os 

workbook = openpyxl.load_workbook(r'C:\Users\admin\Downloads\notes_S1\notes.xlsx', data_only = True)
titres_onglets = workbook.sheetnames
onglet1 = workbook[titres_onglets[0]]
onglet2 = workbook[titres_onglets[1]]
onglet3 = workbook[titres_onglets[2]]
onglet4 = workbook[titres_onglets[3]]
onglet5 = workbook[titres_onglets[4]]
onglet6 = workbook[titres_onglets[5]]
onglet7 = workbook[titres_onglets[6]]
onglet8 = workbook[titres_onglets[7]]
onglet9 = workbook[titres_onglets[8]]
onglet10 = workbook[titres_onglets[9]]
onglet11 = workbook[titres_onglets[10]]
onglet12 = workbook[titres_onglets[11]]
onglet13 = workbook[titres_onglets[12]]
onglet14 = workbook[titres_onglets[13]]
onglet15 = workbook[titres_onglets[14]]
onglet16 = workbook[titres_onglets[15]]
onglet17 = workbook[titres_onglets[16]]
onglet18 = workbook[titres_onglets[17]]
onglet19 = workbook[titres_onglets[18]]
onglet20 = workbook[titres_onglets[19]]
onglet21 = workbook[titres_onglets[20]]
onglet22 = workbook[titres_onglets[21]]
onglet23 = workbook[titres_onglets[22]]
onglet24 = workbook[titres_onglets[23]]
onglet25 = workbook[titres_onglets[24]]

#onglet 1
#Liste des lignes
lignes = []
for row in onglet1.values:
    #ajoute de la ligne row à la liste lignes
    lignes.append(list(row))


#Liste des colonnes
colonnes = []
for column in onglet1.columns:
    colonnes.append([cell.value for cell in column])


#onglet 2
#Liste des lignes
lignes = []
for row in onglet2.values:
    #ajoute de la ligne row à la liste lignes
    lignes.append(list(row))


#Liste des colonnes
colonnes = []
for column in onglet2.columns:
    colonnes.append([cell.value for cell in column])


#onglet 3
#Liste des lignes
lignes = []
for row in onglet3.values:
    #ajoute de la ligne row à la liste lignes
    lignes.append(list(row))


#Liste des colonnes
colonnes = []
for column in onglet3.columns:
    colonnes.append([cell.value for cell in column])


#onglet 4
#Liste des lignes
lignes = []
for row in onglet4.values:
    #ajoute de la ligne row à la liste lignes
    lignes.append(list(row))


#Liste des colonnes
colonnes = []
for column in onglet4.columns:
    colonnes.append([cell.value for cell in column])


#onglet 5
#Liste des lignes
lignes = []
for row in onglet5.values:
    #ajoute de la ligne row à la liste lignes
    lignes.append(list(row))


#Liste des colonnes
colonnes = []
for column in onglet5.columns:
    colonnes.append([cell.value for cell in column])

#ferme le fichier excel
workbook.close()


#Géneration de la page html
def genere_page_web( nom_fichier, titre, corps):    
    with open(nom_fichier, 'w', encoding='utf-8') as f:
        f.write(corps)
    print(f"le fichier {nom_fichier} a été généré avec succès !")
    def main():
        corps = """
            <!DOCTYPE html>
                <html lang="en">
                <head>
                    <meta charset="UTF-8">
                    <link rel="stylesheet" href="style.css">
                    <meta name="viewport" content="width=device-width, initial-scale=1.0">
                    <title>Document</title>
                </head>
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
    main()
