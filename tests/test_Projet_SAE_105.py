import openpyxl

workbook = openpyxl.load_workbook('/workspaces/SAE105_EXCEL_-_VALIDATION_UE/tests/test/notes_S1/Anglais_technique_1.xlsx', data_only = True)
titres_onglets = workbook.sheetnames
onglet1 = workbook[titres_onglets[0]]
#Liste des lignes
lignes = []
for row in onglet1.values:
    #ajoute de la ligne row à la liste lignes
    lignes.append(list(row))
#Liste des colonnes
colonnes = []
for column in onglet1.columns:
    colonnes.append([cell.value for cell in column])
#ferme le fichier excel
workbook.close()
#Géneration de la page html
def genere_page_web( nom_fichier, titre, corps):    

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
    genere_page_web("./index.html", "mon_titre", corps)

if __name__ == "__main__":
    main()