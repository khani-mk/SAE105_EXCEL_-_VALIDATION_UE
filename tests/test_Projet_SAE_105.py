import openpyxl
import os

fichier_de_ref = '/workspaces/SAE105_EXCEL_-_VALIDATION_UE/tests/test/Coef.xlsx'

#CREATION DU PROGRAMME POUR UTILISER APRES LES DONNES QUI SONT DANS LE FICHIER EXCEL
tableau_coef= []

Fichier_coef_S1= openpyxl.load_workbook(fichier_de_ref, data_only = True)
Onglets_coef = Fichier_coef_S1.sheetnames
feuille_active = Fichier_coef_S1[Onglets_coef[0]] 
#On va lire les données qui sont présente dans le fichier pour pouvoir applique les coef sur les fichiers après
for ligne in feuille_active.iter_rows(
    min_row=0, max_row=39, min_col=0, max_col=5,
    values_only=True):
    tableau_coef.append(list(ligne))
    
    def dico_coef(fichier_coef):
        wb = openpyxl.load_workbook(fichier_de_ref, data_only=True)
        Onglets_coef = Fichier_coef_S1.sheetnames
        feuille_active = Fichier_coef_S1[Onglets_coef[0]] 

        entetes = [cell.value for cell in feuille_active[1]]

        valeurs = []
        for ligne in feuille_active.iter_rows(
            min_row=2, max_row=feuille_active.max_row,
            values_only=True):
            valeurs.append(dict(zip(entetes, ligne)))

        return valeurs
    print(dico_coef(fichier_de_ref))

    for ligne in tableau_coef:
        if ligne["4"] == "Initiation_aux_réseaux_informatiques":
            print(ligne["Coefficient"])
    exit()







exit()
#print(tableau_coef)
dicsionaire_notes = {}
for ligne in tableau_coef:
    if ligne[2] is not None and ligne[4] is not None:
        semestre = str(ligne[0]).strip()
        nom_matiere = str(ligne[2]).strip()
        nom_UE = str(ligne[3]).strip()
        valeur_coef = float(ligne[4])

        if nom_matiere not in dicsionaire_notes:
            dicsionaire_notes[nom_matiere] = {}

        dicsionaire_notes[nom_matiere][nom_UE] = valeur_coef

print(dicsionaire_notes)

# ON VA LIRE TOUS LES FICHIERS EXCEL QUI SONT DANS LE DOSSIER notes_S1 
dossier_notes = '/workspaces/SAE105_EXCEL_-_VALIDATION_UE/tests/test/notes_S1'

#ON CREER UNE VARIABLE QUI VA DIRE ECRIRE TOUT LES NOMS DES FICHIERS QUI TERMINE PAR xlxs QUI SONT PRESENTS DANS LE DOSSIER notes_S1 
for nom_fichier in os.listdir(dossier_notes):
    if nom_fichier.endswith('.xlsx'):
        print(f"Fichier trouvé : {nom_fichier}")

#ON CREER UN TABLEAU VIDE POUR QU'IL PUISSE STOCKER PAR LA SUITE LES DONNES DES FICHIERS EXCEL
notes = []


for fichier in os.listdir(dossier_notes):
    Fichier = openpyxl.load_workbook(os.path.join(dossier_notes, fichier), data_only = True)
    Onglets = Fichier.sheetnames
    feuille_active = Fichier[Onglets[0]]
    
    nblignes= feuille_active['A2'].value

    for ligne in feuille_active.iter_rows(
        min_row=2, max_row=nblignes, min_col=2, max_col=4, 
        values_only=True):
        notes.append(list(ligne))
    
#print(notes)

exit()

Gros_tableau=[notes , tableau_coef]
print(Gros_tableau)

notes = Gros_tableau[0]
coefs = Gros_tableau[1]

for ligne_notes in notes:
    prenom = ligne_notes[0]
    nom = ligne_notes[1]
    note = ligne_notes[2]
    matière = nom_fichier
    print(prenom, nom, note , matière)























#Géneration de la page html
def genere_page_web( nom_fichier, titre, corps):    
    with open(nom_fichier, 'w', encoding='utf-8') as f:
        f.write(corps)
    print(f"le fichier {nom_fichier} a été généré avec succès !")
    def main():
        corps = """
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
                <body>
                     <div class="table-container">
        <table class="styled-table">
            <thead>
                <tr>
                    <th>Nom</th>
                    <th>Prénom</th>
                    <th>UE1.1</th>
                    <th>UE1.2</th>
                    <th>UE1</th>
                    <th>Etat UE1</th>
                    <th>UE2.1</th>
                    <th>UE2.2</th>
                    <th>UE2</th>
                    <th>Etat UE2</th>
                    <th>UE3.1</th>
                    <th>UE3.2</th>
                    <th>UE3</th>
                    <th>Etat UE3</th>
                    <th>Etat BUT1</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>Nom1</td>
                    <td>Prénom1</td>
                    <td>10.40</td>
                    <td>12.25</td>
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
    genere_page_web("index.html", "mon_titre", corps) 

 
#if __name__ == "__main__":
#    main() # type: ignore
