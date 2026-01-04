import openpyxl
import os

def lecture_du_fichier_coef_dico(fichier_coef):
    """
    Lit le fichier Excel des coefficients et le transforme en structure utilisable.

    Cette fonction ouvre le fichier Excel spécifié, lit la première feuille,
    récupère les en-têtes et crée une liste de dictionnaires pour chaque ligne.

    :param fichier_coef: Le chemin complet vers le fichier Excel des coefficients (str).
    :return: Une liste de dictionnaires contenant les infos des coefficients (list).
    """
    tableau_coef = []
    Fichier_coef= openpyxl.load_workbook(fichier_coef, data_only = True)

    Onglets_coef = Fichier_coef.sheetnames
    feuille_active = Fichier_coef[Onglets_coef[0]] # la feuille active est le premier onglet du fichier
    
    entetes = [cell.value for cell in feuille_active[1]]
        
    for ligne in feuille_active.iter_rows(
        min_row=2, max_row=feuille_active.max_row,
        values_only=True):
        tableau_coef.append(dict(zip(entetes, ligne)))
    
    return tableau_coef


def lire_fichier_excel(fichier, Dossier_semestre):
    """
    Lit un fichier Excel de notes pour une matière donnée.

    :param fichier: Le nom du fichier Excel à lire (ex: 'R101.xlsx').
    :param Dossier_semestre: Le chemin du dossier contenant ce fichier.
    :return: Une liste de dictionnaires représentant les notes des étudiants pour cette matière.
    """
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


def determiner_etat_ue(note):
    """
    Détermine l'état d'une UE (Validé, Compensable, Échec) selon sa moyenne.

    :param note: La moyenne annuelle de l'UE (float).
    :return: Un tuple contenant le texte à afficher et la classe CSS correspondante (str, str).
    """
    if note >= 10:
        return "VALIDÉ", "ue-validee"
    elif note >= 8:
        return "COMPENSABLE", "ue-compensee"
    else:
        return "NON VALIDÉ", "ue-echec"


def calculer_decision_passage(liste_moyennes_annuelles):
    """
    Calcule la décision finale du jury pour le passage en année supérieure.

    Règles de passage BUT sur les MOYENNES ANNUELLES :
    1. Aucune UE < 8 (Pas de 'NON VALIDÉ').
    2. Au moins 2 UE >= 10.

    :param liste_moyennes_annuelles: Liste des moyennes de toutes les UEs de l'étudiant.
    :return: Un tuple contenant la décision finale et la classe CSS (str, str).
    """
    # S'il n'y a aucune note
    if not liste_moyennes_annuelles:
        return "Incomplet", "decision-fail"

    # 1. Vérification éliminatoire (< 8)
    if any(note < 8 for note in liste_moyennes_annuelles):
        return "REFUSÉ (UE < 8)", "decision-fail"
    
    # 2. Comptage des UE validées (>= 10)
    nb_ue_validees = sum(1 for note in liste_moyennes_annuelles if note >= 10)
    
    # Règle : Il faut au moins 2 UE >= 10
    if nb_ue_validees >= 2:
        return "ADMIS (Passage BUT2)", "decision-ok"
    else:
        return "REFUSÉ (Manque UE > 10)", "decision-fail"


# --- PROGRAMME PRINCIPAL ---
# La ligne ci-dessous empêche Sphinx d'exécuter ce code lors de la lecture des fonctions
if __name__ == "__main__":
    
    #CHEMIN DES FICHIERS EXCEL
    fichier_de_ref = '/workspaces/SAE105_EXCEL_-_VALIDATION_UE/data/coefficients/Coef.xlsx'

    dossiers_notes = []
    dossiers_notes.append('/workspaces/SAE105_EXCEL_-_VALIDATION_UE/data/notes/notes_S1')
    dossiers_notes.append('/workspaces/SAE105_EXCEL_-_VALIDATION_UE/data/notes/notes_S2')

    #INITIALISATION DES TABLEAUX
    tableau_coef = []
    note_matiere = []
    Gros_Tableau_Notes = []
    notes = {}

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
        for matière in tableau_coef:
            matiere_ue = matière["Unité_d_Enseignement"]
            matiere_fichier = matière["Fichier"]
            matiere_coef = float(matière["Coefficient"])
            #SI la matière existe dans l'UE en cours
            if  matiere_ue == ue :
                for eleve in Gros_Tableau_Notes:
                    cle = (eleve["Nom"], eleve["Prénom"],ue)  # identifiant unique
                    if eleve["Fichier_Matière"] == matière["Fichier"] :
                        # Creation du tableau de notes final avec la pondération
                        if cle not in notes:
                            notes[cle] = 0
                        notes[cle] += float(eleve["Note"]) * matiere_coef / 100

    # On détermine les noms des UEs "racines"
    ues_racines = sorted(list({ue.split('.')[0] for ue in liste_ue}))

    # Liste de tous les étudiants (Nom, Prénom) uniques
    etudiants_uniques = sorted(list(set((k[0], k[1]) for k in notes.keys())))

    # GÉNÉRATION DU HTML 
    html_content = """
    <!DOCTYPE html>
    <html lang="fr">
    <head>
        <meta charset="UTF-8">
        <title>BUT1-R&T</title>
        <style>
            body{ font-family: 'Segoe UI', sans-serif; padding: 20px; font-size: 13px; }
            h2 { text-align: center; color: #333; }
            table{ border-collapse: collapse; width: 100%; box-shadow: 0 0 15px rgba(0,0,0,0.1); margin-top: 20px; }
            td, th{ border: 1px solid #ddd; padding: 8px 4px; text-align: center; }
            thead tr:first-child th { background-color: #005f99; color: white; border-right: 1px solid white; }
            thead tr:nth-child(2) th { background-color: #007acc; color: white; font-size: 0.9em; }
            .col-semestre { background-color: #fcfcfc; color: #555; }
            .col-moyenne { background-color: #fff; font-weight: bold; border-left: 2px solid #ccc; }
            .ue-validee { background-color: #d4edda; color: #155724; font-weight: bold; }
            .ue-compensee { background-color: #fff3cd; color: #856404; font-weight: bold; }
            .ue-echec { background-color: #f8d7da; color: #721c24; font-weight: bold; }
            .decision-ok { background-color: #28a745; color: white; font-weight: bold; font-size: 1.1em; }
            .decision-fail { background-color: #dc3545; color: white; font-weight: bold; font-size: 1.1em; }
            tr:hover { background-color: #f1f1f1; }
        </style>
    </head>
    <body>
        <h2>Notes BUT1 - R&T </h2>
        <table>
        <thead>
            <tr>
                <th rowspan="2" style="width: 150px;">Étudiant</th>
    """

    for ue in ues_racines:
        html_content += f'<th colspan="4">{ue} (Annuel)</th>'

    html_content += '<th rowspan="2">DÉCISION</th></tr>'
    html_content += '<tr>'
    for ue in ues_racines:
        html_content += f'<th>{ue}.1</th><th>{ue}.2</th><th>Moy</th><th>État</th>'
    html_content += '</tr></thead><tbody>'

    # --- REMPLISSAGE DU TABLEAU ---
    for (nom, prenom) in etudiants_uniques:
        moyennes_annuelles_etudiant = []
        ligne_html_etudiant = f"<tr><td style='text-align:left; padding-left:10px;'><b>{nom}</b> {prenom}</td>"
        
        for racine in ues_racines:
            nom_ue_s1 = f"{racine}.1"
            nom_ue_s2 = f"{racine}.2"
            note_s1 = notes.get((nom, prenom, nom_ue_s1), 0.0)
            note_s2 = notes.get((nom, prenom, nom_ue_s2), 0.0)
            
            moyenne_annuelle = (note_s1 + note_s2) / 2
            moyennes_annuelles_etudiant.append(moyenne_annuelle)
            
            txt_etat, class_etat = determiner_etat_ue(moyenne_annuelle)
            
            ligne_html_etudiant += f'<td class="col-semestre">{round(note_s1, 2)}</td>'
            ligne_html_etudiant += f'<td class="col-semestre">{round(note_s2, 2)}</td>'
            ligne_html_etudiant += f'<td class="col-moyenne">{round(moyenne_annuelle, 2)}</td>'
            ligne_html_etudiant += f'<td class="{class_etat}">{txt_etat}</td>'

        txt_decision, class_decision = calculer_decision_passage(moyennes_annuelles_etudiant)
        ligne_html_etudiant += f'<td class="{class_decision}">{txt_decision}</td></tr>'
        html_content += ligne_html_etudiant

    html_content += """
        </tbody>
        </table>
    </body>
    </html>
    """

    # Écriture dans le fichier
    with open("/workspaces/SAE105_EXCEL_-_VALIDATION_UE/html/mes_notes.html", "w", encoding="utf-8") as file:
        file.write(html_content)

    print("Le fichier 'mes_notes.html' a été généré avec succès.")