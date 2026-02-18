# Automatisation de Validation d'UE - SAE 105

Ce projet, réalisé dans le cadre de la **SAE 105** en B.U.T. Réseaux et Télécommunications, vise à automatiser le traitement des notes des étudiants à partir de fichiers Excel pour générer un rapport de validation d'UE.

L'objectif est de manipuler des données structurées, d'appliquer des coefficients et de produire un rendu visuel clair sous forme de rapport HTML.

---

## Fonctionnalités
- **Lecture de données** : Extraction des notes et informations des étudiants depuis un fichier `.xlsx`.
- **Calcul automatique** : Calcul des moyennes par module et par UE en respectant les coefficients.
- **Vérification des conditions** : Détermination automatique de l'admission (validation de l'UE, compensation ou redoublement).
- **Génération de rapport** : Création d'un fichier HTML dynamique présentant les résultats de manière structurée.

##  Technologies utilisées
- **Python 3.x**
- **Openpyxl** : Bibliothèque pour la lecture et l'écriture de fichiers Excel.
- **HTML/CSS** : Pour la mise en forme du rapport final.

## Installation et Utilisation

1. **Cloner le dépôt :**
   ```bash
   git clone [https://github.com/khani-mk/SAE105_EXCEL_-_VALIDATION_UE.git](https://github.com/khani-mk/SAE105_EXCEL_-_VALIDATION_UE.git)
   cd SAE105_EXCEL_-_VALIDATION_UE