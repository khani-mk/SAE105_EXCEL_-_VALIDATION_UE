Cahier des Charges
==================
Fonctionnalités Attendues
-------------------------
Le programme doit réaliser les trois étapes suivantes :

Lecture : Extraire les données pertinentes (notes, noms) depuis plusieurs fichiers Excel (un fichier par matière).,
Traitement : Calculer les moyennes de chaque Unité d'Enseignement (UE) et déterminer si l'UE est validée.,
Affichage : Construire une page HTML/CSS présentant les résultats sous forme de tableau.,

Contraintes Techniques
----------------------
Langage : Python (version > 3.7).,
Environnement : Le script doit s'exécuter sous Linux.,
Versionnage : Le projet doit être géré avec Git et GitHub.,
Qualité :
Le code doit être documenté (docstrings).
Une documentation technique doit être générée avec Sphinx.,
Des tests unitaires doivent être présents.,
,
,

Utilisation du Programme
------------------------
Le script principal doit pouvoir être lancé en ligne de commande en spécifiant le dossier des notes et le dossier de sortie :

.. code-block:: bash

   ./nom_projet.py --files-dir dir_notes --output-dir ./../html/

Structure des Fichiers
----------------------
Le projet doit respecter une arborescence précise, séparant :

Le code source (nomprojet/),
Les données (data/),
La documentation (docs/),
Le site généré (html/),
Les tests (tests/)
