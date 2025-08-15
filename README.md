# Outil interactif d'analyse acoustique

## Description

Ce projet propose un outil interactif d'analyse acoustique utilisant **Streamlit**. Il permet de comparer les courbes d'absorption acoustique à partir de données contenues dans deux fichiers Excel. L'application permet à l'utilisateur de télécharger ces fichiers, de sélectionner des paramètres (épaisseur et densité), et de visualiser les courbes d'absorption sous forme de graphiques. De plus, l'utilisateur peut télécharger le graphique généré au format PDF.

## Fonctionnalités

- **Téléchargement de fichiers Excel** : Permet de télécharger deux fichiers Excel contenant des données acoustiques.
- **Sélection de paramètres** : L'utilisateur peut choisir l'épaisseur et la densité des matériaux pour ajuster les courbes d'absorption.
- **Comparaison visuelle** : Affiche les courbes d'absorption des deux fichiers Excel dans un graphique pour comparaison.
- **Téléchargement du graphique** : Permet de télécharger le graphique comparatif au format PDF.

URL Vers l'appli : **https://kundtpy.streamlit.app/**

Utilisation
Télécharger les fichiers Excel : Cliquez sur les boutons "Télécharger le premier fichier Excel" et "Télécharger le deuxième fichier Excel" pour importer les données de vos fichiers Excel.

Sélectionner les paramètres : Utilisez les menus déroulants pour choisir l'épaisseur et la densité du matériau que vous souhaitez analyser.

Voir les courbes d'absorption : Les courbes d'absorption acoustique des deux fichiers seront tracées en fonction des paramètres sélectionnés.

Télécharger le graphique : Un bouton vous permet de télécharger le graphique comparatif sous forme de fichier PDF.

Structure du code
load_data_from_excel(file) : Fonction qui charge les données depuis un fichier Excel et extrait les fréquences, les épaisseurs, les densités et les coefficients d'absorption.
save_as_pdf(fig) : Fonction qui sauvegarde le graphique généré sous forme de fichier PDF.
L'application utilise Matplotlib pour générer les graphiques et Streamlit pour l'interface utilisateur interactive.
Exemple de structure des fichiers Excel
Les fichiers Excel doivent contenir les données suivantes :

La première colonne pour les fréquences (en Hz).
Les autres colonnes pour les coefficients d'absorption pour différentes combinaisons d'épaisseur et de densité.

Contributions
Les contributions sont les bienvenues ! Si vous avez des suggestions ou des améliorations à proposer, n'hésitez pas à ouvrir une issue ou à soumettre une pull request.

Licence
Ce projet est sous licence MIT. Consultez le fichier LICENSE pour plus d'informations.

