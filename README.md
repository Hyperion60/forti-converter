# forti-converter
Conversion des fichiers de logs Fortigate vers le format CSV ou directement vers Excel

## Conversion vers Excel
Pour executer le programme de conversion vers Excel, vous avez besoin d'installer `openpyxl`.

Le package openpyxl est mal optimisé niveau RAM et ce script n'est à utiliser que pour les petits fichiers
(15 Go de RAM pour un fichier de 300 000 lignes de logs).

## Conversion vers CSV
Il faut prévoir 4.5 fois la taille du fichier en RAM pour assurer la bonne tenue de l'execution du script.

## Execution
- La première fenêtre qui s'ouvre permet de selectionner le fichier à convertir.
- La seconde fenêtre permet de selectionner le dossier de sortie dans lequel sera écrit le ou les fichiers produits.
