# QuadraImport
![image](https://github.com/CaptainCoding7/QuadraImport/assets/46071575/b8bfdeb4-cccb-4a58-b984-01aa069ff70b)

Ce logiciel permet de transformer un relevé bancaire (au format excel) en un fichier excel pouvant être importé dans le logiciel Quadra Compta. Voici la procédure à suivre pour importer un rélevé dans Quadra Compta grâce à Excel2Quadra.<br>
1.	Convertir le relevé du format pdf vers excel*<br>
  a.	Ouvrir un navigateur internet et chercher « Adobe convertir  pdf excel »<br>
  b.	Cliquer sur le premier lien<br>
  c.	Choisir le fichier pdf à convertir.<br>
  d.	Si Adobe demande de se connecter, se connecter avec son compte google<br>
  e.	Télécharge le relevé excel et le mettre dans le dossier Bureau/Excel2Quadra/releves_excel_bruts <br><br>
2.	Génération de l’excel transformé avec Excel2Quadra<br>
  a.	Ouvrir Excel2Quadra.exe<br>
  b.	Charger d’abord le fichier excel d’origine (format .xlsx) en cliquant sur le bouton « Sélectionner… ». Il se trouve dans le dossier releves_excel_bruts<br>
  c.	Choisissez les options de génération<br>
  d.	Lancer la génération du nouvel excel transformé en cliquant sur « Transformer le relevé »<br>
  e.	Le fichier a été créé dans le dossier outputs_releves_transformes<br><br>

3.	A partir du nouveau fichier excel, créer un fichier .txt importable dans Quadra Compta :<br>
  a.	Ourvrir l’excel généré<br>
  b.	Cliquer sur l’onglet « Compléments » (nous utiliserons la macro Qimport installé sur le PC)<br>
  c.	Cliquer sur l’icone du petit dossier jaune et choisir le fichier Import_B1.prm<br>
  d.	Valider en cliquant sur la petite encoche verte<br>
  e.	Entrer le numéro du journal de relevé bancaire (B1 en général)<br>
  f.	Choisir euro<br>
  g.	Garder QExport.txt comme nom de fichier et choisir le dossier Bureau/Excel2Quadra/txt pour import quadra ASCII<br>
  h.	Valider et si un message sur une erreur de date apparait, choisir une date par défaut comme 010123<br>
  i.	Le fichier a été créé dans Bureau/Excel2Quadra/txt pour import quadra ASCII<br><br>

4.	Import le relevé txt (ASCII) dans Quadra Compta<br>
  a.	Choisir Suivi dossier=> Import ASCII<br>
  b.	Sélectionner le fichier QExport.txt généré précédemment en allant dans le dossier Bureau/Excel2Quadra/txt pour import quadra ASCII<br>
  c.	Importer <br>

