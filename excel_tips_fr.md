# EXCEL - version française

### Raccourcis clavier
Tous les raccourcis indiqués sont pour Azery sauf indication contraire.

**Ctrl+Plus** : Insérer une ligne

**Ctrl+Minus**: Supprimer une ligne

**Ctrl+Space** : Permet de sélectionner la totalité des cellules de la colonne dans laquelle le curseur est placé

**Ctrl+Space** suivi par **Ctrl+Plus** (**Ctrl+Minus**) permet d'insérer (supprimer) une colonne.

**Alt+"** : Masquer les lignes sélectionnées

**Alt+_** : Afficher les lignes masquées

**Alt+(** : Masquer les colonnes sélectionnées

**Alt+)**: Afficher les colonnes masquées

**Alt+Shift+Flèche de droite** : Grouper les colonnes

**Alt+Shift+Flèche de gauche** : Dégrouper les colonnes


**Ctrl+Shift+L** : 1) Faire apparaître les boutons de filtres aux en-têtes des colonnes du tableau. 2) Pour une table pour laquelle les filtres ont déjà été activées, permett d'effacer les filtres en affichant la totalité du tableau non filtrée.

**Ctrl+Flèche du haut** : Déplacer le curseur à la première ligne de la table (pour utiliser ce raccourci, le curseur doit préalablement être placé dans une cellule de la table concernée)

**Alt+Flèche du bas** : Afficher les options du filtre de l'entête d'une colonne du tableau (le curseur doit se trouver dans la cellule qui contient l'entête de colonne avec filtre)

**Space** : Permet d'activer ou de désactiver les options du filtre. Le déplacement entre les options du filtre se fait avec les flèches de direction. La touche Entrée valide le choix des options du filtre.

**Ctrl+L** :  Ouvre une fenêtre permettant de convertir une plage de cellules en un tableau structuré.

**Alt+=** : Pour un tableau structuré, permet de calculer la somme des valeurs d'une colonne (le curseur doit être placé dans la première cellule vide en bas de la colonne). **Alt+Flèche du bas** permet de choisir un autre calcul (ex. moyenne) au lieu de la somme.

**Ctrl+Shift+1** (Azerty) : Afficher la fenêtre Format de cellule.

**Ctrl+Tab** : Navigation entre les onglets d'une fenêtre de dialogue.
**Ctrl+Shift+Tab** : Retour en arrière dans la navigation entre les ongles d'une fenêtre de dialogue.

**Alt+(Fn)+F1** :  Créer un graphique à partir d'un tableau, sur la même feuille. (Le curseur doit être placé dans une cellule de ce tableau)

**(Fn)+F11** : Créer un graphique à partir d'un tableau, sur une nouvelle feuille.





### RECHERCHEV
(recherche verticale)

**Recherche de correspondance exacte :**
Par exemple, pour trouver dans un tableau un nom ou un salaire correspondant à une matriculé d'un salarié donnée.
Cette fonction demande beaucoup de calculs comme elle travaille sur l'ensemble de la table qui dans laquelle les données sont recherchées.

=RECHERCHEV(G1;A2:D5; 2; FAUX)
**G1** : valeur d'intérêt (ex. identifiant unique permettant de répérer la ligne dans le tableau). Elle peut être indiquée en se référent à l'adresse de la cellule ou en entrant directement la valeur dans la formule. 
Ce paramètre peut également être représenté par une concaténation des valeurs (ex. B72&B73) ou par une partie de valeur (ex. partie 
de la chaîne de caractères) **NB**: la valeur d'intérêt doit être **unique** dans le tableau, sinon la formule va retourner les valeurs correspondantes à la première ligne avec cette valeur d'intérêt
**A2:D5** : plage dans lequel il faut rechercher les valeurs. **NB:** la valeur d'intérêt doit se trouver dans la 1ère colonne (la plus à gauche) de cette plage. Dans cet exemple, les identifiants uniques se trouvent dans la colonne A. S'il y a d'autres colonnes dans la table avant celle des identifiants uniques, il faut veiller à ce que la plage sélectionnée ait la colonne d'identifiants comme première colonne (même si ce n'est pas la première colonne du tableau).
Si le format des données de la valeur d'intérêt (G1) et de la 1ère colonne de la plage (A) ne sont pas le même, cela n'empêche pas la recherche.
**2** : indice de la colonne dans cette plage qui contient la valeur à retourner dans la cellule. Ces colonnes doivent se trouver obligatoirement à droite par rapport à la colonne A qui contient l'identifiant unique.
**FAUX** ou **0**:  paramètre indiquant qu'on cherche la corresondance exacte. Si ce paramètre n'est pas indiqué, Excel considère qu'il cherche la correspondance approximative et que la plage dans laquelle la recherche est effectuée est triée en ordre croissant. Donc si on cherche la valeur "Armoire" et dans la plage de référence les valeurs sont "Chaise", "Bureau", "Armoire", Excel voit que la première est Chaise (commençant par C) et considère donc qu'il n'y a pas de valeur commençant par A dans la table, donc la formule retourne N/A malgrué que la valeur "Armoire" est bien présente dans la table.

=RECHERCHEV(G$1;$A$2:$D$5; 3; FAUX)
S'il faut obtnir les valeurs de différentes colonnes pour la même valeur d'intérêt, on fige la valeur d'intérêt (ici, la coordonnée de la ligne uniquement: **G$1**) et la plage (ici, les coordonnées de lignes uniquement: **$A$2:$D$5**) et on indique le numéro de la colonne dans la plage qui contient la valeur d'intérêt (ici, **3**)

**NB** : La fonction RECHERCHEV ne supporte pas l'insertion ou la suppression des colonnes dans la table où la recherche doit être effectuée, elle doit donc être mise à jour dès que ces modifications sont effectuées.

**Gestion des erreurs:**
=SIERREUR(formule; "Message d'erreur à afficher")
Par exemple:
=SIERREUR(RECHERCHEV(A13;$A$3:$F$7;6;0); "Région non renseignée")
Si la valeur recherchée correspondat à la valeur d'intérêt n'est pas trouvée dans la table (ex. ici la région pour laquelle on recherche le montant des ventes ne figure pas dans la table), le message d'erreur prédéfini s'affiche (ici, "Région non renseignée").


**Recherche de correspondance approximative:**

=RECHERCHEV(B2;F2:G4;2;VRAI)
**VRAI** : paramètre indiquant qu'on cherche une correspondance approximative

Par exemple, pour un tableau qui contient le montant d'achat, retrouver les remises applicables en se référant à un tableau avec colonnes "montant d'achat" et "remise" (plage **F2:G4**) qui contient le pourcentage de remise applicable à partir d'un certain seuil (colonne **2** de la plage F2:G4).

 Par exemple, le tableau de remises contient la remise la ligne suivante: 1000 € (colonne montant d'achat) - 5 % (remise applicable à partir de 1000 €). Dans ce cas la formule va fonctionner de manière suivante:
- Pour un montant d'achat de **999** €: remise de 5 %  **n'est pas applicable**
- pour un montant d'achat de **1000** €: remise de 5 % est **applicable**
- pour un montant d'achat de **10001** € : remise de 5 % est **applicable**
***NB** : pour que les remises soient correctement appliquées, les montants d'achat (colonne F / colonne 1 de la plage F2:G4) doivent être classés par ordre décroissant, par exemple: ligne 1: 0 €, ligne 2: 1000 €, ligne 3: 10 000 € etc

### RECHERCHEX
Variante récente de la fonction RECHERCHEV qui permet de remedier à certains inconvénients de cette dernière:
- recherche s'effecture dans une colonne donnée et non dans la table entière, ce qui permet de limiter les calculs,
- recherche peut s'effectuer tant à la gauche qu'à droite par rapport à la colonne qui contient les valeurs d'intérêt,
- le syntaxe de formule faisant référence aux colonnes directement, la formue se met à jour automatiquement en cas d'ajout ou de suppression des colonnes dans la table dans laquelle la recherche est effectuée.

=RECHERCHEX(G6;B2:B11;C2:C11)
G6 : valeur d'intérêt
B2:B11 : la colonne de la table qui contient la valeur d'intérêt
C2:C11 : la colonne de la table qui conient la valeur à retourner

### RECHERCHEX renvoie la date 00/01/1900 au lieu de la valeur vide dans la colonne des dates
Dans ce cas, utiliser au lieu de format de date "date courte", utiliser le format personnalisé **jj/mm/aaaa;;**

### RECHERCHEX renvoie 0 au lieu de la valeur vide dans la colonne de texte
Dans ce cas, modifier la formule selon le modèle suivant:
- formule initiale:
=RECHERCHEX($A2;Payments!$B:$B;Payments!D:D)
- formule modifiée:
=SI(ESTVIDE(RECHERCHEX($A2;Payments!$B:$B;Payments!D:D));""; RECHERCHEX($A2;Payments!$B:$B;Payments!D:D))

### RECHERCHEH
Recherche horizontale (analogue à RECHERCHEV), mais la recheche s'effectue dans la ligne indiquée.
=RECHERCHEH(A12;A3:F7;5;0)
**A12**: valeur d'intérêt
**A3:F7**: plage dans lequel la recherche est effectuée
**5**: indice de la ligne dans laquelle la valeur cible se trouve
**FAUX** ou **0**:  paramètre indiquant qu'on cherche la corresondance exacte. 

### Utiliser une liste déroulante pour déterminer les valeurs qu'on peut entrer à partir d'une liste existante (validation des données)
Se placer dans la cellule dans laquelle la liste déroulante doit être proposée, ensuite:
Onglet Données -> zone Outils de Données -> Validation des données -> fenêtre Validation des données s'ouvre
Dans la fenêtre, dans l'option Autoriser choisir Liste. Dès que cette option est choisie, la ligne Source apparaît. 
Dans la ligne Source, il faut indiquer les cellules qui contiennent la liste qui doit utilisé en tant que liste déroulante.

### Donner un nom à une plage (plage nommée)
Sélectionner la plage souhaitée, dans la cellule d'adresse indiquer le nom de la plage (ex. _ventesParRegion). 
Par la suite, lors de l'entre d'une formule, on peut utiliser le nom de la plage au lieu des adresses des cellules.
Le nom de la plage ne doit pas utiliser d'espaces ni d'accents. Il est intéressant de commencer les noms des plages avec _ 
afin de pouvoir les sélectionner facilement lors de l'entrée de la formule.

### Trouver l'indice de la colonne de table avec une entête donnée (EQUIV)
=EQUIV(C52;C45:E45;0)
**C52** : cellule qui contient la valeur d'intérêt (ex. ici, la valeur contenue dans la cellule C52 était "Février")
**C45:E45** : plage qui contient les entêtes des colonnes (une seule ligne avec entêtes comme "Janvier", "Février", "Mars")
**0** : paramètre indiquant qu'on recherche la valeur exacte.
La fonction retourne le nombre (ex. 2) qui correspond à l'indice de la colonne dont l'entête correspond à la valeur d'intérêt ("Février)

### Remplacer les valeurs dans une colonne
Ex. remplacer tous les zéros 0 par des valeurs vides ""
Sélectionner la colonne où le remplacement doit être effectué -> Ctrl+f -> dans la fenêtre Rechercher et remplacer, aller dans l'onglet Remplacer, indiquer la valeur à remplacer et la valeur remplaçante, cliquer sur Remplacer tout. Le remplacement sera effectué dans la colonne sélectionnée seulement (et non dans la totalité du tableau/feuille)

### Concaténer les valeurs des cellules
=B61&C61
**B61** : contient la valeur "Produit A"
**&** : opérateur de concaténation
**C61** : contient la valeur "Janvier"
Résultat retourné dans la cellule avec formule: "Produit AJanvier"

### Récuperer la partie droite de la chaîne des caractères
- Si le nombre de caractères dans la partie de gauche est connu et est constant (par exemple, nous avons la colonne avec valeurs de type "Pays A - France", "Pays B - Belgique", "Pays C - Italie" et nous voulons récuperer uniquement les noms des pays "France", "Belgique", "Italie" afin de les utiliser dans une formule).
=DROITE(A88; NBCAR(A88)-9)
=DROITE(Valeur; Nombre de caractères)
**A88** : La cellule dans laquelle se trouver la chaîne de caractères source
**NBCAR(A88)-9** : nombre de caractères à récuperer à partir de la droite.
Détail de cette partie de la formule:
**NBCAR(A88**)** : Fonction calculant le nombre de caractères dans la cellule A88
**9** : Nombre constant des caracteres de gauche correspondant à la partie de la chaîne comme "Pays A - ", "Pays B - ", "Pays C - "

### SIERREUR 
- La fonction SIERREUR peut servir pour afficher si la formule indiquée en tant que 1er paramètre de la fonction retourne une erreur:
=SIERREUR(formule; "Message d'erreur à afficher")
- Au lieu du message d'erreur, il est possible de donner une formule qui sera exécutée en cas d'erreur.
=SIERREUR(RECHERCHEV(A105;$B$98:$C$101;2;0); RECHERCHEV(A105;$E$98:$F$101;2;0))
Ici, les 

### Tableau croisé dynamique ne calcule pas correctement la somme des durées sous format Heure
(Réponse testée depuis du site Microsoft) :
After Pivoting select your column which you want to "SUM" and then open "Format Cells" dialog box.
From "Number" tab select "Custom" from left side navigation bar. Then you will see this [$-x-systime]h:mm:ss AM/PM custom date/time format applied on the selected cells in pivot.
Now simply type [hh]:mm:ss this and Click on Ok.
This will return the summation of your time in pivot.