En cours de construction et d'actualisation :

### `SUB addUA400()`

Rajoute des 400 pour les noms composés à une autorité auteur.

Se base sur sur le champ 200.


### `SUB addUB700S3()`

Remplace la 700 actuelle de la notice bibliographique par une 700 contenant le PPN du presse-papier et le $4 de l'ancienne 700.

Aussi, remplace le $btm des exemplaires du RCR, ou signale la présence de plusieurs exemplaires dans l'ILN.

Pour un bon fonctionnement, le PPN de l'auteur doit être copié dans le presse papier.


### `SUB ChantierTheseAddUB183()`

Ajoute une 183 en fonction de la 215 (notamment des chiffres détectés dans le $a) dans le cadre du chantier thèse.


### `SUB chantierTheseLoopAddUB183()`

Exécute `ChantierTheseAddUB183`, sauf si l'utilisateur refuse l'ajout, sur la liste de PPN présente dans le presse-papier.

Exporte tous les 10 PPN traités puis à la fin.

Pour un bon fonctionnement, la liste de PPN doit provenir d'une colonne Excel et être copiée dans le presse-papier.



### `FUNC CountOccurrences(p_strStringToCheck, p_strSubString, p_boolCaseSensitive)`

Arguments explicites.

Renvoi le nombre d'occurrences.

[Consulter la source](https://www.thoughtasylum.com/2009/07/30/VB-Script-Count-occurrences-in-a-text-string/)


### `SUB ctrlUA103eqUA200f()`

Exporte et compare le $a de UA103 et le $f de UA200 pour chaque PPN de la liste présente dans le presse-papier.

Pour un bon fonctionnement, la liste de PPN doit provenir d'une colonne Excel et être copiée dans le presse-papier.


### `SUB ctrlUB700S3()`

Exporte le premier $ de UB700 pour chaque PPN de la liste présente dans le presse-papier.

Pour un bon fonctionnement, la liste de PPN doit provenir d'une colonne Excel et être copiée dans le presse-papier.


### `FUNC decompUA200enUA400([string]impUA200)`

`impUA200` doit être le champ 200.

Renvoi les champs 400 créés à partir de la décomposition du nom composé du champ 200 importé.


### `SUB exportVar(var, boolAppend)`

`var` est l'information à exporter ; si `boolAppend` est false, réécrit le fichier.

Exporte dans `export.txt` (même emplacement que `winibw.vbs`).

Est utilisée par toutes les procédures qui exporte des données.

[Consulter la source](http://eddiejackson.net/wp/?p=8619)


### `SUB getCoteEx()`

Renvoie dans le presse-papier la cote du document pour ce RCR (malfonctionne s'il y a plusieurs exemplaires de ce RCR)


### `SUB getTitle()`

Renvoie dans le presse papier le titre du document en remplaçant les @ et $e.

Si le titre est entièrement en majuscule, le renvoie en minuscule (sauf première lettre).


### `SUB getUA810b()`

Si un seul UA810 est présent, écrit le $b "né le" à partir des informations de la 103de la notice.

Si plusieurs UA810 sont présents, renvoie le $b dans le presse-papier.

Pour un bon fonctionnement, la 103 doit comprendre AAAAMMJJ.


### `SUB goToTag([string]tag, [string, "none" pour empty]subTag, [bool]toEndOfField, [bool]toFirst, [bool]toLast)`

`subTag` ne doit pas contenir le $ ET est sensible à la casse. Le reste est explicite.

Place le curseur à l'emplacement indiqué par les paramètres. Si plusieurs occurrences sont rencontrées sans que `toFirst` ou `toLast` soit true, une boîte de dialogue s'ouvre pour sélectionner l'occurrence souhaitée.


### `SUB goToTagInputBox()`

Permet d'essayer `goToTag` en indiquant les paramètres voulus.


### `SUB searchExcelPPNList()`

Recherche la liste de PPN contenu dans le presse-papier.

Pour un bon fonctionnement, la liste de PPN doit provenir d'une colonne Excel et être copiée dans le presse-papier.


### `SUB Sleep([int]time)`

`time` en secondes.

Permet de mettre en pause un script pendant t = `time`.

[Consulter la source](https://stackoverflow.com/questions/1729075/how-to-set-delay-in-vbscript#answer-12921137)


### `SUB toEditMode([bool]lgpMode, [bool]save)`

`lgpMode` : si true, passe en mode présentation

Passe en mode édition (ou présentation).
