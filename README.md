# Scripts pour WinIBW

## Introduction

Contractuel dans une bibliothèque universitaire, j'ai eu l'occasion de travailler à nouveau avec WinIBW, notamment dans le cadre de chantiers relativement répétitifs. J'ai donc essayé de pousser plus loin ma familiarité avec ce logiciel en essayant d'utiliser des scripts simples pour effectuer des opérations basiques plus rapidement, comme rechercher un PPN. Au fil des semaines, j'ai voulu en savoir plus sur les possibilités qu'offrent cette fonctionnalité, pour finalement me rendre compte que l'internet ne proposent tant de ressources à ce sujet (mais bien plus que je ne pensais à l'origine).

Je ne suis pas un expert en VBS ou en informatique, loin de-là. Ces scripts sont peut-être une hérésie pour des gens compétents. De plus, certains scripts sont pensés pour répondre à mes besoins dans mon environnement, ce qui veut dire qu'ils ne fonctionnent pas dans toutes les situations imaginables. Toutefois, s'ils peuvent donner des idées à ne serait-ce qu'une personne, j'en serai ravi.

Ces informations en tête, il est, je pense, préférable de bien prendre le temps de lire et comprendre le script avant toute utilisation, et le modifier si nécessaire.

À noter : je vais à terme retravailler mes scripts, parce que certains ne sont pas très très jolis à voir.

## Les notations spéciales
De manière générale, j'essaye d'utiliser une structure similaire entre mes scripts, notamment en terme d'appellation de certaines variables.

### Pour les champs en Unimarc
→ U + _{type de notice}_ + _{champ}_ + _{sous-champ}_

Avec :
* type de notice :
  * "A" pour les notices d'autorité auteur ;
  * "B" pour les notices bibliographiques ;
* champ : le champ sous forme de nombre ;
* sous-champ :
  * lettre minuscule ;
  * "S" + le chiffre.

Exemples :
* `UB200a` → dans une notice bibliographique, le sous-champ `a` de la zone 200 ;
* `UA700S4` → dans une notice d'autorité auteur, le sous-champ `4` de la zone 700.

### Pour les informations à modifier selon son environnement de travail
Certaines informations propres à ma bibliothèque ont été remplacées par des expressions génériques.

→ $\_$#$\_$ + _{l'information}_ + $\_$#$\_$

Exemple :
* `$_$#$_$RCR$_$#$_$` → le RCR.

## Des reliquats de notes personnelles
Certaines parties de mes scripts ne servent qu'à moi et peuvent se trouver dans le code si j'oublie de les retirer. En l'occurrence :
* les notes sur les raccourcis en début de code ;
* la notion `_A_MOD_` en début de code.

De plus, chaque script exportant du contenu contient la notation `$_#_$` au début de l'export, qui sert uniquement à traverser plus vite le document exporté.

## La validation automatique
Il est à noter que normalement, aucun des scripts qui effectueraient des modifications sur une notice ne se termine par une validation automatique de celles-ci [12/06/2021] : je préfère toujours pouvoir vérifier que tout est bon avant validation.

Toutefois, cette validation se met en place très facilement avec l'ajout de `Application.ActiveWindow.SimulateIBWKey "FR"` à la fin du script.

## L'absence de contrôle du type de notice
À l'heure actuelle, les scripts destinés à un type de notice particulier (lecture ou modification) ne contrôlent pas s'ils sont exécutés sur ce type de notice ou sur un autre. Je souhaite définitivement mettre ce contrôle en place, mais nous verrons, si j'y arrive, quand.

## Sources extérieures
Voici les sources des quelques scripts que j'ai récupérés sur l'internet, en espérant n'en avoir oublié aucun :

1. CountOccurrences : [VBScript - Count occurrences in a text string / Stephen Millard, publié le 30 juillet 2009](https://www.thoughtasylum.com/2009/07/30/VB-Script-Count-occurrences-in-a-text-string/) [cons. le 29/05/2021]

1. Sleep : [Réponse de Original Paulie D à la question How to set delay in vbscript de Mark posée le 13 novembre 2009 sur StackOverflow](https://stackoverflow.com/questions/1729075/how-to-set-delay-in-vbscript#answer-12921137) [cons. le 29/05/2021]

1. ExportVar : [VBScript Text Files: Read, Write, Append / MrNetTek, publié le 19 novembre 2015](http://eddiejackson.net/wp/?p=8619) [cons. le 29/05/2021]

## Documentation sur mes scripts

Note : à l'heure actuelle, certaines informations complémentaires se situent dans les notes au début du script, notamment les scripts nécessaires au fonctionnement.

### `SUB addUA400()`

Rajoute des 400 pour les noms composés à une autorité auteur.

Se base sur sur le champ 200.


### `SUB addUB700S3()`

Remplace la 700 actuelle de la notice bibliographique par une 700 contenant le PPN du presse-papier et le $4 de l'ancienne 700.

Aussi, remplace le $btm des exemplaires du RCR, ou signale la présence de plusieurs exemplaires dans l'ILN.

Pour un bon fonctionnement, le PPN de l'auteur doit être copié dans le presse papier.


### `SUB ChantierTheseAddUB183()`

Ajoute une 183 en fonction de la 215 (notamment des chiffres détectés dans le $a) dans le cadre du chantier thèse.

### `SUB chantierTheseAutoriteAuteur()`

Crée une notice d'autorité auteur à partir de la notice dans le presse papier dans le cadre du chantier thèse.

Applique l'ajout des UA400 pour les noms composés.

Pour un bon fonctionnement, la notice doit être copiée dans le presse papier (générée depuis un document Excel particulier)


### `SUB chantierTheseLoopAddUB183()`

Exécute `ChantierTheseAddUB183`, sauf si l'utilisateur refuse l'ajout, sur la liste de PPN présente dans le presse-papier.

Exporte tous les 10 PPN traités puis à la fin.

Pour un bon fonctionnement, la liste de PPN doit provenir d'une colonne Excel et être copiée dans le presse-papier.


### `SUB CollerPPN()`

Recherche le PPN contenu dans le presse papier.


### `FUNC CountOccurrences(p_strStringToCheck, p_strSubString, p_boolCaseSensitive)`

Arguments explicites.

Renvoi le nombre d'occurrences.

[Source](https://www.thoughtasylum.com/2009/07/30/VB-Script-Count-occurrences-in-a-text-string/)


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

[Source](http://eddiejackson.net/wp/?p=8619)


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


### `SUB LastCHE()`

Affiche l'historique des recherches.


### `SUB searchExcelPPNList()`

Recherche la liste de PPN contenu dans le presse-papier.

Pour un bon fonctionnement, la liste de PPN doit provenir d'une colonne Excel et être copiée dans le presse-papier.


### `SUB Sleep([int]time)`

`time` en secondes.

Permet de mettre en pause un script pendant t = `time`.

[Source](https://stackoverflow.com/questions/1729075/how-to-set-delay-in-vbscript#answer-12921137)


### `SUB toEditMode([bool]lgpMode, [bool]save)`

`lgpMode` : si true, passe en mode présentation

Passe en mode édition (ou présentation).
