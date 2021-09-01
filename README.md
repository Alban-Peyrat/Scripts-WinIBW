# Scripts pour WinIBW

Les scripts proposés visent généralement à accélérer des traitements répétitifs dans WinIBW. Certains d'entre eux, classés en tant que concepts, visent à contrôler des données sans devoir les modifier via des outils externes type tableur.

## Contexte du développement

J'ai développé ces scripts lors de mon contrat à la Bibliothèque universitaire des Sciences du Vivant et de la Santé - Josy Reiffers (Bordeaux). Certaines de mes missions impliquaient des tâches répétitives ou plusieurs saisies d'une même donnée.

Ce sont donc posées les questions suivantes :
* si dans 9 cas sur 10 je dois écrire la même information, n'est-il pas possible de l'écrire automatiquement et modifier manuellement le cas restant ?
* si je dois écrire plusieurs fois la même donnée, n'est-il pas possible de l'écrire une fois, la dupliquer et ainsi éviter des erreurs de saisies (et gagner du temps) ?
* si je génère des données plus rapidement, ne m'est-il pas possible de contrôler les erreurs de saisie de certaines de celles-ci sans effectuer un traitement sur Excel ?

## De l'usage de ces scripts

Certains scripts sont pensés pour répondre à mes besoins dans mon environnement, ce qui veut dire qu'ils ne fonctionnent pas dans toutes les situations imaginables.

Ces informations en tête, il est, je pense, préférable de bien prendre le temps de lire et comprendre le script avant toute utilisation, et le modifier si nécessaire, notamment car certains contiennent des données propres à mon établissement.

De plus, certains de ces scripts seront peut-être sujets à des modifications, notamment car ils ne sont pas toujours très jolis à voir.

Rappel : pour installer les scripts dans WinIBW, référez-vous au [guide pour les scripts utilisateurs de l'Abes](http://documentation.abes.fr/sudoc/manuels/logiciel_winibw/scripts/index.html#CreerScriptUtilisateur).

## Organisation des scripts

Les scripts sont divisés entre trois fichiers différents :
* [scripts principaux](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs), qui contient majoritairement les scripts à activer :
  * AddSujetRAMEAU ;
  * addUA400 ;
  * addUB700S3 ;
  * changeExAnom ;
  * ChantierTheseAddUB183 ;
  * ChantierTheseLoopAddUB183 ;
  * decompUA200enUA400 ;
  * getCoteEx ;
  * getDataUAChantierThese ;
  * getTitle ;
  * getUA810b ;
  * getUB310 ;
  * PurifUB200a ;
  * searchExcelPPNList ;
* [scripts ressources](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs), qui contient les scripts qui facilitent l'exécution des autres :
  * appendNote ;
  * CountOccurrences ;
  * exportVar ;
  * goToTag ;
  * goToTagInputBox ;
  * Sleep ;
  * toEditMode ;
  * uCaseNames ;
* [concepts](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/concepts.vbs), qui contient des concepts que je n'utilise pas mais qui théoriquement fonctionnent, ou des scripts de mon bac à sable que je pense utiles à partager. Cerains sont en train d'être remplacés par des procécédes proches de CoCo-SAlma2 dans l'outil ConS*tance* :
  * ctrlUA103eqUA200f ;
  * ctrlUB700S3.

## Notations des champs en Unimarc

De manière générale, j'essaye d'utiliser une structure similaire entre mes scripts, notamment pour les champs UNIMARC :

U + _type de notice_ + _champ_ + _sous-champ_

Avec :
* type de notice :
  * "A" pour les notices d'autorité auteur ;
  * "B" pour les notices bibliographiques ;
* champ : le champ sous forme de nombre ;
* sous-champ :
  * lettre minuscule ;
  * "S" + le chiffre.

Exemples :
* `UB200a` : dans une notice bibliographique, le sous-champ `a` de la zone 200 ;
* `UA700S4` : dans une notice d'autorité auteur, le sous-champ `4` de la zone 700.

## Les informations à modifier selon son environnement de travail

Certaines informations propres à ma bibliothèque sont à remplacer :
* le RCR de ma bibliothèque (330632101) ;
* le chemin d'accès au profil WinIBW (C:\/oclcpica/WinIBW30/Profiles).

## Des reliquats de notes personnelles

Certaines parties de mes scripts ne servent qu'à moi et peuvent se trouver dans le code si j'oublie de les retirer. En l'occurrence :
* les notes sur les raccourcis en début de code ;
* la notion `_A_MOD_` en début de code.

De plus, chaque script exportant du contenu contient la notation `$_#_$` au début de l'export, qui sert uniquement à traverser plus vite le document exporté.

## La validation automatique

Il est à noter que normalement, aucun des scripts qui effectueraient des modifications sur une notice ne se termine par une validation automatique de celles-ci [24/08/2021] : je préfère toujours pouvoir vérifier que tout est bon avant validation.

Toutefois, cette validation se met en place très facilement avec l'ajout de `Application.ActiveWindow.SimulateIBWKey "FR"` à la fin du script.

## L'absence de contrôle du type de notice

À l'heure actuelle, les scripts destinés à un type de notice particulier (lecture ou modification) ne contrôlent pas s'ils sont exécutés sur ce type de notice ou sur un autre. J'envisage à terme d'en configurer un, si j'y arrive.

## Sources extérieures

Voici les sources des quelques scripts que j'ai récupérés sur l'internet, en espérant n'en avoir oublié aucun :

1. CountOccurrences : [VBScript - Count occurrences in a text string / Stephen Millard, publié le 30 juillet 2009](https://www.thoughtasylum.com/2009/07/30/VB-Script-Count-occurrences-in-a-text-string/) [cons. le 29/05/2021]

1. Sleep : [Réponse de Original Paulie D à la question How to set delay in vbscript de Mark posée le 13 novembre 2009 sur StackOverflow](https://stackoverflow.com/questions/1729075/how-to-set-delay-in-vbscript#answer-12921137) [cons. le 29/05/2021]

1. ExportVar : [VBScript Text Files: Read, Write, Append / MrNetTek, publié le 19 novembre 2015](http://eddiejackson.net/wp/?p=8619) [cons. le 29/05/2021]

## Liste des modifications

* le 02/08/2021
  * suppression de `PurifUB200a` car peu d'intérêts à être partagé ;
  * suppression de `CollerPPN` car peu d'intérêts à être partagé ;
  * suppression de `LastCHE` car peu d'intérêts à être partagé.
* le 23/08/2021 :
  * ajout de `AddSujetRAMEAU` pour ajouter des 60X ;
  * ajout de `ctrlTraitementInterne` ;
  * ajout de `getUB310` pour récupérer dans le presse-papier l'information de la première 310 ;
  * ajout de `PurifUB200a` pour adapter un titre à son écriture en UNIMARC ;
  * scission de `addUB700S3` : la partie sur l'exemplaire a été isolée dans un nouveau script, `changeExAnom`.
* le 24/08/2021 :
  * [répartition des scripts entre plusieurs fichiers](https://github.com/Alban-Peyrat/Scripts-WinIBW#organisation-des-scripts) ;
  * actualisation des présentations des scripts, notamment en intégrant les dernières modifications ;
  * adaptation du projet pour être cohérent avec les autres outils.
* le 25/08/2021 :
  * suppression de `ctrlTraitementInterne`, que j'avais dû arrêter en plein milieu du développement ;
  * modification de la description de [concepts](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/concepts.vbs) et ajout de la mention de ConS*tance* ;
* le 01/09/2021 :
  * ajout de `appendNote` pour ajouter à une variable la donnée voulue ;
  * ajout de `getDataUAChantierThese` pour exporter les données d'une thèse dans le cadre d'un chantier sur les thèses ;
  * ajout de `uCaseNames` pour mettre des majuscules aux noms renseignés ;
  * modification de `getCoteEx` dû à une réécriture du script. Détecte désormais l'intégralité des cotes associées au RCR et permet de sélectionner celles voulues, ou toutes ;
  * probable mise à jour prochaine de `decompUA200enUA400` pour être plus efficace et utiliser `uCaseNames`.

## Présentation des scripts

### `SUB AddSujetRAMEAU()`

Ouvre une boîte de dialogue permettant d'insérer des UB60X à partir du PPN.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-addsujetrameau).


### `SUB addUA400()`

Rajoute des UA400 pour les noms composés à une autorité auteur en se basant sur la UA200.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-addua400).


### `SUB addUB700S3()`

Remplace la UB700 actuelle de la notice bibliographique par une UB700 contenant le PPN du presse-papier et le $4 de l'ancienne UB700.

Contient aussi un appel du [script supprimant des anomalies dans les exemplaires](https://github.com/Alban-Peyrat/Scripts-WinIBW#sub-changeexanomnotice).

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-addub700s3).


### `FUNCTION appendNote(var, text)`

Renvoie `var` comme équivalent à `text` si `var` était vide, sinon, renvoie `var` suivi d'un saut de ligne puis de `text`.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#function-appendnotevar-text).


### `SUB changeExAnom(notice)`

Remplace le $btm de la zone eXX associée au RCR par $bx ou signale la présence de plusieurs eXX associées à ce RCR.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-changeexanomnotice).


### `SUB ChantierTheseAddUB183()`

Ajoute une UB183 en fonction de la UB215 (notamment des chiffres détectés dans le $a).

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-chantiertheseaddub183).


### `SUB chantierTheseLoopAddUB183()`

Exécute `ChantierTheseAddUB183`, sauf si l'utilisateur refuse l'ajout, sur la liste de PPN présente dans le presse-papier et exporte un rapport des modifications ou non effectuées.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-chantiertheseloopaddub183).


### `FUNC CountOccurrences(p_strStringToCheck, p_strSubString, p_boolCaseSensitive)`

Renvoi le nombre d'occurrences.

[Consulter la source originale](https://www.thoughtasylum.com/2009/07/30/VB-Script-Count-occurrences-in-a-text-string/), [consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#func-countoccurrencesp_strstringtocheck-p_strsubstring-p_boolcasesensitive).


### `SUB ctrlUA103eqUA200f()`

Exporte et compare le $a de UA103 et le $f de UA200 pour chaque PPN de la liste présente dans le presse-papier.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/concepts.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-ctrlua103equa200f).


### `SUB ctrlUB700S3()`

Exporte le premier $ de UB700 pour chaque PPN de la liste présente dans le presse-papier.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/concepts.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-ctrlub700s3).


### `FUNC decompUA200enUA400([string]impUA200)`

Renvoie les UA400 créés à partir de la décomposition du nom composé du UA200 importé (`impUA200`).

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#func-decompua200enua400stringimpua200).


### `SUB exportVar(var, boolAppend)`

Exporte `var` dans `export.txt` (même emplacement que `winibw.vbs`), réécrivant le fichier si `boolAppend` est false. Est utilisé par toutes les procédures qui exporte des données.

[Consulter la source originale](http://eddiejackson.net/wp/?p=8619), [consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-exportvarvar-boolappend).


### `SUB getCoteEx()`

Renvoie dans le presse-papier la cote du document. Si plusieurs cotes sont présentes, donne le choix entre en sélectionner une, ou toutes les sélectionner, permettant également de choisir le séparateur.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-getcoteex).


### `SUB getDataUAChantierThese()`

Copie dans le presse-papier le PPN, l'année de soutenance, la discipline, le patronyme, le prénom, l'année de naissance, le sexe, le titre et la cote du document, séparés par des tabulations horizontales.

Créé dans le cadre d'un chantier sur les thèses, l'exploitation de ces données se fait dans un tableur Excel particulier.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-getdatauachantierthese).


### `SUB getTitle()`

Renvoie dans le presse-papier le titre du document en remplaçant les @ et $e. Si le titre est entièrement en majuscule, le renvoie en minuscule (sauf première lettre).

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-gettitle).


### `SUB getUA810b()`

Si un seul UA810 est présent, écrit le $b "né le" à partir des informations de la UA103 de la notice, sinon, renvoie le $b dans le presse-papier.

Pour un bon fonctionnement, la UA103 doit comprendre AAAAMMJJ.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-getua810b).


### `SUB getUB310()`

Copie dans le presse-papier la valeur du premier UB310.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-getub310).


### `SUB goToTag([string]tag, [string, "none" pour empty]subTag, [bool]toEndOfField, [bool]toFirst, [bool]toLast)`

Attention, `subTag` ne doit pas contenir le $ ET est sensible à la casse.

Place le curseur à l'emplacement indiqué par les paramètres. Si plusieurs occurrences sont rencontrées sans que `toFirst` ou `toLast` soit true, une boîte de dialogue s'ouvre pour sélectionner l'occurrence souhaitée.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-gototagstringtag-string-none-pour-emptysubtag-booltoendoffield-booltofirst-booltolast).


### `SUB goToTagInputBox()`

Permet d'essayer `goToTag` en indiquant les paramètres voulus.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-gototaginputbox).


### `FUNCTION PurifUB200a(UB200, isUB541)`

Renvoie l'adaptation d'un titre en son écriture en UNIMARC.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#function-purifub200aub200-isub541).


### `SUB searchExcelPPNList()`

Recherche la liste de PPN contenue dans le presse-papier.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-searchexcelppnlist).


### `SUB Sleep([int]time)`

Permet de mettre en pause un script pendant t = `time` (en secondes).

[Consulter la source originale](https://stackoverflow.com/questions/1729075/how-to-set-delay-in-vbscript#answer-12921137), [consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-sleepinttime).


### `SUB toEditMode([bool]lgpMode, [bool]save)`

Passe en mode édition (ou présentation).

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#sub-toeditmodeboollgpmode-boolsave).


### `FUNCTION uCaseNames(noms)`

Renvoie `noms` après avoir mis une majuscule au début de chaque nom renseigné.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs), [consulter la documentation complète](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/documentation.md#function-ucasenamesnoms).
