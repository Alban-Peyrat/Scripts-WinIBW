# Scripts pour WinIBW

__Rappel : pour installer les scripts dans WinIBW, référez-vous au [guide pour les scripts utilisateurs de l'Abes](http://documentation.abes.fr/sudoc/manuels/logiciel_winibw/scripts/index.html#CreerScriptUtilisateur).__

_[Cliquez ici pour atteindre la liste des modifications.](https://github.com/Alban-Peyrat/Scripts-WinIBW#liste-des-modifications)_

_Version du 15/10/2021. Une refonte est en cours de réflexion. Tous les changements apportés le 13/10/2021 ne sont pas encore écrits dans la documentation. Les scripts en revanche sont bien actualisés._

Les scripts proposés visent généralement à accélérer des traitements répétitifs dans WinIBW. Certains d'entre eux, classés en tant que concepts, visent à contrôler des données sans devoir les modifier via des outils externes type tableur.

## De l'usage de ces scripts

Certains scripts sont pensés pour répondre à mes besoins dans mon environnement, ce qui veut dire qu'ils ne fonctionnent pas dans toutes les situations imaginables.

Ces informations en tête, il est, je pense, préférable de bien prendre le temps de lire et comprendre le script avant toute utilisation, et le modifier si nécessaire, notamment car certains contiennent des données propres à mon établissement.

De plus, certains de ces scripts seront peut-être sujets à des modifications, notamment car ils ne sont pas toujours très jolis à voir.

## Notations des champs en Unimarc

De manière générale, j'essaye d'utiliser une structure similaire entre mes scripts, notamment pour les champs UNIMARC :

U + _type de notice_ + _champ_ + _sous-champ_

Avec :
* type de notice :
  * `A` pour les notices d'autorité auteur ;
  * `B` pour les notices bibliographiques ;
* champ : le champ sous forme de nombre ;
* sous-champ :
  * lettre minuscule ;
  * `S` + le chiffre.

Exemples :
* `UB200a` : dans une notice bibliographique, le sous-champ `a` de la zone 200 ;
* `UA700S4` : dans une notice d'autorité auteur, le sous-champ `4` de la zone 700.

## Les informations à modifier selon son environnement de travail

Certaines informations propres à ma bibliothèque sont à remplacer :
* le RCR de ma bibliothèque (330632101) ;
* le chemin d'accès au profil WinIBW (C:\/oclcpica/WinIBW30/Profiles).

## La validation automatique

Il est à noter que normalement, aucun des scripts qui effectueraient des modifications sur une notice ne se termine par une validation automatique de celles-ci : je préfère toujours pouvoir vérifier que tout est bon avant validation.

Toutefois, cette validation se met en place très facilement avec l'ajout de `Application.ActiveWindow.SimulateIBWKey "FR"` à la fin du script.

## L'absence de contrôle du type de notice

À l'heure actuelle, les scripts destinés à un type de notice particulier (lecture ou modification) ne contrôlent pas s'ils sont exécutés sur ce type de notice ou sur un autre. J'envisage à terme d'en configurer un, si j'y arrive.

## Sources extérieures

Voici les sources des quelques scripts que j'ai récupérés sur l'internet, en espérant n'en avoir oublié aucun :

1. CountOccurrences : [VBScript - Count occurrences in a text string / Stephen Millard, publié le 30 juillet 2009](https://www.thoughtasylum.com/2009/07/30/VB-Script-Count-occurrences-in-a-text-string/) [cons. le 29/05/2021]

1. Sleep : [Réponse de Original Paulie D à la question How to set delay in vbscript de Mark posée le 13 novembre 2009 sur StackOverflow](https://stackoverflow.com/questions/1729075/how-to-set-delay-in-vbscript#answer-12921137) [cons. le 29/05/2021]

1. ExportVar : [VBScript Text Files: Read, Write, Append / MrNetTek, publié le 19 novembre 2015](http://eddiejackson.net/wp/?p=8619) [cons. le 29/05/2021]

## Présentation des scripts

### Scripts principaux

Ce fichier réunit majoritairement les scripts à exécuter ou de traitement.

#### `add18XmonoImp`

Ajoute les 181-2-3 pour une monographie imprimée sans illustration.

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà puis insère à la fin de celle-ci :
* `181 ##$P01$ctxt`
* `182 ##$P01$cn`
* `183 ##$P01$anga`
* un retour à la ligne.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `add214Elsevier`

Ajoute une 214 type pour Elsevier (2021).

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà puis insère à la fin de celle-ci :
* `214 #0$aIssy-les-Moulineaux$cElsevier Masson SAS$dDL 2021`
* un retour à la ligne

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `addBibgFinChap`

Ajoute une mention de bibliographie à la fin de chaque chapitre.

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà puis insère à l'emplacement du curseur :
* `Chaque fin de chapitre comprend une bibliographie`

__Malfonctionnement possible : si la notice n'était pas en mode édition, le texte ne s'écrira probablement pas si la grille des données codées n'est pas affichée.__

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `addCouvPorte`

Ajoute le début d'une 312 `La couverture porte en plus`.

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà puis insère à la fin de celle-ci :
* `312 ##$aLa couverture porte en plus : "`

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `addISBNElsevier`

Ajoute une 010 avec le début de l'ISBN d'Elsevier.

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà puis insère à la fin de celle-ci :
* `010 ##$A978-2-294-`

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `AddSujetRAMEAU`

Ouvre une boîte de dialogue permettant d'insérer des UB60X à partir du PPN.

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà, puis lance une boucle qui s'exécutera jusqu'à 1000 fois. À chaque exécution, ouvre une boite de dialogue permettant de coller directement le PPN et montrant la liste des commandes supplémentaires disponibles :
* ajouter `UX` devant le PPN (sans espace) permet de choisir la 60X à insérer :
  * par défaut, le script ajoute une `606 ##` ;
  * `U0` pour insérer une `600 #1` ;
  * `U1` pour insérer une `601 02` ;
  * `U2` pour insérer une `602 ##` ;
  * `U4` pour insérer une `604 ##` ;
  * `U5` pour insérer une `605 ##` ;
  * `U7` pour insérer une `607 ##` ;
  * `U8` pour insérer une `608 ##` ;
* ajouter `_[IndicateurNo1][IndicateurNo2]` après le PPN (sans espace) permet de changer les indicateurs. __Il est obligatoire d'indiquer les 2 indicateurs.__ Cette commande est cumulable avec l'option `UX` ;
* ajouter `$3` devant le PPN permet de rajouter ce PPN en tant que subdivision __au dernier PPN entré durant cette activation du script__ ;
* écrire `ok` (valeur par défaut de la boite de dialogue) permet de sortir de la boucle et de terminer le script.

Une fois la donnée saisie, le script supprime `PPN` suivi d'un espace, `(PPN)`, les espaces, les retours à la ligne et les retours chariot (`chr(10)` et `chr(13)`, ce qui permet notamment d'éviter des problèmes si le PPN est copié depuis une cellule Excel) et place le curseur à la fin de la notice. Dans la suite de l'explication, la donnée saisie par l'utilisateur correspondra au résultat de cette opération de suppression.

Concrètement, si le troisième caractère en partant de la fin est un `_`, les indicateurs prennent la valeur des deux derniers caractères renseignés.
Ensuite, si les deux premiers caractères sont `$3`, le script va réécrire le champ UNIMARC stocké en mémoire (cf ci-après) en insérant avant le neuvième dernier caractère (= avant `$2rameau` et un retour à la ligne (`chr(10)`)) la donnée saisie par l'utilisateur. En clair, il rajoute supposément `$3123456789` avant le `$2`.
En revanche, si les deux premiers caractères ne sont pas `$3`, le script insère à l'emplacement du curseur (= fin de la notice) le champ qu'il a en mémoire (donc rien pour la première occurrence) puis va isoler comme `PPN` les 9 premiers caractères en commençant à partir du troisième caractère (supposément le PPN dans la forme `UX123456789`).
Il regarde ensuite si les deux premiers caractères de la donnée saisie équivalent à un des `UX` précédemment cités. Si oui, il détermine la valeur du `X` de la `60X`associée. Si non, il attribue `6` au `X` et isole alors comme `PPN` les neufs premiers caractère de la donnée saisie. Ainsi, saisir `U9123456789` écrire une `606` avec comme PPN `U91234567`.
Une fois le traitement des commandes terminé, il conserve alors en mémoire un champ :
* `60` + la valeur du `X` + un espace + la valeur des indicateurs + `$3` + le PPN qu'il a isolé + `$2rameau` + un retour à la ligne (`chr(10)`)

Lorsque la donnée saisie est égale à `ok`, il insère le champ en mémoire avant d'achever le script.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `addUA400`

Rajoute des UA400 pour les noms composés en se basant sur la UA200, sinon rajoute une UA400 copiant la UA200. _Ce script n'est pas universel et ne fonctionne qu'en présence d'un `$a` et d'un `$b`._

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà, puis lance le script [`findUA200aUA200b`](https://github.com/Alban-Peyrat/Scripts-WinIBW#findua200aua200b), pour récupérer la 200, la 200 `$a`, la 200 `$b` et la position du premier dollar (ou de la fin du champ) après le `$b`.
Il lance ensuite le script [`decompUA200enUA400`](https://github.com/Alban-Peyrat/Scripts-WinIBW#decompua200enua400) en injectant le `$a` et le `$b` précédemment obtenu pour récupérer les 400 des noms composés.
Il vérifie ensuite si la longueur du champ renvoyé par `decompUA200enUA400` est inférieure à 5 (= si aucune 400 n'a été générée) :
* si c'est le cas, il va copier la 200 précédemment obtenue en supprimant tout ce qui se trouve après la position du premier dollar après le `$b`, puis remplace dans ce qu'il reste `200` par `400` et supprime `$90y`.
Il insère ensuite le nouveau champ à la fin de la notice et place le curseur après le huitième caractère de celui-ci (en théorie, au début du contenu du premier dollar).
* si ce n'est pas le cas, il insère le champ renvoyé à la fin de la notice.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `addUB700S3`

Remplace la UB700 actuelle de la notice bibliographique par une UB700 contenant le PPN du presse-papier et le $4 de l'ancienne UB700.

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà, puis recherche à l'intérieur de celle-ci un retour à la ligne (`chr(10)`) suivi de `700` (supposément, la première 700).
Le script sélectionne ensuite les trois derniers caractères de ce champ (supposément le code fonction) puis génère :
* `700 #1$3` + le contenu du presse-papier + `$4` + la sélection en cours.

Il supprime de ce champ généré les retours à la ligne (`chr(10)`), puis supprime le champ où se trouve le curseur (ancienne 700) et insère à sa place la nouvelle 700 et un retour à la ligne.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `changeExAnom`

Remplace le `$btm` de la zone eXX associée au RCR par `$bx` ou signale la présence de plusieurs eXX associées à ce RCR ou non. __Le mode d'affichage de la notice doit (probablement) être `UNM` pour fonctionner correctement.__ _Ce script vise un objectif assez précis, voir le contexte de développement à la fin de sa documentation._

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà, puis copie l'intégralité de la notice.
Le script exécute ensuite [`CountOccurrences`](https://github.com/Alban-Peyrat/Scripts-WinIBW#countoccurences) pour compter le nombre de `chr(10)` suivi de `e` en tenant compte de la casse (= compte le nombre de notices d'exemplaires dans l'ILN) :
* si une occurrence est détectée, exécute [`goToTag`](https://github.com/Alban-Peyrat/Scripts-WinIBW#gototag) pour se rendre sur le champ 930, puis recule de 1 caractère (bascule sur le champ précédent) et sélectionne les deux prochains caractères sur la gauche (= les deux derniers caractères du champ).
Il compare ensuite si ces deux caractères en minuscule sont égaux à `tm`, auquel cas, ils les remplacent par `x`, récupère le numéro de champ et affiche une infobulle (numéro de champ + `: tm remplacé par x`) ;
* si plus d'une occurrence est détectée, il réexécute `countOccurrences` en comptant cette fois-ci le nombre  `$b` suivi du RCR __(pour utiliser le script sur votre RCR, changez `330632101` en votre RCR)__ :
  * si plus d'une occurrence est trouvée, recherche `$btm` suivi d'un `chr(10)` suivi de `930 ` et récupère le numéro du champ. Si ce numéro commence par `e`, affiche une infobulle (numéro du champ + `à supprimer`, avec comme titre de fenêtre `Exemplaire fictif`), sinon affiche une autre infobulle (`Plusieurs exemplaires réels sur ce RCR. Vérification recommandée.`) ;
  * sinon, affiche une infobulle (`Plusieurs exemplaires réels. Vérification recommandée.`).

_Contexte de développement : dans le cadre d'un chantier sur les thèses, des exemplaires pouvaient avoir en `$b` des `eXX` la mention `TM` (ou `M` supposément, dans la pratique je n'en ai pas vus / je les ai ratés) liée à l'ancien signalement dans téléthèses.
Ainsi, certains exemplaires téléthèses ont été réutilisés sans changer la valeur du `$b`, d'autres sont seulement des exemplaires fictifs en complément de l'exemplaire réel.
Par ailleurs, nous sommes généralement la seule biblitohèque de l'ILN possédant les thèses de ce chantier, ce qui explique les demandes de vérification du script si plusieurs exemplaires sont détectés dans l'ILN._

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `ChantierTheseAddUB183`

Ajoute une UB183 en fonction de la UB215 (notamment des chiffres détectés dans le $a).

_Type de procédure : SUB_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `chantierTheseLoopAddUB183`

Exécute `ChantierTheseAddUB183`, sauf si l'utilisateur refuse l'ajout, sur la liste de PPN présente dans le presse-papier et exporte un rapport des modifications ou non effectuées.

_Type de procédure : SUB_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `decompUA200enUA400`

Renvoie des UA400 créés à partir de la décomposition du nom composé de l'UA200, à l'aide des `$a` et `$b` injectés.

_Type de procédure : FUNCTION_

_Paramètres :_
* `UA200a` : contenu de la 200 `$a` ;
* `UA200b` : contenu de la 200 `$b`.

Le script est une grande boucle `While` qui boucle tant que `UA200a` contient un espace ou un tiret.
À chaque isntance, il détecte quel est le séparateur (en comparant quelle est la plus petite position, 0 exclu, entre l'espace et le tiret).
Il construit ensuite la nouvelle forme, en ajoutant à la fin de `UA200b` (avec un espace si le dernier caractère n'est pas `'` ou `-`) le début de `UA200a` jusqu'au séparateur, supprimant ensuite cette partie (séparateur compris) dans `UA200a`.
Le script analyse ensuite si les caractères au début du nouveau `UA200a` sont les particules rejetées françaises (`de` suivi d'un espace ou `d'`), si c'est le cas, il les retire de `UA200a` et les rajoute à la fin de `UA200b` (sans espace si nécessaire).
Il rajoute ensuite le champ ci-dessous à la valeur qui sera renvoyée (via [`appendNote`](https://github.com/Alban-Peyrat/Scripts-WinIBW#appendnote)) avant de passer à la prochaine instance :
* `400 #1$a` + la valeur actuelle de `UA200a` + `$b` + la valeur actuelle de `UA200b`

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `findUA200aUA200b`

Renvoie la position la UA200, son `$a`, son `$b` et la position du premier dollar suivant le `$b` ou à défaut celle de la fin du champ. __Doit être appelé depuis l'écran de modification pour fonctionner.__

_Type de procédure : SUB_

Récupère le premier champ 200 de la notice puis initie une boucle `While` tant que `UA200fPos` est égal à zéro (sa valeur par défaut), tout en générant un compteur supplémentaire.
À chaque instance de la boucle, en fonction de la valeur du compteur (augmente de 1 à la fin de chaque instance), la script va attribuer à `UA200fPos` la position d'un dollar (0 par défaut, ce qui veut dire que si le dollar n'est pas présent, la boucle continue) :
* compteur = 0 : `$f` ;
* compteur = 1 : `$c` ;
* compteur = 2 : `$x` ;
* compteur = 3 : `$y` ;
* compteur = 4 : `$z` ;
* si le compteur a une autre valeur, assigne à `UA200fPos` la longueur de la 200 __+ 1__ (sinon [`addUA400`](https://github.com/Alban-Peyrat/Scripts-WinIBW#addua400) supprimerait parfois la dernière lettre du prénom).

Il isole ensuite la valeur du `$a` puis du `$b` et renvoie la 200, la `$a` isolé, le `$b` isolé et `UA200fPos` __sous forme d'une seule chaîne de caractères en séparant les différentes valeurs par `;_;`__.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs)

#### `generalLauncher`

Ouvre une boîte de dialogue servant à lancer les scripts (majoritairement de type `add` et `get`). 

_Type de procédure : SUB_


Ouvre une boîte de dialogue contenant la liste des scripts suivants accompagnés de leur identifiant, la liste étant décomposée en plusieurs parties :
* notices bibliographiques :
  * 14 : exécuter [`add18XmonoImp`](https://github.com/Alban-Peyrat/Scripts-WinIBW#add18xmonoimp) ;
  * 1 : exécuter [`addCouvPorte`](https://github.com/Alban-Peyrat/Scripts-WinIBW#addcouvporte) ;
  * 2 : exécuter [`addBibgFinChap`](https://github.com/Alban-Peyrat/Scripts-WinIBW#addbibgfinchap) ;
  * 3 : exécuter [`addEISBN`](https://github.com/Alban-Peyrat/Scripts-WinIBW#addeisbn) ;
  * 4 : exécuter [`AddSujetRAMEAU`](https://github.com/Alban-Peyrat/Scripts-WinIBW#addsujetrameau) ;
  * 15 : placer dans le presse-papier le renvoi de [`addUB700S3`](https://github.com/Alban-Peyrat/Scripts-WinIBW#addub700s3) ;
* Elsevier :
  * 6 : exécuter [`addISBNElsevier`](https://github.com/Alban-Peyrat/Scripts-WinIBW#addisbnelsevier) ;
  * 7 : exécuter [`add214Elsevier`](https://github.com/Alban-Peyrat/Scripts-WinIBW#add214elsevier) ;
* récupérer des informations :
  * 8 : placer dans le presse-papier le renvoi de [`getTitle`](https://github.com/Alban-Peyrat/Scripts-WinIBW#gettitle) ;
  * 9 : placer dans le presse-papier le renvoi de [`getCoteEx`](https://github.com/Alban-Peyrat/Scripts-WinIBW#getcoteex) ;
* thèses
  * 10 : exécuter [`getDataUAChantierThese`](https://github.com/Alban-Peyrat/Scripts-WinIBW#getdatauachantierthese) ;
  * 5 : exécuter `perso_CTaddUB700S3` ;
  * 11 : placer dans le presse-papier le renvoi de [`getUB310`](https://github.com/Alban-Peyrat/Scripts-WinIBW#getub310) ;
* notices d'autorités
  * 12 : exécuter [`addUA400`](https://github.com/Alban-Peyrat/Scripts-WinIBW#addua400) ;
  * 13 : placer dans le presse-papier le renvoi de [`getUA810b`](https://github.com/Alban-Peyrat/Scripts-WinIBW#getua810b) ;
* CorWin :
  * 77 : lance le lanceur de [CorWin](https://github.com/Alban-Peyrat/CorWin).


_Contexte de développement : j'utilise des raccourcis pour la majorité de mes scripts. Or à force de créer de petits scripts, les combinaisons de raccourcis se limitent et m'obligent à retenir beaucoup de raccourcis différents. Le lenceur général permet donc de réduire ce nombre. Aussi, les nombres sont attribués dans l'ordre d'ajout et non pas dans l'ordre où ils sont listés._

#### `getCoteEx`

Renvoie dans le presse-papier la cote du document. Si plusieurs cotes sont présentes, donne le choix entre en sélectionner une, ou toutes les sélectionner, permettant également de choisir le séparateur.

_Type de procédure : FUNCTION_

_Renvoi :_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `getDataUAChantierThese`

Copie dans le presse-papier le PPN, l'année de soutenance, la discipline, le patronyme, le prénom, l'année de naissance, le sexe, le titre et la cote du document, séparés par des tabulations horizontales. Une option permet de réécrire ou d'éditer les champs directement depuis WinIBW.

_Type de procédure : SUB_

_Renvoi :_

Créé dans le cadre d'un chantier sur les thèses, l'exploitation de ces données se fait dans un tableur Excel particulier.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `getTitle`

Renvoie dans le presse-papier le titre du document en remplaçant les @ et $e. Si le titre est entièrement en majuscule, le renvoie en minuscule (sauf première lettre).

_Type de procédure : FUNCTION_

_Renvoi :_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `getUA810b`

Si un seul UA810 est présent, écrit le $b "né le" à partir des informations de la UA103 de la notice, sinon, renvoie le $b dans le presse-papier.

Pour un bon fonctionnement, la UA103 doit comprendre AAAAMMJJ.

_Type de procédure : FUNCTION_

_Renvoi :_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `getUB310`

Copie dans le presse-papier la valeur du premier UB310.

_Type de procédure : FUNCTION_

_Renvoi :_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `PurifUB200a`

Renvoie l'adaptation d'un titre en son écriture en UNIMARC.

_Type de procédure : FUNCTION_

_Renvoi :_

_Paramètres :_
* UB200 : PAS A JOUR
* isUB541 : PAS A JOUR

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_principaux.vbs)

#### `searchExcelPPNList`

Recherche la liste de PPN contenue dans le presse-papier.

_Type de procédure : SUB_

Transforme la liste de PPN du presse-papier en :
* supprimant `(PPN)` ;
* remplançant `chr(10)` par `OR` (avec espace avant et après) ;
* ajoutant au début `che PPN` suivi d'un espace ;
* supprimant les quatre derniers caractères (supposément `OR` avec un espace avant et après).

Place ensuite la requête dans le presse-papier et lance la requête.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs)

### Scripts ressources

Ce fichier contient les scripts facilitant l'exécution des autres, qui sont amenés à être appelés dans de nombreux autres scripts.

#### `appendNote`

Renvoie la variable injectée avec le texte injecté, ajoutant un saut de ligne si la variable n'était pas vide.

_Type de procédure : FUNCTION_

_Paramètres :_
* `var` : variable à laquelle on veut ajouter du texte ;
* `text` : texte à ajouter à la variable.

Regarde si `var` est vide :
* si oui, renvoie le `text` ;
* si non, renvoie `var` + `chr(10)` + `text`.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs)

#### `CountOccurrences`

Renvoi le nombre d'occurrences.

_Type de procédure : FUNCTION_

_Paramètres :_
* `p_strStringToCheck` : variable qui sera fouillée ;
* `p_strSubString` : texte à chercher ;
* `p_boolCaseSensitive` : __bool__ définit si la recherche sera sensible à la casse.

Renvoie le nombre de fois où `p_strSubString` apparait dans `p_strStringToCheck` en comptant le nombre de parties lorsque l'on divise `p_strStringToCheck` en utilisant `p_strSubString` comme séparateur.
Si `p_boolCaseSensitive` est `false`, alors le script passe dans un premier temps les deux autres variables en minuscule.

[Consulter la source originale](https://www.thoughtasylum.com/2009/07/30/VB-Script-Count-occurrences-in-a-text-string/), [consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs)

#### `exportVar`

Exporte le texte injecté dans `export.txt` (même emplacement que `winibw.vbs`). __Pour l'utiliser, pensez à changer la destination du document, et le nom si vous le souhaitez.__

_Type de procédure : SUB_

_Paramètres :_
* `var` : le texte à exporter ;
* `boolAppend` : __bool__ définit si le script doit ajouter à la fin du fichier (`true`) ou réécrire le fichier.

[Consulter la source originale](http://eddiejackson.net/wp/?p=8619), [consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs)

#### `goToTag`

Attention, `subTag` ne doit pas contenir le $ ET est sensible à la casse.

Place le curseur à l'emplacement indiqué par les paramètres. Si plusieurs occurrences sont rencontrées sans que `toFirst` ou `toLast` soit true, une boîte de dialogue s'ouvre pour sélectionner l'occurrence souhaitée.

_Type de procédure : SUB_

_Paramètres :_
* tag : [string] A FAIRE
* subTag : [string, "none" pour empty] A FAIRE
* toEndOfField : [bool] A FAIRE
* toFirst : [bool] A FAIRE
* toLast : [bool] A FAIRE

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs)

#### `goToTagInputBox`

Permet d'essayer `goToTag` en indiquant les paramètres voulus.

_Type de procédure : SUB_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs)

#### `Sleep`

Permet de mettre en pause un script. __Évitez l'utilisation.__

_Type de procédure : SUB_

_Paramètres :_
* `time` : __int__ temps à attendre (en secondes).

[Consulter la source originale](https://stackoverflow.com/questions/1729075/how-to-set-delay-in-vbscript#answer-12921137), [consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs)

#### `toEditMode`

Passe en mode édition (ou présentation).

_Type de procédure : SUB_

_Paramètres :_
* `lgpMode` : __bool__ définit si l'on souhaite passer en mode présentation (`true`) ;
* `save` : __bool__ définit si l'on souhaite sauvegarder les modifications si l'on passe en mode présentation.

Script barbare qui pour le moment essaye de savoir s'il est possible de coller une information dans la notice :
* si l'opération entraîne une erreur (non visible par l'utilisateur), détermine que la notice n'est pas en mode édition ;
* sinon, détermine que la notice est en mode édition.

Il agit ensuite selon trois scénarios :
* il doit passer en mode édition et la notice n'est pas en mode édition, il lance la commande `mod` ;
* il doit passer en mode présentation et sauvegarder la notice, il simule la validation de la notice ;
* il doit passer en mode présentation et ne pas sauvegarder la notice, il simule alors une annulation puis une validation (= pour le message qui apparait en cas de tentative d'annulation alors que des modifications ont été effectuées).

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs)

#### `uCaseNames`

Renvoie les noms injectés avec une majuscule au début de chacun d'entre eux.

_Type de procédure : FUNCTION_

_Paramètres :_
* `noms` : les noms à formatter.

Passe en majuscule le premier caractère de `noms` et en minuscule le reste, avant de lancer une boucle à trois instances.
Pour chaque instance, le script détermine un séparateur qu'il va rechercher dans `noms` (espace puis `-` puis `'`).
Il initie alors une variable `jj` avec la valeur `0` puis lance alors une boucle `While` tant que la recherche du séparateur à l'intérieur de `noms` en commençant à la position `jj + 1` est concluante.
Si la recherche est concluante, `jj` prend la valeur de la position du séparateur identifié, puis le script conserve tel quel tout ce qui se trouve jusqu'à `jj`, puis passe en majuscule le caractère se trouvant en `jj + 1` et conserve tout ce qui le suit tel quel.
Une fois la boucle `While` interrompue, il passe ensuite à la prochaine instance de la première boucle.
Une fois les trois instances de celle-ci terminée, il remplace `De` (espace avant et après) et `D'` (espace avant) par leur équivalent en minuscule, avant de renvoyer le résultat final.


[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts_ressources.vbs)

### Concepts de scripts

Ce fichier contient des concepts que je n'utilise pas mais qui théoriquement fonctionnent, ou des scripts de mon bac à sable que je pense utiles à partager. Certains d'entre eux ont des équivalents dans mes outils, auquel cas, un lien vers ceux-ci sera présent.

#### `ctrlUA103eqUA200f`

___Voir [ConStance CS1](https://github.com/Alban-Peyrat/ConStance#cs1--%C3%A9quivalence-champs-103--200f-idref) pour un outil équivalent.___

Exporte et compare le $a de UA103 et le $f de UA200 pour chaque PPN de la liste présente dans le presse-papier.

_Type de procédure : SUB_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/concepts.vbs)

#### `ctrlUB700S3`

___Voir [ConStance CS2](https://github.com/Alban-Peyrat/ConStance#cs2--pr%C3%A9sence-dun-lien-en-700) pour un outil équivalent et [ConStance CS3](https://github.com/Alban-Peyrat/ConStance#cs3--pr%C3%A9sence-dun-lien-en-7xx) pour un outil équivalent utilisable sur toutes les 700.___

Exporte le premier $ de UB700 pour chaque PPN de la liste présente dans le presse-papier.

_Type de procédure : SUB_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/concepts.vbs)

## Liste des modifications

Le 02/08/2021 :
* suppression de `PurifUB200a` car peu d'intérêts à être partagé ;
* suppression de `CollerPPN` car peu d'intérêts à être partagé ;
* suppression de `LastCHE` car peu d'intérêts à être partagé.

Le 23/08/2021 :
* ajout de `AddSujetRAMEAU` pour ajouter des 60X ;
* ajout de `ctrlTraitementInterne` ;
* ajout de `getUB310` pour récupérer dans le presse-papier l'information de la première 310 ;
* ajout de `PurifUB200a` pour adapter un titre à son écriture en UNIMARC ;
* scission de `addUB700S3` : la partie sur l'exemplaire a été isolée dans un nouveau script, `changeExAnom`.

Le 24/08/2021 :
* répartition des scripts entre plusieurs fichiers ;
* actualisation des présentations des scripts, notamment en intégrant les dernières modifications ;
* adaptation du projet pour être cohérent avec les autres outils.

Le 25/08/2021 :
* suppression de `ctrlTraitementInterne`, que j'avais dû arrêter en plein milieu du développement ;
* modification de la description de [concepts](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/concepts.vbs) et ajout de la mention de [Constance](https://github.com/Alban-Peyrat/ConStance) ;

Le 01/09/2021 :
* ajout de `appendNote` pour ajouter à une variable la donnée voulue ;
* ajout de `getDataUAChantierThese` pour exporter les données d'une thèse dans le cadre d'un chantier sur les thèses ;
* ajout de `uCaseNames` pour mettre des majuscules aux noms renseignés ;
* modification de `getCoteEx` dû à une réécriture du script. Détecte désormais l'intégralité des cotes associées au RCR et permet de sélectionner celles voulues, ou toutes ;
* probable mise à jour prochaine de `decompUA200enUA400` pour être plus efficace et utiliser `uCaseNames` ;

Le 02/09/2021 :
* modification de `getDataUAChantierThese` pour réorganiser l'inputBox, rajouter de la précision à la note sur les noms d'épouse et empêcher des valeurs illégales pour le genre ;
* modification de `getCoteEx` pour afficher le numéro de l'occurrence et de l'exemplaire en cas de cotes multiples, ainsi que de gérer la sélection individuelle de plusieurs cotes.

Le 13/10/2021 :
* sur le document et le dépôt : la liste des modifications est renvoyée au fonds avec un lien hypertexte vers celle-ci en début de page. Par ailleurs, la documentation complète viendra finalement s'ajouter dans ce document ;
* modification de `getTitle` : changement de la détection du champ 200 et de son successeur avec un chr(13) pour éviter le problème des nombres dans le titre (ex: "201 patients") ;
* modification de `getDataUAChantierThese` : mise en place d'une vraie solution pour la détection de genre (ajout d'un `_` devant les données entrées, ce qui empêche la détection du split ; ne prend désormais que le second caractère de l'occurrence au lieu de l'itégralité du texte de celle-ci). Ajout de la  possibilité de modifier les champs Nom, Prénom, et Titre dans une nouvelle boite de dialogue préremplie avec les données détectées. Quelques précisions sur le choix de restreindre cette option uniquement à ces trois champs :
  * discpline requiert une correspondance exacte dans Excel, autant y utiliser la liste de données validées ;
  * années de soutenance et de naissance : se modifient en 4 touches, il est peu nécessaire d'afficher l'information entrée quand réécrire l'année est directement dans Excel est généralement plus rapide ;
  * nom, prénom, titre : certains cas requièrent une modification d'un carcatère dans des textes parfois longs, pouvoir modifier le nom sans devoir le retaper peut être plus agréable ;
  * cote : copier-coller la cote via `getCoteEx` est plus précis que de la retaper ;
  * majsucule : ce n'est pas une information dans l'_output_ final ;
  * notes : force la présence de toutes les notes ajoutées à l'output final ;
  * Sexe : a déjà sa boîte de dialogue ;
* refonte de `addUA400` et `decompUA200enUA400` et ajout de `findUA200aUA200b` :
  * `decompUA200enUA400` prend désormais comme paramètre la $a et le $b originaux au lieu de prendre l'intégralité de la 200, la détection de ces deux sous-champs se fait donc désormais dans `findUA200aUA200b`. (L'externalisation de cette partie du code est liée à des scripts non partagés que j'utilise) ;
  * `addUA400` ne cessera plus de fonctionner s'il n'y a ni $f, ni $c dans le cas de nom non-composé ;
  * `addUA400` prend désormais en compte la présence d'un $x, $y ou $z avant de considérer le champ comme achevé ;
  * `decompUA200enUA400` est maintenant bien plus lisible ;
  * attention toutefois, ce couple de scripts requièrent toujours la présence d'un $a et d'un $b pour pouvoir fonctionner [(voir les modifications prévues)](https://github.com/Alban-Peyrat/Scripts-WinIBW#modifications-prevues).
* refonte des scripts de type `get` et ajout du lanceur général (`generalLauncher`) : création d'une interface pour pouvoir lancer les scripts de type `add` et `get`. L'implentation de ce lanceur a été l'occasion de modifier tous les scripts de type `get` pour qu'ils puissent être utilisables dans d'autres scripts sans devoir stocker le résultat dans le presse-papier. De fait, il n'est plus possible de leur attribuer un raccourci sans créer spécialement un nouveau script qui se contente d'appeler la fonction et de placer le résultat dans le presse-papier. La création de ce lanceur est lié à la multiplication de courts scripts que j'utilise et une multiplication trop importantes des raccourcis associés.
* ajout de `add18XmonoImp` : ajoute une 181 P01 txt, 182 c et 183 nga ;
* ajout de `addISBNElsevier` : ajoute une 010 avec le début d'un IBSN type d'Elsevier ;
* ajout de `add214Elsevier` : ajoute une 214 type pour Elsevier-Masson SAS à Issy-les-Moulineaux avec un DL 2021 ;
* ajout de `addBibgFinChap` : ajoute à l'emplacement du curseur la mention de bibliographie en fin de chapitre ;
* ajout de `addCouvPorte` : ajoute une 312 pour indiquer ce que la couverture porte en plus.

Le 15/10/2021 :
* renommage des scripts ressources et concepts (parce que j'ai découvert que WinIBW peut séparer les scripts en plusieurs catégories) ;
* modification mineuresur le fonctionnement de la boucle de `AddSujetRAMEAU` ;
* ajout de la documentation complète de :
  * `add18XmonoImp` ;
  * `add214Elsevier` ;
  * `addBibgFinChap` ;
  * `addCouvPorte` ;
  * `addISBNElsevier` ;
  * `AddSujetRAMEAU`.

### Modifications prévues

* `getTitle` : permettre son utilisation autant en mode édition que présentation ;
* scripts de type `get` : vérification de l'utilisation du presse-papier et restituer le presse-papier présent avant le lancement du script s'il est réécrit ;
* `addUA400` (et associés) : gérer les autres informations ;
* `decompUA200enUA400` : gérer les cas où l'indicateurs 2 est `0` ainsi que l'absence de `$b` ;
* ajout de `addEISBN` : ajoute une 452 avec un _place holder_ ou le titre s'il est déjà renseigné, ainsi que les trois premières parties de l'ISBN. Apparement j'ai cessé le développement en plein milieu ;
* correction du malfonctionnement probable de `addBibgFinChap` ;
* nettoyage et correction de code et des commentaires de début de script.
