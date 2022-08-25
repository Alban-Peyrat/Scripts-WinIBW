# Scripts pour WinIBW

## Installer les scripts

### En Visual Basic Script (VBS)

_Dans WinIBW, vous retrouverez les scripts dans les `fonctions`._

Procédure d'installation :
* [Téléchargez ce dépôt](https://github.com/Alban-Peyrat/WinIBW/archive/refs/heads/main.zip).
* Au sein de celui-ci, vous trouverez dans le dossier `scripts` un sous-dossier appelé `vbs`.
C'est au sein du sous-dossier `vbs` que se trouvent les fichiers contenant les scripts.
Vous pouvez placer ces scripts où vous le souhaitez (dans votre profil WinIBW semble être une bonne idée.
Par exemple, les miens se trouvent sous `C:\oclcpica\WinIBW30\Profiles\apeyrat001\alp_scripts`).
__Ne placez pas ces scripts dans le même dossier que `winibw.vbs`, sinon WinIBW pourrait charger à l'infini au démarrage.__
__Je vous invite à les placer dans un sous-dossier.__
* Dans WinIBW, ouvrez le menu `Script` puis `Éditer`.
* Sélectionnez `(General)` et `(Declarations)`.
* Puis collez l'intégralité du code ci-dessous.
* __Modifiez ensuite le chemin d'accès à votre dossier dans `sluitMapIn` (première ligne du code que vous avez collé) pour qu'il pointe vers votre dossier.__
* Fermez l'éditeur de script.
* Redémarrez WinIBW pour bien sauvegarder les changemnts et les appliquer.

``` VBScript
sluitMapIn("C:\oclcpica\WinIBW30\Profiles\apeyrat001\alp_scripts\vbs")

' Import all vbs files from a directory'
'From https://cbs-nl.oclc.org/htdocs/winibw/scripts/WinIBW3.installatie.scriptbeheer.html'
Private Sub sluitMapIn(map)
   Set oFSO = CreateObject("Scripting.FileSystemObject")
   If map = "" Then
      msgbox "Aucun nom de dossier renseigné."
   elseif oFSO.FolderExists(map) Then
      Set folder = oFSO.GetFolder(map)
      Set bestanden  = folder.Files
      For each bestand In bestanden
         if lcase(mid(bestand.Name,len(bestand.Name)-3))=".vbs" then
            sluitVBSin(map & "\" & bestand.Name)
         end if
      Next
   else
      msgbox "Le dossier """ + map + """ n'existe pas ?"
   End If
End Sub
```

### En Javascript (JS)

_Dans WinIBW, vous retrouverez les scripts dans les `fonctions standarts`._

Procédure d'installation :
* [Téléchargez ce dépôt](https://github.com/Alban-Peyrat/WinIBW/archive/refs/heads/main.zip).
* Au sein de celui-ci, vous trouverez dans le dossier `scripts` un sous-dossier appelé `js` ainsi qu'un fichier `alp_central_scripts.js`.
C'est au sein du sous-dossier `js` que se trouvent les fichiers contenant les scripts.
Vous pouvez placer ces scripts où vous le souhaitez, mais __`alp_central_scripts.js` est configuré pour charger des fichiers se trouvant dans votre profil utilisateur WinIBW__ (par exemple, les miens se trouvent sous `C:\oclcpica\WinIBW30\Profiles\apeyrat001\alp_scripts`).
* Au sein de `alp_central_scripts.js`, vous devrez éditer la liste des fichiers que vous voulez charger.
Pour ce faire, ouvrez le fichier dans un éditeur de texte et rendez-vous à la ligne 63 du fichier (qui commence par `const alpScripts =`).
Remplacer les noms semi-complets (chemin d'accès au fichier en utilisant comme base votre profil utilisateur WinIBW + nom du fichier + extension du fichier) des fichiers par défaut par ceux que vous voulez charger.
Veillez à respecter la mise en forme déjà présente : à la fin, votre liste doit ressembler à :

``` Javascript
const alpScripts = ["fichier1",
"fichier2",
"fichier3"];
```

* Une fois les modifications effectuées, sauvegardez le fichier et fermez-le.
* Ouvrez WinIBW (ou rallumez-le s'il était ouvert).
* Ouvrez le menu `Script` puis `Éditer`.
* Sélectionnez `(General)` et `(Declarations)`.
* Puis collez cette ligne de code tout en haut de la fenêtre :

``` VBScript
application.writeProfileString "ibw.standardScripts","script.AlP","resource:/Profiles/apeyrat001/alp_scripts/alp_central_scripts.js"
```

* __Modifiez ensuite le nom complet de votre fichier `alp_central_scripts.js` dans le code que vous avez collé pour qu'il pointe vers votre fichier.__
* Fermez l'éditeur de scripts.
* Redémarrez WinIBW pour bien sauvegarder les changements et les appliquer.

## Informations pour scripter dans WinIBW

### Exécuter des scripts utilisateurs depuis les scripts standarts

Pour ce faire, vous devrez utiliser :
* le script `__executeVBScript()` dans `peyrat_ressource.js`,
* le script `__executeUserScript()` dans `peyrat_ressource.js`,
* le script `executeVBScriptFromName` dans `alp_ressources.vbs` et l'avoir sur le raccourci Maj + Ctrl + Alt + L (ou changer dans `__executeUserScript()`la combinaison de touches.

Une fois que tout ceci est prêt, il vous suffit d'appeler `__executeUserScript()` avec comme seul paramètre le nom de la fonction que vous voulez appeler.

### Prise en compte des accents

Pour que les accents soient pris en compte dans vos fichiers, vous devez les encoder en `Western (Windows 1252)`, pas en `UTF-8` (ou autres).

### Validation des notices par script

Il est tout à fait possible de valider via un script les modifications apportées à une notice à l'aide de la commande `Application.ActiveWindow.SimulateIBWKey` avec comme paramètre `"FR"`, mes scripts ne le font pas par choix :

``` Javascript
// Validation des notices en Javascript
application.activeWindow.simulateIBWKey("FR");
```

```VBScript
' Validation des notices en VBScript
Application.ActiveWindow.SimulateIBWKey "FR"
```

### Vérification du type de notice

Mes scripts ne vérifient pas (pour le moment en tout cas) s'ils sont exécutés sur le bon type de notice, en revanche, cette vérification est tout à fait possible à l'aide du script [`getNoticeType()` en VBS](#getnoticetype) ou de [`__getNoticeType()` en JS](#__getnoticetype).

## Présentation des scripts

### Scripts utilisateurs (VBS)

#### Fichier `winibw.vbs`

Contient uniquement les scripts utilisés pour paramétrer WinIBW autant pour son interface que pour récupérer des variables communes à VBS et JS que pour charger les autres scripts en VBS que pour permettre au fichier central de paramétrage de JS d'être chargé.
_[Consulter le fichier](/scripts/winibw.vbs)_

##### Lignes de code hors des fonctions

* `application.writeProfileString "ibw.standardScripts","script.AlP","resource:/Profiles/apeyrat001/alp_scripts/alp_central_scripts.js"` : permet de charger le script central de JS qui permettra de charger par la suite les autres scripts JS.
Changez `resource:/Profiles/apeyrat001/alp_scripts/alp_central_scripts.js` par le chemin d'accès à votre script central de JS.
* `sluitMapIn("C:\oclcpica\WinIBW30\Profiles\apeyrat001\alp_scripts\vbs")` : permet de charger les autres scripts VBS.
Changez `C:\oclcpica\WinIBW30\Profiles\apeyrat001\alp_scripts\vbs` par le chemin d'accès à votre dossier contenant les scripts VBS.
Vous pouvez charger plusieurs dossiers, ou charger un fichier individuellement à l'aide de [la fonction `sluitVBSin()`](#sluitVBSin).
* `Set WSHShell = CreateObject("WScript.Shell")` : permet de créer un objet `WScript.Shell` qui vous permettra de récupérer les informations d'une variable environnementale, à l'aide de `WSHShell.ExpandEnvironmentStrings("%MY_RCR%")`, en remplaçant `MY_RCR` part le nom de la variable.
* Notamment, sont récupérés les noms des chemins spéciaux de WinIBW (`ProfD` par exemple) à savoir :
  * `WINIBW_dwlfile` : le nom complet du fichier de téléchargement ;
  * `WINIBW_prnfile` : le nom complet du fichier d'impression ;
  * `WINIBW_BinDir` : le nom complet du dossier principal de WinIBW ;
  * `WINIBW_ProfD` : le nom complet du dossier de l'utilisateur.

##### `sluitMapIn()`

_Provient de [Installatie van WinIBW3 (3.7) ter ondersteuning van script-beheer - VBScript centraal beheerd](https://cbs-nl.oclc.org/htdocs/winibw/scripts/WinIBW3.installatie.scriptbeheer.html)._

_Paramètre :_
* `map` : chemin d'accès d'un dossier

Permet de charger tous les scripts en `.vbs` du dossier `map` pour pouvoir les exécuter dans WinIBW. _Requiert [la fonction `sluitVBSin()`](#sluitVBSin)_

##### `sluitVBSin()`

_Provient de [Installatie van WinIBW3 (3.7) ter ondersteuning van script-beheer - VBScript centraal beheerd](https://cbs-nl.oclc.org/htdocs/winibw/scripts/WinIBW3.installatie.scriptbeheer.html)._

_Paramètre :_
* `VBSbestand` : chemin d'accès d'un fichier

Permet de charger le fichier `VBSbestand` de scripts en `.vbs` pour pouvoir exécuter les exécuter dans WinIBW.


------------------------------------------------------------


#### Fichier `alp_cat_add.vbs`

Contient tous les scripts permettant de rajouter des informations à une notice d'autorité ou bibliographique.
_[Consulter le fichier](/scripts/vbs/alp_cat_add.vbs)_

##### `add18XmonoImp()`

Passe la notice en mode édition si elle ne l'est pas déjà puis insère les 181-2-3 pour une monographie imprimée sans illustration à la fin de celle-ci suivi d'un retour à la ligne :

``` MARC
181 ##$P01$ctxt
182 ##$P01$cn
183 ##$P01$anga
```

##### `add18XmonoImpIll()`

Passe la notice en mode édition si elle ne l'est pas déjà puis insère les 181-2-3 pour une monographie imprimée avec illustration à la fin de celle-ci suivi d'un retour à la ligne :

``` MARC
181 ##$P01$ctxt
181 ##$P01$csti
182 ##$P01$P02$cn
183 ##$P01$P02$anga
```

##### `add214Elsevier()`

Passe la notice en mode édition si elle ne l'est pas déjà puis insère une 214 type pour Elsevier (2022) à la fin de celle-ci suivi d'un retour à la ligne :

``` MARC
214 #0$aIssy-les-Moulineaux$cElsevier Masson SAS$dDL 2022
```

##### `addAutFromUB()`

Permet de créer un squelette de notice d'autorité à partir d'une notice bibliographique (pour préremplir la 810).

Ouvre une boîte de dialogue demandant d'entrer le patronyme puis une seconde demandant les éléments rejetés.
Le script récupère ensuite le titre du document via la fonction [`getTitle`](#gettitle) ainsi que l'année en utilisant la `100 $c` ou `100 $a`.
Crée ensuite une notice d'autorité sous la forme suivante :

``` MARC
008 $aTp5
106 ##$a0$b1$c0
101 ##$afre
102 ##$aFR
103 ##$a19XX
120 ##$a -----À-COMPLÉTER-MANUELLEMENT-----
200 #1$90y$a{nom}$b{prenom}$f19..-....
340 ##$a -----COMPLÉTER-AVEC-D-AUTRES-INFORMATIONS-----
810 ##$a{titre} / {prenom} {nom}, {annee}
```

Enfin, génère des 400 en cas de nom composé (si le nom comprend un espace ou un `-`) via la fonction [`addUA400`](#addua400).

##### `addBibgFinChap()`

Passe la notice en mode édition si elle ne l'est pas déjà puis insère à l'emplacement du curseur :

``` MARC
Chaque fin de chapitre comprend une bibliographie
```

__Malfonctionnement possible : si la notice n'était pas en mode édition, le texte ne s'écrira probablement pas si la grille des données codées n'est pas affichée.__

##### `addCouvPorte()`

Passe la notice en mode édition si elle ne l'est pas déjà puis insère une 312 indiquant que la couverture porte des informations supplémentaires à la fin de celle-ci, en plaçant le curseur à la fin du champ généré :

``` MARC
312 ##$aLa couverture porte en plus : "
```

##### `addEISBN()`

Permet d'ajouter une 452 avec le titre et les 2 (ISBN 10) ou 3 (ISBN 13) premiers éléments de du premier ISBN renseigné.

Passe la notice en mode édition si elle ne l'est pas déjà puis détermine la position du `@` au sein de la `200 $a` et récupère le titre du document via la fonction [`getTitle`](#gettitle).
Si aucun titre n'est renvoyé, le titre sera égal à `@ -----À-COMPLÉTER-MANUELLEMENT-----`.
Le script récupère ensuite le `$a` ou `$A` de la première `010`, puis supprime les deux derniers éléments de celui-ci en utlisant les `-` comme séparateurs d'éléments.
S'il n'y a aucun `-` dans l'ISBN ou que la récupération de celui-ci a échoué, l'ISBN inséré sera vide.
Enfin, insère à la fin de la notice le champ suivant, en plaçant un `@` à l'emplacement détecté plus tôt et en plaçant le curseur à la fin de ce nouveau champ :

``` MARC
452 ##$t{titre}$y{ISBN modifié}
```

##### `addISBNElsevier()`

Passe la notice en mode édition si elle ne l'est pas déjà puis insère une 010 avec les trois premiers éléments d'un ISBN 13 d'Elsevier à la fin de celle-ci :

``` MARC
010 ##$A978-2-294-
```

##### `addNoteBonISBN()`

Passe la notice en mode édition si elle ne l'est pas déjà puis insère une 310 indiquant que l'ISBN 13 exact provient du service [Nouveautés éditeurs de la Bibliothèque nationale de France](https://nouveautes-editeurs.bnf.fr/) à la fin de celle-ci :

``` MARC
301 ##$aL'ISBN 13 exact provient du service Nouveautés éditeurs de la Bibliothèque nationale de France
```

##### `AddSujetRAMEAU()`

Ouvre une boîte de dialogue permettant d'insérer des UB60X à partir du PPN.

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

##### `addUA400()`

Rajoute des UA400 pour les noms composés en se basant sur la UA200, sinon rajoute une UA400 copiant la UA200. _Ce script n'est pas universel et ne fonctionne qu'en présence d'un `$a` et d'un `$b`._

Passe la notice en mode édition si elle ne l'est pas déjà, puis lance le script [`findUA200aUA200b`](#findua200aua200b), pour récupérer la 200, la 200 `$a`, la 200 `$b` et la position du premier dollar (ou de la fin du champ) après le `$b`.
Il lance ensuite le script [`decompUA200enUA400`](#decompua200enua400) en injectant le `$a` et le `$b` précédemment obtenu pour récupérer les 400 des noms composés.
Il vérifie ensuite si la longueur du champ renvoyé par `decompUA200enUA400` est inférieure à 5 (= si aucune 400 n'a été générée) :
* si c'est le cas, il va copier la 200 précédemment obtenue en supprimant tout ce qui se trouve après la position du premier dollar après le `$b`, puis remplace dans ce qu'il reste `200` par `400` et supprime `$90y`.
Il insère ensuite le nouveau champ à la fin de la notice et place le curseur après le huitième caractère de celui-ci (en théorie, au début du contenu du premier dollar).
* si ce n'est pas le cas, il insère le champ renvoyé à la fin de la notice.

##### `addUB700S3()`

Remplace la UB700 actuelle de la notice bibliographique par une UB700 contenant le PPN du presse-papier et le $4 de l'ancienne UB700.

Passe la notice en mode édition si elle ne l'est pas déjà, puis recherche à l'intérieur de celle-ci un retour à la ligne (`chr(10)`) suivi de `700` (supposément, la première 700).
Le script sélectionne ensuite les trois derniers caractères de ce champ (supposément le code fonction) puis génère :

``` MARC
700 #1$3{contenu du presse papier}$4{sélection en cours}
```

Il supprime de ce champ généré les retours à la ligne (`chr(10)`), puis supprime le champ où se trouve le curseur (ancienne 700) et insère à sa place la nouvelle 700 et un retour à la ligne.

##### `addUB7XX()`

Insère une 7XX avec le PPN présent dans le presse-papier en indiquant le code fonction voulu.
Par défaut, le script est configuré pour insérer un nom de personne, il est toutefois possible d'insérer un nom de collectivité ou un nom de famille, tout comme il est possible de modifier les indicateurs déterminés par défaut.

Valeurs par défaut des indicateurs :
* Nom de personne : `#1` ;
* Collectivité : `02` ;
* Nom de famille : `##`.

Au lancement du script, une boîte de dialogue s'ouvre demandant d'insérer le code de fonction.
C'est cette même boîte de dialogue qui permet de configurer le type de champ que l'on insère ainsi que les indicateurs voulus :
* __Si l'on souhaite insérer autre chose qu'un nom de personne__, le premier caractère de la réponse doit être :
  * `c` pour une collectivité ;
  * `f` pour un nom de famille ;
* __Si l'on souhaite insérer des indicateurs différents de ceux par défaut,__ il faut écrire __les deux indicateurs entre deux espaces avant le code de fonction__ (les deux espaces sont nécessaires même si l'on souhaite renseigné un nom de personne).

__Le code de fonction doit toujours correspondre aux trois derniers caractères de la réponse.__

Une fois la réponse validée, le script récupère comme PPN le contenu du presse-papier puis analyse la réponse donnée.
Dans un premier temps, il détermine la dizaine du numéro de champ à insérer en se basant sur le premier de caractère de la réponse où `c` sera 1 (collectivité), `f` sera 2 (nom de famille) et tout autre caractère sera 0 (nom de personne, la valeur par défaut).
Dans un deuxième temps, il détermine les indicateurs en divisant la réponse en utilisant les espaces comme séparateur : s'il y a 2 espaces au total, il détermine que les indicateurs sont __exactement__ ceux contenus entre les deux espaces, sinon, il prend la valeur par défaut des indicateurs en fonction de la dizaine du numéro de champ (0 = `#1`, 1 = `02`, 2 = `##`).
Dans un troisième temps, il prend les trois derniers caractères de la réponse comme étant le code de fonction (aucune vérification).
Dans un quatrième temps, il détermine l'unité du numéro de champ à partir d'une liste __non exhaustive__, donnant les résultats suivants :
* Unité de champ 0 :
  * `070` (auteur) ;
  * `340` (éditeur scientifique) ;
  * `651` (directeur de publication) ;
  * `730` (traducteur) ;
* Unité de champ 1 :
  * `555` (membre du jury) ;
  * `727` (directeur de thèse) ;
  * `956` (président du jury de soutenance) ;
  * `958` (rapporteur de la thèse) ;
* Unité de champ 2 :
  * `080` (préfacier) ;
  * `440` (illustrateur) ;
* Pour tous les autres cas, l'unité de champ sera remplacée par ` -----COMPLÉTER-MANUELLEMENT-----`.

Dans un cinquième temps, si l'unité de de champ est 0, le script vérifie s'il existe déjà une 7X0, si oui, il transforme l'unité de champ en 1.
Enfin, le script passe la notice en mode édition si elle ne l'est pas déjà puis insère à la fin de celle-ci suivi d'un retour à la ligne :

``` MARC
7{dizaine de champ}{unité de champ} {indicateurs}$3{PPN}$4{code fonction}
```

##### `createItemAvaibleForILL()`

Crée un exemplaire avec la cote voulue et indiquant la disponibilité de celui-ci pour le prêt entre bibliothèque.

Si la notice est une notice bibliographique (déterminé en utilisant la fonction [getNoticeType](#getnoticetype)), ouvre une bopite de dialogue demandant d'écrire la cote du document voulue, sinon, affiche une erreur.
Une fois la cote validée, désactive les données codées puis crée un nouvel exemplaire via la commande `\INV E*` (qui permet de créer un nouvel exemplaire en affichant tous ceux présents dans l'ILN et sans devoir indiquer le numéro de l'exemplaire à créer).
Insère ensuite à la fin de l'écran d'édition :

``` MARC
e* $bx
930 ##$a{cote}$ju
```


------------------------------------------------------------


#### Fichier `alp_cat_get.vbs`

Contient tous les scripts permettant de récupérer des informations depuis WinIBW.
_[Consulter le fichier](/scripts/vbs/alp_cat_get.vbs)_


##### `getCoteEx()`

_Je l'ai codé il y a un certain temps mais je n'ai jamais eu de problèmes avec (très certainement grâce à la méthode de fonctionnement de l'établissement).
Par contre il mériterait vraiment d'être réécrit..._

Renvoie la cote associé à l'exemplaire du RCR (ou l'un ou plusieurs des exemplaires, ou tous les exemplaires).
Le RCR dans le script déposé correspond à une variable globale basée sur la variable environnementale que j'ai définie ([voir les lignes de codes hors des fonctions dans winibw.vbs](#lignes-de-code-hors-des-fonctions)).

Récupère la notice bibliographique à l'aide de la fonction `application.activeWindow.copyTitle` puis la divise en utilisant comme séparateur `$b{RCR}`.
Ensuite, pour chaque partie exceptée la première, isole le numéro d'exemplaire en recherchant dans la partie précédente la dernière occurrence d'un retour à la ligne suivi de `e`, en conservant 3 caractères (ex : `e02`).
Le script détecte dans un second temps s'il y a un `$a` entre le début de la partie et la première occurrence de `A98 ` : si c'est le cas, il conserve comme cote tout ce qui se trouve entre le `$a` et le premier prochain `$` si et seulement si ce prochain `$` se situe avant un retour à la ligne, sinon, il conserve comme cote tout ce qui se trouve entre le `$a` et le premier retour à la ligne.
_(En relisant ce code, il est évident que cette détection peut renvoyer de fausses informations...)_
Si en revanche aucun `$a` n'est détecté entre le début de la partie et le `A98 `, la cote prend la valeur `[Exemplaire sans cote]`.
Le script ajoute enfin à une variable le texte suivant : `[Occ. {numéro de la cote pour ce RCR}] {numéro d'exemplaire ILN} : {cote}`

Une fois chaque exemplaire partie de la notice traitée, le script regarde le nombre de cotes détectées pour le RCR : __s'il y a en a une seule, renvoie la cote__, sinon, il ouvre une boîte de dialogue présentant toutes les cotes trouvées (sous la forme du texte à la fin du paragraphe ci-dessus).
Il est alors possible de sélectionner un, plusieurs ou tous les exemplaires, ainsi que de renseigner un séparateur si l'on souhaite récupérer plusieurs cotes (le séparateur par défaut est un retour à la ligne).
Ainsi l'on répond :
* le numéro de __l'occurrence__ que l'on souhaite (si l'on en veut une seule) ;
* les numéros des occurrences que l'on souhaite séparés par des `_` si l'on en souhaite plusieurs ;
* `all` si l'on souhaite toutes les cotes ;
* et l'on renseigne si l'on souhaite un séparateur autre que celui par défaut en ajoutant à notre réponse :
  * `$$t` pour une tabulation horizontale ;
  * `$$;` pour un point-virgule ;
  * `$$#{un séparateur personnalisé}` pour un séparateur personnalisé (par exemple, `$$##` pour utiliser `#` comme séparateur, `$$#!` pour `!` comme séparateur)

Le script renverra ensuite la/les cotes demandées avec le séparateur indiqué.
_Je ne rentrerai pas plus dans les détails pour cette partie, le code est disponible dans le ficher si jamais il y a un problème, mais cette partie mériterait également d'être retravaillée._

##### `getTitle()`

_Mériterait un affinage voire une réécriture._

Renvoie le titre __supposément__ au format ISBD, mais à utiliser de préférences sur des titres assez simples.

Récupère l'intégralité de la première 200 : s'il n'y en a aucune, renvoie `Aucune 200`.
Sinon, conserve tout ce qui se trouve entre le premier `$a` et le premier `$f`, ou entre le premier `$a` et la fin du champ si aucun `$f` n'est présent.
Supprime ensuite __toutes__ les `@` et remplace tous les `$e` par des ` : ` (entre espaces).
Enfin, regarde si le titre est uniquement en majsucule : si c'est le cas, le passe en minuscule en conservant uniquement la première lettre en majuscule, puis renvoie ce titre modifié.
_À ce stade, vous pouvez, je pense, comprendre le commentaire introduisant ce script._

##### `getUA810b()`

_Comme tous dans ce fichier, mériterait d'être réécrit, notamment parce qu'il devrait se contenter de récupérer une information sans en écrire et qu'il a beaucoup de problèmes possibles._

Génère une `$b` donnant des informations sur la date de naissance à destination d'une 810, puis :
* l'écrit à la fin de la 810 s'il n'y en a qu'une seule ;
* renvoie la `$b` générée si plusieurs 810 se trouve dans la notice.

Passe la notice en mode édition si elle ne l'est pas déjà, puis copie l'intégralité de la notice.
Commence ensuite la construction du `$b` en isolant les 8 derniers caractères de la première 103 (supposément AAAAMMJJ), puis en déterminant le genre à l'aide du dernier caractère de la première 120, avant de générer le `b` suivant :

``` MARC
$bné{e si féminin} le JJ-MM-AAAA
```

À l'aide de la copie de notice précédemment effectué, compte le nombre d'occurrences de `810 ##` puis :
* si une seule occurrence est trouvée, se rend à la fin de se champ et insère le `$b` généré ;
* dans tous les autres cas, renvoie le `$b` généré.

##### `getUB310()`

_Mérite une refonte, particulièrement parce qu'il était majoritairement utilisé pour récupérer les informations sur les droits d'accès aux thèses, qui n'ont plus leur place en 310._

Renvoie le contenu de la `310 $a`.

Passe la notice en mode édition si elle ne l'est pas déjà, puis copie la notice à l'aide de la fonction `application.activeWindow.copyTitle`.
Renvoie ensuite tout ce qui se trouve entre le premier `310 ##$a` et le premier retour à la ligne suivant ce `310 ##$a`.


------------------------------------------------------------


#### Fichier `alp_chantier_theses.vbs`

Contient tous les scripts que j'ai spécialement développé dans le cadre de chantiers sur les thèses à l'Université de Bordeaux.
_[Consulter le fichier](/scripts/vbs/alp_chantier_theses.vbs)_


------------------------------------------------------------


#### Fichier `alp_corwin.vbs`

Contient tous les scripts permettant le fonctionnement du [projet CorWin permettant de contrôler des données dans WinIBW](../../../CorWin).

Ce projet n'est supposé être utilisé que dans les très rares cas où je n'ai pas réussi à récupérer les données que je souhaite via les API ou autres.
Et il est à l'abandon depuis octobre 2021.

Donc outre le fait que je ne l'utilise pas, que je ne suis pas vraiment sûr que ce soit une bonne idée de l'utiliser et que le code est absolument terrible, si jamais cette analyse est nécessaire il faudra probablement revoir le code.

##### `CorWin_CW1()`

Pour chaque PPN présent dans le presse-papier __provenant de CorWin__, exporte les `103 $a` et `103 $b` en utilisant [`Ress_exportVar()`](#ress_exportvar).
Les données exportées pour chaque PPN sont, séparées par des `;_;` :
* le PPN ;
* la `103 $a` ;
* la `103 $b`.

_Il est à noter que certaines parties de cette explication sont des suppositions parce que évidemment mettre des commentaires dans mon code je me suis probablement dit que cela ne servait à rien :)))._

Le script récupère le contenu du presse-papier et le divise en utilisant le retour à la ligne comme séparateur, et initie la variable `storedPPN` comme étant égal à `X`.
Il initie ensuite un fichier à l’emplacement prévu par CorWin, emplacement qui se trouve être le premier PPN de la liste, qu'il remplace ensuite par `XXXXXXXXX`.
Le script commence alors une boucle pour chaque PPN dans la liste.
Si celui-ci n'est ni vide, ni égal à `XXXXXXXXX`, lance la commande `che ppn {PPN}`.
Il compare ensuite la valeur de la variable `P3GPP` à celle de `storedPPN`, pour vérifier que la notice est différente de celle traitée à l'itération précédente, dans le cas où la recherche du PPN ait échoué : si les deux PPN sont identiques, termine l'itération en renvoyant comme PPN `Erreur : {le PPN qui a été utilisé dans la commande che PPN}` et `00000000` comme valeur pour les `103 $a` et `103 $b`.

Si la vérification est réussie, copie l'intégralité de la notice via la fonction `application.activeWindow.copyTitle` puis isole la 103.
Le script isole ensuite le contenu du `$a` et du `$b` s'ils existent, sinon il leur attribue la valeur `00000000`.
_(Sachant que je ne crois pas voir de vérification de la présence d'une 103 mais ignorons les conséquences qu'aurait cette absence.)_
`storedPPN` prend alors la valeur du PPN grâce à une magnifique manœuvre récupérant les 9 caractères précédent un retour à la ligne suivi de `008` suivi d'un espace, puis génère les données à exporter en utilisant `storedPPN`.
Enfin, écrit sur le fichier la réponse générée avant de passer à l'itération suivante.

Une fois tous les PPN traités, ferme le fichier et affiche une infobulle demandant de lancer l'analyse de ce fichier sur CorWin.

##### `CorWin_Launcher()`

Ouvre une boîte de dialogue servant à lancer les scripts de CorWin.
En l'occurrence, un seul traitement existe, il faut donc appeler le traitement `CW1`.


------------------------------------------------------------


#### Fichier `alp_dumas.vbs`

Contient le script utilisé pour l'ancienne version du générateur de notice UNIMARC à partir d'un dépôt [DUMAS](https://dumas.ccsd.cnrs.fr/).
[Le dépôt ub-svs contient plus d'informations à ce sujet](../../../ub-svs).
_[Consulter le fichier](/scripts/vbs/alp_dumas.vbs)_

Sachant que ce script doit être très très très largement repris de [celui de l'Abes pour le script utilisateur IdRef](https://github.com/abes-esr/winibw-scripts/blob/3f374e37151ab686fd1423cc21195b997d7df4b9/user-scripts/idref/IdRef.vbs).

##### `these_catDumas()`

Ce script se connecte à l'URL de mon site qui génère la notice UNIMARC du dépôt DUMAS dont l'URL est contenue dans le presse-papier, puis récupère cette notice pour la coller dans WinIBW.

Le début du script isole le `docid` du document pour pouvoir exécuter la fonction principale de la page de mon générateur avec, une fois que WinIBW a réussi à se connecter au site.
Une fois la fonction lancée, WinIBW "sommeille" pendant 1 seconde (via [`ress_sleep()`](#ress_sleep) pour laisser à mon site le temps d'interroger DUMAS et de créer la notice, puis récupère l'élément contenant la notice.
Il exécute enfin la commande `cre` puis insère le texte contenu dans l'élément récupérer précédemment.


------------------------------------------------------------


#### Fichier `alp_interface.vbs`

Contient tous les scripts développés uniquement pour servir d'interface ou pour générer certaines recherches à partir des informations disponibles dans l'interface ou pour générer des requêtes répétitives.
_[Consulter le fichier](/scripts/vbs/alp_interface.vbs)_

##### `generalLauncher()`

Ouvre une boîte de dialogue servant à lancer les scripts.

Chaque script se voit attribuer un numéro d'index qui sert à le faire s'exécuter.
Pour les scripts de type _Function_, le résultat renvoyé dans le presse-papier.
La mise à jour de la liste autant des actions à effectuer que du texte affiché dans l'interface doit se faire manuellement.
_Il est probablement possible de rendre le script un peu plus lisible._

Je ne vais pas établir la liste des scripts possibles, vous pouvez vous référer à la boîte de dialogue qui s'ouvre lors de son exécution ou au code.

##### `goToWorkRecord()`

Ouvre la notice d'autorité de l’œuvre associée.

Récupère la notice via la variable `P3CLIP` puis la divise en utilisant un retour à la ligne comme séparateur.
Pour chaque ligne, regarde si les trois premiers caractères sont `579` : lorsqu'une ligne correspond, active la commande `che ppn {les 9 caractères suivants le $3 sur cette ligne}` puis arrête l'exécution du script.
Si aucune ligne ne correspond, affiche un message d'erreur.

##### `searchDoublonPossible()`

S'utilise lorsque WinIBW affiche le message `Doublon possible` après la création d'une notice : cherche le PPN indiqué comme doublon potentiel par WinIBW.

Récupère le premier message affiché dans la zone des messages de WinIBW.
Si celui-ci contient `PPN `, active la commande `che ppn {les 9 caractères suivants le "PPN " dans le message}`.
Si le message ne contient pas `PPN ` (ou si aucun message n'est affiché), affiche une erreur.

_Il est possible que ce script puisse être utilisé pour rechercher le PPN indiqué par WinIBW dans d'autres messages que celui des doublons possibles si leur forme correspond. _

##### `searchExcelPPNList()`

Recherche la liste de PPN contenue dans le presse-papier.

Transforme la liste de PPN du presse-papier en :
* supprimant `(PPN)` ;
* remplaçant `chr(10)` (retour à la ligne) par ` OR ` (avec espace avant et après) ;
* ajoutant au début `che PPN` suivi d'un espace ;
* supprimant les quatre derniers caractères (supposément `OR` avec un espace avant et après).

Place ensuite la requête dans le presse-papier et lance la requête.


------------------------------------------------------------


#### Fichier `alp_PEB.vbs`

Contient tous les scripts développés à destination du module PEB de WinIBW.
_[Consulter le fichier](/scripts/vbs/alp_PEB.vbs)_

_[Voir le document dédié](./PEB.md)_


------------------------------------------------------------


#### Fichier `alp_ressources.vbs`

Contient tous les scripts ressources que j'utilise au sein des autres scripts.
_[Consulter le fichier](/scripts/vbs/alp_ressources.vbs)_

Je recommande fortement de l'installer pour pouvoir utiliser la plupart des scripts en VBS.
Par ailleurs, certains scripts conservent une ancienne notation en commençant par `Ress_` que j'utilisais lorsque tous mes scripts se trouvaient dans `winibw.vbs`.
Comme les scripts ressources sont utilisés dans beaucoup d'autres scripts, je n'ai pas supprimé le préfixe.
Toutefois, ils seront ici classés en ignorant le préfixe.

##### `Ress_appendNote()`

_Dans l'idéal je devrais la réécrire pour qu'elle soit une copie exacte de [`__addTextToVar` en JS](#__addtexttovar), mais comme indiqué au-dessus, il faudrait probablement que je modifie beaucoup de scripts._

Renvoie la variable injectée avec le texte injecté, ajoutant un saut de ligne si la variable n'était pas vide.

_Paramètres :_
* `var` : variable à laquelle on veut ajouter du texte ;
* `text` : texte à ajouter à la variable.

Regarde si `var` est vide :
* si oui, renvoie le `text` ;
* si non, renvoie `var` + `chr(10)` + `text`.

##### `Ress_CountOccurrences()`

_[Créé par Stephen Millard, publié le 30 jullet 2009 sur ThoughtAsylum.](https://www.thoughtasylum.com/2009/07/30/VB-Script-Count-occurrences-in-a-text-string/)_

Renvoi le nombre d'occurrences.

_Paramètres :_
* `p_strStringToCheck` : variable qui sera fouillée ;
* `p_strSubString` : texte à chercher ;
* `p_boolCaseSensitive` : __bool__ définit si la recherche sera sensible à la casse.

Renvoie le nombre de fois où `p_strSubString` apparait dans `p_strStringToCheck` en comptant le nombre de parties lorsque l'on divise `p_strStringToCheck` en utilisant `p_strSubString` comme séparateur.
Si `p_boolCaseSensitive` est `false`, alors le script passe dans un premier temps les deux autres variables en minuscule.

##### `decompUA200enUA400()`

Renvoie des 400 créés en décomposant les noms composés pour les autorités nom de personne à partir du `$a` et du `$b` renseignés (en suivant _supposément_ [la forme des points d'accès pour les noms de personne françaises disponible sur le site de l'IFLA](https://www.ifla.org/g/cataloguing/names-of-persons/)).
Par exemple, `$a = Poisson-Truite $b = Jacques` et `$a = La Vallière de La Truite Dorée $b = Louise Françoise` créeront respectivement :

``` MARC
400 #1$aTruite$bJacques Poisson-

400 #1$aVallière de La Truite Dorée$bLouise Françoise La
400 #1$aLa Truite Dorée$bLouise Françoise La Vallière de
400 #1$aTruite Dorée$bLouise Françoise La Vallière de La
400 #1$aDorée$bLouise Françoise La Vallière de La Truite
```

_Paramètres :_
* `UA200a` : `200 $a` de l'autorité nom de personne voulue ;
* `UA200b` : `200 $b` de l'autorité nom de personne voulue.

Ce script fonctionne en boucle tant qu'un espace ou un `-` est détecté dans le `$a`.
S'il n'y en a pas, ne renvoie rien.
Détermine dans un premier temps l'emplacement du séparateur détecté, puis rajoute à la fin du `$b` un espace (sauf si le dernier caractère du `$b` est un `-` ou une `'`) suivi de tout ce qui se trouve entre le début du `$a` et le séparateur (inclus), retirant le séparateur si c'est un espace.
Supprime ensuite du `$a` ce qui a été transféré dans le `$b`.

Le script performe ensuite une vérification afin de rejeter les prépositions `de` et `d'`.
Ainsi, si les trois premiers caractères du nouveau `$a` sont `de ` (suivi d'un espace) ou si les deux premiers sont `d'`, les prépositions sont retirés de `$a` et ajoutées à la fin du `$b` selon le même principe qu'au-dessus.

Avant de terminer cette itération de la boucle, rajoute à la variable qui sera renvoyée, via l'utilisation de [`Ress_appendNote()`](#ress_appendnote), le champ suivant :

``` MARC
400 #1$a{le $a transformé}$b{le $b transformé}
```

##### `delEspaceB4Tag()`

Supprime les espaces se trouvant avant un numéro de champ.

Tant qu'il existe des champs commençant par un espace à l'aide de la fonction `application.activeWindow.title.findTag`, se rend au début du premier de ces champs puis sélectionne ensuite tous les espaces précédent le premier mot à l'aide de la fonction `application.activeWindow.title.wordRight` et supprime la sélection.

##### `executeVBScriptFromName()`

_[Élaboré grâce à la page dédié à la fonction `Execute` du ss64.com par Simon Sheppard.](https://ss64.com/vb/execute.html)_


__En état expérimental, utilisation fortement non recommandée.
Son but est d'être utilisé depuis un script standart, pas depuis un script utilisateur.__

##### `Ress_exportVar()`

_[Original créé par MrNetTek, publié le 19 novembre 2015 sur Lab Core | The Lab of MrNetTek.](http://eddiejackson.net/wp/?p=8619)_

_Il faudrait le revoir._

Exporte le texte injecté dans `export.txt` qui se trouve dans le profil WinIBW de l'utilisateur (même emplacement que `winibw.vbs`).

_Paramètres :_
* `var` : le texte à exporter ;
* `boolAppend` : __bool__ définit si le script doit ajouter à la fin du fichier (`true`) ou réécrire le fichier.

##### `findUA200aUA200b()`

_Mériterait d'être revue voire supprimer [par le nouveau `getTag`](#ress_gettag)._

Renvoie, depuis l'écran de modification d'une notice d'autorité (ne vérifie pas le type de notice), les informations suivantes, séparées par des `;_;` :
* la 200 complète ;
* la `200 $a` ;
* la `200 $b` ;
* la position du premier caractère suivant la fin du `$b`. 

Récupère la 200 à l'aide de la fonction `application.activeWindow.title.findTag`, puis identifie le premier caractère suivant le `$b` en cherchant s'il existe, dans l'ordre :
* un `$f` ;
* un `$c` ;
* un `$x` ;
* un `$y` ;
* un `$z` ;
* si aucun n'existe, détermine que rien ne suit le `$b`.

Sont ensuite isolés le `$a` comme étant tout ce qui se trouve entre le premier `$a` et le premier `$b`, puis le `$b` comme étant tout ce qui se trouve entre le premier `$b` et le premier caractère suivant le `$b` précédemment identifié, avant de renvoyer les données comme expliqué ci-dessus.

##### `getNoticeType()`

Renvoie l'entier :
* `0` si c'est une notice d'autorité,
* `1` si c'est une notice bibliographique,
* `2` si ce n'est aucune des deux.

Pour déterminer cette information, il se base sur la variable `P3VMC` qui correspond au type de document (`008 position 1 et 2`) et, si la première variable n'a pas de valeur, sur la variable `scr` qui correspond au code de l'écran.
Pour `scr`, son utilisation n'est supposée avoir lieu que si le script est utilisé dans le cadre d'une création de notice _ex-nihilo_ ou sur un écran autre qu'une notice.
Le script vérifie donc uniquement si `scr` est égal à `II` (création de notice d'autorité) ou `IT` (création de notice bibliographique).

##### `Ress_getTag()`

Renvoie un champ ou un sous-champ __depuis l'écran de modification ou l'écran de présentation__.
De mémoire, certaines fonctionnalités avancées ne fonctionnent pas à la perfection.

__Ce script est trop complexe et pas assez efficace pour son ambition.
De fait, je ne l'expliquerai pas car il serait juste plus efficace de le recréer à partir de zéro, tout en le scindant en `getField` et `getSubfield`.
Dans l'idéal, la nouvelle version s'approcherait de [la version JS développée par la GBV (voir `__getFields`)](#__getfields) et serait développée autant en VBS qu'en JS.
Toujours dans l'idéal, elle prendrait comme argument une expression régulière, comme celle de la GBV, et renverrait une liste de résultats (contrairement à la GBV), soit de champs au format chaîne de caractère, soit de dictionnaires.
La fonction pour les sous-champs serait assez similaire, prenant comme argument un champ / dictionnaire correspondant au champ.__

_Paramètres :_
* `tag` : numéro du champ recherché ;
* `forceOcc` : la/les occurrence à retourner :
  * `no` : si plusieurs champs existent, ouvre une boîte de dialogue demandant de choisir l'occurrence voulue ;
  * `last` : renvoie la dernière occurrence ;
  * `all` : renvoie toutes les occurrences, séparées par des `;_;_;` ;
  * un nombre : le numéro de l'occurrence voulue, si le nombre est plus grand que le nombre d'occurrences, renvoie la dernière occurrence ;
* `subtag` : le sous-champ à renvoyer, __doit prendre la valeur `none` si l'on souhaite récupérer le champ complet__ ;
* `forceOccSub` : la/les occurrence du sous-champ à retourner :
  * `no` : si plusieurs champs existent, ouvre une boîte de dialogue demandant de choisir l'occurrence voulue ;
  * `last` : renvoie la dernière occurrence ;
  * `all` : renvoie toutes les occurrences, séparées par des `;_;_;` ;
  * un nombre : le numéro de l'occurrence voulue, si le nombre est plus grand que le nombre d'occurrences, renvoie la dernière occurrence.

##### `Ress_goToTag()`

En mode édition, place le curseur à l'emplacement voulu.

__Ce script est trop complexe et pas assez efficace pour son ambition.
De fait, je ne l'expliquerai pas car il serait juste plus efficace de le recréer à partir de zéro.__

Voici les informations que j'avais pu écrire dessus par le passé :
* Attention, `subTag` ne doit pas contenir le $ ET est sensible à la casse.
* Si plusieurs occurrences sont rencontrées sans que `toFirst` ou `toLast` soit true, une boîte de dialogue s'ouvre pour sélectionner l'occurrence souhaitée.

_Paramètres :_
* `tag` : chaîne de caractères
* `subTag` : chaîne de caractères, `none` si on ne veut pas
* `toEndOfField` : booléen
* `toFirst` : booléen
* `toLast` : booléen

##### `Ress_goToTagInputBox()`

Interface permettant d'essayer [`Ress_goToTag()`](#ress_gototag) en indiquant les paramètres voulus.

##### `PurifUB200a()`

_Mériterait probablement d'être refait. Je pense qu'il a été créé pour faciliter la conversion de métadonnées depuis DUMAS._

Renvoie un titre originellement au format ISBD au format UNIMARC, pour une 200 ou une 541.
Voir les exemples ci-dessous, avec en premier le champ original et en second le champ renvoyé, d'abord pour une 200, puis pour une 541 :

``` MARC
200 #1$aPoisson : vie de truite$fLouise Françoise
200 #1$a@Poisson$evie de truite$fLouise Françoise

541 ##$aLe grand Jacques : sa vie : 1950-1985$zfre
541 ##$aLe @grand Jacques$esa vie$e1950-1985$zfre
```

_Paramètres :_
* `UB200` : la 200 (ou 541) sous forme de chaîne de caractères ;
* `isUB541` : `false` si c'est une 200, `true` si c'est une 541.

_Trop compliqué pour ce que c'est je pense, je vais donc survoler certains points._

Isole le contenu du `$a` en utilisant la position du `$f` si c'est une 200 ou du `$z` si c'est une 541.
Remplace ensuite tous les ` : ` (entre espaces) et `: ` (espace uniquement après) par des `$e`.
Ajoute enfin l'`@` au début du `$a`, sauf si celui-ci commence par les caractères suivants, auquel cas il ajoute l'`@` après (note : pour tous à l'exception de ceux se terminant par `'`, le script vérifie s'il y a un espace après) :
* `De la ` ;
* `De l'` ;
* `Les ` ;
* `Des ` ;
* `Une ` ;
* `The ` ;
* `Le ` ;
* `La ` ;
* `Un ` ;
* `An ` ;
* `De ` ;
* `Du ` ;
* `A ` ;
* `L'` ;
* `D'`.

Renvoie ensuite le champ original en remplaçant le `$a` original par celui qu'il a généré.

##### `Ress_Sleep()`

_[Créé par Paulie D, publié le 16 octobre 2012 sur Stackoverflow en réponse à la question _How to set delay in vbscript_ de Mark, postée le 13 novembre 2009.](https://www.thoughtasylum.com/2009/07/30/VB-Script-Count-occurrences-in-a-text-string/)_

Permet de mettre en pause un script. __Évitez l'utilisation.__

_Paramètre :_
* `time` : temps à attendre (en secondes).

##### `Ress_toEditMode()`

_Mériterait une petite actualisation_

Passe en mode édition (ou présentation).

_Paramètres :_
* `lgpMode` : `true` pour passer en mode présentation, `false` pour passer en mode édition ;
* `save` : en cas de passage en mode présentation, `true` pour sauvegarder les modifications, `false` pour ne pas enregistrer.

Script barbare qui pour le moment essaye de savoir s'il est possible de coller une information dans la notice :
* si l'opération entraîne une erreur (non visible par l'utilisateur), détermine que la notice n'est pas en mode édition ;
* sinon, détermine que la notice est en mode édition.

Il agit ensuite selon trois scénarios :
* il doit passer en mode édition et la notice n'est pas en mode édition, il lance la commande `mod` ;
* il doit passer en mode présentation et sauvegarder la notice, il simule la validation de la notice ;
* il doit passer en mode présentation et ne pas sauvegarder la notice, il simule alors une annulation puis une validation (= pour le message qui apparait en cas de tentative d'annulation alors que des modifications ont été effectuées).

##### `Ress_uCaseNames()`

_Pourrait être revue._

Renvoie les noms injectés avec une majuscule au début de chacun d'entre eux.

_Paramètre :_
* `noms` : les noms à formater.

Passe en majuscule le premier caractère de `noms` et en minuscule le reste, avant de lancer une boucle à trois itérations.
Pour chaque itération, le script détermine un séparateur qu'il va rechercher dans `noms` (espace puis `-` puis `'`).
Il initie alors une variable `jj` avec la valeur `0` puis lance alors une boucle `While` tant que la recherche du séparateur à l'intérieur de `noms` en commençant à la position `jj + 1` est concluante.
Si la recherche est concluante, `jj` prend la valeur de la position du séparateur identifié, puis le script conserve tel quel tout ce qui se trouve jusqu'à `jj`, puis passe en majuscule le caractère se trouvant en `jj + 1` et conserve tout ce qui le suit tel quel.
Une fois la boucle `While` interrompue, il passe ensuite à la prochaine itération de la première boucle.
Une fois les trois itérations de celle-ci terminée, il remplace `De` (espace avant et après) et `D'` (espace avant) par leur équivalent en minuscule, avant de renvoyer le résultat final.


------------------------------------------------------------


#### Fichier `alp_theses2win.vbs`

Contient tous les scripts développés à destination du projet de création de notices UNIMARC dans WinIBW à partir d'une base externe.
_[Consulter le fichier](/main/scripts/js/alp_theses2win.js)_

_[Voir le document dédié](./imp2Win.md)_

_Dans l'idéal il faudrait scinder le fichier en `alp_imp2win.vbs` et `alp_UBtheses2win.vbs` pour bien distinguer les fonctions ressources et les fonctions propres au projet développé à l'Université de Bordeaux, mais je n'avais pas le temps, donc un jour peut-être._

_Comme présenté dans le document dédié, le projet a été théorisé et les tests techniques ont été effectués.
Le code est comme je l'ai laissé par manque de temps, de fait il manque par exemple une fonction qui permet d'analyser l'URL donnée par l'utilisateur pour décider de lancer le script pour DUMAS ou OSKAR Bordeaux, la fonction pour DUMAS ne prend pas de paramètre alors qu'elle devrait prendre une URL comme paramètre, cette même fonction n'a pas le bon procédé technique, etc._

##### Les procédés employés

Deux méthodes différentes sont employées dans ce projet :
* l'XML en utilisant l'objet `XMLDOM` ([ressource utilisée : Manipuler des fichiers XML en VBScript avec XPath par Baptiste Wicht, publié originnellement le 19 décembre 2007](https://baptiste-wicht.developpez.com/tutoriels/microsoft/vbscript/xml/xpath/)) ;
* l'HTML en utilisant l'objet [`Internet Explorer`](https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752084(v=vs.85)) ([en utilisant le code de l'Abes pour le script utilisateur IdRef](https://github.com/abes-esr/winibw-scripts/blob/3f374e37151ab686fd1423cc21195b997d7df4b9/user-scripts/idref/IdRef.vbs))

Puisque nous utilisons des objets `InternetExplorer`, on ne peut pas utiliser de JSON car Internet Explorer ne peut pas afficher du JSON, il lance un téléchargement.
Il devrait être possible d'utiliser de tout de même utiliser des JSON en téléchargeant le document, par exemple avec [`httpdownload.vbs`](#fichier-httpdownload.vbs), mais cette possibilité n'a pas été explorée pour le moment.
Il y a peut-être d'autre solutions pour utiliser du JSON, j'ai pas forcément eu le temps d'explorer tout ce qui était possible.

##### Notes et ressources

Voici quelques notes que je me suis marqué pour interagir avec ces objets :
* utiliser `set` pour les objets ;
* pour l'HTML, utiliser :
  * `getElementsByTagName`,
  * `getElementsByClassName`,
  * `getElementsById`,
  * `getAttribute`,
  * `.innertext` ;
* pour le XML, utiliser :
  * `selectSingleNode`,
  * `selectNodes`,
  * `getAttribute`,
  * `.text`.

Enfin, voici quelques ressources que je me suis également marquées :
* pour les dictionnaires : [Dictionnary sur le site Dot Net Perls, par Sam Allen](https://www.dotnetperls.com/dictionary-vbnet) ;
* pour les éléments : [Element par MDN (developer.mozilla.org)](https://developer.mozilla.org/en-US/docs/Web/API/Element)
* pour les expressions régulières :
  * [la méthode `Replace` de l'objet `RegExp` en VBScript par O'Reilly](https://www.oreilly.com/library/view/vbscript-in-a/1565927206/re155.html) ;
  * [la page consacrée au VBScript de Regular-Expressions.info](https://www.regular-expressions.info/vbscript.html).

##### `dumasXMLDOM()`

Une fonction de test de récupération d'information via l'objet `XMLDOM`.

Dans l'idéal, je pense qu'elle devait devenir l'équivalent de [`getIEObjectDocument()`](#getieobjectdocument) mais pour les objets `XMLDOM`.

Dans son état actuel, affiche tous les titres de l'XML-TEI (`/TEI/text/body/listBibl/biblFull/titleStmt/title`) de la thèse avec l'identifiant HAL `dumas-01911186` dans une infobulle différente pour chaque titre.

##### `thesesDumas2winibw()`

La fonction principale pour la conversion de DUMAS vers WinIBW.

Supposément, prend comme paramètre :
* `url` : l'URL DUMAS du document.

Elle fonctionne encore avec un objet `InternetExplorer.Document`, probablement parce que j'ai découvert comment gérer les objets `XMLDOM` après.

Dans son état actuel, récupère l'objet `InternetExplorer.Document` de la thèse avec l'identifiant HAL `dumas-01911186` puis crée une notice ~~bibliographique~~ d'autorité avec la commande `cre e` (_magnifique erreur_) si l'objet a bien été retourné par `getIEObjectDocument()`, sinon affiche une erreur et arrête l'exécution du programme.

Récupère ensuite le premier tag HTML `licence` via [`getIEDocTag()`](#getiedoctag) et affiche dans une infobulle :
* son `.textContent` ;
* son `.innerText` ;
* le résultat de [`getXmlTextContent()`](#getxmltextcontent).

##### `getIEObjectDocument()`

_[Original créé par l'Abes dans le script utilisateur IdRef](https://github.com/abes-esr/winibw-scripts/blob/3f374e37151ab686fd1423cc21195b997d7df4b9/user-scripts/idref/IdRef.vbs)_

Renvoie [un objet `InternetExplorer.Document`](https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752052(v=vs.85)).

_Ce qui est un problème parce que je sais pas ce que devient le IE qui ne se ferme probablement jamais...
À étudier_

_Paramètre :_
* `url` : l'url de la page.

##### `getIEDocTag()`

Renvoie une liste d'éléments HTML obtenus via [`getElementsByTagName`](https://developer.mozilla.org/en-US/docs/Web/API/Element/getElementsByTagName).

_Paramètres :_
* `IEDoc` : l'objet `InternetExplorer.Document` (voir [`getIEObjectDocument()`](#getieobjectdocument)) ;
* `tag` : la balise HTML voulue.

##### `getElemAttr()`

Renvoie la valeur de l'attribut voulu via [`getAttribute`](https://developer.mozilla.org/en-US/docs/Web/API/Element/getAttribute)

_Paramètres :_
* `elem` : l'élément (voir [`getIEDocTag()`](#getiedoctag)) ;
* `attr` : l'attribut voulu.

##### `getXmlTextContent()`

_[Expression régulière originale créée par jcomeau_ictx, originellement publiée le 19 juillet 2011 sur StackOverflow en réponse à la question How to get the pure text without HTML element using JavaScript? posée par John le 19 juillet 2011.](https://stackoverflow.com/questions/6743912/how-to-get-the-pure-text-without-html-element-using-javascript#6744068')_

Renvoie l'`innerText` d'un tag XML sans son tag, si l'XML a été chargé comme un objet `InternetExplorer.Document` _(je pense)_.

_Paramètre :_
* `elem` : l'élément (voir [`getIEDocTag()`](#getiedoctag)).

Utilise l'expression régulière `<[^>]*>` en mode global pour supprimer les tags HTML obtenu par le `.textContent` d'`elem`. 

##### `oskar2winibw()`

La fonction principale pour la conversion d'OSKAR Bordeaux vers WinIBW.

_Paramètre :_
* `url` : l'URL OSKAR Bordeaux du document.

Dans son état actuel, récupère l'objet `InternetExplorer.Document` d'une thèse (il faut afficher les métadonnées complètes en avance je pense) puis crée une notice ~~bibliographique~~ d'autorité avec la commande `cre e` (_magnifique erreur_) si l'objet a bien été retourné par `getIEObjectDocument()`, sinon affiche une erreur et arrête l'exécution du programme.

Initie ensuite un dictionnaire, puis récupère la table contenant toutes les métadonnées.
Pour chaque ligne dans la table, récupère la première et seconde colonne (respectivement le nom de la métadonnée et la valeur de celle-ci), puis analyse quelle est la métadonnée afin de déterminer l'affichage final : si elle est considéré comme utile, son nom français s'affichera proprement, sinon c'est la flèche suivante qui s'affichera  `---------->` (`dc.contributor.author`, `dc.contributor.advisor`, `dc.date` sont les seules considérées comme utiles).
Enfin, insère dans la notice bibliographique la ligne suivi d'un retour à la ligne  :
`{le nom français / la flèche} {nom de la métadonnée} : {valeur de la métadonnée}`

Dans la partie commentée au bas du code, il y a une utilisation plus pratique du code.
Dans les grandes lignes, basée sur le document à l'URI `https://oskar-bordeaux.fr/handle/20.500.12278/23589`, récupère dans le dictionnaire le prénom et le nom de famille du directeur de thèse et de l'auteur, en passant en minuscule les noms de famille, ainsi que le titre de la thèse et l'année de soutenance.
Pour le titre de la thèse, il est divisé en titre / sous-titre à l'emplacement du `?`.
Génère ensuite une 200 et une 214 avant d'afficher dans une infobulle 6 informations (toutes sauf le titre propre apparemment) :

``` MARC
200 1#$a@{titre}$e{sous-titre}$f{prénom auteur} {nom auteur}$gsous la direction de {prénom directeur de thèse} {nom directeur de thèse}
214 #1$a{année}
```


------------------------------------------------------------


### Scripts standarts (JS)

#### Fichier `alp_central_scripts.js`

Contient des paramètres qui seront définis à l'initialisation de WinIBW (ou lors de l'actualisation des scripts standards).
Ces paramètres sont soit des variables pour les scripts standards uniquement, soit des variables environnementales pour permettre leur utilisation en VBS, soit des fonctions permettant de charger les autres scripts standards.

_[Consulter le fichier](/scripts/alp_central_scripts.js)_

#####  Lignes de code hors des fonctions

* Initialisation des constantes `thePrefs` et `theEnv` pour les scripts standards. J'ai un doute sur `thePrefs`, `theEnv` permet notamment d’interagir avec des variables environnementales.
Les deux ont été récupérés des scripts de la [GBV (plus d'informations à ce sujet dans la partie consacrée à leurs scripts)](#fichier-gbvjs).
* Initialisation des variables environnementales des noms des chemins spéciaux de WinIBW (`ProfD` par exemple) pour les rendre accessibles en VBS, à savoir :
  * `WINIBW_dwlfile` : le nom complet du fichier de téléchargement ;
  * `WINIBW_prnfile` : le nom complet du fichier d'impression ;
  * `WINIBW_BinDir` : le nom complet du dossier principal de WinIBW ;
  * `WINIBW_ProfD` : le nom complet du dossier de l'utilisateur.
* __Initialisation de la variable `alpScripts`.__
C'est dans cette variable qu'il faut indiquer le chemin d'accès à vos scripts standards __se trouvant dans votre profil utilisateur.__
Indiquer le chemin d'accès, en utilisant des `/`, pas des `\`.
Par exemple, mon fichier contenant mes scripts principaux se trouve dans le sous-dossier `alp_scripts` puis `js`, ce qui donne `"alp_scripts/js/peyrat_main.js"`.
Immédiatement après l'initialisation de cette constante, une boucle va ajouter aux préférences utilisateurs de votre profil les scripts indiqués dans `alpScripts`, en leur attribuant le nom `AlP` suivi de leur index (commençant à 0).
__L'ordre dans lequel vous classez les scripts dans `alpScripts` correspond à leur ordre de chargement__ (je crois).
Enfin, si vous décidez de supprimer un script sans en rajouter (=vous chargiez 4 fichiers, vous n'en chargez plus que 3), ce qui réduirait le nombre total de scripts, __pensez à vous rendre dans le `user_pref.js` dans votre profil utilisateur pour supprimer le `ibw.standardScripts.script.AlP` avec l'index le plus grand__ qui ne sera pas automatiquement supprimé.
Ci-dessous, l'exemple de ma variable `alpScripts` :

``` Javascript
const  alpScripts  = ["alp_scripts/js/NE_PAS_DIFFUSER.js",
"alp_scripts/js/peyrat_ressources.js",
"alp_scripts/js/GBV.js",
"alp_scripts/js/peyrat_main.js",
"alp_scripts/js/SCOOP.js",
"alp_scripts/js/peyrat_peb.js",
"alp_scripts/python-winibw/pythWinIBW.js",
"alp_scripts/python/python.js",
"alp_xul/xul_test.js"];
```

 
 ------------------------------------------------------------


#### Fichier `GBV.js`

_[Consulter le fichier](/main/scripts/js/GBV.js)_

__Ce fichier contient des scripts que j'ai récupérés parmi ceux proposés par la [Gemeinsame Bibliotheksverbund](https://www.gbv.de/).__
Je les ai récupérés de la page [d'informations de version de WinIBW](https://wiki.k10plus.de/display/K10PLUS/SWB-WinIBW-Versionsinformationen).
Il n'est pas exclu que certaines fonctions soient similaires à certaines que j'aurais pu développer avant d'analyser ces fichiers.

##### `alert()`

_Provient de `Update_2022_10/scripts/k10_public.js - function alert()`._

_Paramètre :_
* `meldungstext` : une chaîne de caractères

Ouvre une boîte de dialogue avec l'icône d'alerte, le titre `Alerte` et `meldungstext` comme texte.

##### `__warning()`

_Provient de `Update_2022_10/scripts/k10_public.js - function __warnung()`._

_Paramètre :_
* `meldungstext` : une chaîne de caractères

Ouvre une boîte de dialogue avec l'icône d'alerte, le titre `Attention` et `meldungstext` comme texte.

##### `__error()`

_Provient de `Update_2022_10/scripts/k10_public.js - function __fehler()`._

_Paramètre :_
* `meldungstext` : une chaîne de caractères

Ouvre une boîte de dialogue avec l'icône d'erreur, le titre `Erreur` et `meldungstext` comme texte.

##### `__msg()`

_Provient de `Update_2022_10/scripts/k10_public.js - function __meldung()`._

_Paramètre :_
* `meldungstext` : une chaîne de caractères

Ouvre une boîte de dialogue avec l'icône d'information, le titre `Message` et `meldungstext` comme texte.

##### `__question()`

_Provient de `Update_2022_10/scripts/k10_public.js - function __frage()`._

_Paramètre :_
* `meldungstext` : une chaîne de caractères

Ouvre une boîte de dialogue avec l'icône de question, le titre `Question` et `meldungstext` comme texte.

##### `__getMsgs()`

_Provient de `Update_2022_10/scripts/k10_public.js - function __alleMeldungen()`._

Renvoie l'intégralité des messages affichés par WinIBW séparés par des retours à la ligne.
Ne renvoie pas le type du message.

##### `GBVgetMsgsClipboard()`

_Provient de `Update_2022_10/scripts/k10_public.js - function meldungenKopieren()`._

Copie dans le presse papier l'intégralité des messages affichés par WinIBW.
Utilise la fonction [`__getMsgs()`](#__getmsgs) pour la récupération.

##### `__delFields()`

_Provient de `Update_2022_10/scripts/k10_public.js - function felderLoeschen()`._

_Paramètre :_
* `regexpFelder` : une expression régulière correspondant au numéro du champ voulu **(pas une chaîne de caractères)**.
Exemple : `/70[0-9]|606/` si l'on veut supprimer n'importe quel champ en 70X ou le champ 606.

En mode édition, supprime tous les champs dont le numéro de la zone correspond à `regexpFelder`.
Le script passe sur chacune des lignes de la notice et vérifie si le numéro du champ correspond à l'expression régulière : si c'est le cas, supprime l'intégralité du champ.

##### `__delFieldsContent()`

_Provient de `Update_2022_10/scripts/k10_public.js - function feldInhaltLoeschen()`._

_Paramètre :_
* `regexFeld` : une expression régulière correspondant au numéro du champ voulu **(pas une chaîne de caractères)**.
Exemple : `/70[0-9]|606/` si l'on veut supprimer le contenu de n'importe quel champ en 70X ou le champ 606.

En mode édition, supprime le contenu de tous les champs dont le numéro de la zone correspond à `regexFeld`.
Le script passe sur chacune des lignes de la notice et vérifie si le numéro du champ correspond à l'expression régulière : si c'est le cas, supprime l'intégralité du champ à l'exception du premier mot.
__Attention, la détection du premier mot se fait avec la fonction `application.activeWindow.title.wordRight` de WinIBW, ce qui veut dire que si l'une des lignes de la notice ne contient qu'un seul mot et qu'elle correspond à l'expression régulière renseignée, la ligne suivant sera entièrement supprimée sans vérification de correspondance avec `regexFeld`.__
Pour tester les séparateurs de mots, il est possible d'utiliser la commande `Ctrl+{Flèche directionelles latérales}` pour voir le comportement de WinIBW.
Voici quelques informations (les listes sont non exhaustives) :
* caractères non séparateurs (WinIBW les considère comme faisant partie du mot et ne s'arrête ni avant ni après eux) :
  * `.`
  * `-`
* caractères séparateurs compris dans le mot (WinIBW comprend ces caractères dans les mots mais s'arrête avant le prochain caractère ne faisant pas partie de cette liste) :
  * espace
  * retour à la ligne
* caractères séparateurs exclus du mot (WinIBW s'arrête avant ceux-ci) :
  * `_`
  * `/`
  * `\`
  * `:`
  * `;`
  * `!`
  * `?`
  * `(`
  * `[`
  * `{`
  * `#`

##### `__insFieldIfInexistant()`

_Provient de `Update_2022_10/scripts/k10_public.js - function feldEinfuegenNummerisch()`._

_Paramètres :_
* `ergaenzeFeld` : le numéro de champ que l'on souhaite insérer ;
* `strInhalt` : le contenu du champ que l'on souhaite insérer.

En mode édition, ajoute le champ `ergaenzeFeld` avec le contenu `strInhalt` à l'emplacement correspondant s'il n'existe aucun champ avec ce numéro.
Le script vérifie dans un premier temps si un champs existe déjà avec le numéro `ergaenzeFeld`, à l'aide de la fonction `application.activeWindow.title.findTag`.
__Attention, cette fonction de WinIBW renvoie le premier champ qui commence par la chaîne de caractère renseignée,__ ainsi `70` renverra la première 70X, `610 2` renverra la première 610 dont le premier indicateur est 2, même s'il y a des `610 1` avant.
Si un champ est renvoyé, le script s'arrête, sinon parcourt l'intégralité des champs et insère `ergaenzeFeld` et `strInhalt` séparés par un espace à l'emplacement correspondant (ou à la fin de la notice).
Pour déterminer l'emplacement correspondant, compare pour chaque champ le numéro avec `ergaenzeFeld` pour déterminer lequel est le plus grand.

##### `__insField()`

_Provient de `Update_2022_10/scripts/k10_public.js - function feldEinfuegenNummerischOhnePruefung()`._

_Paramètres :_
* `ergaenzeFeld` : le numéro de champ que l'on souhaite insérer ;
* `strInhalt` : le contenu du champ que l'on souhaite insérer.

En mode édition, ajoute le champ `ergaenzeFeld` avec le contenu `strInhalt` à l'emplacement correspondant.
Insère `ergaenzeFeld` et `strInhalt` séparés par un espace à l'emplacement correspondant (ou à la fin de la notice).
Pour déterminer l'emplacement correspondant, compare pour chaque champ le numéro avec `ergaenzeFeld` pour déterminer lequel est le plus grand : s'il existe déjà des champs avec ce numéro, il sera inséré à après ceux existants.

##### `__getFields()`

_Provient de `Update_2022_10/scripts/k10_public.js - function felderSammeln()`._

_Paramètre :_
* `regexpFelder` : une expression régulière correspondant au numéro du champ voulu **(pas une chaîne de caractères)**.
Exemple : `/70[0-9]|606/` si l'on veut supprimer le contenu de n'importe quel champ en 70X ou le champ 606.

En mode édition, renvoie une chaîne de caractères contenant tous les champs dont le numéro de champ correspond à `regexpFelder`, séparés par des retours à la ligne.
Le script passe sur chacune des lignes de la notice et vérifie si le numéro du champ correspond à l'expression régulière : si c'est le cas, ajoute à la variable renvoyée un retour à la ligne suivant du champ complet.
__Cela signifie que le premier caractère sera toujours un retour à la ligne.__

##### `__dateYMD()`

_Provient de `Update_2022_10/scripts/k10_public.js - function __datum()`._

Renvoie la date actuelle au format `AAAA.MM.JJ`.

##### `__dateDMY`

_Provient de `Update_2022_10/scripts/k10_public.js - function __datumTTMMJJJJ()`._

Renvoie la date actuelle au format `JJ.MM.AAAA`.

##### `__dateHours`

_Provient de `Update_2022_10/scripts/k10_public.js - function __datumUhrzeit()`._

Renvoie la date et l'heure actuelles au format `AAAAMMJJHHMMSS`.

##### `__trim()`

_Provient de `Update_2022_10/scripts/k10_public.js - function stringTrim()`._

_Paramètre :_
* `meinString` : une chaîne de caractères.

Renvoie la chaîne de caractère en supprimant les espaces devant et derrière le texte.
Tant que le script détecte une correspondance entre `meinString` et l'expression régulière `/^ | $/` (le premier caractère est un espace ou le dernier caractère est un espace), supprime le caractère qui a provoqué la correspondance.

##### `GBVhackSystemVariables`

_Provient de `Update_2022_10/scripts/k10_public.js - function hackSystemVariables()`._

_De ce que je comprends, le script est d'OCLC.
De plus, elle fait supposément le même travail que la fonction `hackSystemvariables` déjà présente dans la version WinIBW de l'Abes, à l'exception près que celle de l'Abes située dans le fichier `scripts/AFF.JS` est mal codée et renvoie très justement une erreur (sur les ordinateurs que j'ai vu en tout cas)._

Copie dans le presse-papier la plupart des variables (nom et valeurs) de WinIBW.
Le script regarde pour chacune des combinaisons ci-dessous s'il existe une variable non nulle dans WinIBW : si elle existe, ajoute à la variable renvoyée `- {nom de la variable}: {valeur de la variable}` suivi d'un retour à la ligne.

Les combinaisons testées sont : `P3G` / `P3V` / `P3G` suivis de deux variables qui prennent comme valeur soit `!`, soit un chiffre, soit une lettre en majuscule.
L'intégralité des combinaisons possibles sont testées une par une (`P3G!!`, puis `P3G!0`, etc.), ce qui représente 1369 variables pour chaque `P3`, mais moins d'une centaine sont renvoyées.
Toutes les variables ne sont donc pas renvoyées, par exemple `P3CLIP`, qui contient la notice renvoyée par la fonction `Copier notice`, n'est pas renvoyée, mais la majorité des variables prennent la forme testée par cette fonction.


------------------------------------------------------------


#### Fichier `peyrat_main.js`

_[Consulter le fichier](/scripts/js/peyrat_main.js)_

__Ce fichier ne contient pas de scripts publics actuellement.__


------------------------------------------------------------


#### Fichier `peyrat_peb.js`

Contient tous les scripts développés à destination du module PEB de WinIBW.
_[Consulter le fichier](/main/scripts/js/peyrat_peb.js)_

_[Voir le document dédié](./PEB.md)_


------------------------------------------------------------


#### Fichier `peyrat_ressources.js`

Contient tous les scripts ressources que j'utilise au sein des autres scripts.
_[Consulter le fichier](/scripts/js/peyrat_ressources.js)_

##### `__AbesDelTitleCreated()`

_Inspirée des fonctions de `scripts/standart_copy.js` de l'Abes, utilise la fonction `suptag()`._

En mode édition, supprime de la notice tous champs commençant par `Cré`.
Sert à supprimer les informations de création d'une notice copiée via la fonction dédiée dans WinIBW (et qui auraient été collées sans passer par la fonction `Coller notice`).

##### `__AbesDelItemData()`

_Inspirée des fonctions de `scripts/standart_copy.js` de l'Abes, utilise la fonction `suptag()`._

En mode édition, supprime de la notice tous champs commençant par `A`, `9`, `E` ou `e`.
Sert à supprimer les informations d'exemplaires d'une notice __en affichage UNM__ copiée via la fonction dédiée dans WinIBW (et qui aurait été collée sans passer par la fonction `Coller notice`).

##### `__addTextToVar()`

_Paramètres :_
* `vari` : la variable originale ;
* `text` : le texte à rajouter à `vari` ;
* `sep` : le séparateur à placer entre `vari` et `text`.

Renvoie `vari` avec `sep` puis `text` ajoutés à la fin de `vari`.
Si `vari` est une chaîne de caractères vide, renvoie `text`.

##### `__connectBaseProd()`

Se connecte à la base de production du Sudoc.

##### `__connectBaseTest()`

Se connecte à la base de test du Sudoc.

##### `__createWindow()`

Crée une nouvelle fenêtre dans WinIBW.

##### `__dateToYYYYMMDD_HHMM()`

_Paramètre :_
* `date` : un objet Javascript `date`.

Renvoie la `date` sous forme de chaîne de caractères au format `YYYYMMDD_HHMM`.

##### `__deconnect()`

Ferme __l'intégralité__ des fenêtres ouvertes dans WinIBW.

##### `__executeUserScript()`

__En état expérimental, utilisation fortement non recommandée.
Et ce n'est probablement pas une solution efficace si l'on souhaite reprendre le script standart ensuite.__

_Inspiré des travaux de Philippe Combot, notamment ceux pour [lancer des applciations depuis WinIBW](https://web.archive.org/web/20140815070455/http://combot.univ-tln.fr:80/winibw/interactions/applis.html)_

_Paramètre :_
* `fctName` : le nom de la fonction utilisateur voulue ;
* `sleep` : supposément le nombre de millisecondes de pause.
Au vu du script, cette option n'a pas été implémentée, probablement parce que la fonction `sleep` empêcherait la bonne exécution du script utilisateur.

Exécute un script utilisateur en créant un fichier VBS qui sera ensuite exécuter par WinIBW.
Nécessite que la fonction [`executeVBScriptFromName` de `alp_ressources.vbs`](#executeVBScriptFromName) soit installée et __est un raccourci clavier associé.__
Si ce dernier n'est pas `Ctrl+Shift+Alt+L`, il est nécessaire de changer la première fonction `sendKeys` dans la variable `vbsCodeLines`.

Le script écrase (ou crée) le fichier `execute_VBS_from_JS.vbs` dans le profil WinIBW de l'utilisateur avec le code suivant (`vbs_test` prenant la valeur de `fctName`) :

``` VBS
Dim oShell
Set oShell = CreateObject("WScript.Shell")
oShell.AppActivate("WinIBW")
oShell.SendKeys "+^%l"
WScript.Sleep 100
oShell.SendKeys "vbs_test"
WScript.Sleep 100
oShell.SendKeys "{Enter}"
Set oShell = Nothing

```

Ensuite, utilise la fonction [`__executeVBScript`](#__executeVBScript) pour exécuter ce fichier.

##### `__executeVBScript()`

_Paramètre :_
* `filePath` : le chemin d'accès complet d'un fichier.

Exécute le fichier indiqué.

##### `__findExactText()`

_Paramètre :_
* `txt` : le texte à rechercher.

Recherche la première occurrence de `txt` (sensible à la casse) dans la notice et la sélectionne.

##### `__getEnvVar()`

_Paramètre :_
* `varName` : le nom de la variable environnementale voulue.

Renvoie la valeur de la variable environnementale `varName` si elle existe, sinon renvoie `false`.

##### `__getNoticeType()`

Renvoie l'entier :
* `0` si c'est une notice d'autorité,
* `1` si c'est une notice bibliographique,
* `2` si ce n'est aucune des deux.

Pour déterminer cette information, il se base sur la variable `P3VMC` qui correspond au type de document (`008 position 1 et 2`) et, si la première variable n'a pas de valeur, sur la variable `scr` qui correspond au code de l'écran.
Pour `scr`, son utilisation n'est supposée avoir lieu que si le script est utilisé dans le cadre d'une création de notice _ex-nihilo_ ou sur un écran autre qu'une notice.
Le script vérifie donc uniquement si `scr` est égal à `II` (création de notice d'autorité) ou `IT` (création de notice bibliographique).

##### `__hasWarningMsg()`

Renvoie tous les messages d'alerte (messages de type `2`) actuellement affichés dans la fenêtre active, séparés par des `;`.
S'il n'y a aucun message d'alerte, renvoie une chaîne de caractères vide.

##### `__insertText()`

_Paramètre :_
* `txt` : le texte à insérer

Insère `txt` à la fin de la notice.

##### `__isTitle()`

Renvoie `true` ou `false` selon si la fenêtre active a un `title` ou non.

##### `__logIn()`

_Paramètre :_
* `identifiants` : la paire identifiant / mot de passe séparée par un espace.

S'identifie à la base (en utilisant la commande `log`).

##### `__parseDocLine()`

_Paramètre :_
* `line` : la ligne à diviser.

Renvoie sous forme d'_array_ `line` en utilisant les tabulations horizontales comme séparateur. 

##### `__removeAccents()`

_Vieux script, il y a possiblement plus efficace._

_Paramètre :_
* `str` : le texte à modifier.

Renvoie `str` en retirant les accents des voyelles, la cédille des `C` et en séparant en deux lettres `Æ` et `Œ` (majuscules et minuscules).

##### `__serializeArray()`

_Paramètres :_
* `vari` : l'_array_ à transformer ;
* `sep` : le séparateur à employer.

Renvoie `vari` sous forme de chaîne de caractères en utilisant `sep` comme séparateur entre chaque élément.

##### `__sleep()`

_[Provient de la réponse de BeNdErR à la question _JavaScript sleep/wait before continuing_ sur _StackOverflow_, consultée le 12/04/2022](https://stackoverflow.com/questions/16873323/javascript-sleep-wait-before-continuing#16873849)._

_Paramètre :_
* `milliseconds` : le nombre de millisecondes à attendre.

Met en pause le script durant `milliseconds` millisecondes.

##### `__timerToReal()`

_Paramètres :_
* `start` : un objet Javascript `date` correspondant au début d'un intervalle ;
* `end` : un objet Javascript `date` correspondant à la fin d'un intervalle.

Renvoie la différence entre `start` et `end` sous forme d'une chaîne de caractères au format `X minute(s) X seconde(s)`.


------------------------------------------------------------















# Ancienne doc

#### `changeExAnom`

Remplace le `$btm` de la zone eXX associée au RCR par `$bx` ou signale la présence de plusieurs eXX associées à ce RCR ou non. __Le mode d'affichage de la notice doit (probablement) être `UNM` pour fonctionner correctement.__ _Ce script vise un objectif assez précis, voir le contexte de développement à la fin de sa documentation._

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà, puis copie l'intégralité de la notice.
Le script exécute ensuite [`CountOccurrences`](#countoccurences) pour compter le nombre de `chr(10)` suivi de `e` en tenant compte de la casse (= compte le nombre de notices d'exemplaires dans l'ILN) :
* si une occurrence est détectée, exécute [`goToTag`](#gototag) pour se rendre sur le champ 930, puis recule de 1 caractère (bascule sur le champ précédent) et sélectionne les deux prochains caractères sur la gauche (= les deux derniers caractères du champ).
Il compare ensuite si ces deux caractères en minuscule sont égaux à `tm`, auquel cas, ils les remplacent par `x`, récupère le numéro de champ et affiche une infobulle (numéro de champ + `: tm remplacé par x`) ;
* si plus d'une occurrence est détectée, il réexécute `countOccurrences` en comptant cette fois-ci le nombre  `$b` suivi du RCR __(pour utiliser le script sur votre RCR, changez `330632101` en votre RCR)__ :
  * si plus d'une occurrence est trouvée, recherche `$btm` suivi d'un `chr(10)` suivi de `930 ` et récupère le numéro du champ. Si ce numéro commence par `e`, affiche une infobulle (numéro du champ + `à supprimer`, avec comme titre de fenêtre `Exemplaire fictif`), sinon affiche une autre infobulle (`Plusieurs exemplaires réels sur ce RCR. Vérification recommandée.`) ;
  * sinon, affiche une infobulle (`Plusieurs exemplaires réels. Vérification recommandée.`).

_Contexte de développement : dans le cadre d'un chantier sur les thèses, des exemplaires pouvaient avoir en `$b` des `eXX` la mention `TM` (ou `M` supposément, dans la pratique je n'en ai pas vus / je les ai ratés) liée à l'ancien signalement dans téléthèses.
Ainsi, certains exemplaires téléthèses ont été réutilisés sans changer la valeur du `$b`, d'autres sont seulement des exemplaires fictifs en complément de l'exemplaire réel.
Par ailleurs, nous sommes généralement la seule biblitohèque de l'ILN possédant les thèses de ce chantier, ce qui explique les demandes de vérification du script si plusieurs exemplaires sont détectés dans l'ILN._

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `ChantierTheseAddUB183`

Ajoute une UB183 en fonction de la UB215 (notamment des chiffres détectés dans le $a).

_Type de procédure : SUB_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `chantierTheseLoopAddUB183`

Exécute `ChantierTheseAddUB183`, sauf si l'utilisateur refuse l'ajout, sur la liste de PPN présente dans le presse-papier et exporte un rapport des modifications ou non effectuées.

_Type de procédure : SUB_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `getDataUAChantierThese`

Copie dans le presse-papier le PPN, l'année de soutenance, la discipline, le patronyme, le prénom, l'année de naissance, le sexe, le titre et la cote du document, séparés par des tabulations horizontales. Une option permet de réécrire ou d'éditer les champs directement depuis WinIBW.

_Type de procédure : SUB_

_Renvoi :_

Créé dans le cadre d'un chantier sur les thèses, l'exploitation de ces données se fait dans un tableur Excel particulier.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)
