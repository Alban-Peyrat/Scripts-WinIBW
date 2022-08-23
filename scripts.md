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
Vous pouvez placer ces scripts où vous le souhaitez (dans votre profil WinIBW semble être une bonne idée.
Par exemple, les miens se trouvent sous `C:\oclcpica\WinIBW30\Profiles\apeyrat001\alp_scripts`).
* Au sein de `alp_central_scripts.js`, vous devrez éditer la liste des fichiers que vous voulez charger.
Pour ce faire, ouvrez le fichier dans un éditeur de texte et rendez-vous à la ligne 28 du fichier (qui commence par `const alpScripts =`).
Remplacer les noms complets (chemin d'accès au fichier + nom du fichier + extension du fichier) des fichiers par défaut par ceux que vous voulez charger (si les fichiers se trouvent au sein de WinIBW, vous pouvez remplacer le début par `resource:/`).
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

Contient uniquement les scrripts utilisés pour paramétrer WinIBW autant pour son interface () que pour récupérer des variables communes à VBS et JS que pour charger les autres scripts en VBS que pour permettre au fichier central de paramétrage de JS d'être chargé.
_[Consulter le fichier](/scripts/winibw.vbs)_

##### Lignes de code hors des fonctions

* `application.writeProfileString "ibw.standardScripts","script.AlP","resource:/Profiles/apeyrat001/alp_scripts/alp_central_scripts.js"` : permet de charger le script central de JS qui permettra de charger par la suite les autres scripts JS.
Changez `resource:/Profiles/apeyrat001/alp_scripts/alp_central_scripts.js` par le chemin d'accès à votre script central de JS.
* `sluitMapIn("C:\oclcpica\WinIBW30\Profiles\apeyrat001\alp_scripts\vbs")` : permet de charger les autres scripts VBS.
Changez `C:\oclcpica\WinIBW30\Profiles\apeyrat001\alp_scripts\vbs` par le chemin d'accès à votre dossier contenant les scripts VBS.
Vous pouvez charger plusieurs dossiers, ou charger un fichier individuellement à l'aide de [la fonction `sluitVBSin()`](#sluitVBSin).
* `Set WSHShell = CreateObject("WScript.Shell")` : permet de créer un objet `WScript.Shell` qui vous permettra de récupérer les informations d'une variable environnementale, à l'aide de `WSHShell.ExpandEnvironmentStrings("%MY_RCR%")`, en remplaçant `MY_RCR` part le nom de la variable.

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
Le script récupère ensuite le titre du document via la fonction [`getTitle`](`#gettitle`) ainsi que l'année en utilisant la `100 $c` ou `100 $a`.
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

Passe la notice en mode édition si elle ne l'est pas déjà puis détermine la position du `@` au sein de la `200 $a` et récupère le titre du document via la fonction [`getTitle`](`#gettitle`).
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


#### Fichier `alp_concepts.vbs`

Contient tous les concepts de scripts que j'ai pu développer pour contrôler le signalement.
_[Consulter le fichier](/scripts/vbs/alp_concepts.vbs)_

##### `ctrlUA103eqUA200f()`

___Voir [ConStance CS1](../../../ConStance#cs1--équivalence-champs-103--200f-idref) pour un outil équivalent.___

Exporte et compare le $a de UA103 et le $f de UA200 pour chaque PPN de la liste présente dans le presse-papier.

##### `ctrlUB700S3()`

___Voir [ConStance CS2](../../../ConStance#cs2--présence-dun-lien-en-700) pour un outil équivalent et [ConStance CS3](../../../ConStance#cs3--pr%ésence-dun-lien-en-7xx) pour un outil équivalent utilisable sur toutes les 700.___

Exporte le premier $ de UB700 pour chaque PPN de la liste présente dans le presse-papier.


------------------------------------------------------------


#### Fichier `alp_corwin.vbs`

Contient tous les scripts permettant le fonctionnement du [projet CorWin permettant de contrôler des données dans WinIBW](../../../CorWin).
_[Consulter le fichier](/scripts/vbs/alp_corwin.vbs)_


------------------------------------------------------------


#### Fichier `alp_dumas.vbs`

Contient tous les scripts développés en lien avec [DUMAS](https://dumas.ccsd.cnrs.fr/).
[Le dépôt ub-svs contient plus d'informations à ce sujet](../../../ub-svs).
_[Consulter le fichier](/scripts/vbs/alp_dumas.vbs)_


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

##### `getNoticeType()`

Renvoie l'entier :
* `0` si c'est une notice d'autorité,
* `1` si c'est une notice bibliographique,
* `2` si ce n'est aucune des deux.

Pour déterminer cette information, il se base sur la variable `P3VMC` qui correspond au type de document (`008 position 1 et 2`) et, si la première variable n'a pas de valeur, sur la variable `scr` qui correspond au code de l'écran.
Pour `scr`, son utilisation n'est supposée avoir lieu que si le script est utilisé dans le cadre d'une création de notice _ex-nihilo_ ou sur un écran autre qu'une notice.
Le script vérifie donc uniquement si `scr` est égal à `II` (création de notice d'autorité) ou `IT` (création de notice bibliographique).


------------------------------------------------------------


### Scripts standarts (JS)

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

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `add18XmonoImpIll`

Ajoute les 181-2-3 pour une monographie imprimée avec illustration.

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà puis insère à la fin de celle-ci :
* `181 ##$P01$ctxt`
* `181 ##$P01$csti`
* `182 ##$P01$P02$cn`
* `183 ##$P01$P02$anga`
* un retour à la ligne.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `add214Elsevier`

Ajoute une 214 type pour Elsevier (2021).

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà puis insère à la fin de celle-ci :
* `214 #0$aIssy-les-Moulineaux$cElsevier Masson SAS$dDL 2021`
* un retour à la ligne

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `addBibgFinChap`

Ajoute une mention de bibliographie à la fin de chaque chapitre.

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà puis insère à l'emplacement du curseur :
* `Chaque fin de chapitre comprend une bibliographie`

__Malfonctionnement possible : si la notice n'était pas en mode édition, le texte ne s'écrira probablement pas si la grille des données codées n'est pas affichée.__

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `addCouvPorte`

Ajoute le début d'une 312 `La couverture porte en plus`.

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà puis insère à la fin de celle-ci :
* `312 ##$aLa couverture porte en plus : "`

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `addISBNElsevier`

Ajoute une 010 avec le début de l'ISBN d'Elsevier.

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà puis insère à la fin de celle-ci :
* `010 ##$A978-2-294-`

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

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

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `addUA400`

Rajoute des UA400 pour les noms composés en se basant sur la UA200, sinon rajoute une UA400 copiant la UA200. _Ce script n'est pas universel et ne fonctionne qu'en présence d'un `$a` et d'un `$b`._

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà, puis lance le script [`findUA200aUA200b`](#findua200aua200b), pour récupérer la 200, la 200 `$a`, la 200 `$b` et la position du premier dollar (ou de la fin du champ) après le `$b`.
Il lance ensuite le script [`decompUA200enUA400`](#decompua200enua400) en injectant le `$a` et le `$b` précédemment obtenu pour récupérer les 400 des noms composés.
Il vérifie ensuite si la longueur du champ renvoyé par `decompUA200enUA400` est inférieure à 5 (= si aucune 400 n'a été générée) :
* si c'est le cas, il va copier la 200 précédemment obtenue en supprimant tout ce qui se trouve après la position du premier dollar après le `$b`, puis remplace dans ce qu'il reste `200` par `400` et supprime `$90y`.
Il insère ensuite le nouveau champ à la fin de la notice et place le curseur après le huitième caractère de celui-ci (en théorie, au début du contenu du premier dollar).
* si ce n'est pas le cas, il insère le champ renvoyé à la fin de la notice.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `addUB700S3`

Remplace la UB700 actuelle de la notice bibliographique par une UB700 contenant le PPN du presse-papier et le $4 de l'ancienne UB700.

_Type de procédure : SUB_

Passe la notice en mode édition si elle ne l'est pas déjà, puis recherche à l'intérieur de celle-ci un retour à la ligne (`chr(10)`) suivi de `700` (supposément, la première 700).
Le script sélectionne ensuite les trois derniers caractères de ce champ (supposément le code fonction) puis génère :
* `700 #1$3` + le contenu du presse-papier + `$4` + la sélection en cours.

Il supprime de ce champ généré les retours à la ligne (`chr(10)`), puis supprime le champ où se trouve le curseur (ancienne 700) et insère à sa place la nouvelle 700 et un retour à la ligne.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

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
Il rajoute ensuite le champ ci-dessous à la valeur qui sera renvoyée (via [`appendNote`](#appendnote)) avant de passer à la prochaine instance :
* `400 #1$a` + la valeur actuelle de `UA200a` + `$b` + la valeur actuelle de `UA200b`

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

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
* si le compteur a une autre valeur, assigne à `UA200fPos` la longueur de la 200 __+ 1__ (sinon [`addUA400`](#addua400) supprimerait parfois la dernière lettre du prénom).

Il isole ensuite la valeur du `$a` puis du `$b` et renvoie la 200, la `$a` isolé, le `$b` isolé et `UA200fPos` __sous forme d'une seule chaîne de caractères en séparant les différentes valeurs par `;_;`__.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_ressources.vbs)

#### `generalLauncher`

Ouvre une boîte de dialogue servant à lancer les scripts (majoritairement de type `add` et `get`). 

_Type de procédure : SUB_


Ouvre une boîte de dialogue contenant la liste des scripts suivants accompagnés de leur identifiant, la liste étant décomposée en plusieurs parties :
* notices bibliographiques :
  * 14 : exécuter [`add18XmonoImp`](#add18xmonoimp) ;
  * 1 : exécuter [`addCouvPorte`](#addcouvporte) ;
  * 2 : exécuter [`addBibgFinChap`](#addbibgfinchap) ;
  * 3 : exécuter [`addEISBN`](#addeisbn) ;
  * 4 : exécuter [`AddSujetRAMEAU`](#addsujetrameau) ;
  * 15 : placer dans le presse-papier le renvoi de [`addUB700S3`](#addub700s3) ;
* Elsevier :
  * 6 : exécuter [`addISBNElsevier`](#addisbnelsevier) ;
  * 7 : exécuter [`add214Elsevier`](#add214elsevier) ;
* récupérer des informations :
  * 8 : placer dans le presse-papier le renvoi de [`getTitle`](#gettitle) ;
  * 9 : placer dans le presse-papier le renvoi de [`getCoteEx`](#getcoteex) ;
* thèses
  * 10 : exécuter [`getDataUAChantierThese`](#getdatauachantierthese) ;
  * 5 : exécuter `perso_CTaddUB700S3` ;
  * 11 : placer dans le presse-papier le renvoi de [`getUB310`](#getub310) ;
* notices d'autorités
  * 12 : exécuter [`addUA400`](#addua400) ;
  * 13 : placer dans le presse-papier le renvoi de [`getUA810b`](#getua810b) ;
* CorWin :
  * 77 : lance le lanceur de [CorWin](https://github.com/Alban-Peyrat/CorWin).


_Contexte de développement : j'utilise des raccourcis pour la majorité de mes scripts. Or à force de créer de petits scripts, les combinaisons de raccourcis se limitent et m'obligent à retenir beaucoup de raccourcis différents. Le lenceur général permet donc de réduire ce nombre. Aussi, les nombres sont attribués dans l'ordre d'ajout et non pas dans l'ordre où ils sont listés._

#### `getCoteEx`

Renvoie dans le presse-papier la cote du document. Si plusieurs cotes sont présentes, donne le choix entre en sélectionner une, ou toutes les sélectionner, permettant également de choisir le séparateur.

_Type de procédure : FUNCTION_

_Renvoi :_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `getDataUAChantierThese`

Copie dans le presse-papier le PPN, l'année de soutenance, la discipline, le patronyme, le prénom, l'année de naissance, le sexe, le titre et la cote du document, séparés par des tabulations horizontales. Une option permet de réécrire ou d'éditer les champs directement depuis WinIBW.

_Type de procédure : SUB_

_Renvoi :_

Créé dans le cadre d'un chantier sur les thèses, l'exploitation de ces données se fait dans un tableur Excel particulier.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `getTitle`

Renvoie dans le presse-papier le titre du document en remplaçant les @ et $e. Si le titre est entièrement en majuscule, le renvoie en minuscule (sauf première lettre).

_Type de procédure : FUNCTION_

_Renvoi :_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `getUA810b`

Si un seul UA810 est présent, écrit le $b "né le" à partir des informations de la UA103 de la notice, sinon, renvoie le $b dans le presse-papier.

Pour un bon fonctionnement, la UA103 doit comprendre AAAAMMJJ.

_Type de procédure : FUNCTION_

_Renvoi :_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `getUB310`

Copie dans le presse-papier la valeur du premier UB310.

_Type de procédure : FUNCTION_

_Renvoi :_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `PurifUB200a`

Renvoie l'adaptation d'un titre en son écriture en UNIMARC.

_Type de procédure : FUNCTION_

_Renvoi :_

_Paramètres :_
* UB200 : PAS A JOUR
* isUB541 : PAS A JOUR

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `searchDoublonPossible`

Recherche le PPN qualifié de doublon possible par WinIBW.

_Type de procédure : SUB_

Récupère le premier message affiché, si celui-ci contient `PPN` suivi d'un espace, isole les neuf caractères suivant cette expression et lance la recherche `che ppn` avec le PPN isolé.
Si l'expression n'est pas trouvée, renvoie un message d'erreur.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

#### `searchExcelPPNList`

Recherche la liste de PPN contenue dans le presse-papier.

_Type de procédure : SUB_

Transforme la liste de PPN du presse-papier en :
* supprimant `(PPN)` ;
* remplançant `chr(10)` par `OR` (avec espace avant et après) ;
* ajoutant au début `che PPN` suivi d'un espace ;
* supprimant les quatre derniers caractères (supposément `OR` avec un espace avant et après).

Place ensuite la requête dans le presse-papier et lance la requête.

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_principaux.vbs)

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

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_ressources.vbs)

#### `CountOccurrences`

Renvoi le nombre d'occurrences.

_Type de procédure : FUNCTION_

_Paramètres :_
* `p_strStringToCheck` : variable qui sera fouillée ;
* `p_strSubString` : texte à chercher ;
* `p_boolCaseSensitive` : __bool__ définit si la recherche sera sensible à la casse.

Renvoie le nombre de fois où `p_strSubString` apparait dans `p_strStringToCheck` en comptant le nombre de parties lorsque l'on divise `p_strStringToCheck` en utilisant `p_strSubString` comme séparateur.
Si `p_boolCaseSensitive` est `false`, alors le script passe dans un premier temps les deux autres variables en minuscule.

[Consulter la source originale](https://www.thoughtasylum.com/2009/07/30/VB-Script-Count-occurrences-in-a-text-string/), [consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_ressources.vbs)

#### `exportVar`

Exporte le texte injecté dans `export.txt` (même emplacement que `winibw.vbs`). __Pour l'utiliser, pensez à changer la destination du document, et le nom si vous le souhaitez.__

_Type de procédure : SUB_

_Paramètres :_
* `var` : le texte à exporter ;
* `boolAppend` : __bool__ définit si le script doit ajouter à la fin du fichier (`true`) ou réécrire le fichier.

[Consulter la source originale](http://eddiejackson.net/wp/?p=8619), [consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_ressources.vbs)

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

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_ressources.vbs)

#### `goToTagInputBox`

Permet d'essayer `goToTag` en indiquant les paramètres voulus.

_Type de procédure : SUB_

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_ressources.vbs)

#### `Sleep`

Permet de mettre en pause un script. __Évitez l'utilisation.__

_Type de procédure : SUB_

_Paramètres :_
* `time` : __int__ temps à attendre (en secondes).

[Consulter la source originale](https://stackoverflow.com/questions/1729075/how-to-set-delay-in-vbscript#answer-12921137), [consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_ressources.vbs)

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

[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_ressources.vbs)

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


[Consulter le script](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/scripts/scripts_ressources.vbs)

