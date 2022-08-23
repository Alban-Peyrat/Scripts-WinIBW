# Scripts pour le PEB

__À noter : le nom réel des scripts est précédé de `AlP_PEB`. Le préfixe a été ici retiré pour faciliter la lecture.__

## Installation des scripts

Il existe deux manières d'installer ces scripts, les résultats sont supposément les mêmes, hormis le [launcher](#launcher) qui diffère entre les deux.
__Toutefois__, [`triRecherche`](#trirecherche), [`searchInSuDb`](#searchinsudb) et [`askFromSu`](#askfromsu) n'existent qu'en __`JS`__.


### En tant que scripts utilisateurs (VBS)

[Voir la procédure d'installation des scripts utilisateurs dans WinIBW](./scripts.md#en-visual-basic-script-vbs).
Le fichier se trouve dans le sous-dossier `scripts` puis `vbs`, sous le nom `alp_PEB.vbs`.

### En tant que scripts standarts (JS)

[Voir la procédure d'installation des scripts standarts dans WinIBW](./scripts.md#en-javascript-js).
Le fichier se trouve dans le sous-dossier `scripts` puis `js`, sous le nom `peyrat_peb.js`.

_Uniquement pour [`triRecherche`](#trirecherche) : vous nécessiterez également le fichier `peyrat_ressources.js`._

## Utiliser les scripts

Rappels :
* si vous avez installé les scripts en tant que scripts utilisateurs, vous les retrouverez dans la catégorie `Fonctions`, si vous les avez installés en tant que scripts standarts, vous les retrouverez dans la catégorie `Standart Fonctions` ;
* les scripts sont précédés de la mention `AlP_PEB`.


### Via boutons

Ouvrez le menu `Options` puis `Personnaliser...` puis `Commandes` et retrouvez le script qui vous intéresse dans la liste des commandes de la bonne catégorie.
Effectuez ensuite un cliqué-déposé vers votre barre de menu.
__Si vous sélectionnez d'afficher du texte, vous pouvez modifier le texte du bouton via l'option `Texte bouton :`.__

### Via raccourcis

Ouvrez le menu `Options` puis `Personnaliser...` puis `Clavier`.
Retrouvez le script qui vous intéresse dans la liste des commandes de la bonne catégorie et affectez-y un raccourci.

## Présentation des scripts

### `getNumDemande()`

Renvoie dans le presse-papier le numéro de la demande PEB.

__Détails :__ Renvoie la variable `P3GA*`.

### `getNumDemandePostValidation()`

Renvoie dans le presse-papier le numéro de la demande PEB __venant d'être effectuée__.

__Détails :__ Le script récupère le premier message affiché dans la barre des messages et renvoie les dix premiers caractères en suivant l'expression `no.` (suivi d'un espace).

### `getPPN()`

Renvoie dans le presse-papier le PPN de la demande PEB.

__Détails :__ Renvoie la variable `P3VTA`.

### `getRCRDemandeur()`

Renvoie dans le presse-papier du RCR demandeur.

__Détails :__ Renvoie la variable `P3VF0` si celle-ci et la variable `P3VF0` sont identiques (tous les cas que j'ai pu voir pour le moment).
Toutefois, si les deux variables existent, il y a probablement une raison, que je ne connais pas.
Aussi, si jamais les deux variables sont différentes, le script ouvre une boîte une boîte de dialogue demandant de sélectionner le RCR voulu, ou aucun des deux.

### `getRCRFournisseurOnHold()`

Renvoie dans le presse-papier du RCR fournisseur dont une réponse est attendue.

__Détails :__ Divise en plusieurs parties la variable `P3VCA` (l'`iframe` contenant la liste des fournisseurs) en utilisant le retour chariot comme séparateur.
Pour chaque fournisseur, recherche un caractère d'échappement suivi de `E` suivi d'un caractère d'échappement suivi de `LRT` et isole ce qui suit jusqu'au prochain caractère d'échappement (supposément, l'information contenue dans la colonne "Commentaire").
Si cette donnée isolée correspond à `En attente de réponse`, il recherche alors la même expression que précédemment en remplaçant `LRT` par `LSS` (supposément le RCR de la bibliothèque) et place dans le presse-papier les 9 caractères suivant (puisque les RCR font 9 caractères).
Le script force ensuite l'arrêt de la boucle.

### `getTitleAuth()`

Renvoie dans le presse-papier, séparés par un retour à la ligne, le titre du document, l'auteur du document, le titre de la partie et l'auteur de la partie, tels que visibles sur la demande.
Attention, si une de ces informations n'existe pas, cela ne supprime pas le retour à la ligne pour autant.
Ainsi, si vous collez le résultat dans Excel, la 3e ligne sera toujours le titre de la partie, qu'un auteur pour le document existe ou non.

__Détails :__ Renvoie les variables `P3VTC`, `P3VTD`, `P3VAB`, `P3VAA`, séparées par des retours à la ligne (`vblf` (VBS) / `\n` (JS)).

### `Launcher()`

Ouvre une boîte de dialogue qui permet de lancer l'exécution d'un des autres scripts de PEB que j'ai développés.

__Détails :__ la boîte de dialogue varie selon si l'on utilise les scripts en VBS ou en JS. __En JS__, la boite de dialogue propose des options cliquables, qui exécuteront les scripts associés . __En VBS__, la boîte de dialogue demande d'entrer le numéro associé au script :
* 0 (VBS) / `Get no demande PEB` (JS) : exécuter [`getNumDemande`](#getnumdemande) ;
* 1 (VBS) / `Get no demande PEB post-validation` (JS) : exécuter [`getNumDemandePostValidation`](#getnumdemandepostvalidation) ;
* _exclusif JS :_ `Trier recherche` (JS) : exécuter [`triRecherche`](#trirecherche) ;
* 2 (VBS) / `Get PPN` (JS) : exécuter [`getPPN`](#getppn) ;
* 3 (VBS) / `Get RCR demandeur` (JS) : exécuter [`getRCRDemandeur`](#getrcrdemandeur) ;
* 4 (VBS) / `Get RCR fournisseur en attente` (JS) : exécuter [`getRCRFournisseurOnHold`](#getrcrfournisseuronhold) ;
* 5 (VBS) / `Get titre et auteur document` (JS) : exécuter [`getTitleAuth`](#getTitleAuth).

### `triRecherche()`

Ouvre un fichier Excel permettant de trier et filtrer les résultats d'une requête.

__Détails :__ le script va créer un fichier `triPEB.xls` (ou effacer ses données s'il existe déjà) dans le profil WinIBW de l'utilisateur (au même emplacement que les fichiers de scripts), puis y écrire les en-têtes.
_Le fichier ne doit pas être ouvert avant l'exécution du script pour un bon fonctionnement de celui-ci._
Le script récupère ensuite le numéro du lot affiché, lance une requête `\too s{NUMÉRO DE LOT} k` pour rafficher le lot et récupérer le nombre de résultats.

Il va ensuite commencer une boucle tant que la ligne traitée est strictement inférieure au nombre de résultats.
À chaque instance de la boucle, le script lance la requête `\too s{NUMÉRO DE LOT} {LIGNE TRAITÉE + 1 } k` et récupère la variable `P3VKZ`, qui contient l'intégralité du segment de la liste des résultats contenant la ligne traitée + 1 [(voir plus d'informations à ce sujet)](./etude_fonctionement_WinIBW.md).
Le script divise ensuite cette variable en utilisant les retours chariots comme séparateur.
Pour chaque chaque notice identifiée, il supprime les guillemets, puis remplace les caractères accentués et autres caractères spéciaux par des lettres classiques [(voir la fonction associée)](./).
Cette étape est nécessaire pour faciliter la lecture du fichier puisque ces caractères seront "corrompus" si laissés tels quels (certains caractères accentués poseront quand même problèmes).
Il isole ensuite plusieurs informations en recherchant le caractère d'échappement suivi de `L` suivi de l'identifiant de l'information, puis en conservant tout ce qui se trouve jsuqu'au prochain caractère d'échappement suivi de `E`.
Voici la liste des informations récupérées, dans l'ordre des colonnes sur le fichier final, avec leur identifiant dans l'affichage par défaut :
* PPN (`PP`) ;
* auteur (`V0`) ;
* titre (`V1`) ;
* éditeur (`V2`) ;
* édition (`V3`) ;
* année (`V4`) ;
* numéro dans le lot (`NR`, pas affiché dans Excel, correspond à la ligne traitée sus-mentionnée).

Une fois ces six premières informations récupérées, il les ajoute dans une nouvelle ligne sur le fichier puis passe à la notice suivante.

Une fois toutes les notices du segment traité, il ajoute 1 à une variable qui casse la boucle `While` si elle atteint 9999 puis passe à la prochaine instance de celle-ci.
Une fois cette boucle achevée, le script ferme le fichier, puis l'ouvre en utilisant l'application configurée par défaut pour ce type de fichier.

### `searchInSuDb()`

_Fonctionne de pair avec [`askFromSu`](#askfromsu) pour permettre de faciliter les opérations de recherche de documents en permettant par exemple de choisir d'autres options de tri.
Toutefois, cette méthode fonctionne en recherchant dans la base de catalogage, or il n'est pas à ma connaissance possible d'exclure de cette base les documents n'étant localisés dans aucune bibliothèque du réseau, ce que fait la base de PEB._

Lance la recherche actuelle dans la base de catalogage.
__Ne connaissant pas de méthode plus efficace pour récupérer l'exacte requête qui a créé un jeu de résultat, j'ai fait le choix de ne faire fonctionner le script uniquement si la requête qui a créé le jeu commençait par `che`.__
Il est probablement possible de prendre en compte des requêtes réaffineées à l'aide de `ET`, `OU` ainsi que de prendre en compte celles provenant de liens, mais je n'ai pas eu le temps de le faire.

__Détails :__ récupère à l'aide de la variable `P3GAD` les limitations utilisée dans la requête, puis la requête elle-même avec `P3LCO`, qui correspond en réalité au texte situé après `Recherche effectuée` (et qui explique les problèmes cités plus haut).
Vérifie si le début de `P3LCO` est égal à `recherche`, si oui, remplace `recherche` par `\zoe`, si non, arrête l'exécution du script en indiquant que ce type de recherche n'est pas pris en compte.
Se connecte ensuite au catalogue dans la base de catalogage, en recherchant la requête récupérée avec les limitations récupérées et en les affichant au format ISBD (via la commande `\sys 1;\bes 1;{limitations + requête};\too i`).
Vérifie ensuite si la recherche a fonctionnée à l'aide de la variable `P3GSY` qui correspond à la base sélectionnée, si ce n'est pas le cas, cesse l'exécution du script et affiche une erreur.
__Après vérification, le test actuellement implanté est (probablement) partiellement erroné :__ en effet, il est possible que la commande échoue après s'être connectée à la base de catalogage sans que le script détecte que l'exécution a échoué.
Dans le doute, il est simplement possible d'appuyer sur Échap pour vérifier la base.
Enfin, exécute la commande `\too k 1` pour afficher le contenu de la page (sinon WinIBW affiche une page vide).

### `askFromSu()`

_Dans l'idéal, il faut rajouter d'autres vérifications avant d'utiliser ce script (par exemple, vérifier si WinIBW affiche une alerte lors du lancement de la recherche)._

_Fonctionne de pair avec [`searchInSuDb`](#searchinsudb) pour permettre de faciliter les opérations de recherche de documents en permettant par exemple de choisir d'autres options de tri.
Toutefois, cette méthode fonctionne en recherchant dans la base de catalogage, or il n'est pas à ma connaissance possible d'exclure de cette base les documents n'étant localisés dans aucune bibliothèque du réseau, ce que fait la base de PEB._

Récupère le PPN d'une notice de la base de catalogage et crée une demande de PEB dans la base SUPEB (__= clique sur le bouton `PEB`__, ne complète pas d'autres informations.

__Détails :__ récupère le PPN via la variable `P3GPP`.
Si cette variable est vide, ouvre une erreur et cesse l'exécution du script.
Autrement, vérifie si la notice n'est pas une notice d'autorité à l'aide de la variable `P3VMC`, qui contient la position 1 et 2 de la `008 $a`, retournant une erreur et cessant l'exécution du script si la position 1 est égale à la valeur `T`.
(Je suis à peu près sûr que cette vérification est une copie provenant de `scripts/standart_copy.js` de l'Abes.)
Se connecte ensuite au catalogue dans la base de PEB, en recherchant le PPN récupéré et en les affichant au format ISBD (via la commande `\sys 2;\bes 1;\zoe ppn {ppn};\too i`).
Vérifie ensuite si la recherche a fonctionnée à l'aide de la variable `P3GSY` qui correspond à la base sélectionnée, si ce n'est pas le cas, cesse l'exécution du script et affiche une erreur.
__Après vérification, le test actuellement implanté est (probablement) partiellement erroné :__ en effet, il est possible que la commande échoue après s'être connectée à la base de PEB sans que le script détecte que l'exécution a échoué.
Notamment, si aucune bibliothèque est localisée, un message d'erreur devrait apparaître ici, mais il n'y en a pas.
Le script exécute alors le 7e bouton de la barre (via `application.activeWindow.simulateIBWKey("F9")`), puis vérifie l'écran qui est affiché.
Cette vérification se fait grâce à la variable `scr` qui doit avoir la valeur `AA` si c'est une demande de PEB.
Si ce n'est pas le cas, affiche un message d'erreur.
