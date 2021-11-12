# Liste des modifications

## Le 02/08/2021 :
* suppression de `PurifUB200a` car peu d'intérêts à être partagé ;
* suppression de `CollerPPN` car peu d'intérêts à être partagé ;
* suppression de `LastCHE` car peu d'intérêts à être partagé.

## Le 23/08/2021 :
* ajout de `AddSujetRAMEAU` pour ajouter des 60X ;
* ajout de `ctrlTraitementInterne` ;
* ajout de `getUB310` pour récupérer dans le presse-papier l'information de la première 310 ;
* ajout de `PurifUB200a` pour adapter un titre à son écriture en UNIMARC ;
* scission de `addUB700S3` : la partie sur l'exemplaire a été isolée dans un nouveau script, `changeExAnom`.

## Le 24/08/2021 :
* répartition des scripts entre plusieurs fichiers ;
* actualisation des présentations des scripts, notamment en intégrant les dernières modifications ;
* adaptation du projet pour être cohérent avec les autres outils.

## Le 25/08/2021 :
* suppression de `ctrlTraitementInterne`, que j'avais dû arrêter en plein milieu du développement ;
* modification de la description de [concepts](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/concepts.vbs) et ajout de la mention de [Constance](https://github.com/Alban-Peyrat/ConStance) ;

## Le 01/09/2021 :
* ajout de `appendNote` pour ajouter à une variable la donnée voulue ;
* ajout de `getDataUAChantierThese` pour exporter les données d'une thèse dans le cadre d'un chantier sur les thèses ;
* ajout de `uCaseNames` pour mettre des majuscules aux noms renseignés ;
* modification de `getCoteEx` dû à une réécriture du script. Détecte désormais l'intégralité des cotes associées au RCR et permet de sélectionner celles voulues, ou toutes ;
* probable mise à jour prochaine de `decompUA200enUA400` pour être plus efficace et utiliser `uCaseNames` ;

## Le 02/09/2021 :
* modification de `getDataUAChantierThese` pour réorganiser l'inputBox, rajouter de la précision à la note sur les noms d'épouse et empêcher des valeurs illégales pour le genre ;
* modification de `getCoteEx` pour afficher le numéro de l'occurrence et de l'exemplaire en cas de cotes multiples, ainsi que de gérer la sélection individuelle de plusieurs cotes.

## Le 13/10/2021 :
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

## Le 15/10/2021 :
* renommage des scripts ressources et concepts (parce que j'ai découvert que WinIBW peut séparer les scripts en plusieurs catégories) ;
* modification mineuresur le fonctionnement de la boucle de `AddSujetRAMEAU` ;
* ajout de la documentation complète de :
  * `add18XmonoImp` ;
  * `add214Elsevier` ;
  * `addBibgFinChap` ;
  * `addCouvPorte` ;
  * `addISBNElsevier` ;
  * `AddSujetRAMEAU`.

## Le 26/10/2021 :
* ajout de la documentation pour les scripts (`get`, chantiers thèse, concepts et `goToTag` exclus, qui pour une majorité nécessitent des modifications importantes) ;
* modifications mineures de `addUA400` et nettoyage du code ;
* `findUA200aUA200b` : correction d'un bug qui supprimait la dernière lettre de la partie rejetée si la 200 ne comprenait pas de `$` après le `$b` ;
* `addUB700S3` :
  * n'utilise plus le presse papier ;
  * suppression de la partie du code spécifique au chantier thèse. Un script personnel regroupant celui-ci et la partie supprimée a été créé ;
  * nettoyage du code ;
* `generalLauncher` : modification de l'ID de `addUB700S3` (pour moi, cela correspond à conserver le même ID pour le même script, la nouvelle version de `addUB700S3` étant techniquement un nouveau script) ;
* `changeExAnom` : récupère la notice par lui-même, n'utilise plus le presse-papier et nettoyage du code ainsi que suppression de parties inutiles ;
* `exportVar` : nettoyage du code ;
* `uCaseNames` : force désormais la majuscule sur la première lettre (le script était conçu pour des noms entièrement en maajuscule).

## Le 04/11/2021 :
* la liste de modifications se trouve désormais dans un fichier à part ;
* ajout [des scripts de PEB et de leur documentation](https://github.com/Alban-Peyrat/Scripts-WinIBW/blob/main/PEB.md) ;
* ajout des scripts :
  * `addAutFromUB` : génère un squelette de notice auteur à partir d'une notice bibliographique ;
  * `addUB7XX` : ajoute une `7XX`, avec un fonctionnement similaire à `addSujetRAMEAU` ;
  * `chantierThese_getJuryForExcel` : récupère des données de la thèse pour l'exporter vers Excel ;
  * `chantierThese_addJuryFromExcel` : ajoute à la notice bibliographique une `200$g`, une `314` et les `701` prévues par Excel ;
  * `chantierThese_addJuryAut`: ajoute une notice d'autorité auteur à partir de l'extraction de `chantierThese_getJuryforExcel` ;
  * `ress_getTag` : renvoie la valeur du/des champs/sous-champs voulus (une phase de test est prévue) ;
* correction de `addUA400` : détecte désormais correctement les particules rejetées quelle que soit la casse.

## Le 05/11/2021 :
* ajout de `PEBgetPPN` et `PEBgetRCRFournisseurOnHold` ;
* ajout des accents dans les scripts PEB en JS ;
* ajout de sécurités et réponses en cas d'échec dans les scripts PEB ;
* correction des préfixes des scripts PEB (problème pour les scripts utilisateurs).

## Le 15/11/2021 :

### Scripts pour le PEB :
* ajout de `PEBtriRecherche` ;
* mise à jour de `PEBLauncher` ;
* mise à jour de la procédure d'installation et de l'introduction.

### Scripts en VBS :
* 

### Scripts en JS :
* création du fichier `peyrat_js_scripts.js` ;
* ajout de `AlP_js_removeAccents`.

# Modifications prévues

* `getTitle` : permettre son utilisation autant en mode édition que présentation ;
* scripts de type `get` : vérification de l'utilisation du presse-papier et restituer le presse-papier présent avant le lancement du script s'il est réécrit ;
* `addUA400` (et associés) : adapter à `getTag`, mettre des majuscules si nécessaire, gérer les autres informations ;
* `decompUA200enUA400` : gérer les cas où l'indicateurs 2 est `0` ainsi que l'absence de `$b` ;
* ajout de `addEISBN` : ajoute une 452 avec un _place holder_ ou le titre s'il est déjà renseigné, ainsi que les trois premières parties de l'ISBN. Apparement j'ai cessé le développement en plein milieu ;
* correction du malfonctionnement probable de `addBibgFinChap` ;
* nettoyage et correction de code et des commentaires de début de script ;
* séparation des scripts de chantier thèses.
