# Étude du fonctionnement de WinIBW

_À noter : une partie de ces informations sont des observations et des inductions de ma part, il n'est pas exclu que j'ai mal interprété et tiré de mauvaises conclusions._

## L'affichage d'une liste de résultats

Les listes de résultats en présentation courte (_short presentation list_ dans la documentation [_Scripting in WinIBW3 : getting started (v1.17) / OCLC (lien direct)_](http://www.zeitschriftendatenbank.de/fileadmin/user_upload/ZDB/pdf/winibw/Scripting_in_WinIBW3_V_1_17.pdf)) sont renvoyées dans WinIBW via la variable `P3VKZ`.
Toutefois, la variable ne contient pas l'intégralité de la liste, mais seulement un segment de celle-ci.
Dans cette étude, nous allons voir comment se comporte cette variable `P3GKZ` lorsque WinIBW nous affiche une liste de résultats.

### Vocabulaire utilisé dans l'étude

Pour éviter une confusion sur la terminologie employée dans cette étude, je vais me permettre de définir ce que j'entends pour certains termes :
* `configuration` désigne tout autant les paramètres physiques d'un ordinateur (taille de l'écran par exemple) que la taille de la fenêtre de WinIBW, le nombre de barres d'outils dans WinIBW, que la police d'écriture, la taille de cette dernière, etc. ;
* `entrée` désigne une notice dans un lot de notices, son numéro est celui indiqué dans la première colonne de l'affichage en présentation courte.
Il n'est pas exclu que j'emploie également le terme `notice` ;
* `ligne` désigne une ligne dans la liste visible sur WinIBW, le nombre de celles-ci est fixe et dépend de la configuration.
Par exemple, mon écran professionnel avec les barres d'outils que j'ai dans WinIBW me permettent d'afficher 33 lignes simultanément lorsque je suis en plein écran.

### Quelques mots sur les lots

_À écrire_

### Segmentation du lot de notices

Lorsqu'une requête est entrée dans la barre de commande de WinIBW comme `aff s14`, `aff s14 153 k` ou `che mti poisson`, il arrive que WinIBW envoie également d'autres requêtes au serveur sans que l'utilisateur ne les voit.
Elles ont pour but d'afficher proprement la liste.
Ces requêtes envoyées ont la forme `\too s14 17 k`, qui est le langage interne de WInIBW (basé sur le néerlandais, ici `too` veut dire `toon` (`afficher` en français)).
Grâce à ces requêtes, nous pouvons mieux comprendre comment sont générées ces listes.

Celles-ci sont divisées en plusieurs segments, qui possèdent les caractéristiques suivantes :
* la divison du lot s'occurre à partir du début de la liste, ainsi, le premier segment contiendra toujours les 16 premières notices du lot, tout comme le treizième segment contiendra toujours les notices 193 à 208 ;
* le dernier segment du lot est le seul à pouvoir contenir moins de 16 notices.
Il contiendra le nombre de notices restants.
Par exemple, si le lot contient 18 notices, le second segment ne contiendra que 2 notices, la 17 et la 18.

### Renvoi dans WinIBW des segments du lots

Maintenant que nous avons établi comment sont divisés ces segments, nous allons nous intéresser à comment ils sont renvoyés dans WinIBW par le serveur.
Tout comme la segmentation du lot, le renvoi de segment de lot suit quelques règles :
* sont uniquement renvoyés l'intégralité des segments nécessaires à l'affichage immédiat de la liste (c'est-à-dire les segments contenant les entrées qui seront visibles) ;
* le premier segment renvoyé (que j'appelerai `Sn`) est toujours celui contenant la notice mentionnée. Si aucune notice n'est mentionnée, comme c'est le cas lorsque on lance une commande `che` ou `tri`, la notice sélectionnée est par défaut 1 (c'est pour cela que la première entrée est sélectionnée par défaut) ;
* les autres segments renvoyés le sont dans l'ordre numérique.
Ainsi, si `Sn+1` et `Sn+2` doivent être renvoyés, `Sn+1` sera le premier, si `Sn+1` et `Sn-1` doivent être renvoyés, `Sn-1` sera le premier, si `Sn-1` et `Sn-2` doivent être renvoyés, `Sn-2` sera le premier.

### Le comportement de WinIBW sur la ligne sélectionnée

Avant d'aller plus loin, il convient de comprendre le comportement de WinIBW lié à l'emplacement de la ligne sélectionnée dans la liste :
* pour l'affichage initial, l'entrée sélectionnée sera celle indiquée dans la requête (ou la première entrée si aucun numéro de notice n'est transmise dans la requête) ;
* cette ligne sera la cinquième si cela est possible (c'est-à-dire si la notice demandée est la cinquième ou plus)...
* ...sauf si la configuration ne permet pas d'afficher plus de 5 lignes simultanément, auquel cas, la ligne sélectionnée sera :
  * l'avant dernière ligne si la configuration permet d'afficher 4 ou 5 lignes ;
  * la troisième ligne si la configuration ne permet pas d'afficher plus de trois lignes.
__Même si cela implique que la ligne ne soit pas visible__, dans le cas où seules 1 ou 2 lignes sont visibles.
* concernant le défilement de l'écran, la ligne sélectionnée sera toujours la ligne __non-rognée__ située à l'extrémité de la liste.
Soit toujours la première ligne en défilant vers le haut, et généralement l'avant-dernière ligne en défilant vers le bas car la dernière ligne est généralement en partie rognée.
À noter que, suivant la logique de renvoi des segments précédemment évoquée, le segment contenant l'entrée associée à la ligne rognée est renvoyé __dès qu'elle est supposée s'afficher, même si elle est rognée__.

### Note sur l'affichage de la liste

Il est à noter que WinIBW ne laisse aucune ligne vide s'il peut l'éviter.
C'est-à-dire que tant que le nombre total d'entrées est supérieur au nombre de lignes affichables par la configuration, WinIBW n'affichera aucune ligne vide non-rognée.
(Si la dernière ligne est rognée, il affichera évidemment une ligne vide sous la dernière entrée lorsque celle-ci est sélectionnée).

### Déterminer les segments renvoyés

En suivant toutes les règles évoquées jusqu'à maintenant, nous pouvons donc déterminer par avance quels seront les segments renvoyés lors de l'affichage initial de la liste et dans quel ordre ils le seront :
* si le numéro de notice sélectionné se trouve être l'une des 4 premières entrées du segment, le premier segment renvoyé sera `Sn`, suivi de `Sn-1`, suivi de tous les `Sn+X` nécessaires...
* ...sauf si `Sn-1` n'existe pas (c'est-à-dire que `Sn` est le premier segment du lot), auquel cas, seront renvoyés après `Sn` tous les `Sn+X` nécessaires ;
* si le numéro de notice sélectionné se trouve être entre la 5e et la 16e entrée comprise, le premier segment renvoyé sera `Sn`, puis seront renvoyés tous les `Sn+X` nécessaires...
* ...sauf si un ou plusieurs de ces `Sn+X` n'existent pas (c'est-à-dire si `Sn` ou si l'un des `Sn+X` qui n'est pas le dernier à devoir être affiché se trouve être le dernier segment du lot), auquel cas, seront renvoyés après `Sn` tous les `Sn-X` nécessaires suivis de tous les `Sn+X` jusqu'au dernier segment.

Après l'affichage initial, les segments sont renvoyés au fur et à mesure de leur nécessité d'être affichés.

Prenons quelques exemples pour illustrer ces propos, notamment le dernier point de la liste qui peut être difficile à comprendre.
Nous nous baserons sur deux configurations, `maConfig` qui est capable d'afficher 33 lignes (soit 3 segments) et `largeConfig` qui est capable d'afficher 100 lignes (soit 7 segments).
Soit la requête :
* `che aut leger` (qui deviendra le lot numéro 1), renvoyant 169 résultats :
  * `maConfig` affichera `S1` puis `S2` puis `S3` ;
  * `largeConfig` affichera `S1` puis `S2` [...] jusqu'à `S7` ;
* `aff s1 34 k` : 
  * `maConfig` affichera `S3` (lignes 33 à 48) puis `S2` (lignes 17 à 32) puis `S3` ;
  * `largeConfig` affichera `S3` (lignes 33 à 48) puis `S2` (lignes 17 à 32) puis `S4` [...] jusqu'à `S8` ;
  * en effet, la notice demandée, 34, est comprise dans les 4 premières entrées du segment `S3`, il est donc nécessaire d'afficher le segment précédent ;
* `aff s1 127 k` :
  * `maConfig` affichera `S8` (lignes 113 à 128) puis `S9` (lignes 129 à 144) puis `S10` (lignes 145 à 160) ;
  * `largeConfig` affichera `S8` (lignes 113 à 128) puis `S5` jusqu'à `S7` puis `S9` (lignes 129 à 144) puis `S10` (lignes 145 à 160) puis `S11` (lignes 161 à 169)
  * pour `maConfig`, le résultat est assez simple, la notice 127 est la quinzième notice de son segment donc il n'est pas nécessaire d'afficher le segment précédent, il affiche donc les segments suivants.
En revanche, pous `largeConfig`, c'est un peu plus compliqué.
WinIBW doit afficher les trois mêmes segments que `maconfig`, puis essayer d'afficher la suite des résultats.
Or, avec `S11`, il atteint la dernière entrée du lot, sauf qu'il n'a que 57 lignes des 100 qu'il doit afficher.
Il doit donc aller chercher les 43 lignes restantes avant `S8`, soit les trois segments antérieur, à savoir `S5`, `S6` et `S7`.
Étant donné que le premier renvoi est forcément le segment contenant la notice demandée, il renvoie `S8` dans un premier temps, puis il renvoie dans un second temps les 6 segments restants dans l'ordre numérique, soit `S5`, `S6`, `S7`, `S9`, `S10`, `S11`.

Il est également à noter que le serveur renvoie la page à afficher en même temps que le premier segment.

### `P3VKZ` et les scripts

Venons-en maintenant à l'application de toute cette théorie dans le cadre de scripts.
Lorsque l'on envoie une requête via un script, __il semblerait__ que WinIBW attende la fin de l'exécution dus script pour d'afficher la liste en appelant le reste des segments nécessaires.
__Mais__, il semblerait également que si nous appelons la variable `P3VKZ` d'une quelconque manière __depuis un script standart, WinIBW n'essaeira pas d'obtenir les autres segments nécessaires à l'affichage__, ce qui aura pour effet de ne pas afficher la liste présentation courte (mais tout le reste de l'interface).
__Ainsi, il me semble tout à fait possible d'utiliser `P3VKZ` autant dans un script standart que dans un script utilisateur afin de récupérer un segment spécifique d'une liste présentation courte.__
Si je suis définitivement convaincu pour les scripts standarts, une tentative de passage de [`AlP_PEBtriRecherche`](./PEB.md#trirecherche) en `VBS` devrait me fixer définitivement sur cette idée.
 
De fait, puisqu'il est possible de récupérer un segment spécifique, il est donc possible de récupérer entièrement la liste puisque le nombre de notices dans un lot est connu et qu'il est accessible via la variable `P3GSZ`.
Il suffit donc de générer une boucle qui récupère un à un les segments en envoyant une nouvelle requête lorsque l'on a achevé de traiter la dernière entrée du segment occupant actuellement `PEVKZ`.

Par ailleurs, si depuis le début j'utilise `k` comme mode d'affichage, tout ce qui est dit semble également vrai pour le mode d'affichage `k:{zoneUNIMARC}`, ce qui signifie qu'il est également possible de traiter en masse directement via WinIBW sans devoir passer par les notices une par une.

### Les différentes colonnes

Avant de terminer cette étude, il semble important de mentionner la forme que prend `P3VKZ`.

Il existe deux caractères qui sont omniprésents dans cette variable à savoir :
* la caractère d'échappement (`27` en ASCII, `001B` en Unicode) ;
* le retour chariot / retour à la ligne (`13` en ASCII, `000D` en Unicode).

Ce dernier est utilisé comme séparateur entre chaque entrée.
En ce qui concerne le premier, je dois avouer être moins certain de quelques unes de mes affirmations.
Toutefois, voici les conclusions que j'ai pu tirer :
* suivi de `H`, il signifie que la prochaine colonne est cachée.
Il se situe toujours juste avant l'information.
(`H` pour _hidden_ ?) ;
* suivi de `E`, il signifie la fin de la colonne, y compris la dernière de la ligne (il y en a donc un avant le retour chariot) (`E` pour _end_ ?) ;
* suivi de `D`, il sert à basculer de nouveau en affichage visible ?
J'émets cette hypothèse car il semble apparaître toujours après le `E` d'une colonne contenant un `H`.
Dans le cadre de cette hypothèse, il signifierai _display_ ?, ce qui voudrait dire que le statut `caché` est une bascule avec `H` comme activateur et `D` comme désactivateur.
* suivi de `L`, il est forcément suivi de deux autres caractères majuscules / numériques.
`L` sert à introduire une variable, les deux caractères suivant son l'identifiant de cette variable.

Ainsi, nous pouvons récupérer les informations dans la variable `P3VKZ` grâce à leur identifiant :
* `PP` pour le PPN ;
* `MB` pour [le type de notice autorité](http://documentation.abes.fr/sudoc/formats/unma/zones/008.htm) ou le [type de document](http://documentation.abes.fr/sudoc//formats/unmb/DonneesCodees/CodesZone008.htm#pos1-2) selon le type de notice (position 1 et 2 de la `008`) ;
* `MC` pour le type de document sous une autre forme, basé sur `MB` ?
En effet, en parcourant un peu les résultats, j'ai pu identifier un certain nombre de valeurs en `MC` (la première valeur dans la liste ci-dessous) qui correspondaient à chaque fois, je pense, à des valeurs de `MB` (la deuxième valeur dans la liste ci-dessous) :
  * `art` pour `As` (_article_ ?) ;
  * `avi` pour `Ba` (_audio video interleave_ ?) ;
  * `mon` pour `Aa` (_monograph_ ?) ;
  * `olr` pour `Oa` (_online ressource_ ?) ;
  * `per` pour `Ad` et `Ab`(_periodical_ ?) ;
  * `snd` pour `Ga` (_sound_ ?)
  * `ths` pour `Tb`, `Td`, `Tg`, `Tp` et `Tq` (_thesaurus_ ? Je ne suis pas un expert en histoire des catalogues informatiques, mais j'ai cru comprendre qu'à une époque les notices d'autorité auteurs n'existaient pas, de fait, lors de leur implémentation, il leur aurait été attribué le même identifiant que les autres autorités, à savoir les autorités sujet).
* `NR` pour le numéro de l'entrée dans le lot ;
* `MA` pour :
  * en affichage `k` (par défaut) et `k:zonesUNM` : la `008` position 1 à 3, précédé d'une astérisque pour une raison justifiée que je n'ai pas pu déterminer ;
  * en affichage `PPN` : le PPN ;
* `V0` pour :
  * en affichage `k` (par défaut) : une adaptation sous forme internationale de la `7X0` (premier auteur) ;
  * en affichage `k:zonesUNM` (quelque soit le nombre de zones renseignées) : l'intégralité des champs UNIMARC spécifiés, séparés par un point virgule ;
  * inexistant en affichage `PPN` ;
* `V1` pour une adaptation au format ISBD de la `200` `$a` à `$e` (titre, exclusif à l'affichage `k`) ;
* `V2` pour la `205 $a` (mention d'édition, exclusif à l'affichage `k`) ;
* `V3` pour une adaptation des `210` ou `214` `$c` (éditeur, exclusif à l'affichage `k`) ;
* `V4` pour la `100 $a` (année, exclusif à l'affichage `k`).

_Note : je pars du principe que la signification des mots est anglaise et non néerlandaise (langue native de WinIBW) car il est bien indiqué `CR` pour_ carriage return _et non son équivalent néerlandais._
