# Notices UNIMARC dans WinIBW à partir des dépôts des thèses d’exercice

*L'ébauche de code peut être trouvé [en cliquant sur ce lien](../../../../WinIBW/tree/main/scripts/vbs/alp_theses2win.vbs).*

Ce projet fait suite au [projet de générateur de notice UNIMARC d'une thèse d'exercice déposée sur DUMAS créé en fin 2021](../../../../ub-svs/blob/main/dumas/poc-cat-DUMAS-WinIBW.md).
Il est à noté que le projet était pensé spécifiquement pour la Bibliothèque universitaire des Sciences du Vivant et de la Santé - Josy Reiffers, même si j'aurais préféré le rendre plus facilement ouvert à d'autres bibliothèques.

Parmi les défauts de la précédente version, deux étaient notables :
* il était obligatoire de passer par mon site web pour manuellement récupérer la notice plutôt que de passer uniquement par WinIBW,
* les thèses déposées dans OSKAR ("archive ouverte institutionnelle des partenaires du site universitaire de Bordeaux") n’étaient pas concernées.

Cette nouvelle version répond à ces deux défauts, en abandonnant les essais d’automatisation d’ajout des autorités si une correspondance assez importante était établie.
Toutefois, cette nouvelle version est __théorique et non codée__.
Les essais techniques ont eu lieu, il est donc techniquement possible de la programmer si voulu, mais je ne l'ai pas développé.

## Fonctionnement

Cette nouvelle version s’utilise intégralement sur WinIBW, plus besoin de passer par mon site web pour générer la notice.

Au démarrage, le script demande à l’usager de renseigner dans une boîte de dialogue l’URL du document.
À partir de cette information, le script déterminera s’il doit procéder pour DUMAS ou pour OSKAR.
Pour ce qui concerne DUMAS, le script a recours à l’export de données au format TEI d’une notice d’un document et à[ l’API de Recherche d’HAL](https://api.archives-ouvertes.fr/docs/search).
Pour ce qui concerne OSKAR, faute de moyens techniques plus satisfaisant, le script à recours à la page HTML contenant l’intégralité des métadonnées.


Le script se connecte à ces services, en extrait les données qu’il souhaite, puis crée une notice bibliographique en générant les champs selon ce qu’il a récupéré.
(Pour créer la notice, il utilise la commande `sys 1; fic 1; cre`)

Comme pour la version précédente, certains champs contiennent un code de trois lettres entre le numéro du champ et les indicateurs, qui empêchent la validation de la notice sans supprimer ces codes :
* `MOD` : pour modifier, les métadonnées ne permettent pas de remplir précisément ce champ (`029, 320 et 303` __(uniquement pour OSKAR)__)
* `VER` : pour vérifier, les métadonnées générées dans ce champ doivent être vérifiées car elles peuvent être mal générées dû à la complexité des modifications à apporter (`200`)
* `DEL` : pour supprimer, le champ est temporaire et sert uniquement à afficher sur la notice des informations (`610`)

## DUMAS vers UNIMARC

### Récupération des données

#### Export TEI de la notice

* `Text/body/listBibl/biblFull/titleStmt/title` :
  * `[xml:lang]` = langue du titre
  * Contenu = titre
* `Text/body/listBibl/biblFull/titleStmt/author[role="aut"]` :
  * `/forename` = prénom
  * `/surname` = patronyme
* `Text/body/listBibl/biblFull/editionStmt/edition[type="current"]` :
  * `/date[type="whenSubmitted"]` = date de dépôt (pour l’embargo)
  * `/date[type="whenEndEmbargoed"]` = date de fin d’embargo
* `Text/body/listBibl/biblFull/publicationStmt/idno[type="halUri"]` = URI
* `Text/body/listBibl/biblFull/notesStmt/note[type="degree"][n]` = identifiant du type de diplôme
* `Text/body/listBibl/biblFull/sourceDesc/biblStruct/monogr/imprint` :
  * `/biblScope[unite="pp"]` = nombre de pages
  * `/date[type="dateDefended"]` = date de soutenance
* `Text/body/listBibl/biblFull/sourceDesc/biblStruct/monogr/authority[type="supervisor"]` = directeurs de thèse
* `Text/body/listBibl/biblFull/profileDesc/langUsage/language[ident]` = langue du document
* `Text/body/listBibl/biblFull/profileDesc/textClass/keywords/term[xml:lang="fr"]` = mots-clefs (français uniquement)
* `Text/body/listBibl/biblFull/profileDesc/abstract` :
  * `[xml:lang]` = langue du résumé
  * `/p` = résumé

#### API de recherche

* `Response/result/doc/arr[name="dumas_degreeSpeciality_s"]/str` = identifiant de la spécialité

### Notice UNIMARC

* 008 : `008 $aOax3`
* 029 : `029 MOD ##$aFR$e` + l'année + `BORD` + `M`, `3`, `P` ou `O` selon le type de mémoire et la spécialité + `XXX` à modifier manuellement en le numéro de la thèse
* 100 : `100 0#$a` + l'année
* 101 : `101 0#$a`, `$c`, `$d` avec les langues du document et les langues des résumés
* 102 : `102 ##$aFR`
* 104 : `104 ##$ak$by$cy$dba$e0$ffre`
* 105 : `105 ##$ay$bm$ba$c0$d0$e0$fy$gy`
* 135 : `135 ##$ad$br`
* 181 : `181 ##$P01$ctxt`
* 182 : `182 ##$P01$cc`
* 183 : `183 ##$P01$aceb`
* 200 : `200 VER 1#$a` + le titre dont la langue correspond à la première langue renvoyée avec les `:` remplacés en `$e` et le `@` placé en début de titre sauf si présence d'un article rejeté, auquel cas, il est placé après l'article + `$f` + les prénoms suivi d'un espace suivi du noms des auteurs renvoyés, séparés par des virgules + `$gsous la direction de` + les directeurs de thèses renvoyés
* 214 : `214 #1$a` + l'année
* 230 : `230 ##$aDonnées textuelles`
* 303 : `303 ##$aL'impression du document génère` + le nombre de page entouré d'espace + `f.`
* 304 : `304 ##$aTitre provenant de l'écran-titre`
* 320 : `320 MOD ##$aBibliogr. XX réf. Annexes` à modifier manuellement le nombre de références et supprimer les annexes s'il n'y en a pas
* 328 : `328 #0$bThèse d'exercice$c` + domaine + spécialité si indiquée + `$eBordeaux$d` + l'année
* 330 : `330 ##$a` + résumé (pour chaque résumé entré)
* 337 : `337 ##$aConfiguration requise : un logiciel capable de lire un fichier au format : application/pdf`
* 371 : `371 0#$aThèse sous embargo jusqu'au` JOUR mois ANNÉE (si la date de fin d’embargo est différente de la date de publication)
* 541 : `541 ##$a` + les autres titres renvoyés (même changements que celui en `200`) + `$z` + la langue correspondante
* 610 : `610 DEL 0#$a` + mot-clef (pour chaque mot-clef en français). Substitut temporaire à la `606`
* 608 : `608 ##$3027253139$2rameau`
* 700 : `700 #1$a` + nom de l'auteur + `$b` + prénoms de l'auteur + `$4070` (pour chaque auteur renvoyé (`701` pour les auteurs numéro 2 et plus)
* 701 : `701 #1$a` + nom du directeur de thèse + `$b` + prénoms du directeur de thèse + `$4727` (pour chaque directeur de thèse renvoyé)
* 711 : `711 02$3175206562$4295`
* 856 : `856 4#$qPDF$u` + l'URI du document renvoyé + `$2Accès au texte intégral`
* Exemplaire : `e* $bx`

## OSKAR vers UNIMARC

### Récupération des données OSKAR

* `dc.contributor.advisor` = directeurs de thèse
* `dc.contributor.author` = auteurs
* `dc.date` = date de publication
* `dc.identifier.uri` = URI
* `dc.description.abstract` = résumé français
* `dc.description.abstractEn` = résumé anglais
* `dc.language.iso` = langue du document
* `dc.subject` (`dc.subject.en` est ignoré) = mots-clefs
* `dc.title` = titre français
* `dc.title.en` (ou autre langue) = titre langue étrangère
* `bordeaux.thesis.type` = domaine de la thèse
* `bordeaux.thesis.discipline` = spécialité de la thèse

### Notice UNIMARC

* 008 : `008 $aOax3`
* 029 : `029 MOD ##$aFR$e` + l'année + `BORD` + `M`, `3`, `P` ou `O` selon le type de mémoire et la spécialité + `XXX` à modifier manuellement en le numéro de la thèse
* 100 : `100 0#$a` + l'année
* 101 : `101 0#$a`, `$c`, `$d` avec la langue du document et les langues des résumés
* 102 : `102 ##$aFR`
* 104 : `104 ##$ak$by$cy$dba$e0$ffre`
* 105 : `105 ##$ay$bm$ba$c0$d0$e0$fy$gy`
* 135 : `135 ##$ad$br`
* 181 : `181 ##$P01$ctxt`
* 182 : `182 ##$P01$cc`
* 183 : `183 ##$P01$aceb`
* 200 : `200 VER 1#$a` + le titre dont la langue correspond à la langue renvoyée avec les `:` remplacés en `$e` et le `@` placé en début de titre sauf si présence d'un article rejeté, auquel cas, il est placé après l'article + `$f` + les prénoms suivi d'un espace suivi du noms des auteurs renvoyés, séparés par des virgules + `$gsous la direction de` + les directeurs de thèses renvoyés
* 214 : `214 #1$a` + l'année
* 230 : `230 ##$aDonnées textuelles`
* 303 : `303 MOD ##$aL'impression du document génère XX f.` à modifier manuellement
* 304 : `304 ##$aTitre provenant de l'écran-titre`
* 320 : `320 MOD ##$aBibliogr. XX réf. Annexes` à modifier manuellement le nombre de références et supprimer les annexes s'il n'y en a pas
* 328 : `328 #0$bThèse d'exercice$c` + domaine + spécialité si indiquée + `$eBordeaux$d` + l'année
* 330 : `330 ##$a` + résumé (pour chaque résumé entrée)
* 337 : `337 ##$aConfiguration requise : un logiciel capable de lire un fichier au format : application/pdf`
* 371 : `371 0#$L'accès à la ressource est réservé aux usagers de la communauté universitaire de Bordeaux`
* 541 : `541 ##$a` + les autres titres renvoyés (même changements que celui en `200`) + `$z` + la langue correspondante
* 610 : `610 DEL 0#$a` + mot-clef (pour chaque mot-clef en français). Substitut temporaire à la `606`
* 608 : `608 ##$3027253139$2rameau`
* 700 : `700 #1$a` + nom de l'auteur + `$b` + prénoms de l'auteur + `$4070` (pour chaque auteur renvoyé (`701` pour les auteurs numéro 2 et plus)
* 701 : `701 #1$a` + nom du directeur de thèse + `$b` + prénoms du directeur de thèse + `$4727` (pour chaque directeur de thèse renvoyé)
* 711 : `711 02$3175206562$4295`
* Exemplaire : `e* $bx`
* E856 : `E856 4#$qPDF$u` + l'URI du document renvoyé + `$zAccès réservé à la communauté universitaire`
