Sub add18XmonoImp()
'Ajout une 181 txt, 182 n 183 nga pour P01
	Ress_toEditMode false, false
	
	Application.activeWindow.title.endOfBuffer
	Application.activeWindow.title.insertText	"181 ##$P01$ctxt" & chr(10) & "182 ##$P01$cn" & chr(10) & "183 ##$P01$anga" & chr(10)
	
End Sub

Sub add214Elsevier()
'Ajoute une 214 type pour Elsevier
	
	Ress_toEditMode false, false
	
	Application.activeWindow.title.endOfBuffer
	Application.activeWindow.title.insertText	"214 #0$aIssy-les-Moulineaux$cElsevier Masson SAS$dDL 2021" & chr(10)
	
End Sub

Sub addBibgFinChap()
	Ress_toEditMode false, false
	Application.activeWindow.title.insertText	"Chaque fin de chapitre comprend une bibliographie"
End Sub

Sub addCouvPorte()
	
	Ress_toEditMode false, false
	
	Application.activeWindow.title.endOfBuffer
	Application.activeWindow.title.insertText	"312 ##$aLa couverture porte en plus : """
End Sub

Sub addISBNElsevier()
'Ajoute une 010 avec le début de l'ISBN d'Elsevier
	
	Ress_toEditMode false, false
	
	Application.activeWindow.title.endOfBuffer
	Application.activeWindow.title.insertText	"010 ##$A978-2-294-"
End Sub

Sub AddSujetRAMEAU()
'Permet d'ajouter des 606
'Raccourci : Ctrl Shift (
'Requis : rien
	dim PPN, UB606, inds, Xvalue, PPNclean
	
	Ress_toEditMode false, false
	
With Application.activeWindow

	For ii = 0 To 999
		inds = "##"
		PPN = Inputbox("Écrire le PPN à ajouter (en 606 (##) par défaut)"_
		& chr(10) & chr(10) & "-> $3{PPN} pour ajouter une subdivision au précédent"_
		& chr(10) & "-> Écrire '_{ind.1}{ind.2}' après le PPN SANS ESPACE pour changer les indicateurs par défaut"_
		& chr(10) & chr(10) & "NB : les indicateurs par défaut sont ceux entre parenthèses"_
		& chr(10) & chr(10) & "U0{PPN} pour ajouter 600 (#1)"_
		& chr(10) & "U1{PPN} pour ajouter 601 (02)"_
		& chr(10) & "U2{PPN} pour ajouter 602"_
		& chr(10) & "U4{PPN} pour ajouter 604"_
		& chr(10) & "U5{PPN} pour ajouter 605"_
		& chr(10) & "U7{PPN} pour ajouter 607"_
		& chr(10) & "U8{PPN} pour ajouter 608"_
		, "Ajouter une 60X:", "ok")
		PPN = Replace(PPN, "PPN ", "")
		PPN = Replace(PPN, "(PPN)", "")
		PPN = Replace(PPN, " ", "")
		PPN = Replace(PPN, chr(10), "")
		PPN = Replace(PPN, chr(13), "")
		
		.Title.EndOfbuffer
		If PPN = "ok" Then
			.Title.InsertText UB606
			Exit For
		Else
			If Left(Right(PPN, 3), 1) = "_" Then
				inds = Right(PPN, 2)
			End If
			If Left(PPN, 2) = "$3" Then
				UB606 = Left(UB606, Len(UB606)-9) & PPN & right(UB606, 9)
			Else
				.Title.InsertText UB606
				PPNclean = Mid(PPN, 3, 9)
				Select Case UCase(Left(PPN, 2))
					Case "U0" Xvalue = "0"
						If inds = "##" Then
							inds = "#1"
						End If
					Case "U1" Xvalue = "1"
						If inds = "##" Then
							inds = "02"
						End If
					Case "U2" Xvalue = "2"
					Case "U4" Xvalue = "4"
					Case "U5" Xvalue = "5"
					Case "U7" Xvalue = "7"
					Case "U8" Xvalue = "8"
					Case Else Xvalue = "6"
						PPNclean = Left(PPN, 9)
				End Select
				UB606 = "60" & Xvalue & " " & inds & "$3" & PPNclean & "$2rameau" & chr(10)
			End If
		End If
	Next
	
End With

End Sub

Sub addUA400()
'Ajoute un/des champs 400 à une notice d'autorité auteur
'Requis : decompUA200enUA400, Ress_toEditMode
'PAS UNIVERSEL. Fonctionne uniquement s'il y a un $a et un $b au moins

    dim UA200, UA200a, UA200b, UA200fPos, UA400, temp
    
    Ress_toEditMode false, false
    
With Application.activeWindow.Title
	
	temp = findUA200aUA200b
	temp = Split(temp, ";_;")
	UA200 = temp(0)
	UA200a = temp(1)
	UA200b = temp (2)
	UA200fPos = temp(3)

	UA400 = decompUA200enUA400(UA200a, UA200b)
	
	.endofbuffer
	
'Ajoute une 400 à modifier si decompUA200enUA400 n'a pas renvoyé de 400
	If Len(UA400) < 5 Then
		UA400 = Left(UA200, UA200fPos)
		If Right(UA400, 1) = "$" Then
			UA400 = Left(UA400, Len(UA400)-1)
		End If
		UA400 = replace(UA400, "200", "400")
		UA400 = replace(UA400, "$90y", "")
		
		.InsertText vblf & UA400
		.startoffield
		.CharRight 8
	Else
		.InsertText vblf & UA400
	End If
    
End With

End Sub

Sub addUB700S3()
'Remplace la 700 actuelle de la notice bibliographique par une 700 contenant le PPN du presse-papier et le $4 de l'ancienne 700
'Requis : Ress_toEditMode
	
	dim UB700
	
	Ress_toEditMode false, false

With Application.ActiveWindow.Title
	
	.Find(chr(10) & "700 ")
	.EndOfField
	.CharLeft 3, true
	UB700 = "700 #1$3" & application.activeWindow.clipboard & "$4" & .selection
	UB700 = replace(UB700, chr(10), "")
	.deleteLine
	
	.InsertText UB700 & vblf
    
End With
    
End Sub

Sub changeExAnom()

Dim notice, nbOcc, nbOccRCR, exSB

With Application.activeWindow.Title
	.SelectAll
	notice = .selection

	nbOcc = Ress_CountOccurrences(notice, chr(10) & "e", true)
	If nbOcc = 1 Then
		Ress_goToTag "930", "none", false, false, false
		.charLeft(1)
		.charLeft 2, true
		If LCase(.selection) = "tm" Then
			.InsertText "x"
			exSB = .tag
			MsgBox exSB & " : tm remplacé par x"
		End If
	ElseIf nbOcc > 1 Then
		nbOccRCR = Ress_CountOccurrences(notice, "$b330632101", true)
		If nbOccRCR > 1 Then
			.Find("$btm" & chr(10) & "930 ")
			exSB = .tag
			if Left(exSB, 1) = "e" Then
				MsgBox exSB & " à supprimer", , "Exemplaire fictif"
			Else
				MsgBox "Plusieurs exemplaires réels sur ce RCR." & chr(10) &  chr(10) & "Vérification recommandée."
			End If
		Else
			MsgBox "Plusieurs exemplaires réels." & chr(10) & chr(10) & "Vérification recommandée."
		End If
	End If
End With
End Sub

Sub ChantierTheseAddUB183
'Ajoute une 183 en fonction de la 215 (notamment des chiffres détectés dans le $a) dans le cadre du chantier thèse
'Raccourci : Texte only
'Requis : Ress_goToTag, Ress_toEditMode
'_A_MOD_

	dim UB215, z, pages, numPages
	dim y(99)
	dim notice, nbSP, output, nbVblf, count
	
	notice = application.activeWindow.copyTitle
	Ress_toEditMode false, false

With Application.activeWindow.title
	
		'Détermine le $a à ecrire
		UB215 = .FindTag("215")
		z = split(UB215,"$")
		for each x in z
			if Left(x, 1) = "a" Then
				pages = x
			End If
		next
		If InStr(pages, "vol") <> 0 Then
			pages = Right(pages, Len(pages) - InStr(pages, "vol")-2)
		End If
		For i = 0 to Len(pages)
			y(i) = Mid(pages, i+1, 1)
			If isNumeric(y(i)) = true Then
				numPages = numPages & y(i)
			End If
		Next
		
	'determine le nb de $P
		'if Ress_CountOccurrences(notice, "181 ##", true) >= Ress_CountOccurrences(notice, "181 ##", true) Then
		'	nbSP = Ress_CountOccurrences(notice, "181 ##", true)
		'Else
		'	nbSP = Ress_CountOccurrences(notice, "182 ##", true)
		'End If
		'If nbSP = 1 Then
		'	output = "181 ##$P01"
		'Else 
	'Tant que je modifie pas
		output = "183 ##$P01"		



		if numPages < 49 Then
			output = output & "$angb"
		Else
			output = output & "$anga"
		End If
		
		Ress_goToTag "200", "none", false, true, false
		.InsertText output & vblf
		Ress_goToTag "215", "none", true, true, false
		
End With

End Sub

Sub ChantierTheseLoopAddUB183
'Exécute ChantierTheseAddUB183, sauf si l'utilisateur refuse l'ajout, sur la liste de PPN présente dans le presse-papier
'Raccourci : texte only
'Requis : ChantierTheseAddUB183, Ress_exportVar

	dim output, PPNList, statut, ListeStatuts, wrongPPN, count
	
	ListeStatuts = "ok" & chr(10) & "pb" & chr(10) & "no p" & chr(10) & "d f" & chr(10) & "$$stop"
	count = 0
	output = "$_#_$ Chantier thèse ajout UB183 : " & FormatDateTime(Now) & vblf & "PPN;Statut" & vblf
	
With Application.activeWindow

	PPNList = split(.clipboard, Chr(10))
	
	For each PPN in PPNList
		count = count +1
		wrongPPN = false
    		.command "che ppn " & PPN
    		If .Messages.Count > 0 Then
	    		If .messages.Item(0).Text = "PPN erroné" Then
	    			MsgBox "PPN erroné"
	    			wrongPPN = true
	    			statut = "PPN erroné"
	    		End If
	    	End if
	    	
	    	If wrongPPN = false Then
	    	
			chantierTheseAddUB183
			
	    		statut = Inputbox(.title.findtag("215") & chr(10) & ListeStatuts, "Définir le statut (PPN n°"&count&":", "ok")
	    		
	    		If statut = "ok" Then
				.SimulateIBWKey "FR"
			Else
				Select Case statut
					Case "pb" statut = "Problème"
					Case "no p" statut = "Pas de pagination"
					Case "d f" statut = "Déjà fait"
					Case "$$stop" statut = "Arrêt forcé"
					Case Else statut = "Statut invalide"
				End Select
				.SimulateIBWKey "FE"
				.SimulateIBWKey "FR"
	    		End If
	    	
	    	End If

    		output = output & PPN & ";" & statut & chr(10)
    		If Fix(count/10) = count/10 Then
    			output = Left(output, Len(output)-1)
    			Ress_exportVar output, true
    			output = ""
    		End If
    		If statut = "Arrêt forcé" Then
    			Exit For
    		End If
	Next
    
End With

	Ress_exportVar output, true

End Sub

Function decompUA200enUA400(UA200a, UA200b)
'Renvoi les champs 400 créés à partir de la décomposition du nom composé du champ 200 importé
'Requis : RIEN

	dim separateur
	
	While (InStr(UA200a, " ") <> 0) OR (InStr(UA200a, "-") <> 0)
'Détermine le séparateur
		If (InStr(UA200a, " ") > 0) AND (InStr(UA200a, "-") = 0 OR (InStr(UA200a, " ") < InStr(UA200a, "-"))) Then
			separateur = InStr(UA200a, " ")
		ElseIf (InStr(UA200a, "-") > 0) AND (InStr(UA200a, "0") = 0 OR (InStr(UA200a, " ") > InStr(UA200a, "-"))) Then
			separateur = InStr(UA200a, "-")
		End If
	
'Construit la nouvelle forme
		If (Right(UA200b, 1) = "-") OR (Right(UA200b, 1) = "'") Then
			UA200b = RTrim(UA200b & Left(UA200a, separateur))
		Else
			UA200b = RTrim(UA200b & " " & Left(UA200a, separateur))
		End If
		UA200a = Right(UA200a, Abs(separateur-Len(UA200a)))
				
'Rejet du "de"
		If Left(UA200a, 3) = "de " Then
			UA200a = Mid(UA200a, 4, Len(UA200a))
			If Right(UA200b, 1) = "-" OR Right(UA200b, 1) = "'" Then
				UA200b = UA200b & "de"
			Else
				UA200b = UA200b & " de"
			End If
		End If
'Rejet du "d'"
		If Left(UA200a, 2) = "d'" Then
			UA200a = Mid(UA200a, 3, Len(UA200a))
			If Right(UA200b, 1) = "-" OR Right(UA200b, 1) = "'" Then
				UA200b = UA200b & "d'"
			Else
				UA200b = UA200b & " d'"
			End If
		End If
		
'Ajout à la notice
		decompUA200enUA400 = Ress_appendNote(decompUA200enUA400, "400 #1$a" & UA200a & "$b" & UA200b)
	Wend

End Function

Function findUA200aUA200b()
'Identifie la position du $a et du $b dans la 200UA. Doit être appelé depuis écran de modification
	Dim UA200, UA200fPos, UA200a, UA200b, ii

	UA200 = Application.activeWindow.Title.FindTag ("200")
	UA200fPos = 0
	ii = 0
	While UA200fPos = 0
		Select Case ii
			case 0
				UA200fPos = inStr(UA200, "$f")
			case 1
				UA200fPos = inStr(UA200, "$c")

			case 2
				UA200fPos = inStr(UA200, "$x")
			case 3
				UA200fPos = inStr(UA200, "$y")
			case 4
				UA200fPos = inStr(UA200, "$z")
			case Else
				UA200fPos = Len(UA200) + 1
		End Select
		ii = ii +1
	Wend

	UA200a = Mid(UA200, InStr(UA200, "$a")+2, InStr(UA200, "$b") - InStr(UA200, "$a")-2)
	UA200b = Mid(UA200, InStr(UA200, "$b")+2, UA200fPos - InStr(UA200, "$b")-2)

	findUA200aUA200b = UA200 & ";_;" & UA200a & ";_;" & UA200b & ";_;" & UA200fPos
End Function

Sub generalLauncher()
'Ouvre un input box pour lancer les scripts (add et get)
Dim num

num = Inputbox("Écrire le numéro du script à exécuter"_
	& chr(10) & chr(10) & chr(09) & "Notices bibg :"_
	& chr(10) & "[14] Ajouter 18X mongraphie imprimée"_
	& chr(10) & "[1] Ajouter couverture porte"_
	& chr(10) & "[2] Ajouter bibg en fin de chapitre"_
	& chr(10) & "[3] Ajouter e-ISBN"_
	& chr(10) & "[4] Ajouter sujet RAMEAU"_
	& chr(10) & "[15] Ajouter 700 $3"_
	& chr(10)& chr(10) & chr(09) & "Elsevier"_
	& chr(10) & "[6] Ajouter ISBN Elsevier"_
	& chr(10) & "[7] Ajouter 214 Elsevier"_
	& chr(10)& chr(10) & chr(09) & "Récupérer informations"_
	& chr(10) & "[8] Récupérer le titre"_
	& chr(10) & "[9] Récupérer la cote"_
	& chr(10)& chr(10) & chr(09) & "Thèses"_
	& chr(10) & "[10] Récupérer les données chantier autorités"_
	& chr(10) & "[5] Ajouter 700 $3 & vérif. ex."_
	& chr(10) & "[11] Récupérer la note disponibilité (310)"_
	& chr(10) & chr(10) & chr(09) & "Notices autorité :"_
	& chr(10) & "[12] Ajouter 400"_
	& chr(10) & "[13] Récupérer 810 $b date de naissance"_
	& chr(10) & chr(10) & chr(09) & "[77] Lanceur de CorWin"_
	, "Exécuter un script :", 99)
Select Case num
	case 14
		add18XmonoImp
	case 1
		addCouvPorte
	case 2
		addBibgFinChap
	case 3
		addEISBN
	case 4
		AddSujetRAMEAU
	case 5
		perso_CTaddUB700S3
	case 6
		addISBNElsevier
	case 7
		add214Elsevier
	case 8
		application.activeWindow.clipboard	= getTitle
	case 9
		application.activeWindow.clipboard	= getCoteEx
	case 10
		getDataUAChantierThese
	case 11
		application.activeWindow.clipboard	= getUB310
	case 12
		addUA400
	case 13
		application.activeWindow.clipboard	= getUA810b
	case 15
		addUB700S3
	case 77
		CorWin_Launcher
	case else
		MsgBox "Aucun script correspondant."
End Select

End Sub

Function getCoteEx()
'Renvoie dans le presse-papier la cote du document pour ce RCR (malfonctionne s'il y a plusieurs exemplaires de ce RCR)
'PEUT-ÊTRE je ferai une option pour choisir des cotes spécifiques si j'ai le temps parce que ça m'a l'air compliqué encore
'Raccourci : Ctrl+Shift+D
'Requis : Ress_appendNote

dim notice, cote(98, 2), UEa, ans, temp, separateur, occNb, coteDisplay, ii, ansSplit

notice = Application.activeWindow.copyTitle
notice = split(notice, "$b330632101")

occNb = -1
For Each occ in notice
	occNb = occNb+1
'Ignore la première occurrence
	If occNb > 0 Then
		cote(occNb, 1) = Mid(notice(occNb-1), Instr(notice(occNb-1), chr(13) & "e")+1, 3)
		UEa = InStr(occ, "$a")
'Détecte s'il y a une cote
		If (UEa > 0) AND (UEa < InStr(occ, "A98 ")) Then
'Isole la cote
			occ = Mid(occ, InStr(occ, "$a")+2, len(occ))
			If InStr(occ, "$") < InStr(occ, chr(13)) Then
				cote(occNb, 2) = Mid(occ, 1, InStr(occ, "$")-1)
			Else
				cote(occNb, 2) = Mid(occ, 1, InStr(occ, chr(13))-1)
			End If
		Else
			cote(occNb, 2) = "[Exemplaire sans cote]"
		End If
	coteDisplay = Ress_appendNote(coteDisplay, "[Occ. " & occNb & "] " & cote(occNb, 1) & " : " & cote(occNb, 2))
	End If
Next

'Détecte s'il y a plusieurs exemplaires en mémoire
If occNb > 1 Then
'Ne peut pas excéder 10 cotes différentes atm
	ans = InputBox("Plusieurs cotes pour ce RCR :" & chr(10)_
	& coteDisplay & chr(10) & chr(10)_
	& "Choisissez les numéro d'occurrences voulues (séparer les numéros par _ si nécessaire, 'all' pour toutes)" & chr(10) & chr(10)_
	& "Saut de ligne comme séparateur par défaut, pour en choisir un autre :" & chr(10)_
	& "[$$t] pour une tabulation horizontale" & chr(10)_
	& "[$$;] pour un point-virgule" & chr(10)_
	& "[$$#{votre-choix}] pour un séparateur personnalisé (sans les {})" & chr(10)_
	, "Choisir la cote :", "1")
	coteDisplay = ""
'Cotes individuelles
	If InStr(ans, "all") = 0 Then
		ans = "_" & ans
		ansSplit = Split(ans, "_")
		For each chosenOcc in ansSplit
			If chosenOcc <> "" Then
				If InStr(chosenOcc, "$$") = 0 Then
					temp = chosenOcc
				Else
					temp = Left(chosenOcc, InStr(chosenOcc, "$$")-1)
				End If
'Vérifie si c'est une occurrence valide
				If isNumeric(temp) = true Then
					If (CInt(temp) < occNb+1) AND (CInt(temp) > 0) Then
						coteDisplay = Ress_appendNote(coteDisplay, cote(temp, 2))
					Else
						coteDisplay = Ress_appendNote(coteDisplay, "[Occ. choisie (" & temp &") invalide]")
					End if
				Else
					coteDisplay = Ress_appendNote(coteDisplay, "[" & temp & " n'est pas une occ.]")
				End If
			End If
		Next
'Toutes les cotes
	Else
		For ii = 1 to occNb
			coteDisplay = Ress_appendNote(coteDisplay, cote(ii, 2))
		Next
	End If
	separateur = InStr(ans, "$$")
	If separateur > 0 Then
		separateur = Mid(ans, InStr(ans, "$$")+2, len(ans))
		Select Case Left(separateur, 1)
			case "t"
				coteDisplay = replace(coteDisplay, chr(10), chr(09))
			case ";"
				coteDisplay = replace(coteDisplay, chr(10), ";")
			case "#"
				coteDisplay = replace(coteDisplay, chr(10), Right(separateur, len(separateur)-1))
		End Select
	End If
'S'il n'y a qu'une seule cote en mémoire
Else
	coteDisplay = cote(1, 2)
End If

getCoteEx = coteDisplay
End Function

Sub getDataUAChantierThese()
'Génère le squelette de la notice d'autorité à partir de la notice bibliographique (DANS LE CADRE DU CHANTIER)
'Raccourci : Ctrl Shift J
'Requis : Ress_appendNote, Ress_uCaseNames

	dim PPN_B, notice
	dim year, discipline, nom, prenom, bday, titre, sexe, cote, note
	dim theseData(10, 2)
	dim temp, tableau(999), ii, capsLock, output, ansSplit, jj, sepCheck, kk
	
	capsLock = false
	notice = Application.activeWindow.copyTitle
	
'Déjà une UB700S3
	temp = Mid(notice, InStr(notice, chr(13) & "700")+1, len(notice))
	If Mid(temp, InStr(temp, "$")+1, 1) = "3" Then
		MsgBox "Déjà fait"
		Application.activeWindow.Clipboard = "Déjà fait"
		Exit Sub
	End If
	
'Gestion PPN + 328
	PPN_B = Application.activeWindow.variable("P3GPP")
	temp = Mid(notice, InStr(notice, "328 #"), Len(notice))
	year = Mid(temp, InStr(temp, "$d")+2, 4)
	discipline = Mid(temp, InStr(temp, "$c")+2, InStr(temp, "$e")-InStr(temp, "$c")-2)
'Adaptation de la discipline à mon fichier
'Peut-être jsute limiter la transformation à méd, méd gé, pharma
	Select Case discipline
		Case "Sciences de la vie"
			discipline = "1 - sciences de la vie"
		Case "Médecine générale"
			discipline = "2 - médecine générale"
		Case "Pharmacie"
			discipline = "3 - pharmacie"
		Case "Médecine"
			discipline = "5 - médecine"
		Case "Sciences biologiques et médicales. Biologie-Santé"
			discipline = "6 - biologie - santé"
		Case "Sciences biologiques et médicales. Neurosciences et neuropharmacologie"
			discipline = "7 - neurosciences et neuropharmacologie"
		Case "Sciences biologiques et médicales"
			discipline = "8 - sciences biologiques et médicales"
		Case "Sciences odontologiques"
			discipline = "9 - sciences odontologiques"
		Case "Sciences biologiques et médicales. Epidémiologie et intervention en santé publique"
			discipline = "4 - épidémiologie et intervention en santé publique"
		Case "Sciences biologiques et médicales. Sciences pharmaceutiques"
			discipline = "A - sciences pharmaceutiques"
		Case Else
			note = Ress_appendNote(note, "Sélectionner manuellement la discipline")
	End Select
	
	'Gestion du nom
	temp = Mid(notice, InStr(notice, "700 #"), Len(notice))
	nom = Mid(temp, InStr(temp, "$a")+2, InStr(temp, "$b")-InStr(temp, "$a")-2)
	If UCase(nom) = nom Then
		nom = Ress_uCaseNames(nom)
		capsLock = true
	End If
	
	'Gestion de la bday
	prenom = Mid(temp, InStr(temp, "$b")+2, InStr(temp, "$4")-InStr(temp, "$b")-2)
	bday = ""
	For ii = 0 to Len(prenom)
		tableau(ii) = Mid(prenom, ii+1, 1)
		If isNumeric(tableau(ii)) = true Then
			bday = bday & tableau(ii)
		End If
	Next
	
	'Gestion du prénom
	If InStr(prenom, "$f") > 0 Then
		prenom = Left(prenom, InStr(prenom, "$f")-1)
	End If
	If UCase(prenom) = prenom Then
		prenom = Ress_uCaseNames(prenom)
		capsLock = "-----> CAPS LOCK <-----"
	End If
	
	'Gestion titre + cote
	titre = getTitle
	cote = getCoteEx
		
'Gestion de la note
'UB101 <> fre
	temp = Mid(notice, InStr(notice, chr(13) & "101")+1, len(notice))
	temp = Mid(temp, InStr(temp, "$a"), InStr(temp, chr(13)) - InStr(temp, "$a"))
	If temp <> "$afre" Then
		note = Ress_appendNote(note, "/!\ 101 " & temp)
	End If
'UB102 <> FR
	temp = Mid(notice, InStr(notice, chr(13) & "102")+1, len(notice))
	temp = Mid(temp, InStr(temp, "$a"), InStr(temp, chr(13)) - InStr(temp, "$a"))
	If temp <> "$aFR" Then
		note = Ress_appendNote(note,  "/!\ 102 " & temp)
	End If
'Présence POSSIBLE de nom d'épouse / jeune fille
	temp = Mid(notice, InStr(notice, chr(13) & "200")+1, len(notice))
	temp = Mid(temp, InStr(temp, "$f")+2, InStr(temp, chr(13)) - InStr(temp, "$f")-1)
	If InStr(temp, "$") <> 0 Then
		temp = Left(temp, InStr(temp, "$")-1)
	End If
	If (InStr(temp, "ép.") > 0) OR (InStr(temp, "épouse") > 0) OR (InStr(temp, " fille") > 0) OR (InStr(temp, " naissance") > 0) OR (InStr(temp, " née") > 0) Then
		note = Ress_appendNote(note, "Possiblement un nom d'épouse")
	End If
'Présence POSSIBLE de deux auteurs
	If InStr(temp, " et ") Then
		note = Ress_appendNote(note, "Possiblement deux auteurs")
	End If
	
	'Détermine le sexe + si la cote à un pb)
	sexe = InputBox ("[$$d]     Discipline : " & discipline & chr(10)_
		& "[$$y]     An : " & year & chr(10) & chr(10)_
		& "[$$n$_] Nom : " & nom  & chr(10)_
		& "[$$p$_] Prénom : " & prenom  & chr(10)_
		& "[$$w]     Naissance : " & bday & chr(10) & chr(10)_
		& "[$$t$_] Titre : " & titre & chr(10) & chr(10)_
		& "[$$z]     Cote : " & cote & chr(10) & chr(10)_
		& "Majuscule verrouillée : " & capsLock & chr(10) & chr(10)_
		& "Notes : " & note & chr(10) & chr(10)_
		& "Pour réécrire manuellement un champ, ajouter $${lettre du champ}{nouvelle information} collé au reste de l'input."& chr(10) & chr(10)_
		& "Pour modifier un champ, ajouter $${lettre du champ}$_ collé au reste de l'input, ce qui affichera une nouvelle boîte de dialogue."& chr(10),_
		"Choisir le sexe :", "u")
	sexe = "_" & sexe & "$$"
'Gestion des changements manuels
	ansSplit = Split(sexe, "$$")
	sexe = Mid(ansSplit(0), 2, 1)
	If (sexe <> "a") AND (sexe <> "b") AND (sexe <> "u") Then
		sexe = "u"
	End If
	For Each occ in ansSplit
		Select Case Left(occ, 1)
			case "y"
				year = Right(occ, Len(occ)-1)
			case "d"
				discipline = Right(occ, Len(occ)-1)
			case "n"
				If Left(occ, 3) = "n$_" Then
					nom = Inputbox("Entrer le nouveau nom : " & nom, "Modifier le nom :", nom)
				Else
					nom = Right(occ, Len(occ)-1)
				End If
			case "p"
				If Left(occ, 3) = "p$_" Then
					prenom = Inputbox("Entrer le nouveau prénom : " & prenom, "Modifier le prénom :", prenom)
				Else
					prenom = Right(occ, Len(occ)-1)
				End If
			case "w"
				bday = Right(occ, Len(occ)-1)
			case "t"
				If Left(occ, 3) = "t$_" Then
					titre = Inputbox("Entrer le nouveau titre : " & titre, "Modifier le titre :", titre)
				Else
				titre = Right(occ, Len(occ)-1)
				End If
			case "z"
				cote = Right(occ, Len(occ)-1)
		End Select
	Next
	
	
	note = Replace(note, chr(10), " ; ")
	
	output = PPN_B & chr(09) & year & chr(09) & discipline & chr(09) & nom & chr(09) & prenom & chr(09) & bday & chr(09) & sexe & chr(09) & titre & chr(09) & chr(09) & chr(09) & chr(09) & note & chr(09) & cote
	Application.activeWindow.clipboard = output
End Sub

Function getTitle()
'Renvoie dans le presse papier le titre du document en remplaçant les @ et $e
'Raccourci : Ctrl Shift Q
'Requis : RIEN
'_A_MOD_

	dim z, y, x, i, posUB200, posUB2XX, posA, posF
    
With Application.activeWindow

	z = .copyTitle
	'Trouve le prochain champ pour délimiter la 200
	posUB200 = InStr(z, chr(13) & "200 ")
	i = 201
	posUB2xx = 0
	While posUB2XX = 0
	    x = chr(13) & i & " "
	    posUB2XX = InStr(z, x)
	    i = i + 1
	Wend
	
	z = Mid(z, posUB200, posUB2xx-posUB200)
	posA = InStr(z,"$a")+2
	posF = InStrRev(z, "$f")
	.clipboard = z
	z = Mid(z, posA, posF-posA)
	z = replace(z, "@", "")
	z = replace(z, "$e", " : ")
	y = UCase(z)
	'Si le titre est uniquement en majuscule, le renovie en minuscule pour modifications
	if z = y Then
	     output = Left(z, 1) & Right(LCase(z), Len(z)-1)
	Else
	    output = z
	End If
	getTitle = output

End With

End Function

Function getUA810b()
'Si un seul UA810 est présent, écrit le $b "né le" à partir des informations de la 103de la notice
'Si plusieurs UA810 sont présents, renvoie le $b dans le presse-papier
'Raccourci : Ctrl+Shift+G
'Requis : Ress_CountOccurrences, Ress_goToTag, Ress_toEditMode

	dim z, date, sexe, notice
	
	Ress_toEditMode false, false
	
With Application.activeWindow.title
	
	.selectAll
	.copy
	notice = Application.activeWindow.Clipboard
	
	'Construit le $b
	z = .FindTag ("103")
	z = Right(z, 8)
	sexe = .FindTag ("120")
	if Right(sexe, 1) = "a" Then
	sexe = "$bnée"
	Else
	sexe = "$bné"
	End If
	date = sexe & " le " & Right(z, 2) & "-" & Mid(z, 5,2) & "-" & Left(z, 4)
	
	'Compte le nombre de UA810 pour coller OU mettre dans presse papier
	If Ress_CountOccurrences(notice, "810 ##", false) = 1 Then
	 Ress_goToTag "810", "none", true, true, false
	 .insertText date
	Else
		  .selectNone
	    getUA810b = date
	End If
    
End With

End Function

Function getUB310()
'Si une 310 est présente, récupère son information
'Raccourci : Ctrl+Shift++
'Requis : Ress_CountOccurrences

	dim z, posUB310
	
	Ress_toEditMode true, false
	
With Application.activeWindow
	
	'Récupère la valeur du UB310
	z = .copyTitle
	z = Mid(z, InStr(z, "310 ##$a")+8)
	z = Left(z, InStr(z, chr(13))-1)
	getUB310 = z
End With

End Function

Function PurifUB200a(UB200, isUB541)
'Requis : none
'_A_MOD_ -> mieux handle la provenance

	dim UB200a, UB200aPos, UB200fPos
	UB200aPos = InStr(UB200, "$a")+2
	If isUB541 = false Then
		UB200fPos = InStr(UB200, "$f")
	Else
		UB200fPos = InStr(UB200, "$z")
	End If
	UB200a = Mid(UB200, UB200aPos, UB200fPos - UB200aPos)
	UB200 = Replace(UB200, UB200a, "")
	UB200a = replace(UB200a, " : ", "$e")
	UB200a = replace(UB200a, ": ", "$e")
	'Ajoute le @
	If Left(UB200a, 6)="De la " Then
		UB200a = Left(UB200a, 6) & "@" & Mid(UB200a, 7, Len(UB200a))
	ElseIf Left(UB200a, 5)="De l'" Then
		UB200a = Left(UB200a, 5) & "@" & Mid(UB200a, 6, Len(UB200a))
	ElseIf Left(UB200a, 4)="Les "_
	OR Left(UB200a, 4)="Des "_
	OR Left(UB200a, 4)="Une "_
	OR Left(UB200a, 4)="The " Then
		UB200a = Left(UB200a, 4) & "@" & Mid(UB200a, 5, Len(UB200a))
	ElseIf Left(UB200a, 3)="Le "_
	OR Left(UB200a, 3)="La "_
	OR Left(UB200a, 3)="Un "_
	OR Left(UB200a, 3)="An "_
	OR Left(UB200a, 3)="De "_
	OR Left(UB200a, 3)="Du " Then
		UB200a = Left(UB200a, 3) & "@" & Mid(UB200a, 4, Len(UB200a))
	ElseIf Left(UB200a, 2)="A "_
	OR Left(UB200a, 2)="L'"_
	OR Left(UB200a, 2)="D'"  Then
		UB200a = Left(UB200a, 2) & "@" & Mid(UB200a, 3, Len(UB200a))
	Else
		UB200a = "@" & UB200a
	End If
	PurifUB200a = Left(UB200, UB200aPos-1) & UB200a & Mid(UB200, UB200aPos, Len(UB200))
	
End Function

Sub searchExcelPPNList()
'Recherche la liste de PPN contenu dans le presse-papier
'Raccourci : texte only
'Requis : RIEN

    Dim query
    
With Application.activeWindow
    
    query = "che ppn " & replace(replace(.Clipboard, "(PPN)", ""), Chr(10), " OR ")
    query = Left(query, Len(query)-4)
    .Clipboard = query
    .Command query
    
End With

End Sub
