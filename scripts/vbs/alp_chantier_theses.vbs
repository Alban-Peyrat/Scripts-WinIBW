' Scripts pour les chantiers thèses'

Sub perso_CTaddUB700S3()

	addUB700S3
	'Remplace le $btm des exemplaires du RCR ou signale la présence de plusieurs exemplaires dans l'ILN
	changeExAnom
	Ress_goToTag "101", "none", false, true, false

End Sub

Private Sub changeExAnom()

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
		nbOccRCR = Ress_CountOccurrences(notice, "$b" & MY_RCR, true)
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

Sub chantierThese_addCodedData()
	Dim field, u105, ii, jj, skip, ppn, actWinID, authTtl
'Use findTag with occ tio g to the rgiht one
	Ress_toEditMode false, false


'106 : ajoute si absente, ne vérifie pas si elle est correcte sinon
	field = application.activeWindow.title.findTag("106", , , True)
	If field <> Empty Then
		msgbox "Gérer le cas d'une 106 déjà présente"
	Else
		application.activeWindow.title.endOfField	
		application.activeWindow.title.insertText	vblf & "106 ##$ar"
	End If

'105 :
'Force la valeur  $bm$c0$d0$fy
' Si 215$c contient "ill" alors $aa
	u105 = "105 ##$bm$c0$d0$fy"
	' Dans l'idéal faudrait check toutes les 215 :'(
	field = application.activeWindow.title.findTag("215")
	If field <> Empty Then
		For Each subfield in Split(field, "$")
			If Left(subfield, 1) = "c" and InStr(2, subfield, "ill", 1) > 0 Then
				u105 = Left(u105, 6) & "$aa" & Right(u105, 12)
			End If
		Next
	End If

	field = application.activeWindow.title.findTag("105", , , True)
	If field <> Empty Then
		msgbox "Gérer le cas d'une 105 déjà présente"
	Else
		application.activeWindow.title.endOfField
		application.activeWindow.title.insertText	vblf & u105
	End If

'104 :
'Force la valeur $by$cba
	field = application.activeWindow.title.findTag("104", , , True)
	If field <> Empty Then
		msgbox "Gérer le cas d'une 104 déjà présente"
	Else
		application.activeWindow.title.endOfField
		application.activeWindow.title.insertText	vblf & "104 ##$by$cba"
	End If

'100 :
'Si ya une différence 100$a, 214$d, 328$d, définir la valeur selon 328$d

'702 to 701
'Si ya 702 $4727, la passe en 701'
	ii = 0
	Do Until application.activeWindow.title.findTag("702", ii) = Empty
		field = application.activeWindow.title.findTag("702", ii, True, True)
		If InStr(field, "$4727") Then
			application.activeWindow.title.startOfField
			application.activeWindow.title.find "702", False, True
			application.activeWindow.title.InsertText "701"
			ii = ii - 1
		End If
		ii = ii + 1
	Loop

'Si ya pas de 200$g mais une / plusieurs 701$3XXXXXXXX$4727 :
' rechercher dans une nouvelle fenêtre le PPN, récupérer $a + $b
' Puis retourner dans la notice et rajouter à la fin de 200 $gsous la dir [] + $b +a
	field = application.activeWindow.title.findTag("200")
	skip = False
	If field <> Empty Then
		For Each subfield in Split(field, "$")
			If Left(subfield, 1) = "g" Then
				skip = True
			End If
		Next

		If not skip Then
			Dim dirs(4) '1st cell = nom, 2nd prenom 'Ca peut marcher ce tyabkleau ça me gave
			ii = 0
			jj = 0
			Do Until application.activeWindow.title.findTag("701", ii) = Empty
				field = application.activeWindow.title.findTag("701", ii, True, True)
				If InStr(field, "$4727") > 0 and InStr(field, "$3") > 0 Then
					ppn = Mid(field, InStr(field, "$3") + 2, 9)
					dirs(jj) = chantierThese_getDirNames(ppn)
					jj = jj + 1
					'' = Split(test3, chr(29))
					'dirs(UBound(dirs)) <> Empty '
				End If
				ii = ii + 1
			Loop
		End If
	End If


End Sub

sub chantierThese_addDirEstPsdt()
'Ajout une 314 X X est également psdt de jury
	Ress_toEditMode false, false
	
	Application.activeWindow.title.endOfBuffer
	Application.activeWindow.title.insertText	"314 ##$aX est également président de jury" & chr(10)
	
End Sub

Sub chantierThese_addJuryAut()
'crée le squelette d'une notice autorité auteur pour une nouvelle notice
	Dim xlLine, juryFct, juryNom, juryPrenom, jury(99, 2), ii	
	Dim inst, chosenOcc, fct, univ, UA200output

	xlLine = Split(application.activeWindow.clipboard, chr(09))

	juryFct = Split(xlLine(10), ";")
	juryNom = Split(xlLine(8), ";")
	juryPrenom = Split(xlLine(9), ";")
	inst = "Choisir la notice à créer :" & chr(10)
	For ii = 0 to UBound(juryNom)
		jury(ii, 0) = juryNom(ii)
		jury(ii, 1) = juryPrenom(ii)
		jury(ii, 2) = juryFct(ii)
		inst = ress_appendNote(inst, "[" & ii & "] : " & jury(ii, 0) & ", " & jury(ii, 1) & " (" & jury(ii, 2) & ")")
	Next
	
	chosenOcc = InputBox(inst, "Choisir l'auteur :", "99")
	
	If Not (CInt(chosenOcc) > UBound(juryNom) OR CInt(chosenOcc) < 0) Then
		Select Case jury(chosenOcc, 2)
			case "dir"
				fct = "Directeur"
			case "psdt"
				fct = "Président de jury"
			case "mem"
				fct = "Membre du jury"
			case "rapp"
				fct = "Rapporteur"
		End Select
		
		If CInt(xlLine(3)) < 1971 Then
			univ = "l'Université de Bordeaux"
		ElseIf CInt(xlLine(3)) < 2014 Then
			univ = "Bordeaux 2"
		Else
			univ = "Bordeaux"
		End if

		application.activeWindow.Command "cre e"
		
		application.activeWindow.Title.InsertText "008 $aTp5" & vblf &_
			"106 ##$a0$b1$c0" & vblf &_
			"101 ##$afre" & vblf &_
			"102 ##$aFR" & vblf &_
			"103 ##$a19XX" & vblf &_
			"120 ##$a -----À-COMPLÉTER-MANUELLEMENT-----" & vblf &_
			"200 #1$90y$a" & jury(chosenOcc, 0) & "$b" & jury(chosenOcc, 1) & "$f19..-...." & vblf & _
			"340 ##$a" & fct & " d'une thèse de " & xlLine(4) & " soutenue à " & univ & " en " & xlLine(3) & vblf &_
			"340 ##$a -----COMPLÉTER-AVEC-D-AUTRES-INFORMATIONS-DE-LA-PAGE-DE-REMERCIEMENT-PAR-EXEMPLE-----" & vblf & _
			"810 ##$a" & xlLine(7) & " / " & xlLine(6) & " " & xlLine(5) & ", " & xlLine(3) & " [thèse]$b" & jury(chosenOcc, 1) & " " & jury(chosenOcc, 0) & ", " & LCase(fct)

	'Ajoute UA400
		If (InStr(jury(chosenOcc, 0), " ") > 0) OR (InStr(jury(chosenOcc, 0), "-") > 0) Then
		    	addUA400
		End If
	Else
		MsgBox "Numéro choisi invalide"
	End if
	
End Sub

Sub chantierThese_addJuryFromExcel()
	Dim xlLine, juryPPN, juryFct, juryNom, juryPrenom, jury(99, 4), ii
	Dim mentResp, nomDir, dirNoms, nonDirCount
	Dim UB314, UB701S3, UB701a
	Dim temp, output
	
	xlLine = Split(application.activeWindow.clipboard, chr(09))

	juryPPN = Split(xlLine(11), ";")
	juryFct = Split(xlLine(10), ";")
	juryNom = Split(xlLine(8), ";")
	juryPrenom = Split(xlLine(9), ";")
	nonDirCount = 0
	For ii = 0 to UBound(juryPPN)
		jury(ii, 0) = juryPPN(ii)
		jury(ii, 1) = juryFct(ii)
		nonDirCount = nonDirCount + 1
		Select Case jury(ii, 1)
			case "dir"
				nomDir = juryNom(ii)
				nonDirCount = nonDirCount - 1
				dirNoms = ress_appendNote(dirNoms, juryPrenom(ii) & " " & juryNom(ii))
				jury(ii, 2) = 727
			case "psdt"
				jury(ii, 2) = 956
				jury(ii, 3) = "Président"
				jury(ii, 4) = juryPrenom(ii) & " " & juryNom(ii)
			case "mem"
				jury(ii, 2) = 555
				jury(ii, 3) = "Membre"
				jury(ii, 4) = juryPrenom(ii) & " " & juryNom(ii)
			case "rapp"
				jury(ii, 2) = 958
				jury(ii, 3) = "Rapporteur"
				jury(ii, 4) = juryPrenom(ii) & " " & juryNom(ii)
		End Select
	Next
	
	application.activeWindow.command "CHE PPN " & xlLine(2)
	
	mentResp = ress_getTag("200", "1", "f", "all") & " ; " & ress_getTag("200", "1", "g", "all")
	UB314 = ress_getTag("314", "all", "a", "all")
	UB701S3 = ress_getTag("701", "all", "3", "1")
	UB701a = ress_getTag("701", "all", "a", "1")
	UB702S3 = ress_getTag("702", "all", "3", "1")
	UB702a = ress_getTag("702", "all", "a", "1")

	ress_toEditMode false, false

	output = "VÉRIFIER :"
'200
	if InStr(mentResp, nomDir) > 0 Then
		output = ress_appendNote(output, "200 $g")
	End If
	ress_goToTag "200", "none", true, false, false
	application.activeWindow.title.insertText	"$g[sous la direction de] " & replace(dirNoms, chr(10), ", ")
	
'314
	application.activeWindow.title.endOfBuffer
	If UB314 <> "Aucune 314" Then
		output = ress_appendNote(output, "314")
	End if
	If nonDirCount = 1 Then
		application.activeWindow.title.insertText "314 ##$aAutre contribution : "
	ElseIf nonDirCount > 1 Then
		application.activeWindow.title.insertText "314 ##$aAutres contributions : "
	End If
	For ii = 0 to UBound(juryPPN)
		If jury(ii, 1) <> "dir" Then
			temp = ress_appendNote(temp, jury(ii, 4))
			If jury(ii, 1) <> jury(ii + 1, 1) Then
				If InStr(temp, chr(10)) > 0 Then
					If jury(ii, 1) <> "rapp" Then
						temp = replace(temp, chr(10), ", ") & " (" & jury(ii, 3) & "s du jury) ; "
					Else
						temp = replace(temp, chr(10), ", ") & " (" & jury(ii, 3) & "s) ; "
					End If
				Else
					If jury(ii, 1) <> "rapp" Then
						temp = replace(temp, chr(10), ", ") & " (" & jury(ii, 3) & " du jury) ; "
					Else
						temp = replace(temp, chr(10), ", ") & " (" & jury(ii, 3) & ") ; "
					End If
				End If
				application.activeWindow.title.insertText temp
				temp = ""
			End If
		End If
	Next
	application.activeWindow.title.charLeft 3, true
	application.activeWindow.title.deleteSelection
	application.activeWindow.title.insertText vblf

'701
	For ii = 0 to UBound(juryPPN)
		application.activeWindow.title.insertText "701 #1$3" & jury(ii, 0) & "$4" & jury(ii, 2) & vblf
		If InStr(UB701S3, juryPPN(ii)) > 0 OR InStr(UB701a, juryNom(ii)) > 0 Then
			output = Ress_appendNote(output, juryPPN(ii) & " - " & juryNom(ii))
		End If
		If InStr(UB702S3, juryPPN(ii)) > 0 OR InStr(UB702a, juryNom(ii)) > 0 Then
			output = Ress_appendNote(output, juryPPN(ii) & " (en 702) - " & juryNom(ii))
		End If
	Next
	
	if InStr(output, chr(10)) > 0 Then
		MsgBox output
	End If

End Sub

Sub chantierThese_auteurGlobalGet()
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
	discipline = chantierThese_getDiscipline(temp)
	If Left(discipline, 3) = ";_;" Then
		note = Ress_appendNote(note, "Sélectionner manuellement la discipline")
		discipline = Right(discipline, Len(discipline)-3)
	End if

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

Function chantierThese_getDirNames(ppn)
	'Renvoie le nom et le prénom sous forme de tableau'
	'Assume que le PPN est correct'
	Dim actWinID, authTtl, field, nom, prenom

	' Window management'
	actWinID = application.activeWindow.windowID
	application.newWindow
	

	application.activeWindow.command "che ppn " & ppn '045211213'
	authTtl = application.activeWindow.variable("P3CLIP")
	field = Mid(authTtl, InStr(authTtl, chr(13) & "200")+1, InStr(InStr(authTtl, chr(13) & "200")+1, authTtl, chr(13))- InStr(authTtl, chr(13) & "200"))
	For Each subfield in Split(field, "$")
		If Left(subfield, 1) = "a" Then
			nom = Right(subfield, Len(subfield)-1)
		ElseIf Left(subfield, 1) = "b" Then
			prenom = Right(subfield, Len(subfield)-1)
		End If
	Next

	Application.ActivateWindow actWinID
	chantierThese_getDirNames = nom & chr(29) & prenom
End Function

Function chantierThese_getDiscipline(temp)
'Renvoie le code Excel de la discipline de la thèse
	
	Dim discipline

	discipline = Mid(temp, InStr(temp, "$c")+2, InStr(temp, "$e")-InStr(temp, "$c")-2)

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
			discipline = ";_;" & discipline
	End Select

	chantierThese_getDiscipline = discipline

End Function

Sub chantierThese_getJuryForExcel()
	Dim cote, PPNB, annee, discipline, nomAut, prenomAut, titre
	Dim output, boxMsg, exceptions
	
	Perso_collerPPN
	
	PPNB = application.activeWindow.variable("P3GPP")
	annee = Left(ress_getTag("328", "1", "d", "1"), 4)
	discipline = LCase(ress_getTag("328", "1", "c", "1"))
	titre = getTitle
	cote = getCoteEx
	nomAut = ress_getTag("700", "1", "a", "1")
	If InStr(nomAut, "Aucun $a dans cette ") > 0 Then
		prenomAut = ress_getTag("700", "1", "3", "1")
		application.activeWindow.command "che ppn " & prenomAut
		nomAut = ress_getTag("200", "1", "a", "1")
		prenomAut = ress_getTag("200", "1", "b", "1")
		application.activeWindow.command "che ppn " & PPNB
	Else
		prenomAut = ress_getTag("700", "1", "b", "1")
		If UCase(nomAut) = nomAut Then
			nomAut = Ress_uCaseNames(nomAut)
			prenomAut = Ress_uCaseNames(prenomAut)
			boxMsg = ress_appendNote(boxMsg, "Caps lock")
		End If
	End If
	
	'Temporaire
	prenomAut = replace(prenomAut, chr(10), "")
	'Fin du temporaire
	output = cote & chr(09) & PPNB & chr(09) & annee & chr(09) & discipline & chr(09) & nomAut & chr(09) & prenomAut & chr(09) & titre
	
	application.activeWindow.clipboard	= output
	
	exceptions = ress_getTag("200", "1", "f", "all")
	'import de getDataUAChantierThese (27/10/2021)
	If (InStr(exceptions, "ép.") > 0) OR (InStr(exceptions, "épouse") > 0) OR (InStr(exceptions, " fille") > 0) OR (InStr(exceptions, " naissance") > 0) OR (InStr(exceptions, " née") > 0) Then
		boxMsg = Ress_appendNote(boxMsg, "Possiblement un nom d'épouse")
	End If
	'Présence POSSIBLE de deux auteurs
	If InStr(exceptions, " et ") Then
		boxMsg = Ress_appendNote(boxMsg, "Possiblement deux auteurs")
	End If
	
	If boxMsg <> "" Then
		msgbox boxMsg
	End if
End Sub

sub chantierThese_noDirAddPsdt200f
	Ress_toEditMode false, false
	
	ress_goTotag "200", "none", true, true, false
	Application.activeWindow.title.insertText	"$gprésident du jury de soutenance "
End Sub

Sub ChantierThese_AddUB183
'Ajoute une 183 en fonction de la 215 (notamment des chiffres détectés dans le $a)

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

Sub ChantierThese_LoopAddUB183
'Exécute ChantierTheseAddUB183, sauf si l'utilisateur refuse l'ajout, sur la liste de PPN présente dans le presse-papier

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

Sub excelImpAutAuteur()
'Crée une notice d'autorité auteur à partir de la notice dans le presse papier dans le cadre du chantier thèse
'Raccourci : Ctrl+Shift+&
'Requis : decompUA200enUA400
'_A_MOD_

	dim notice, leftStr, rightStrPos, UA200a, UA200b, UA200output
    
With Application.activeWindow

	notice = .clipboard
	
	'Corrige défauts d'imports depuis excel
	leftStr = replace(left(notice, 5), chr(034), "")
	rightStrPos = InStrRev(notice, "106")
	notice = leftStr & mid(notice, 6, rightStrPos) & "$a0$b1$c0"
	
	'Ajoute UA400
	'UA200aPos = InStr(notice, "200 #1$90y$a")
	'UA400 = Mid(notice, UA200aPos)
	'UA200fPos = InStr(UA400, "$f")-1
	'UA400 = Left(UA400, UA200fPos)   
	'UA400Output = decompUA200enUA400(UA400)
	'notice = notice & vblf & UA400Output
	
	.Command "cre e"
	
	.Title.InsertText notice

'Ajoute UA400
	UA200output = findUA200aUA200b
	UA200output = Split(UA200output, ";_;")
	UA200a = UA200output(1)
	If (InStr(UA200a, " ") > 0) OR (InStr(UA200a, "-") > 0) Then
	    	addUA400
	End If
End With

End Sub

Sub Perso_excelImpBibg()
'Crée un notice bibliographique en important depuis Excel

	dim notice, z
    
With Application.activeWindow

	notice = .clipboard

	.Command "cre"
	.simulateIBWKey "FR"
	.title.selectAll	
	.Title.InsertText notice

	Ress_Sleep 1
	
	'Corrige défauts d'imports depuis excel
	.Title.StartOfBuffer
	.Title.CharRight 1, true
	.title.deleteSelection
	.Title.Find "$_#_$"
	.Title.EndOfBuffer true
	.Title.deleteSelection
	
	
	'Perfectionne le titre
	Ress_goToTag "200", "none", false, true, false
	.Title.EndOfField true
	.Title.InsertText purifUB200a(.title.findtag("200"), false) & chr(10)
	.title.Find("541 ##")
	If .title.Selection = "541 ##" Then
		Ress_goToTag "541", "none", false, true, false
		.Title.EndOfField true
		.Title.InsertText purifUB200a(.title.findtag("541"), true) & chr(10)
	End if

	.title.selectall
	.Title.ReplaceAll chr(10) & chr(10), chr(10)
	.Title.StartOfBuffer

    
End With

End Sub