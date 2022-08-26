'Scripts that add text to a title'

Private Sub add18XmonoImp()
'Ajout une 181 txt, 182 n 183 nga pour P01
	Ress_toEditMode false, false
	
	Application.activeWindow.title.endOfBuffer
	Application.activeWindow.title.insertText	"181 ##$P01$ctxt" & chr(10) & "182 ##$P01$cn" & chr(10) & "183 ##$P01$anga" & chr(10)
	
End Sub

Private Sub add18XmonoImpIll()
'Ajout une 181 txt, 182 n 183 nga pour P01
	Ress_toEditMode false, false
	
	Application.activeWindow.title.endOfBuffer
	Application.activeWindow.title.insertText	"181 ##$P01$ctxt" & chr(10) & "181 ##$P02$csti" & chr(10) & "182 ##$P01$P02$cn" & chr(10) & "183 ##$P01$P02$anga" & chr(10)
	
End Sub

Private Sub add214Elsevier()
'Ajoute une 214 type pour Elsevier
	Ress_toEditMode false, false
	
	Application.activeWindow.title.endOfBuffer
	Application.activeWindow.title.insertText	"214 #0$aIssy-les-Moulineaux$cElsevier Masson SAS$dDL 2022" & chr(10)
	
End Sub

Private Sub addAutFromUB()

	Dim nom, prenom, annee, titre

	nom = Inputbox("Nom :")
	prenom = Inputbox("Prénom :")
	titre = getTitle
	annee = ress_getTag("100", "no", "c", "no")
	If InStr(annee, "Aucun") > 0 Then
		annee = ress_getTag("100", "no", "a", "no")
	End If
	
	application.activeWindow.Command "cre e"
	
	application.activeWindow.Title.InsertText "008 $aTp5" & vblf &_
		"106 ##$a0$b1$c0" & vblf &_
		"101 ##$afre" & vblf &_
		"102 ##$aFR" & vblf &_
		"103 ##$a19XX" & vblf &_
		"120 ##$a -----À-COMPLÉTER-MANUELLEMENT-----" & vblf &_
		"200 #1$90y$a" & nom & "$b" & prenom & "$f19..-...." & vblf & _
		"340 ##$a -----COMPLÉTER-AVEC-D-AUTRES-INFORMATIONS-----" & vblf & _
		"810 ##$a" & titre & " / " & prenom & " " & nom & ", " & annee

'Ajoute UA400
	If (InStr(nom, " ") > 0) OR (InStr(nom, "-") > 0) Then
	    	addUA400
	End If

End Sub

Private Sub addBibgFinChap()
	Ress_toEditMode false, false
	Application.activeWindow.title.insertText	"Chaque fin de chapitre comprend une bibliographie"
End Sub

Private Sub addCouvPorte()
	
	Ress_toEditMode false, false
	
	Application.activeWindow.title.endOfBuffer
	Application.activeWindow.title.insertText	"312 ##$aLa couverture porte en plus : """
End Sub

Private Sub addEISBN()
	Dim atPos, title, ISBN
	
	Ress_toEditMode false, false

'Titre
	atPos = InStr(ress_getTag("200", "1", "a", "1"), "@")
	title = getTitle
	If title = "Aucune 200" Then
		title = " -----À-COMPLÉTER-MANUELLEMENT-----"
		atPos = 1
	End If
'ISBN
	ISBN = ress_getTag("010", "1", "A", "1")
	If ISBN = "Aucun $A dans cette 010" Then
		ISBN = ress_getTag("010", "1", "a", "1")
	End If
	If InStr(ISBN, "-") > 0 Then
		ISBN = Left(ISBN, InStrRev(ISBN, "-")-1)
		ISBN = Left(ISBN, InStrRev(ISBN, "-"))
	Else
		ISBN = ""
	End If

'Output
	Application.activeWindow.title.endOfBuffer
	Application.activeWindow.title.insertText	"452 ##$t"& title & "$y" & ISBN
	application.activeWindow.title.startOfField
	application.activeWindow.title.charRight 7 + atPos
	Application.activeWindow.title.insertText "@"
	Application.activeWindow.title.endOfField
	
End Sub

Private Sub addISBNElsevier()
'Ajoute une 010 avec le début de l'ISBN d'Elsevier
	
	Ress_toEditMode false, false
	
	Application.activeWindow.title.endOfBuffer
	Application.activeWindow.title.insertText	"010 ##$A978-2-294-"
End Sub

Private Sub addNoteBonISBN()
' Ajoute une 301 avec comme provenance le service nouveaux éditeurs de la BnF'
	
	Ress_toEditMode false, false
	
	Application.activeWindow.title.endOfBuffer
	Application.activeWindow.title.insertText	"301 ##$aL'ISBN 13 exact provient du service Nouveautés éditeurs de la Bibliothèque nationale de France"
End Sub

Private Sub AddSujetRAMEAU()
'Permet d'ajouter des 606
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

Private Sub addUA400()
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

Private Sub addUB700S3()
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

Private Sub addUB7XX()
	Dim codeFct, autType, WEMI, inds, PPN, temp, ii

	codeFct = inputBox("Code fonction :" & chr(10) & "Ajouter “c” ou “f” en première position pour insérer une 71X ou une 72X"_
	& chr(10) & "Ajouter les indicateurs entre espaces avant le code pour les choisir"_
	& chr(10) & chr(09) & "Valeurs des indicateurs par défaut"_
	& chr(10) & "Personne : #1"_
	& chr(10) & "Collectivité : 02"_
	& chr(10) & "Famille : ##")

	PPN = application.activeWindow.clipboard

'Dizaine
	If Left(codeFct, 1) <> "c" AND Left(codeFct, 1) <> "f" Then 
		autType = 0
	ElseIf Left(codeFct, 1) = "c" Then
		autType = 1
	ElseIf Left(codeFct, 1) = "f" Then
		autType = 2
	End If

'Indicateurs
	temp = Split(codeFct, " ")
	If UBound(temp) = 2 Then
		inds = temp(1)
	Else
		Select Case autType
			Case 0
				inds = "#1"
			Case 1
				inds = "02"
			Case 2
				inds = "##"
		End Select
	End If

'Code fonction
	codeFct = Right(codeFct, 3)
		
'Unité
	Select Case codeFct
		Case "070", "340", "651", "730"
			WEMI = "0"
		Case "555", "727", "956", "958"
			WEMI = "1"
		Case "080", "440"
			WEMI = "2"
		Case Else
			WEMI = " -----COMPLÉTER-MANUELLEMENT-----"
	End Select
	
	If WEMI = "0" Then
		For ii = 0 to 2
			If InStr(ress_getTag("7" & ii & "0", "no", "3", "1"), "Aucun") = 0 Then
				WEMI = 1
				Exit For
			End If
		Next
	'check si ya un 7X0
	End If
	
'Écriture
	ress_toEditMode false, false

	application.activeWindow.Title.endOfBuffer
	application.activeWindow.Title.InsertText "7" & autType	& WEMI & " " & inds & "$3" & PPN & "$4" & codeFct & vblf
End Sub

Sub createItemAvaibleForILL()
' Creates an item with available for ILL
	Dim cote

	If getNoticeType() = 1 Then
		cote = Inputbox("Écrivez la cote du document :", "Créer un exemplaire :", "")
		Application.ActiveWindow.CodedData = false 'Needs to be set false before the command
		' Otherwise it will input "e* $bx" twice if there are no items in the ILN
		Application.ActiveWindow.Command("\INV E*")
		Application.ActiveWindow.Title.EndOfBuffer
		Application.ActiveWindow.Title.InsertText "e* $bx" & vblf & "930 ##$a" & cote & "$ju" 
	Else
		MsgBox "Vous ne pouvez utiliser ce script que sur une notice bibliographique."
	End If

End Sub