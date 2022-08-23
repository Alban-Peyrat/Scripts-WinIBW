'Scripts that get data from a title'

Private Function getCoteEx()
'Requis : Ress_appendNote

dim notice, cote(98, 2), UEa, ans, temp, separateur, occNb, coteDisplay, ii, ansSplit

notice = Application.activeWindow.copyTitle
notice = split(notice, "$b" & MY_RCR)

occNb = -1
For Each occ in notice
	occNb = occNb+1
'Ignore la premi�re occurrence
	If occNb > 0 Then
		' Il m'a fallu du temps mais voil� ce que fait cette ligne du d�mon :
		' D�tecte le num�ro d'exemplaire.
		cote(occNb, 1) = Mid(notice(occNb-1), InStrRev(notice(occNb-1), chr(13) & "e")+1, 3)
		UEa = InStr(occ, "$a")
'D�tecte s'il y a une cote
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

'D�tecte s'il y a plusieurs exemplaires en m�moire
If occNb > 1 Then
'Ne peut pas exc�der 10 cotes diff�rentes atm
	ans = InputBox("Plusieurs cotes pour ce RCR :" & chr(10)_
	& coteDisplay & chr(10) & chr(10)_
	& "Choisissez les num�ro d'occurrences voulues (s�parer les num�ros par _ si n�cessaire, 'all' pour toutes)" & chr(10) & chr(10)_
	& "Saut de ligne comme s�parateur par d�faut, pour en choisir un autre :" & chr(10)_
	& "[$$t] pour une tabulation horizontale" & chr(10)_
	& "[$$;] pour un point-virgule" & chr(10)_
	& "[$$#{votre-choix}] pour un s�parateur personnalis� (sans les {})" & chr(10)_
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
'V�rifie si c'est une occurrence valide
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
'S'il n'y a qu'une seule cote en m�moire
Else
	coteDisplay = cote(1, 2)
End If

getCoteEx = coteDisplay
End Function

Private Function getTitle()
'Renvoie dans le presse papier le titre du document en rempla�ant les @ et $e
'Requis : ress_getTag

	dim UB200, titre, temp

	UB200 = ress_getTag("200", "1", "none", "all")
	If UB200 = "Aucune 200" Then
		output = UB200
	Else
		posA = InStr(UB200,"$a")+2
		posF = InStr(UB200, "$f")
		If posF = 0 Then
			posF = Len(UB200)
		End if
		titre = Mid(UB200, posA, posF-posA)
		titre = replace(titre, "@", "")
		titre = replace(titre, "$e", " : ")
		temp = UCase(titre)
		'Si le titre est uniquement en majuscule, le renovie en minuscule pour modifications
		if titre = temp Then
		     output = Left(titre, 1) & Right(LCase(titre), Len(titre)-1)
		Else
		    output = titre
		End If
	End If
	getTitle = output

End Function

Private Function getUA810b()
'Si un seul UA810 est pr�sent, �crit le $b "n� le" � partir des informations de la 103de la notice
'Si plusieurs UA810 sont pr�sents, renvoie le $b dans le presse-papier
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
	sexe = "$bn�e"
	Else
	sexe = "$bn�"
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

Private Function getUB310()
'Si une 310 est pr�sente, r�cup�re son information
'Requis : Ress_CountOccurrences

	dim z, posUB310
	
	Ress_toEditMode true, false
	
With Application.activeWindow
	
	'R�cup�re la valeur du UB310
	z = .copyTitle
	z = Mid(z, InStr(z, "310 ##$a")+8)
	z = Left(z, InStr(z, chr(13))-1)
	getUB310 = z
End With

End Function