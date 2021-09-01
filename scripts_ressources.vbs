Function appendNote(var, text)
'Importer de ConStance [01/09/2021]
    If var = "" Then
        var = text
    Else
        var = var & Chr(10) & text
    End If
    appendNote = var
End Function

Function CountOccurrences(p_strStringToCheck, p_strSubString, p_boolCaseSensitive)
'Renvoie le nombre d'occurrences
'Source : https://www.thoughtasylum.com/2009/07/30/VB-Script-Count-occurrences-in-a-text-string/ [cons. le 26/04/2021]
'Requis : RIEN

    Dim arrstrTemp
    Dim strBase, strToFind

    If p_boolCaseSensitive Then
        strBase = p_strStringToCheck
        strToFind = p_strSubString
    Else
        strBase = LCase(p_strStringToCheck)
        strToFind = LCase(p_strSubString)
    End If

    arrstrTemp = Split(strBase, strToFind)
    CountOccurrences = UBound(arrstrTemp)
End Function

Sub exportVar(var, boolAppend)
'Exporte dans export.txt (même emplacement que winibw.vbs)
'Source : eddiejackson.net/wp/?p=8619
'Notes
'OpenTextFile parameters:
'IOMode
'1=Read
'2=write
'8=Append
'Create (true,false)
'Format (-2=System Default,-1=Unicode,0=ASCII)
'J'ai rajouté la var mode pour sélectionner entre append et write
'Requis : RIEN

	dim mode
	
	If boolAppend = true Then
		mode = 8
	Else
		mode = 2
	End If

	'À la BU
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\/oclcpica/WinIBW30/Profiles/apeyrat001/export.txt",mode,true)
	'En TVW
	'Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\/oclcpica/WinIBW30/Profiles/utilisateur/export.txt",mode,true)
	objFileToWrite.WriteLine(var)
	objFileToWrite.Close
	Set objFileToWrite = Nothing

End Sub

Sub goToTag(tag, subTag, toEndOfField, toFirst, toLast)
'Place le curseur à l'empalcement indiqué par les paramètres. Si plusieurs occurrences sont rencontrées sans que toFirst ou toLast soit true, une boîte de dialogue s'ouvre pour sélectionner l'occurrence souhaitée
'La fonction countOccurences est nécesssaire ( https://www.thoughtasylum.com/2009/07/30/VB-Script-Count-occurrences-in-a-text-string/ [cons. le 26/04/2021]
'Tag = [str] champ
'subtag = [str] [case sensitive] sous-champ. Ne pasz mettre le $. Si vide, mettre "none"
'toEndOfField = [bool] place le curseur à la fin du champ OU du sous-champ
'toFirst = [bool] si plusieurs occurences du CHAMP, sélectionne le premier prioritaire sur toLast
'toLast = [bool] si plusieurs occurences du CHAMP, sélectionne le dernier
'Requis : countOccurrences
'_A_MOD_

	Dim notice, nbVblf, nbOcc, occurrences, choseOcc, count
	Dim selectedTag, nbDollar, fromDollToEnd, nextDollar
	Dim clipboardSave    

With Application.activeWindow

	clipboardSave = .clipboard
	
	.title.selectAll
	.title.copy
	notice = .clipboard

End With

	choseOcc = false
	count = 1
	tag = CStr(tag)
	nbOcc = countOccurrences(notice, chr(10) & tag, false)
	If nbOcc > 1 Then
	  If toFirst = true Then
	  	count = 1
		ElseIf toLast = true Then
			count = nbOcc
		ElseIf toFirst = false AND toLast = false Then
			choseOcc = true
		End If
	ElseIf nbOcc = 0 Then
			MsgBox "Le champ " & tag & " n'a pas été trouvé dans la notice"
		Exit Sub
	End If
	
	'Old way
	nbVblf = split(notice, Chr(10))
	for each x in nbVblf
	    If Left(x, 3) = tag Then
	    	If choseOcc = false Then
	        	Exit For
	        Else
	        	occurrences = occurrences & vblf & count & " : " & x
	        	count = count + 1
	        End If
	    End If
	Next
	
	If choseOcc = true Then
		count = inputBox(occurrences, "Choisir le numéro de l'occurence")
		If isNumeric(count) = false Then
			MsgBox "Erreur. Choisir un NOMBRE. Relancer le script."
			Exit Sub
	  End If
	End If

With Application.activeWindow.Title

	    .find(chr(10) & tag)
	    If count > 1 Then
	    	for i = 2 to count
	    		.endOfField
	    		.LineDown(1)
	    	next
	    Else
	    	'.LineDown(1)
		.EndOfField
	    End If
	    
	    'Gestion du $
	    If subTag = "none" Then
		    If toEndOfField = true Then
		    	.EndOfField
		    Else
		    	.StartOfField
		    End If
	    Else
	      .StartOfField
	    	selectedTag = .currentField
	    	occurrences = ""
	      count = 0
	    	nbOcc = countOccurrences(selectedTag, "$" & subTag, true)
	    	If nbOcc = 0 Then
	    		MsgBox "Erreur. Pas de $" & subTag & " dans l'occurrence sélectionnée."
		      If toEndOfField = true Then
		      	.EndOfField
		      Else
		      	.StartOfField
		      End If
	    	ElseIf nbOcc = 1 Then
	      	.Find "$" & subTag, true, true
	      	.charRight(1)
	      	.charLeft(1)
		      If toEndOfField = true Then
		      	fromDollToEnd = Mid(selectedTag, InStr(selectedTag, "$" & subTag)+2, Len(selectedTag))
		      	nextDollar = InStr(fromDollToEnd, "$")
		      	If nextDollar = 0 Then
		      		.EndOfField
		      	Else
		      		.charRight(nextDollar-1)
		      	End If
		      End If
	      Else
		    'Si plusieurs occurrences du $   
		    nbDollar = split(selectedTag, "$")
		    for each x in nbDollar
		        If Left(x, 1) = subTag Then
		        	occurrences = occurrences & vblf & count & " : " & x
		        End If
		        count = count + 1
		    Next
		    count = inputBox(occurrences, "Choisir le numéro de l'occurence")
		    If isNumeric(count) = false Then
			MsgBox "Erreur. Choisir un NOMBRE. Première occurrence sélectionnée."
		  	.Find "$" & subTag, true, true
		  	.charRight(1)
		  	.charLeft(1)
		    End If
		  	.startOfField
			for i = 0 to count-1
				.charRight(Len(nbDollar(i))+1)
			next
			.charRight(1)
		      If toEndOfField = true Then
		      	.charRight(Len(nbDollar(count))-1)
		      End If
	    	End If
	    End If
End With
	    
    Application.activeWindow.clipboard = clipboardSave
    
End Sub

Sub goToTagInputBox()
'Permet d'essayer goToTag en indiquant les paramètres voulus.
'Requis : goToTag

	dim z, y, x, w, v
	z = inputbox("tag")
	v = inputbox("subTag", ,"none")
	y = inputbox("toEndOfField", , "false")
	x = inputbox("goFirst", , "false")
	w = inputbox("goLast", , "false")
	y = CBool(y)
	x = CBool(x)
	w = CBool(w)
	goToTag z, v, y, x, w
	'goToTag "606", "", true, "", ""
End Sub

Sub Sleep(time)
'Source : Original Paulie D comment : https://stackoverflow.com/questions/1729075/how-to-set-delay-in-vbscript
'EVITER L'UTILISATION
	Dim dteWait
	time = CInt(time)
	dteWait = DateAdd("s", time, Now())
	Do Until (Now() > dteWait)
	Loop
End Sub

Sub toEditMode(lgPMode, save)
'Passe en mode édition (ou présentation)
'lgPMode [bool] : true = passer en mode présentation
'save [bool] : si lgPMode =true, alors sauvegarder les changements ou non
'Requis : RIEN

dim z, editMode

With Application.activeWindow

	On Error Resume Next
	z = .title.canPaste

	if Err then
		editMode = false
	Else
		editMode = true
	End If

	If lgPMode = false Then
		If editMode = false Then
			.command "mod"
		End If
	Else
		If editMode = false Then
			If save = true Then
				.SimulateIBWKey "FR"
			Else
				.SimulateIBWKey "FE"
				.SimulateIBWKey "FR"
			End If
		End If
	End If
End With
End Sub

Function uCaseNames(noms)

Dim kk, jj, sepCheck

noms = Left(noms, 1) & LCase(Mid(noms, 2, Len(noms)))
For kk = 0 to 3
	Select Case kk
		Case 0
			sepCheck = " "
		Case 1
			sepCheck = "-"
		Case 2
			sepCheck = "'"
	End Select
	jj = 1
	While jj <> 0
		jj = InStr(jj+1, noms, sepCheck)
		On Error Resume Next
		noms = Left(noms, jj) & UCase(Mid(noms, jj+1, 1)) & Right(noms, Len(noms)-jj-1)
	Wend
Next

noms = Replace(noms, " De ", " de ", 1, -1, 0)
noms = Replace(noms, " D'", " d'", 1, -1, 0)

uCaseNames = noms

End Function
