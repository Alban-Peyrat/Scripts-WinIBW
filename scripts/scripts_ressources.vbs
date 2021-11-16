Function Ress_appendNote(var, text)
'Importer de ConStance [01/09/2021]
    If var = "" Then
        var = text
    Else
        var = var & Chr(10) & text
    End If
    Ress_appendNote = var
End Function

Function Ress_CountOccurrences(p_strStringToCheck, p_strSubString, p_boolCaseSensitive)
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
    Ress_CountOccurrences = UBound(arrstrTemp)
End Function

Sub Ress_exportVar(var, boolAppend)
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

	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\/oclcpica/WinIBW30/Profiles/apeyrat001/export.txt",mode,true)
	objFileToWrite.WriteLine(var)
	objFileToWrite.Close
	Set objFileToWrite = Nothing

End Sub

Function Ress_getTag(tag, forceOcc, subTag, forceOccSub)
'Récupère l'information
'forceOcc = "no"-> non ; "last" -> dernier ; "all" -> toutes ; tout le reste un string de nb
' subtag = "none" pour le champ entier
'forceOccSub idem que forceOcc
'Pour forceOcc et forceOccSub, si la valeur est trop grande, prend le dernier champ


	Dim temp, temp2, editMode, notice, occList, chosenTag, chosenSubTag
	Dim allOcc, chosenOcc, occ, ii, nbDollar, dollarOcc(99)

	application.activeWindow.codedData = false

'Détecte si on est en edit mode
'À changer à terme
	On Error Resume Next
	temp = application.activeWindow.Title.canPaste
	if Err then
		editMode = false
	Else
		editMode = true
	End If

	

	If editMode = false Then
		temp2 = application.activeWindow.clipboard
		notice = Application.activeWindow.copyTitle
		application.activeWindow.clipboard = temp2
	Else
		With application.activeWindow.title
			'temp = .startOfField(true)
			'temp = .selection
			'temp = Len(temp)+1
			'temp2 = .currentLineNumber
			.SelectAll
			notice = .Selection
			.selectNone
			.StartOfBuffer
			'.lineDown temp2
			'.charRight temp
			'.InsertText "-----TEST-----"
		End With
		
	End If

	notice = replace(notice, chr(13), chr(10))
	notice = ";_;" & chr(10) & notice

	While InStr(notice, chr(10) & chr(10)) > 0
		notice = replace(notice, chr(10) & chr(10), chr(10))
	Wend
'Récupère le tag
	occList = Split(notice, chr(10) & tag)

	If UBound(occList) = 0 Then
		chosenTag = "Aucune " & tag
	ElseIf UBound(occList) = 1 Then
		chosenOcc = occList(1)
	ElseIf forceOcc = "last" Then
		chosenOcc = occList(UBound(occList))
	ElseIf forceOcc = "all" Then
		for ii = 1 to UBound(occList)
			If InStr(occList(ii), chr(10)) > 0 Then
				chosenTag = chosenTag & ";_;_;" & tag & Left(occList(ii), InStr(occList(ii), chr(10)))
				If Left(chosenTag, 5) = ";_;_;" Then
					chosenTag = Mid(chosenTag, 6, len(chosenTag))
				End If
			Else
				chosenTag = chosenTag & ";_;_;" &  tag & occList(ii)
			End If
		Next
	ElseIf forceOcc = "no" Then 
		for ii = 1 to UBound(occList)
			If ii <> UBound(occList) Then
				allOcc = Ress_appendNote(allOcc, "[" & ii & "] " & tag & occList(ii))
			Else
				If InStr(occList(ii), chr(10)) > 0 Then
					allOcc = Ress_appendNote(allOcc, "[" & ii & "] " & tag & Left(occList(ii), InStr(occList(ii), chr(10))))
				Else
					allOcc = Ress_appendNote(allOcc, "[" & ii & "] " & tag & occList(ii))
				End If
			End If

		Next
		temp = Inputbox(allOcc, "Choisir le numéro de l'occurrence", 1)
		If CInt(temp) > UBound(occList) OR CInt(temp) < 1 Then
			MsgBox "Cette occurrence n'existe pas"
			Exit function
		End If
		chosenOcc = occList(CInt(temp))
	Else
		If Cint(forceOcc) > UBound(occList) Then
			chosenOcc = occList(UBound(occList))
		Else
			chosenOcc = occList(CInt(forceOcc))
		End If
	End If

'Gestion output'
	If UBound(occList) = 0 OR (UBound(occList) > 1 AND forceOcc = "all") Then
		'skip la suite de l'instruction
	ElseIf InStr(chosenOcc, chr(10)) > 0 Then
		chosenTag = tag & Left(chosenOcc, InStr(chosenOcc, chr(10)))
	Else
		chosenTag = tag & Left(chosenOcc, Len(chosenOcc))
	End If

	if subTag <> "none" Then
		temp2 = Split(chosenTag, ";_;_;")
		For each occ in temp2
			chosenSubTag = ""
			occList = Split(occ, "$")
			If UBound(occList) = 0 Then
				chosenSubTag = chosenTag
			ElseIf UBound(occList) = 1 Then
				chosenOcc = occList(1)
				nbDollar = 1
			ElseIf forceOccSub = "last" Then
				chosenOcc = occList(UBound(occList))

			ElseIf forceOccSub = "all" Then
				for ii = 1 to UBound(occList)
					If Left(occList(ii), 1) = subTag Then
						'If InStr(occList(ii), chr(10)) > 0 Then
						'	chosenSubTag = chosenSubTag & ";_#_;" & Mid(occList(ii), 2, InStr(occList(ii), chr(10))-1)
						'Else
							chosenSubTag = chosenSubTag & ";_#_;" & Mid(occList(ii), 2, Len(occList(ii)))
						'End If
					End if
				Next
				If Left(chosenSubTag, 5) = ";_#_;" Then
					chosenSubTag = Mid(chosenSubTag, 6, len(chosenSubTag))
				End If
			ElseIf forceOccSub = "no" Then 
			allOcc = ""
			nbDollar = 0
			Erase dollarOcc
				for ii = 1 to UBound(occList)
					If Left(occList(ii), 1) = subTag Then
						nbDollar = nbDollar + 1
						dollarOcc(nbDollar) = ii
						If ii <> UBound(occList) Then
							allOcc = Ress_appendNote(allOcc, "[" & nbDollar & "] $" & occList(ii))
						Else
							If InStr(occList(ii), chr(10)) > 0 Then
								allOcc = Ress_appendNote(allOcc, "[" & nbDollar & "] $" & Left(occList(ii), InStr(occList(ii), chr(10))-1))
							Else
								allOcc = Ress_appendNote(allOcc, "[" & nbDollar & "] $" & occList(ii))
							End If
						End If
					End If
				Next
				If InStr(allOcc, chr(10)) > 0 Then
					temp = Inputbox(allOcc, "Choisir le numéro de l'occurrence", 1)
					If CInt(temp) > nbDollar OR CInt(temp) < 1 Then
						MsgBox "Cette occurrence n'existe pas"
						Exit function
					End If
					chosenOcc = occList(dollarOcc(CInt(temp)))
				Else
					chosenOcc = occList(dollarOcc(CInt(Mid(allOcc, 2, 1))))
				End If
			Else
				nbDollar = 0
				Erase dollarOcc
				for ii = 1 to UBound(occList)
					If Left(occList(ii), 1) = subTag Then
						nbDollar = nbDollar + 1
						dollarOcc(nbDollar) = ii
					End If
				Next
				If CInt(forceOccSub) > nbDollar Then
					chosenOcc = occList(dollarOcc(nbDollar))
				Else
					chosenOcc = occList(dollarOcc(CInt(forceOccSub)))
				End If
			End If

'Gestion output'
			If UBound(occList) = 0 OR (InStr(chosenSubtag, ";_#_;") > 0 AND forceOccSub = "all") Then
				'skip la suite de l'instruction
			ElseIf nbDollar = 0 Then
				chosenSubTag = "Aucun $" & subtag & " dans cette " & tag
			ElseIf InStr(chosenOcc, chr(10)) > 0 Then
				chosenSubTag = chosenSubTag & Mid(chosenOcc, 2, InStr(chosenOcc, chr(10)))
			Else
				chosenSubTag = chosenSubTag & Mid(chosenOcc, 2, Len(chosenOcc))
			End If
			ress_getTag = ress_getTag & ";_;_;" & chosenSubTag
		Next

		ress_getTag = Mid(ress_getTag, 6, Len(ress_getTag))
	Else
		ress_getTag = chosenTag
	End If
	
End Function

Sub Ress_goToTag(tag, subTag, toEndOfField, toFirst, toLast)
'Place le curseur à l'empalcement indiqué par les paramètres. Si plusieurs occurrences sont rencontrées sans que toFirst ou toLast soit true, une boîte de dialogue s'ouvre pour sélectionner l'occurrence souhaitée
'La fonction countOccurences est nécesssaire ( https://www.thoughtasylum.com/2009/07/30/VB-Script-Count-occurrences-in-a-text-string/ [cons. le 26/04/2021]
'Tag = [str] champ
'subtag = [str] [case sensitive] sous-champ. Ne pasz mettre le $. Si vide, mettre "none"
'toEndOfField = [bool] place le curseur à la fin du champ OU du sous-champ
'toFirst = [bool] si plusieurs occurences du CHAMP, sélectionne le premier prioritaire sur toLast
'toLast = [bool] si plusieurs occurences du CHAMP, sélectionne le dernier
'Requis : Ress_CountOccurrences
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
	nbOcc = Ress_CountOccurrences(notice, chr(10) & tag, false)
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
	    	nbOcc = Ress_CountOccurrences(selectedTag, "$" & subTag, true)
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

Sub Ress_goToTagInputBox()
'Permet d'essayer Ress_goToTag en indiquant les paramètres voulus.
'Requis : Ress_goToTag

	dim z, y, x, w, v
	z = inputbox("tag")
	v = inputbox("subTag", ,"none")
	y = inputbox("toEndOfField", , "false")
	x = inputbox("goFirst", , "false")
	w = inputbox("goLast", , "false")
	y = CBool(y)
	x = CBool(x)
	w = CBool(w)
	Ress_goToTag z, v, y, x, w
	'Ress_goToTag "606", "", true, "", ""
End Sub

Sub Ress_Sleep(time)
'Source : Original Paulie D comment : https://stackoverflow.com/questions/1729075/how-to-set-delay-in-vbscript
'EVITER L'UTILISATION
	Dim dteWait
	time = CInt(time)
	dteWait = DateAdd("s", time, Now())
	Do Until (Now() > dteWait)
	Loop
End Sub

Sub Ress_toEditMode(lgPMode, save)
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

Function Ress_uCaseNames(noms)

Dim kk, jj, sepCheck

noms = UCase(Left(noms, 1)) & LCase(Mid(noms, 2, Len(noms)))

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

Ress_uCaseNames = noms

End Function
