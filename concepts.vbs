Sub ctrlUA103eqUA200f()
'Exporte et compare le $a de UA103 et le $f de UA200 pour chaque PPN de la liste présente dans le presse-papier.
'Requis : exportVar, getTag (quand implémenter obsvly)
'/!\ EVITER UTILISATION
'/!\ ATTENTION, WinIBW va sembler cesser de fonctionner, laissez-le faire

Dim PPNList, count, notice, UA103a, UA200f
Dim storedPPN, output

With Application.activeWindow

	count = 0
	storedPPN = "X"
	output = "$_#_$ Contrôle équivalence UA103 et UA200$f : " & FormatDateTime(Now) & vblf & "PPN;Result;UA103a;UA200f;Note" & vblf
	PPNList = split(.clipboard, Chr(10))
	
	For each PPN in PPNList
		count = count +1
		.command "che PPN " & PPN
		If .variable("P3GPP") <> storedPPN Then
			notice = .copyTitle
			UA103a = Mid(notice, InStr(notice, chr(13) & "103 ##")+9, 4)
			UA200f = Mid(notice, InStr(notice, chr(13) & "200 ")+6, Len(notice))
			UA200f = Mid(UA200f, InStr(UA200f, "$f") +2, 4)
			If UA103a = UA200f Then
				output = output & PPN & ";OK;" & UA103a & ";" & UA200f & ";#n/a" & vblf
			Else
				output = output & PPN & ";Diff;" & UA103a & ";" & UA200f & ";Pas corr. dates" & vblf
			End If
			storedPPN = Mid(notice, Instr(notice, chr(13) & "008 ")-9, 9)
			 
		Else
			output = output & PPN & ";ERROR;#n/a;#n/a;Recherche non about." & vblf
		End If

	Next
	exportVar output, true

End With

End Sub

Sub ctrlUB700S3()
'Exporte le premier $ de UB700 pour chaque PPN de la liste présente dans le presse-papie
'Requis : exportVar
'/!\ EVITER UTILISATION
'/!\ ATTENTION, WinIBW va sembler cesser de fonctionner, laissez-le faire

Dim PPNList, storedPPN, notice, UB700S_Occ1, count

With Application.activeWindow

	PPNList = Split(.clipboard, chr(10))
	count = 0
	storedPPN = "X"
	output = "$_#_$ Contrôle présence de lien UB700 : " & FormatDateTime(Now) & vblf & "PPN;Result;UB700S_Occ1" & vblf
	
	For Each PPN in PPNList
		count = count+1
		.command "che ppn " & PPN
		If .variable("P3GPP") <> storedPPN Then
			notice = .copyTitle
			UB700S_Occ1 = Mid(notice, InStr(notice, chr(13) & "700 ")+7, 2)
			If UB700S_Occ1 = "$3" Then
				output = output & PPN & ";OK;" & UB700S_Occ1 & vblf
			Else
				output = output & PPN & ";Diff;" & UB700S_Occ1 & vblf
			End If
			storedPPN = Mid(notice, Instr(notice, chr(13) & "008 ")-9, 9)
			 
		Else
			output = output & PPN & ";ERROR;#n/a;Recherche non about." & vblf
		End If
	Next
	exportVar output, true

End With

End Sub
