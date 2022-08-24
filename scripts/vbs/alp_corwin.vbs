' Scripts for CorWin'

Private Sub CorWin_CW1()

'Exporte le $a et le $b de UA103 pour chaque PPN de la liste présente dans le presse-papier.
'Faire les requêtes une par une permet d'éviter la limite de résultats
'Cf Ress_exportVar pour la source de l'export

Dim PPNList, notice, UA103, UA103a, UA103b
Dim storedPPN, output

With Application.activeWindow

	count = 0
	storedPPN = "X"
	PPNList = split(.clipboard, Chr(10))
	
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(PPNlist(0) & "/export_WinIBW.txt",2,true)
	
	PPNList(0) = "XXXXXXXXX"
	
	For each PPN in PPNList
		If PPN <> "XXXXXXXXX" AND PPN <> "" Then
			.command "che PPN " & PPN
			If .variable("P3GPP") <> storedPPN Then
				notice = .copyTitle
				UA103 = Mid(notice, InStr(notice, chr(13) & "103 ##")+1, Len(notice))
				UA103 = Left(UA103, InStr(UA103, chr(13))-1)
				If InStr(UA103, "$a") > 0 Then
					UA103a = Mid(UA103, InStr(UA103, "$a")+2, Len(UA103))
					If InStr(UA103a, "$") > 0 Then
						UA103a = Left(UA103a, InStr(UA103a, "$")-1)
					End If
				Else
					UA103a = "00000000"
				End If
				If InStr(UA103, "$b") > 0 Then
					UA103b = Mid(UA103, InStr(UA103, "$b")+2, Len(UA103))
					If InStr(UA103b, "$") > 0 Then
						UA103b = Left(UA103b, InStr(UA103b, "$")-1)
					End If
				Else
					UA103b = "00000000"
				End If
				storedPPN = Mid(notice, Instr(notice, chr(13) & "008 ")-9, 9)
				output = storedPPN & ";_;" & UA103a & ";_;" & UA103b
			Else
				output = "ERREUR : " & PPN & ";_;00000000;_;00000000"
			End If
			objFileToWrite.WriteLine(output)
		End If
	Next
	
	objFileToWrite.Close
	Set objFileToWrite = Nothing

End With

MsgBox "Traitement sur WinIBW terminé" & vblf & "Vous pouvez lancer l'analyse sur CorWin"
End Sub

Sub CorWin_Launcher()
'Ouvre un input box pour lancer les scripts (add et get), parce que j'ai pas non plus une infinité de touches raccorucis
Dim CWX

CWX = Inputbox("Écrire l'ID du traitement à exécuter"_
	& chr(10) & chr(10) & "[CW1] Vérification format UA103 "_
	, "Lancer un traitement CorWin :", "CW0")
Select Case CWX
	case "CW1"
		CorWin_CW1
	case else
		MsgBox "Aucun traitement correspondant."
End Select

End Sub
