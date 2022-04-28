'Some scripts are missing

Sub executeVBScriptFromName
	' This was created to executed user scripts from standart scripts
	' The shortcut needs to be Shift + Ctrl + Alt + L
	Dim fctName
	fctName = InputBox("Exécuter une fonction VBS", "Écrire le nom de la procédure ou fonction (argument inclus) :")

	If fctName = "" Then
		MsgBox "Aucun script renseigné"
	Else
		Execute fctName
	End If
End Sub
