Sub AlP_PEB_getNumDemande()
	application.activeWindow.clipboard = application.activeWindow.variable("P3GA*")
End Sub

Sub AlP_PEB_getNumDemandePostValidation()
         dim msg
    msg = application.activeWindow.messages.item(0).text
    msg = Mid(msg, InStr(msg, "no. ") + 4, 10)
    application.activeWindow.clipboard = msg
End Sub

sub AlP_PEB_getRCRDemandeur()
	application.activeWindow.clipboard = application.activeWindow.variable("libID")
End Sub

Sub AlP_PEB_Launcher()
	dim num

	num = InputBox("Écrire le numéro du script à exécuter :"_
		& chr(10) & chr(10) & chr(09) & "Récupérer des données :"_
		& chr(10) & "[0] Récupérer le numéro de demande PEB"_
		& chr(10) & "[1] Récupérer le numéro de demande de PEB après validation d'une demande"_
		& chr(10) & "[3] Récupérer le RCR demandeur", "Exécuter un script de PEB :", 99)
		
'	num = InputBox("Écrire le numéro du script à exécuter :"_
'		& chr(10) & chr(10) & chr(09) & "Récupérer des données :"_
'		& chr(10) & "[0] Récupérer le numéro de demande PEB"_
'		& chr(10) & "[1] Récupérer le numéro de demande de PEB après validation d'une demande"_
'		& chr(10) & "[2] Récupérer le PPN"_
'		& chr(10) & "[3] Récupérer le RCR demandeur"_
'		& chr(10) & "[4] Récupérer le RCR fournisseur en attente de réponse"_
'		, "Exécuter un script de PEB :", 99)
		
	Select Case num
		Case 0
			AlP_PEB_getNumDemande
		Case 1
			AlP_PEB_getNumDemandePostValidation
		'Case 2
		'	AlP_PEB_getPPN
		Case 3
			AlP_PEB_getRCRDemandeur
		'Case 4
			'AlP_PEB_getRCRFournisseurOnHold
		Case Else
			MsgBox "Aucun script correspondant."
	End Select

End Sub