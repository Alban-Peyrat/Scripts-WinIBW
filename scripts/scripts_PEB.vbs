Sub AlP_PEBgetNumDemande()
	application.activeWindow.clipboard = application.activeWindow.variable("P3GA*")
End Sub

Sub AlP_PEBgetNumDemandePostValidation()
    dim msg
    msg = ""
    On Error Resume Next
    msg = application.activeWindow.messages.item(0).text
    If InStr(msg, "no. ") > 0 Then
    	msg = Mid(msg, InStr(msg, "no. ") + 4, 10)
    	application.activeWindow.clipboard = msg
    Else
    	msgbox "Le message de création de demande n'est pas affiché."
	End If
End Sub

Sub AlP_PEBgetPPN()
	application.activeWindow.clipboard = application.activeWindow.variable("P3VTA")
End Sub

sub AlP_PEBgetRCRDemandeur()
	Dim VF1, VF0, num
	VF1 = application.activeWindow.variable("P3VF1")
	VF0 = application.activeWindow.variable("P3VF0")
	
	if VF0 <> VF1 Then
		num = InputBox("Quel RCR (écrire le numéro) :"_
			& chr(10)& chr(10) & " - [0] " & VF0 _
			& chr(10) & " - [1] " & VF1, "Quel RCR choisir :", 0)
			
		Select Case num
			Case 0
				application.activeWindow.clipboard = VF0
			Case 1
				application.activeWindow.clipboard = VF1
			Case Else
				MsgBox "Aucun RCR copié"
		End Select
	Else
		application.activeWindow.clipboard = VF0
	End If
End Sub

sub AlP_PEBgetRCRFournisseurOnHold()
	Dim fournisseurs, ii, comment, LRTpos
	fournisseurs = Split(application.activeWindow.variable("P3VCA"), chr(13))
	For ii = 0 to UBound(fournisseurs)
		If ii = UBound(fournisseurs) Then
			MsgBox "Les bibliothèques ont répondu."
			Exit for
		End If
		LRTpos = InStr(fournisseurs(ii), chr(27) & "E" & chr(27) & "LRT") + 6
		comment = Mid(fournisseurs(ii), LRTpos, InStr(LRTpos, fournisseurs(ii), chr(27) & "E") - LRTpos)
		If comment  = "En attente de réponse" Then
			application.activeWindow.clipboard = Mid(fournisseurs(ii), InStr(fournisseurs(ii), chr(27) & "E" & chr(27) & "LSS") + 6, 9)
			Exit For
		End If
	Next 
End Sub

Sub AlP_PEBgetTitleAuth()
    dim titre, article, auteur, auteurArt
    titre = application.activeWindow.variable("P3VTC")
    auteur = application.activeWindow.variable("P3VTD")
    article = application.activeWindow.variable("P3VAB")
    auteurArt = application.activeWindow.variable("P3VAA")
    application.activeWindow.clipboard = titre & vblf & auteur & vblf & article & vblf & auteurArt
End Sub

Sub AlP_PEBLauncher()
	dim num

	num = InputBox("Écrire le numéro du script à exécuter :"_
		& chr(10) & chr(10) & chr(09) & "Récupérer des données :"_
		& chr(10) & "[0] Récupérer le numéro de demande PEB"_
		& chr(10) & "[1] Récupérer le numéro de demande de PEB après validation d'une demande"_
		& chr(10) & "[2] Récupérer le PPN"_
		& chr(10) & "[3] Récupérer le RCR demandeur"_
		& chr(10) & "[4] Récupérer le RCR fournisseur en attente de réponse"_
		& chr(10) & "[5] Récupérer le titre et l'auteur du document demandé", "Exécuter un script de PEB :", 99)
		
	Select Case num
		Case 0
			AlP_PEBgetNumDemande
		Case 1
			AlP_PEBgetNumDemandePostValidation
		Case 2
			AlP_PEBgetPPN
		Case 3
			AlP_PEBgetRCRDemandeur
		Case 4
			AlP_PEBgetRCRFournisseurOnHold
		Case 5
			AlP_PEBgetTitleAuth
		Case Else
			MsgBox "Aucun script correspondant."
	End Select

End Sub
