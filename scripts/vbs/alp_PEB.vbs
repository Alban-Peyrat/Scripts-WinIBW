'Scripts pour le PEB
'Scripts for ILL

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
    	msgbox "Le message de cr�ation de demande n'est pas affich�."
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
		num = InputBox("Quel RCR (�crire le num�ro) :"_
			& chr(10)& chr(10) & " - [0] " & VF0 _
			& chr(10) & " - [1] " & VF1, "Quel RCR choisir :", 0)
			
		Select Case num
			Case 0
				application.activeWindow.clipboard = VF0
			Case 1
				application.activeWindow.clipboard = VF1
			Case Else
				MsgBox "Aucun RCR copi�"
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
			MsgBox "Les biblioth�ques ont r�pondu."
			Exit for
		End If
		LRTpos = InStr(fournisseurs(ii), chr(27) & "E" & chr(27) & "LRT") + 6
		comment = Mid(fournisseurs(ii), LRTpos, InStr(LRTpos, fournisseurs(ii), chr(27) & "E") - LRTpos)
		If comment  = "En attente de r�ponse" Then
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

	num = InputBox("�crire le num�ro du script � ex�cuter :"_
		& chr(10) & chr(10) & chr(09) & "R�cup�rer des donn�es :"_
		& chr(10) & "[0] R�cup�rer le num�ro de demande PEB"_
		& chr(10) & "[1] R�cup�rer le num�ro de demande de PEB apr�s validation d'une demande"_
		& chr(10) & "[2] R�cup�rer le PPN"_
		& chr(10) & "[3] R�cup�rer le RCR demandeur"_
		& chr(10) & "[4] R�cup�rer le RCR fournisseur en attente de r�ponse"_
		& chr(10) & "[5] R�cup�rer le titre et l'auteur du document demand�", "Ex�cuter un script de PEB :", 99)
		
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
		case 5
			 AlP_PEBgetTitleAuth
		Case Else
			MsgBox "Aucun script correspondant."
	End Select

End Sub