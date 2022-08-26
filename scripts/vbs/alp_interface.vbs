'Scripts qui intéragissent avec l'interface

Sub collerPPN()
'Recherche le PPN contenu dans le presse papier
'Raccourci : Ctrl+Shift+V
     
With Application.activeWindow

    If Left(.Clipboard,3) = "PPN" Then
        .Command "che ppn " & Right(.Clipboard,9)
    Else
         .Command "che ppn " & .Clipboard
    End If
    
End With

End Sub

Sub generalLauncher()
'Ouvre un input box pour lancer les scripts
Dim num

num = Inputbox("Écrire le numéro du script à exécuter"_
	& chr(10) & chr(10) & chr(09) & "Général :"_
	& chr(10) & "[18] Rechercher le doublon possible"_
	& chr(10) & chr(10) & chr(09) & "Notices bibg :"_
	& chr(10) & "[14] Ajouter 18X mongraphie imprimée"_
	& chr(10) & "[19] Ajouter 18X mongraphie imprimée illustrée"_
	& chr(10) & "[25] Ajouter note de provenance du bon ISBN"_
	& chr(10) & "[1] Ajouter couverture porte"_
	& chr(10) & "[2] Ajouter bibg en fin de chapitre"_
	& chr(10) & "[3] Ajouter e-ISBN"_
	& chr(10) & "[4] Ajouter sujet RAMEAU"_
	& chr(10) & "[15] Ajouter 700 $3"_
	& chr(10) & "[17] Ajouter une autorité auteur"_
	& chr(10)& chr(10) & chr(09) & "Elsevier"_
	& chr(10) & "[6] Ajouter ISBN Elsevier"_
	& chr(10) & "[7] Ajouter 214 Elsevier"_
	& chr(10)& chr(10) & chr(09) & "Récupérer informations"_
	& chr(10) & "[8] Récupérer le titre"_
	& chr(10) & "[9] Récupérer la cote"_
	& chr(10)& chr(10) & chr(09) & "Thèses"_
	& chr(10) & "[10] Récupérer les données chantier autorités"_
	& chr(10) & "[5] Ajouter 700 $3 & vérif. ex."_
	& chr(10) & "[11] Récupérer la note disponibilité (310)"_
	& chr(10) & "[20] (jury) Récupérer les données"_
	& chr(10) & "[21] (jury) Rajouter informations à la notice"_
	& chr(10) & "[22] (jury) Créer autorité"_
	& chr(10) & "[23] (jury) Ajouter une 314 pour un directeur président de jury"_
	& chr(10) & "[24] (jury) Ajouter 200$g pour président de jury"_
	& chr(10) & chr(10) & chr(09) & "Notices autorité :"_
	& chr(10) & "[16] Créer une notice d'autorité auteur pour cette notice"_
	& chr(10) & "[12] Ajouter 400"_
	& chr(10) & "[13] Récupérer 810 $b date de naissance"_
	& chr(10) & chr(10) & chr(09) & "[77] Lanceur de CorWin"_
	& chr(10) & chr(10) & chr(09) & "[88] Lanceur PEB"_
	, "Exécuter un script :", 99)
Select Case num
	case 1
		addCouvPorte
	case 2
		addBibgFinChap
	case 3
		addEISBN
	case 4
		AddSujetRAMEAU
	case 5
		perso_CTaddUB700S3
	case 6
		addISBNElsevier
	case 7
		add214Elsevier
	case 8
		application.activeWindow.clipboard	= getTitle
	case 9
		application.activeWindow.clipboard	= getCoteEx
	case 10
		chantierThese_auteurGlobalGet
	case 11
		application.activeWindow.clipboard	= getUB310
	case 12
		addUA400
	case 13
		application.activeWindow.clipboard	= getUA810b
	case 14
		add18XmonoImp
	case 15
		'addUB700S3
		perso_CTaddUB700S3 'pour enlever les tm
	case 16
		addAutFromUB
	case 17
		addUB7XX
	case 18
		searchDoublonPossible
	case 19
		add18XmonoImpIll
	case 20
		chantierThese_getJuryForExcel
	case 21
		chantierThese_addJuryFromExcel
	case 22
		chantierThese_addJuryAut
	case 23
		chantierThese_addDirEstPsdt
	case 24
		chantierThese_noDirAddPsdt200f
	case 25
		addNoteBonISBN
	case 77
		CorWin_Launcher
	case 88
		AlP_PEBLauncher
	case else
		MsgBox "Aucun script correspondant."
End Select
End Sub

Sub goToWorkRecord()
'Ouvre la page de l'oeuvre associée document. S'il n'y a pas de 579, affiche un message

	Dim field, fields, ii
	field = application.activeWindow.variable("P3CLIP")
	fields = Split(field, chr(13))

	For ii = 1 to UBound(fields)
		If Left(fields(ii), 3) = "579" Then
			application.activeWindow.command "che ppn " + Mid(fields(ii), InStr(fields(ii), "$3")+2, 11)
			Exit Sub
		End If
	Next
	MsgBox "Pas de notice d'oeuvre liée à cette notice bibliographique (absence de 579)."
End Sub

Sub lastCHE()
'Raccourci : Ctrl+Shift+Ret.Arriere

    Application.ActiveWindow.Command "HIS"

End Sub

Sub searchDoublonPossible()
'Recherche le PPN indiqué dans le message "Doublon possible" après création d'une notice
	 dim msg
	 msg = ""
	 On Error Resume Next
	 msg = application.activeWindow.messages.item(0).text
	 If InStr(msg, "PPN ") > 0 Then
	 	msg = Mid(msg, InStr(msg, "PPN ") + 4, 9)
	 	application.activeWindow.command "che ppn " & msg
	 Else
	 	msgbox "Le message de doublon possible n'est pas affiché."
	End If
End Sub

Sub searchExcelPPNList()
'Recherche la liste de PPN contenu dans le presse-papier
	Dim query
    
With Application.activeWindow
	query = "che ppn " & replace(replace(.Clipboard, "(PPN)", ""), Chr(10), " OR ")
	query = Left(query, Len(query)-4)
	.Clipboard = query
	.Command query
End With

End Sub
