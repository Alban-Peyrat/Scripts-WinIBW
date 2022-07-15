'Ce fichier est uniquement utilis� pour param�trer WinIBW autant pour son interface (?)
'que pour r�cup�rer des variables communes � VBS et JS
'que pour charger les autres scripts en VBS
'que pour permettre au fichier central de param�trage de JS d'�tre charg�

'Charge le script central de JS qui permet � mes scripts JS d'�tre charg�s. Voir le fichier pour plus d'informations
application.writeProfileString "ibw.standardScripts","script.AlP","resource:/Profiles/apeyrat001/alp_scripts/alp_central_scripts.js"

'Permet � mes scripts VBS d'�tre charg�s.
sluitMapIn("C:\oclcpica\WinIBW30\Profiles\apeyrat001\alp_scripts\vbs")

' R�cup�re des variables d'environnement que j'ai d�fini dans un autre fichier (JS)
Dim WSHShell, MY_RCR
Set WSHShell = CreateObject("WScript.Shell")
MY_RCR = WSHShell.ExpandEnvironmentStrings("%MY_RCR%")


Private Sub alp_param()
	'Suppos�ment permet d'appliquer mes param�tres dans WinIBW
'Ces trois param�tres sont d�finis dans un fichier JS
'application.protectedColor = "0x66FFF8";
'application.ignoredColor = "0xFFC000";
'application.addSyntaxColor("UNM", "(?:[^\\$]|^)(?:\\$\\$)*(\\$[^\\$ ])", 0x227711);

' Jamais essay� mais devrait marcher
	application.writeProfileString "winibw.shortpresentationscreen","background","#121212"
	application.writeProfileString "winibw.shortpresentationscreen","font.color","#EBE0EB"
	application.writeProfileInt "winibw.shortpresentationscreen","font.size",16
		
	application.writeProfileString "winibw.presentationscreen","background","#121212"
	application.writeProfileString "winibw.presentationscreen","font.color","#EBE0EB"
	application.writeProfileInt "winibw.presentationscreen","font.size",16

	application.writeProfileString "winibw.diacriticsbar","font.color","#EBE0EB"
	application.writeProfileInt "winibw.diacriticsbar","font.size",16

	application.writeProfileString "winibw.editscreen","background","#660000"
	application.writeProfileString "winibw.editscreen","font.color","#EBE0EB"
	application.writeProfileInt "winibw.editscreen","font.size",16
	
	application.writeProfileString "browser","anchor_color","#117722"
	
	application.writeProfileString "ibw.presentation","syntaxcolor.UNM.format.1","$1<span class=""presunm"" style=""color:#117722"">&lrm;$2</span>"	
	application.writeProfileString "ibw.presentation","syntaxcolor.UNMA.format.1","$1<span class=""presunm"" style=""color:#117722"">&lrm;$2</span>"	
End Sub

'Import all vbs files from a directory'
'Seuls les textes des messages sont traduits
'From https://cbs-nl.oclc.org/htdocs/winibw/scripts/WinIBW3.installatie.scriptbeheer.html'
Private Sub sluitMapIn(map)
   Set oFSO = CreateObject("Scripting.FileSystemObject")
   If map = "" Then
      msgbox "Aucun nom de dossier renseign�."
   elseif oFSO.FolderExists(map) Then
      Set folder = oFSO.GetFolder(map)
      Set bestanden  = folder.Files
      For each bestand In bestanden
         if lcase(mid(bestand.Name,len(bestand.Name)-3))=".vbs" then
            sluitVBSin(map & "\" & bestand.Name)
         end if
      Next
   else
      msgbox "Le dossier """ + map + """ n'existe pas ?"
   End If
End Sub

'Import a signle VBS file'
'Seuls les textes des messages sont traduits
'From https://cbs-nl.oclc.org/htdocs/winibw/scripts/WinIBW3.installatie.scriptbeheer.html'
Private Sub sluitVBSin(VBSbestand)
   Dim f, s, oFSO
   Set oFSO = CreateObject("Scripting.FileSystemObject")
   On Error Resume Next
   If oFSO.FileExists(VBSbestand) Then
      Set f = oFSO.OpenTextFile(VBSbestand)
      s = f.ReadAll
      f.Close
      ExecuteGlobal s
   else
      msgbox "Le fichier """ + VBSbestand + """ � int�grer n'existe pas ?"
      exit sub
   End If
   On Error Goto 0
   Set f = Nothing
   Set oFSO = Nothing
End Sub

'Je suis pas totalement s�r de pourquoi c'est l�. Peut-�tre un exemple pour une msgbox avec un VbYesNo ?
'Semble rechercher un auteur, afficher le premier r�sultat et demander si l'on veut consulter le second
Private Sub Pauze2()
'From Pica Handleiding Script de 2002

 Application.ActiveWindow.Command "rec t ; z aut baantjer"
 Application.ActiveWindow.Command "t s1 1"
 antwoord = (msgbox ("Wil u de tweede treffer uit de set zien?", VbYesNo))

 If antwoord = VBYes Then
 Application.ActiveWindow.Command "t s1 2"
 Else
 MsgBox "Jammer!"

 End If

 End Sub