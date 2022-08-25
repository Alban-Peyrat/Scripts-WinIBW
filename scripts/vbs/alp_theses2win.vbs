' Two different ways are used here :
' - XML : using the XMLDOM object : https://baptiste-wicht.developpez.com/tutoriels/microsoft/vbscript/xml/xpath/
' - HTML : using the InternetExplorer object : Original Script IdRef (Abes) https://github.com/abes-esr/winibw-scripts/blob/3f374e37151ab686fd1423cc21195b997d7df4b9/user-scripts/idref/IdRef.vbs
'As we use an Internet Explorer object, we can't use JSON through IE object
'Use set for objects
'Use getElementsByTagName / ByClassName / ById / getAttribute/ .innertext for HTML
'Use selectSingleNode / selectNodes / getAttribute / .text for XML

'Some ressources :
'Dictionnaries : https://www.dotnetperls.com/dictionary-vbnet'
'Elements : https://developer.mozilla.org/en-US/docs/Web/API/Element'
'Regex : https://www.oreilly.com/library/view/vbscript-in-a/1565927206/re155.html'
'Regex : https://www.regular-expressions.info/vbscript.html'


'A faire : créer une fonction create XML DOM object, qui sera bien plus pratique que l'autre'

Sub dumasXMLDOM()
	Dim url, urlSolr
	' Entre la fin mai/début juin 2022 et la fin aout 2022 ils ont changé le fonctionnement de docid / des id ???
	' Quand j'ai créé le script ça fonctionnait avec le 0 dévant mais maintenant je dois le supprimer ????????
	' [08/2022 comment] url = "https://api.archives-ouvertes.fr/search/dumas/?q=docid:01911186&wt=xml-tei"
	' [08/2022 comment] urlSolr = "https://api.archives-ouvertes.fr/search/dumas/?q=docid:01911186&wt=xml&fl=dumas_degreeSpeciality_s"
	
	url = "https://api.archives-ouvertes.fr/search/dumas/?q=docid:1911186&wt=xml-tei"
	urlSolr = "https://api.archives-ouvertes.fr/search/dumas/?q=docid:1911186&wt=xml&fl=dumas_degreeSpeciality_s"

	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	xmlDoc.Async = "false"
	xmlDoc.Load(url)

	 For Each personneElement In xmlDoc.selectNodes("/TEI/text/body/listBibl/biblFull/titleStmt/title")
		MsgBox personneElement.text
	 Next


	Set xmlDoc = Nothing
End Sub


Sub thesesDumas2winibw()
	'Main function
	'Uses XML-TEI because it has the most precise information'
	'And XML from Solr because XML-TEI soesn't have everything we want


' DELETE
	Dim url, urlSolr
'url = "https://api.archives-ouvertes.fr/search/dumas/?q=docid:01911186&wt=xml-tei&fl=language_s,uri_s,authFirstName_s,authLastName_s,director_s,authStructId_i,dumas_degreeSpeciality_s,dumas_degreeType_s,publicationDateY_i,title_s,keyword_s,abstract_s,page_s,openAccess_bool"
'[08/2022 comment] url = "https://api.archives-ouvertes.fr/search/dumas/?q=docid:01911186&wt=xml-tei"
'[08/2022 comment] urlSolr = "https://api.archives-ouvertes.fr/search/dumas/?q=docid:01911186&wt=xml&fl=dumas_degreeSpeciality_s"
url = "https://api.archives-ouvertes.fr/search/dumas/?q=docid:1911186&wt=xml-tei"
urlSolr = "https://api.archives-ouvertes.fr/search/dumas/?q=docid:1911186&wt=xml&fl=dumas_degreeSpeciality_s"

'DELETE'

	'Gets the input'
	'A CODER'

	'Gets the dumas record'
	Set docDumas = getIEObjectDocument(url)

	'If docDumas is not nothing, creates a title. Else, quit.'
	If Not docDumas Is Nothing Then 
		Application.ActiveWindow.Command "cre e"
	Else
		MsgBox "WinIBW n'a pas réussi à se connecter à DUMAS."
		Exit Sub
	End If


	'Check if the result is a record'
	'A CODER'


	'A CODER : des trucs qui checkent si l'élément / l'attrivvut existent 
	Dim Test, test2, test3
	' [08/2022 comment] Set Test = getIEDocTag(docDumas, "link")(0)
	Set Test = getIEDocTag(docDumas, "licence")(0)

	Msgbox test.textContent & chr(10) & test.innerText & chr(10) & getXmlTextContent(test)
	'MsgBox test.getAttribute("status")
	'MsgBox getElemAttr(test, "status")

	'MsgBox docDumas.getElementsByTagName("notesStmt")(0).getElementsByTagName("note")(0).getAttributeNames()(0)

	' [08/2022 comment : c'est l'IE qu'il faut quit, pas le IE.doc....] docDumas.Quit
	Set docDumas = Nothing 
End Sub

Private Function getIEObjectDocument(url)
	'Creates an IE object and returns the document property'
	'Original Script IdRef (Abes) https://github.com/abes-esr/winibw-scripts/blob/3f374e37151ab686fd1423cc21195b997d7df4b9/user-scripts/idref/IdRef.vbs'
	Dim IE  
	set IE = nothing
	set shapp=createobject("shell.application")
	    on error resume next
	    'pour ouvrir si pas ouvert
	    For Each owin In shapp.Windows
	         if left(owin.document.location.href,len(url))=url then
	            if err.number = 0 then
	                    set IE = owin
	                    'MsgBox "ok"
	              end if
	        end if
	    err.clear
	    Next

	    on error goto 0
	    if IE is nothing then
	         Set IE = CreateObject("InternetExplorer.Application")
	    end if

	    IE.Navigate2 url
		Do While IE.readystate <> 4  
	    Loop  
	    set getIEObjectDocument = IE.document
End Function

Private Function getIEDocTag(IEDoc, tag)
	'Returns an array of object using getElementsByTagName'
	Set getIEDocTag = IEDoc.getElementsByTagName(tag)
End Function

Private Function getElemAttr(elem, attr)
	'Returns the value of the attribute for that element'
	getElemAttr = elem.getAttribute(attr)
End Function

Private Function getXmlTextContent(elem)
	'Returns the innerText of a XMl element without the tag'
	'Regex pattern : https://stackoverflow.com/questions/6743912/how-to-get-the-pure-text-without-html-element-using-javascript#6744068'
	Set reg = New RegExp
	reg.Pattern = "<[^>]*>"
	reg.Global = True

	getXmlTextContent = reg.Replace(elem.textContent, "")
End Function

Private Sub oskar2winibw(url)

	'Pour des tests'
	' Dim url
	' url = "https://oskar-bordeaux.fr/handle/20.500.12278/23589?show=full"


	'Gets the dumas record'
	Dim docOskar
	Set docOskar = getIEObjectDocument(url)

	'If docDumas is not nothing, creates a title. Else, quit.'
	If Not docOskar Is Nothing Then 
		Application.ActiveWindow.Command "cre e"
	Else
		MsgBox "WinIBW n'a pas réussi à se connecter à Oskar."
		Exit Sub
	End If

	' Création dictionaire'
	Dim values
	Set values = CreateObject("Scripting.Dictionary")
	
	'Récupération de la table'
	Dim table, trs
	Set table = getIEDocTag(docOskar, "table")(0)
	Set trs = getIEDocTag(table, "tr")

	'For each line'
	For Each tr in trs
		'Gets the columns'
		Dim cols, mtdt, val, useful
		Set cols = getIEDocTag(tr, "td")

		'Gets values from the columns'
		mtdt = cols(0).textContent
		val = cols(1).textContent


	Select Case mtdt
		case "dc.contributor.author"
			useful = "Auteur : "
		case "dc.contributor.advisor"
			useful = "Directeur de thèse : "
		case "dc.date"
			useful = "Date de publication : "
		case else
			useful = "----------> "
	End Select


		Application.ActiveWindow.Title.InsertText useful & mtdt & " : " & val & vblf
	Next


'Full comment jusqu'à la fin, sauf le docOskar quit

	' 	'OSKAR TEST'
	' ' Création dictionaire'
	' Dim values
	' Set values = CreateObject("Scripting.Dictionary")
	
	' 'Récupération de la table'
	' Dim table
	' Set table = getIEDocTag(docOskar, "table")(0)

	' 'récupération nom du dir'
	' Dim ligneDir, dirTd, preDir, patrDir
	' Set ligneDir = getIEDocTag(table, "tr")(1)
	' Set dirTd = getIEDocTag(ligneDir, "td")(1)
	' patrDir = dirTd.innerText

	' preDir = Mid(patrDir, InStr(patrDir, ",")+2, Len(patrDir))
	' patrDir = Left(patrDir, InStr(patrDir, ",")-1)
	' patrDir = Left(patrDir, 1) & LCase(Right(patrDir, Len(patrDir)-1))
	
	' values.Add "preDir", preDir
	' values.Add "patrDir", patrDir

	' 'récupération nom de l'auteur
	' Dim ligneAut, autTd, preAut, patrAut
	' Set ligneAut = getIEDocTag(table, "tr")(2)
	' Set autTd = getIEDocTag(ligneAut, "td")(1)
	' patrAut = autTd.innerText

	' preAut = Mid(patrAut, InStr(patrAut, ",")+2, Len(patrAut))
	' patrAut = Left(patrAut, InStr(patrAut, ",")-1)
	' patrAut = Left(patrAut, 1) & LCase(Right(patrAut, Len(patrAut)-1))
	
	' values.Add "preAut", preAut
	' values.Add "patrAut", patrAut

	' 'récupération du titre
	' Dim ligneTitre, titreTd, titre, sousTitre
	' Set ligneTitre = getIEDocTag(table, "tr")(16)
	' Set titreTd = getIEDocTag(ligneTitre, "td")(1)
	' titre = titreTd.innerText

	' sousTitre = Mid(titre, InStr(titre, "?")+2, Len(titre))
	' titre = Left(titre, InStr(titre, "?"))

	' values.Add "titre", titre
	' values.Add "sousTitre", sousTitre
 	

	' 'récupération de la Date
	' Dim ligneDate, dateTd
	' Set ligneDate = getIEDocTag(table, "tr")(3)
	' Set dateTd = getIEDocTag(ligneDate, "td")(1)
	' values.Add "date", Left(dateTd.innerText, 4)

	' 'Création de l'UNIMARC
	' 'U200'
	' Application.ActiveWindow.Title.InsertText "200 1#$a@" & values.item("titre") & "$e" & values.item("sousTitre")_
	' & "$f" & values.item("preAut") & " " & values.item("patrAut") & "$gsous la direction de "_ 
	' & values.item("preDir") & " " & values.item("patrDir") & chr(10)

	' 'U214'
	' Application.ActiveWindow.Title.InsertText "214 #1$a" & values.item("date")

	' 'Msgbox values.item("preDir") & chr(10) &values.item("patrDir") & chr(10) & values.item("preAut") & chr(10) &values.item("patrAut") & chr(10) &values.item("sousTitre") & chr(10) & values.item("date")

	' [08/2022 comment : c'est l'IE qu'il faut quit, pas le IE.doc....] docOskar.Quit
	Set docOskar = Nothing 

End Sub