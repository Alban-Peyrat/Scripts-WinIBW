'Scripts for DUMAS'

Sub these_catDumas()
	
	dim docId
	docId = application.activeWindow.clipboard
	docId = replace(docId, chr(13), "")
	docId = replace(docId, chr(10), "")
	docId = replace(docId, " ", "")
	docId = Right(docId, 8)
	
'Source : script IdRef de l'Abes
	set IE = nothing
    set shapp=createobject("shell.application")
     Dim InputTexte
	'MsgBox  "==>" + IE.Visible
    on error resume next
    'pour ouvrir si pas ouvert
    For Each owin In shapp.Windows
         'if left(owin.document.location.href,len("https://api.archives-ouvertes.fr/search/dumas/?q=docid:01911186&wt=xml&fl=language_s,title_s,authFirstName_s,authLastName_s,uri_s,director_s,keyword_s,uri_s,abstract_s,page_s,dumas_degreeSpeciality_s,publicationDateY_i,dumas_degreeSubject_s"))="https://api.archives-ouvertes.fr/search/dumas/?q=docid:01911186&wt=xml&fl=language_s,title_s,authFirstName_s,authLastName_s,uri_s,director_s,keyword_s,uri_s,abstract_s,page_s,dumas_degreeSpeciality_s,publicationDateY_i,dumas_degreeSubject_s" then
         if left(owin.document.location.href,len("https://alban-peyrat.github.io/outils/ub-svs/dumas/generateur-notice.html"))="https://alban-peyrat.github.io/outils/ub-svs/dumas/generateur-notice.html" then
            if err.number = 0 then
                    set IE = owin
                    'MsgBox "ok"
              end if
        end if
    err.clear
    Next

    on error goto 0
    if IE is nothing then
        'MsgBox  "Window Not Open"
         Set IE = CreateObject("InternetExplorer.Application")
    end if

    'IE.Navigate2 "https://api.archives-ouvertes.fr/search/dumas/?q=docid:01911186&wt=xml&fl=language_s,title_s,authFirstName_s,authLastName_s,uri_s,director_s,keyword_s,uri_s,abstract_s,page_s,dumas_degreeSpeciality_s,publicationDateY_i,dumas_degreeSubject_s"    
	IE.Navigate2 "https://alban-peyrat.github.io/outils/ub-svs/dumas/generateur-notice.html"    
	Do While IE.readystate <> 4  
    Loop  
    Set IEDoc = IE.document
    
    'Set inputURL = IEDoc.getElementById("inputURL")
    'inputURL = "https://api.archives-ouvertes.fr/search/dumas/?q=docid:"+docId+"&wt=json&fl=language_s,title_s,authFirstName_s,authLastName_s,uri_s,director_s,keyword_s,uri_s,abstract_s,page_s,dumas_degreeSpeciality_s,publicationDateY_i,dumas_degreeSubject_s"
    'inputURL = "https://dumas.ccsd.cnrs.fr/dumas-01911186"

    Call IEDoc.parentWindow.execScript("main('"&docId&"')","JavaScript")

'Permet de stopper le script WinIBW pour laisser le temps au générateur
	ress_sleep 1

    Set notice = IEDoc.getElementById("notice")
'Fin de l'abes
	
	application.activeWindow.Command "cre"
	'Faire de la détection de ;_;ERREUR;_; + de si la notice est vide dire de réessayer puis si plusierus échec de passer par le site
	application.activeWindow.Title.InsertText notice.innerText
	

'L'abes encore parce que c'est quand même bien mieux de fermer Internet Explorer
'Sauf qu'en fait ça augmente le nombre de renvois vide ahahahahah
'Sauf si je sleep le script WinIBW pour laisser le temps à mon générateur ?
	IE.Quit
	Set IE = Nothing 
End Sub