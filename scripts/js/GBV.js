// All scripts here come from the Gemeinsame Bibliotheksverbund scripts for WinIBW
// See https://wiki.k10plus.de/display/K10PLUS/SWB-WinIBW-Versionsinformationen
// I only changed functions name and text

// From Update_2022_10/scripts/k10_public.js - function alert()
function alert(meldungstext){
	//alert ist eine JavaScript-Funktion, kommt aber im Scripting von WinIBW nicht vor.
	application.messageBox("Alerte", meldungstext, "alert-icon");
}

// From Update_2022_10/scripts/k10_public.js - function __warnung()
function __warning(meldungstext){
	application.messageBox("Attention", meldungstext, "alert-icon");
}

// From Update_2022_10/scripts/k10_public.js - function __fehler()
function __error(meldungstext){
	application.messageBox("Erreur", meldungstext, "error-icon");
}

// From Update_2022_10/scripts/k10_public.js - function __meldung()
function __msg(meldungstext){
	application.messageBox("Message", meldungstext, "message-icon");
}

// From Update_2022_10/scripts/k10_public.js - function __frage()
function __question(meldungstext)
{
	application.messageBox("Question", meldungstext, "question-icon");
}

// From Update_2022_10/scripts/k10_public.js - function __alleMeldungen()
function __getMsgs(){
	var msgAnzahl, msgText, msgType;
	var i;
	var alleMeldungen = "";
	msgAnzahl = application.activeWindow.messages.count;
	for (i=0; i<msgAnzahl; i++)
	{
		msgText = application.activeWindow.messages.item(i).text;
		msgType = application.activeWindow.messages.item(i).type;
		alleMeldungen += msgText + "\n";
	}
	//application.messageBox("Bitte beachten Sie die Meldungen!", alleMeldungen, "message-icon");
	return alleMeldungen;
}

// From Update_2022_10/scripts/k10_public.js - function meldungenKopieren()
function GBVgetMsgsClipboard(){
	var strMeldung = __getMsgs();
	//am Ende Zeilenumbruch entfernen:
	strMeldung = strMeldung.substr(0, strMeldung.length-1);
	application.activeWindow.clipboard = "\u0022" + strMeldung + "\u0022"; // u0022 = Quot
	application.activeWindow.appendMessage("Les messages ont été copiés dans le presse-papier.", 3);
}

// From Update_2022_10/scripts/k10_public.js - function felderLoeschen()
function __delFields(regexpFelder){
// Deletes all fields matching regex
	//Ersatz für title.ttl in Sonderfällen:
	//Beispiel Funktionsaufruf: übergeben werden reguläre Ausdrücke, z.B.
	//felderLoeschen(/417[0-9]|418[0-9]|4201|4218|4088|4236|4712/);

	var n= 0;
	var letzteZeile;

	//wieviele Zeilen sollen geprüft werden?
	application.activeWindow.title.endOfBuffer (false);
	letzteZeile = application.activeWindow.title.currentLineNumber;
	application.activeWindow.title.startOfBuffer (false);

	for (n=0; n<= letzteZeile; n++) {
		if (regexpFelder.test(application.activeWindow.title.tag)){
			application.activeWindow.title.deleteLine (1);
		} else {
			application.activeWindow.title.endOfField(false);//wichtig bei mehrzeiligen Inhalten!
			application.activeWindow.title.lineDown (1, false);
		}
	}
}

// From Update_2022_10/scripts/k10_public.js - function feldInhaltLoeschen()
function __delFieldsContent(regexFeld){
// For all fields matching Regex, deletes the field content
	//Ersatz für title.ttl in Sonderfällen:
	//Beispiel Funktionsaufruf: übergeben werden reguläre Ausdrücke, z.B.
	//feldInhaltLoeschen(/417[0-9]|418[0-9]|4201|4700|7100/);
	//Geht durch alle Zeilen und löscht nur den Inhalt
	var n= 0;
	var letzteZeile;
	//wieviele Zeilen sollen geprüft werden?
	application.activeWindow.title.endOfBuffer (false);
	letzteZeile = application.activeWindow.title.currentLineNumber;
	application.activeWindow.title.startOfBuffer (false);

	for (n=0; n<= letzteZeile; n++) {
		if (regexFeld.test(application.activeWindow.title.tag)){
			application.activeWindow.title.startOfField(false);
			application.activeWindow.title.wordRight(1, false);
			application.activeWindow.title.deleteToEndOfLine();
		}
		application.activeWindow.title.endOfField(false);//wichtig bei mehrzeiligen Inhalten!
		application.activeWindow.title.lineDown (1, false);
	}
}

// From Update_2022_10/scripts/k10_public.js - function feldEinfuegenNummerisch()
function __insFieldIfInexistant(ergaenzeFeld, strInhalt){
// Cherche si le champ existe : si non, l'inséère avec strInhalt comme contenuau bon endroti de la notice
	//Wenn das Feld (ergaenzeFeld) nicht vorkommt, wird sie und der Inhalt (strInhalt) an der passenden Stelle eingefügt.
	//Aufruf mit Feld als String,
	//Beispiel: feldEinfuegenNummerisch("1505", "$erda");
	//Beispiel: feldEinfuegenNummerisch("1505", " ");
	var lZeile = 0, i=0, vorhandenesFeld="";
	//application.activeWindow.title.startOfBuffer(false);
	if (application.activeWindow.title.findTag(ergaenzeFeld, 0, true, false, true) == "") {
		//dann soll die richtige Stelle zum einfügen gefunden werden
		application.activeWindow.title.endOfBuffer(false);
		lZeile = application.activeWindow.title.currentLineNumber;
		application.activeWindow.title.startOfBuffer(false);
		for (i=0; i<=lZeile; i++){
			vorhandenesFeld = application.activeWindow.title.tag;
			//fügt Feld vor der nächst höheren oder in der letzten Zeile ein:
			if ((!isNaN(vorhandenesFeld) && vorhandenesFeld > ergaenzeFeld) || application.activeWindow.title.currentLineNumber == lZeile){
				//alert(ergaenzeFeld);
				application.activeWindow.title.startOfField(false);
				application.activeWindow.title.insertText(ergaenzeFeld + " " + strInhalt + "\n");
				break;
			}
			application.activeWindow.title.endOfField(false);
			application.activeWindow.title.lineDown(1, false);
			//alert(application.activeWindow.title.currentLineNumber);
		}
	}
}

// From Update_2022_10/scripts/k10_public.js - function feldEinfuegenNummerischOhnePruefung()
function __insField(ergaenzeFeld, strInhalt){
// Insère le champ avec strInhalt comem contenu au bon endroti de la notice
	//Ohne Prüfung wird Feld (ergaenzeFeld) mit Inhalt (strInhalt) an der passenden Stelle eingefügt.
	//Aufruf mit Feld als String, z. B. feldEinfuegenNummerischOhnePruefung("3210", "irgendwas");
	var lZeile = 0, i=0, vorhandenesFeld="";
	application.activeWindow.title.endOfBuffer(false);
	lZeile = application.activeWindow.title.currentLineNumber;
	application.activeWindow.title.startOfBuffer(false);
	for (i=0; i<=lZeile; i++){
		vorhandenesFeld = application.activeWindow.title.tag;
		//fügt Feld vor der nächst höheren oder in der letzten Zeile ein:
		if ((!isNaN(vorhandenesFeld) && vorhandenesFeld > ergaenzeFeld) || application.activeWindow.title.currentLineNumber == lZeile){
			//alert(ergaenzeFeld);
			application.activeWindow.title.startOfField(false);
			application.activeWindow.title.insertText(ergaenzeFeld + " " + strInhalt + "\n");
			break;
		}
		application.activeWindow.title.endOfField(false);
		application.activeWindow.title.lineDown(1, false);
	}
}

// From Update_2022_10/scripts/k10_public.js - function felderSammeln()
function __getFields(regexpFelder){
// Returns every field matching Regex
	//Im Edit-Schirm sammelt diese Funktion alle Vorkommnisse der genannten Felder ein.
	//Beispiel Funktionsaufruf: übergeben werden reguläre Ausdrücke.
	//Beispiel: Alle Vorkommnisse von 2275, 2276 und 2277 sollen ausgegeben werden.
	//felderSammeln(/2275|2276|2277/);
	var n= 0;
	var dieZeile, letzteZeile;
	var rueckgabe = "";
	application.activeWindow.title.endOfBuffer(false);
	letzteZeile = application.activeWindow.title.currentLineNumber;
	application.activeWindow.title.startOfBuffer(false);

	for (n=0; n<= letzteZeile; n++) {
		dieZeile = application.activeWindow.title.currentField;
		if(regexpFelder.test(application.activeWindow.title.tag) == true){
			rueckgabe = rueckgabe + "\n" + dieZeile;
		}
		application.activeWindow.title.endOfField(false);//wichtig bei mehrzeiligen Inhalten!
		application.activeWindow.title.lineDown (1, false);
	}
	return rueckgabe;
}

// From Update_2022_10/scripts/k10_public.js - function __datum()
function __dateYMD(){
	//Form: JJJJ.MM.TT
	var heute = new Date();

	var strMonat = heute.getMonth();
	strMonat = strMonat + 1;
	if (strMonat <10){strMonat = "0" + strMonat};

	var strTag = heute.getDate();
	if (strTag <10){strTag = "0" + strTag};

	var datum = heute.getFullYear() + "." + strMonat + "." + strTag;
	return datum;
}

// From Update_2022_10/scripts/k10_public.js - function __datumTTMMJJJJ()
function __dateDMY(){
	//Form: TT.MM.JJJJ
	var heute = new Date();
	var strMonat = heute.getMonth();
	strMonat = strMonat + 1;
	if (strMonat <10){strMonat = "0" + strMonat};

	var strTag = heute.getDate();
	if (strTag <10){strTag = "0" + strTag};

	var datum = strTag + "." + strMonat + "." + heute.getFullYear();
	return datum;
}

// From Update_2022_10/scripts/k10_public.js - function __datumUhrzeit()
function __dateHours(){
// Returns YYYYMMDDHHMMSS
	//das Datum und die Uhrzeit wird Bestandteil des Dateinamens
	var jetzt = new Date();
	var jahr = jetzt.getFullYear();
	var monat = jetzt.getMonth() + 1;
	var strTag = jetzt.getDate();
	var stunde = jetzt.getHours();
	var minute = jetzt.getMinutes();
	var sekunde = jetzt.getSeconds();
	if (monat<10){monat = "0" + monat};
	if (strTag<10){strTag = "0" + strTag} ;
	if (stunde<10){stunde = "0" + stunde};
	if (minute<10){minute = "0" + minute};
	if (sekunde<10){sekunde = "0" + sekunde} ;
	return jahr.toString() + monat.toString() + strTag.toString() + stunde.toString() + minute.toString() + sekunde.toString();
}

// From Update_2022_10/scripts/k10_public.js - function stringTrim()
function __trim(meinString){
// Trim
	//Lösche Blanks am Anfang und am Ende des Strings:
	var regexpBlank = /^ | $/;
	while (regexpBlank.test(meinString) == true){
		meinString = meinString.replace(regexpBlank,"");
	}
	return meinString;
}

// From Update_2022_10/scripts/k10_public.js - function hackSystemVariables()
// According to what I understand, it's OCLC's script
// And according to what I understand, Abes's hackSystemvariables is not coded properly
function GBVhackSystemVariables(){
// Returns in clipboard all variables (names + value)
	//Clemens Buijs:
	var i, j, varName, varValue, reportG = "", reportV = "", reportL = "";
	alpha = "!0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

	//G:
	for (i = 0; i <= alpha.length; i++) {
		for (j = 0; j <= alpha.length; j++) {
		//	Use P3G for global, P3L for local, P3V for field variables
			varName = "P3G" + alpha.charAt(i) + alpha.charAt(j);
			varValue = application.activeWindow.getVariable(varName);
			if (varValue) reportG = reportG + "- " + varName + ": " + varValue + "\r\n";
		}
	}
	//application.messageBox("G-Variable:", reportG, "message-icon");
	//V:
	for (i = 0; i <= alpha.length; i++) {
		for (j = 0; j <= alpha.length; j++) {
		//	Use P3G for global, P3L for local, P3V for field variables
			varName = "P3V" + alpha.charAt(i) + alpha.charAt(j);
			varValue = application.activeWindow.getVariable(varName);
			if (varValue) reportV = reportV + "- " + varName + ": " + varValue + "\r\n";
		}
	}
	//application.messageBox("V-Variable:", reportV, "message-icon");
	//L:
	for (i = 0; i <= alpha.length; i++) {
		for (j = 0; j <= alpha.length; j++) {
		//	Use P3G for global, P3L for local, P3V for field variables
			varName = "P3L" + alpha.charAt(i) + alpha.charAt(j);
			varValue = application.activeWindow.getVariable(varName);
			if (varValue) reportL = reportL + "- " + varName + ": " + varValue + "\r\n";
		}
	}
	//application.messageBox("L-Variable:", reportL, "message-icon");
	// Output to clipboard
	application.activeWindow.clipboard = reportG + reportV + reportL;
	application.messageBox("GBVhackSystemVariables", "Toutes les variables sont désormais dans le presse-papier.", "message-icon");
	//application.activeWindow.appendMessage("Alle Variablen befinden sich jetzt im Zwischenspeicher", 2);
}
