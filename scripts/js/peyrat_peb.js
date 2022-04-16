// Scripts pour le module PEB
// Scripts for ILL module

function AlP_PEBgetNumDemande(){
// Returns the ILL request Number (must be used on a ILL request)
	application.activeWindow.clipboard = application.activeWindow.getVariable("P3GA*")
}

function AlP_PEBgetNumDemandePostValidation(){
// Returns the ILL request Number that was jsut created
// Gets it from the messages, probably only works for french catalog
    var msg;
    try{
		msg = application.activeWindow.messages.item(0).text;
		if(msg.indexOf("no. ") > -1){
			msg = msg.substring(msg.indexOf("no. ")+4, msg.indexOf('no. ')+14);
			application.activeWindow.clipboard = msg;
		}else{
			throw false;
		}
    }catch(e){
    	application.messageBox("Erreur", "Le message de cr\u00E9ation de demande n'est pas affich\u00E9.", "alert-icon");
    }   
}

function AlP_PEBgetRCRDemandeur(){
// Returns the requesting library's RCR of an ILL (must be used on a ILL request)
	var VF0 = application.activeWindow.getVariable("P3VF0");
	var VF1 = application.activeWindow.getVariable("P3VF1");
	
	if(VF1 != VF0){
		var prompter = utility.newPrompter();
		var ans = prompter.confirmEx("Quel RCR choisir", "Quel RCR (cliquer sur le bouton)", "Aucun", VF0, VF1, null, null)
		switch (ans){
			case 0:
				application.messageBox("Erreur", "Aucun RCR copié","alert-icon");
				break;
			case 1:
				application.activeWindow.clipboard = VF0;
				break;
			case 2:
				application.activeWindow.clipboard = VF1;
				break;
			default:
				application.messageBox("Erreur", "Aucun RCR copié","alert-icon");
			}
	}else{
		application.activeWindow.clipboard = application.activeWindow.getVariable("P3VF0");
	}
}

function AlP_PEBgetRCRFournisseurOnHold(){
	var comment;
	var proc = false;
	var fournisseurs = application.activeWindow.getVariable("P3VCA").split("\u000D");
	for(var ii = 0; ii < fournisseurs.length-1; ii++){
		comment = fournisseurs[ii].substring(fournisseurs[ii].indexOf("\u001BE\u001BLRT")+6, fournisseurs[ii].indexOf("\u001BE", fournisseurs[ii].indexOf("\u001BE\u001BLRT")+6));
		if(comment === "En attente de r\u00E9ponse"){
			application.activeWindow.clipboard = fournisseurs[ii].substring(fournisseurs[ii].indexOf("\u001BE\u001BLSS")+6, fournisseurs[ii].indexOf("\u001BE\u001BLSS")+15);
			proc = true;
			break;
		}
	}
	if(proc === false){
		application.messageBox("Erreur", "Les biblioth\u00E8ques ont r\u00E9pondu.", "alert-icon");
	}
}

function AlP_PEBgetPPN(){
// Returns the PPN of the wanted document on an ILL (must be used on a ILL request)
	application.activeWindow.clipboard = application.activeWindow.getVariable("P3VTA");
}

function AlP_PEBgetTitleAuth(){
// Returns the document title, document author, article title and article author
// Separated by carriage return. If one of them doesn't exists, returns an empty stirng instead (must be used on a ILL request)
    var titre = application.activeWindow.getVariable("P3VTC");
    var auteur = application.activeWindow.getVariable("P3VTD");
    var article = application.activeWindow.getVariable("P3VAB");
    var auteurArt = application.activeWindow.getVariable("P3VAA");
    application.activeWindow.clipboard = titre + "\n" + auteur + "\n" + article + "\n" + auteurArt
}

function AlP_PEBLauncher(){
// Opens a launcher for all the scripts in this file
	var thePrompter = utility.newPrompter();
	var ans = thePrompter.select("Ex\u00E9cuter un script du PEB :", "Choisir le script \u00E0 ex\u00E9cuter",
		"Get no demande PEB" +
		"\nGet no demande PEB post-validation" +
		"\nTrier recherche" +
		"\nGet PPN" +
		"\nGet RCR demandeur" +
		"\nGet RCR fournisseur en attente");

	switch (ans){
		case "Get no demande PEB":
			AlP_PEBgetNumDemande();
			break;
		case "Get no demande PEB post-validation":
			AlP_PEBgetNumDemandePostValidation();
			break;
		case "Trier recherche":
			AlP_PEBtriRecherche();
			break;
		case "Get PPN":
			AlP_PEBgetPPN();
			break;
		case "Get RCR demandeur":
			AlP_PEBgetRCRDemandeur();
			break;
		case "Get RCR fournisseur en attente":
			AlP_PEBgetRCRFournisseurOnHold();
			break;
		case "Get titre et auteur document":
			AlP_PEBgetTitleAuth();
			break;
		default:
			application.messageBox("Erreur", "Script s\u00E9lectionn\u00E9 pas pris en charge","alert-icon");
	}
}

function AlP_PEBtriRecherche(){
// Extract and open an excel file with all entries in a short presentation
// list of records (must be used on a short presentation list)
var theOutputFile = utility.newFileOutput();
theOutputFile.createSpecial("ProfD", "triPEB.xls");
theOutputFile.setTruncate(true);
var path = theOutputFile.getPath();
theOutputFile.writeLine("PPN\u0009Auteur\u0009Titre\u0009Edition\u0009Editeur\u0009Annee");
	var resTable = new Array();
	var noLot = application.activeWindow.getVariable("P3GSE");
	application.activeWindow.command("\\too s"+noLot+" k", false);
	var nbRes = application.activeWindow.getVariable("P3GSZ");

//v2
	var row = 0;
	var sec = 0
	while(row < nbRes){
		application.activeWindow.command("\\too s"+noLot+" "+(row+1)+" k", false);
		var segP3VKZ = application.activeWindow.getVariable("P3VKZ");
		var records = segP3VKZ.split("\u000D");
		for(var jj = 0; jj < records.length-1;jj++){
			var record = records[jj];
			record = record.replace(/\"/g, "");
			record = __removeAccents(record);
			var PPN = record.substring(record.indexOf("\u001BLPP")+4, record.indexOf("\u001BE", record.indexOf("\u001BLPP")+4));
			var auteur = record.substring(record.indexOf("\u001BLV0")+4, record.indexOf("\u001BE", record.indexOf("\u001BLV0")+4));
			var titre = record.substring(record.indexOf("\u001BLV1")+4, record.indexOf("\u001BE", record.indexOf("\u001BLV1")+4));
			var edition = record.substring(record.indexOf("\u001BLV2")+4, record.indexOf("\u001BE", record.indexOf("\u001BLV2")+4));
			var editeur = record.substring(record.indexOf("\u001BLV3")+4, record.indexOf("\u001BE", record.indexOf("\u001BLV3")+4));
			var annee = record.substring(record.indexOf("\u001BLV4")+4, record.indexOf("\u001BE", record.indexOf("\u001BLV4")+4));
			theOutputFile.writeLine(PPN+"\u0009"+auteur+"\u0009"+titre+"\u0009"+edition+"\u0009"+editeur+"\u0009"+annee);
			row = parseInt(record.substring(record.indexOf("\u001BLNR")+4, record.indexOf("\u001BE", record.indexOf("\u001BLNR")+4)).replace(" ", ""));
		}
//Empêche la boucle While de tourner à l'infini
		sec++;
		if(sec > 9999){
			break;
		}
	}
	theOutputFile.close();
	application.shellExecute(path, 9, "edit", "");
}




function AlP_PEBsearchInSuDb(){
/* Marche pas pour le moment si ça vient d'un lien*/

	// Gets the limitations parameters
	application.activeWindow.command("\\too \\adi", false);
	var lim = application.activeWindow.messages.item(0).text;

	// Gets the query
	var query = application.activeWindow.getVariable("P3VCO");
	if (query.substring(0, 11).indexOf("recherche") > -1){
	    query = query.replace("recherche", "\\zoe");
	// Exits if the query is not initiated with "che"
	}else {
	    application.messageBox("Erreur", "Ce type de recherche n'est pas pris en compte.", "error-icon");
	    return
	}


	// Connects to Sudoc catalog, launches the search aff k
	application.activeWindow.command("\\sys 1;\\bes 1;"+lim+query, false);

	// Checks if the search worked
	if (application.activeWindow.getVariable("P3GSY") != "SU") {
	    application.messageBox("Erreur", "La recherche a échoué. Vous vous trouvez actuellement dans la base " + application.activeWindow.getVariable("P3GSY") + ".\nRéférez-vous aux messages de WinIBW pour plus d'informations.", "error-icon");
	    return
	}

	// Checks if the default display is ISBD
	// Actually the parameter is reseted when switching database so the check is kinda useless
	var affDl = application.activeWindow.getVariable("P3GDL", "I");
	if (affDl !== "I"){
	    // Opens a new window to set default parameters
	    application.activeWindow.command("\\mut \\par", true);
	    application.activeWindow.setVariable("P3VDL", "I");
	    application.activeWindow.simulateIBWKey("FR");
	    application.activeWindow.closeWindow();
	}

	// Without this, WinIBW won't display the list
	application.activeWindow.command("\\too k 1", false);
}

function AlP_PEBaskFromSu(){
	var ppn = application.activeWindow.getVariable("P3GPP");
	// Checks if there's a PPN
	if (ppn == "") {
	    application.messageBox("Erreur", "Veuillez sélectionner une notice.", "error-icon");
	    return
	// Checks if it's a bibliographic record
	}else if (application.activeWindow.getVariable("P3VMC").charAt(0) == "T") {
	    application.messageBox("Erreur", "Ceci est une notice d'autorité. Veuillez sélectionner une notice bibliographique.", "error-icon");
	    return
	}

	application.activeWindow.command("\\sys 2;\\bes 1;\\zoe ppn "+ppn+";\\too i", false);
	// Checks if the search worked
	if (application.activeWindow.getVariable("P3GSY") != "SU PEB") {
	    application.messageBox("Erreur", "La recherche a échoué. Vous vous trouvez actuellement dans la base " + application.activeWindow.getVariable("P3GSY") + ".\nRéférez-vous aux messages de WinIBW pour plus d'informations.", "error-icon");
	    return
	}

	application.activeWindow.simulateIBWKey("F9");
	// Checks if the ILL request started
	if (application.activeWindow.getVariable("scr") != "AA") {
	    application.messageBox("Erreur", "La demande de PEB a échoué. Vous vous trouvez actuellement dans la base " + application.activeWindow.getVariable("P3GSY") + ".\nRéférez-vous aux messages de WinIBW pour plus d'informations.", "error-icon");
	    return
	}
}
