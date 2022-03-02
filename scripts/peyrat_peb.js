function AlP_PEBgetNumDemande(){
	application.activeWindow.clipboard = application.activeWindow.getVariable("P3GA*")
}


function AlP_PEBgetNumDemandePostValidation(){
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

function AlP_PEBgetPPN(){
	application.activeWindow.clipboard = application.activeWindow.getVariable("P3VTA");
}

function AlP_PEBgetRCRDemandeur(){
	var VF0 = application.activeWindow.getVariable("P3VF0");
	var VF1 = application.activeWindow.getVariable("P3VF1");
	
	if(VF1 != VF0){
		var prompter = Components.classes["@oclcpica.nl/scriptpromptutility;1"]
				.createInstance(Components.interfaces.IPromptUtilities);
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

function AlP_PEBLauncher(){
	const utility = {
		newPrompter: function() {
			return Components.classes["@oclcpica.nl/scriptpromptutility;1"]
			.createInstance(Components.interfaces.IPromptUtilities);
		}
	};
	
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
		default:
			application.messageBox("Erreur", "Script s\u00E9lectionn\u00E9 pas pris en charge","alert-icon");
	}
}

function AlP_PEBtriRecherche(){
const utility = {
	newFileOutput: function() {
		return Components.classes["@oclcpica.nl/scriptoutputfile;1"]
		.createInstance(Components.interfaces.IOutputTextFile);
	}
};
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
			record = AlP_js_removeAccents(record);
			var PPN = record.substring(record.indexOf("\u001BH\u001BLPP")+6, record.indexOf("\u001BE", record.indexOf("\u001BE\u001BLPP")+6));
			var auteur = record.substring(record.indexOf("\u001BE\u001BLV0")+6, record.indexOf("\u001BE", record.indexOf("\u001BE\u001BLV0")+6));
			var titre = record.substring(record.indexOf("\u001BE\u001BLV1")+6, record.indexOf("\u001BE", record.indexOf("\u001BE\u001BLV1")+6));
			var edition = record.substring(record.indexOf("\u001BE\u001BLV2")+6, record.indexOf("\u001BE", record.indexOf("\u001BE\u001BLV2")+6));
			var editeur = record.substring(record.indexOf("\u001BE\u001BLV3")+6, record.indexOf("\u001BE", record.indexOf("\u001BE\u001BLV3")+6));
			var annee = record.substring(record.indexOf("\u001BE\u001BLV4")+6, record.indexOf("\u001BE", record.indexOf("\u001BE\u001BLV4")+6));
			theOutputFile.writeLine(PPN+"\u0009"+auteur+"\u0009"+titre+"\u0009"+edition+"\u0009"+editeur+"\u0009"+annee);
			row = parseInt(record.substring(record.indexOf("\u001BD\u001BLNR")+6, record.indexOf("\u001BE", record.indexOf("\u001BD\u001BLNR")+6)).replace(" ", ""));
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
