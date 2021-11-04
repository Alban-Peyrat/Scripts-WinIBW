//Pour pouvoir afficher un prompter
// const utility = {
// 	newFileInput: function() {
// 		return Components.classes["@oclcpica.nl/scriptinputfile;1"]
// 		.createInstance(Components.interfaces.IInputTextFile);
// 	},
// 	newFileOutput: function() {
// 		return Components.classes["@oclcpica.nl/scriptoutputfile;1"]
// 		.createInstance(Components.interfaces.IOutputTextFile);
// 	},
// 	newPrompter: function() {
// 		return Components.classes["@oclcpica.nl/scriptpromptutility;1"]
// 		.createInstance(Components.interfaces.IPromptUtilities);
// 	}
// };

function AlP_PEBgetNumDemande(){
	application.activeWindow.clipboard = application.activeWindow.getVariable("P3GA*")
}


function AlP_PEBgetNumDemandePostValidation(){
    var msg;
    msg = application.activeWindow.messages.item(0).text;
    msg = msg.substring(msg.indexOf("no. ")+4, msg.indexOf('no. ')+14);
    application.activeWindow.clipboard = msg;
}

function AlP_PEBgetPPN(){
	application.activeWindow.clipboard = application.activeWindow.getVariable("P3VTA");
}

function AlP_PEBgetRCRDemandeur(){
	application.activeWindow.clipboard = application.activeWindow.getVariable("libID");
}

function AlP_PEBLauncher(){
	const utility = {
		newPrompter: function() {
			return Components.classes["@oclcpica.nl/scriptpromptutility;1"]
			.createInstance(Components.interfaces.IPromptUtilities);
		}
	};
	
	var thePrompter = utility.newPrompter();
	var ans = thePrompter.select("Executer un script du PEB :", "Choisir le script a executer",
		"Get no demande PEB" +
		"\nGet no demande PEB post-validation" +
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
			application.messageBox("Erreur", "Script sélectionné pas pris en charge","alert-icon");
	}
}
