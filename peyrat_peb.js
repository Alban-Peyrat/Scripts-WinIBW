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

function AlP_PEB_getNumDemande(){
	application.activeWindow.clipboard = application.activeWindow.getVariable("P3GA*")
}


function AlP_PEB_getNumDemandePostValidation(){
    var msg;
    msg = application.activeWindow.messages.item(0).text;
    msg = msg.substring(msg.indexOf("no. ")+4, msg.indexOf('no. ')+14);
    application.activeWindow.clipboard = msg;
}

function AlP_PEB_getRCRDemandeur(){
	application.activeWindow.clipboard = application.activeWindow.getVariable("libID");
}

function AlP_PEB_Launcher(){
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
		"\nGet RCR demandeur" +
		"\nGet RCR fournisseur en attente");

	switch (ans){
		case "Get no demande PEB":
			AlP_PEB_getNumDemande();
			break;
		case "Get no demande PEB post-validation":
			AlP_PEB_getNumDemandePostValidation();
			break;
		case "Get RCR demandeur":
			AlP_PEB_getRCRDemandeur();
			break;
		case "Get RCR fournisseur en attente":
			AlP_PEB_getRCRFournisseurOnHold();
			break;
		default:
			application.messageBox("Erreur", "Script sélectionné pas pris en charge","alert-icon");
	}
}