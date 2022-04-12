// ------------------------------ Initialisation of consts ------------------------------

// Already declared by the Abes I think
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

const thePrefs =
Components.classes["@mozilla.org/preferences-service;1"]
.getService(Components.interfaces.nsIPrefBranch);

const theEnv = Components.classes["@mozilla.org/process/environment;1"]
.getService(Components.interfaces.nsIEnvironment);

// ------------------------------ Load scripts ------------------------------

const alpScripts = ["resource:/Profiles/apeyrat001/alp_scripts/js/NE_PAS_DIFFUSER.js",
"resource:/Profiles/apeyrat001/alp_scripts/js/peyrat_ressources.js",
"resource:/Profiles/apeyrat001/alp_scripts/js/peyrat_main.js",
"resource:/Profiles/apeyrat001/alp_scripts/js/SCOOP.js",
"resource:/Profiles/apeyrat001/alp_scripts/js/peyrat_peb.js",
"resource:/Profiles/apeyrat001/alp_scripts/python-winibw/pythWinIBW.js",
"resource:/Profiles/apeyrat001/alp_scripts/python/python.js"];

for (var ii = 0; ii < alpScripts.length;ii++){
	application.writeProfileString("ibw.standardScripts", "script.AlP"+ii, alpScripts[ii]);
}