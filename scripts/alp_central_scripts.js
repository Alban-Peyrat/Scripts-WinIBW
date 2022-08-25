// ------------------------------ Initialisation of consts ------------------------------

// Utility is declared in standart_utility.js (Abes)
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

// Application is already defined
// var application = Components.classes["@oclcpica.nl/kitabapplication;1"]
//           .getService(Components.interfaces.IApplication);

const thePrefs =
Components.classes["@mozilla.org/preferences-service;1"]
.getService(Components.interfaces.nsIPrefBranch);

const theEnv = Components.classes["@mozilla.org/process/environment;1"]
.getService(Components.interfaces.nsIEnvironment);

// ------------------------------ Env. variables for directories ------------------------------

// Marche pas hihihihihihihihi
// From Update_2022_10/scripts/k10_public.js - function getSpecialDirectory()
// From Update_2022_10/scripts/k10_public.js - function getSpecialPath()
// See GBV.js for more information
// This was probably OCLC's
// function getSpecialPath(theDirName, theRelativePath){
// 	//gibt den Pfad als String aus
// 	const nsIProperties = Components.interfaces.nsIProperties;
// 	var dirService = Components.classes["@mozilla.org/file/directory_service;1"]
// 							.getService(nsIProperties);
// 	var theFile = dirService.get(theDirName, Components.interfaces.nsILocalFile);
// 	theFile.appendRelativePath(theRelativePath);
// 	return theFile.path;
// }
// function getSpecialDirectory(name){
// 	const nsIProperties = Components.interfaces.nsIProperties;
// 	var dirService = Components.classes["@mozilla.org/file/directory_service;1"].getService(nsIProperties);
// 	return dirService.get(name, Components.interfaces.nsIFile);
// }

//theEnv.set("WINIBW_ProfD", getSpecialDirectory("profD"));
//theEnv.set("WINIBW_BinDir", getSpecialDirectory("BinDir"));
var getThoseDamnSpecialPaths = utility.newFileInput();
theEnv.set("WINIBW_dwlfile", getThoseDamnSpecialPaths.getSpecialPath("dwlfile", "a").slice(0, -2));
theEnv.set("WINIBW_prnfile", getThoseDamnSpecialPaths.getSpecialPath("prnfile", "a").slice(0, -2));
theEnv.set("WINIBW_BinDir", getThoseDamnSpecialPaths.getSpecialPath("BinDir", "\z").slice(0, -2));
theEnv.set("WINIBW_ProfD", getThoseDamnSpecialPaths.getSpecialPath("ProfD", "\z").slice(0, -2));
getThoseDamnSpecialPaths.close();

// ------------------------------ Load scripts ------------------------------

const alpScripts = ["alp_scripts/js/NE_PAS_DIFFUSER.js",
"alp_scripts/js/peyrat_ressources.js",
"alp_scripts/js/GBV.js",
"alp_scripts/js/peyrat_main.js",
"alp_scripts/js/SCOOP.js",
"alp_scripts/js/peyrat_peb.js",
"alp_scripts/python-winibw/pythWinIBW.js",
"alp_scripts/python/python.js",
"alp_xul/xul_test.js"];

for (var ii = 0; ii < alpScripts.length;ii++){
	application.writeProfileString("ibw.standardScripts", "script.AlP"+ii, "resource:"+theEnv.get("WINIBW_ProfD").replace(theEnv.get("WINIBW_BinDir"), "").replace("\\", "/")+"/"+alpScripts[ii]);
}