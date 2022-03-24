//In WinIBW's folder, defaults/pref/setup.js, add at the top of the file (uncommented) :
//pref("ibw.standardScripts.script.python", "resource:/SCOOP/scripts/python/python.js");

// If this const isn't already declared, declare it
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

function pythonWinIBWexample(){
	// Gets the type of document
	var isAut = application.activeWindow.getVariable("P3VMC");

	// Checks if the script can properly execute
	if (isAut == ""){
		application.messageBox("Python-WinIBW example failed", "Please select a record before executing the script.", "error-icon")
		return
	}else if (isAut.charAt(0) == "T"){
		application.messageBox("Python-WinIBW example failed", "This is record is an authority record, not a document record.", "error-icon")
		return
	}

	// Gets the PPN
	var PPN = application.activeWindow.getVariable("P3GPP");

	// Writes the PPN in js_to_pyth_file
	var jsToPyth = utility.newFileOutput();
	jsToPyth.setTruncate(true); // The path will be created in writing-mode and not in append-mode
	jsToPyth.create(pythPar["js_to_pyth_file"]);
	jsToPyth.write(PPN);
	jsToPyth.close();

	// Executes the python script
	__execute_python("C:\\oclcpica\\WinIBW30\\SCOOP\\scripts\\python\\python-winibw_example.py");

	// Defines the FileInput
	var pythToJs = utility.newFileInput();
	
	// Waits for pyth_to_js_file to exist AND to be opened 
	var maxWait = 10; // in sec
	var success = false;
	var ii = 0;
	while((ii < maxWait*2) && (success == false)){
		try {
			success = pythToJs.open(pythPar["pyth_to_js_file"]); // Opens pyth_to_js_file
		}
		catch (e){}
		ii++;
		__sleep(500) // This waits for 0,5 sec
	}
	if(ii==maxWait){ // This only occurs if the file could not be opened
		application.messageBox("Python-WinIBW example failed", "WinIBW could not open pyth_to_js_file after "+maxWait+" seconds. The script will delete temporary files and stop.", "error-icon");
		__clean_python_temp_file();
		return
	}
		
	// Retrieves the data returned by the python script	
	var record = "";
	while(!pythToJs.isEOF()){
		record += pythToJs.readLine()+"\n";
	}
	pythToJs.close();
	
	// Display the data
	application.messageBox("Python-WinIBW example result", record, "message-icon");

	// Deletes the temporary files
	__clean_python_temp_file();
}

//Réponse de BeNdErR à : https://stackoverflow.com/questions/16873323/javascript-sleep-wait-before-continuing
function __sleep(milliseconds) {
  var start = new Date().getTime();
  for (var i = 0; i < 1e7; i++) {
    if ((new Date().getTime() - start) > milliseconds){
      break;
    }
  }
}