//In WinIBW's folder, defaults/pref/setup.js, add at the top of the file (uncommented) :
//pref("ibw.standardScripts.script.pythWinIBW", "resource:/SCOOP/scripts/python-winibw/pythWinIBW.js");

// To define parameters for both python scripts and WinIBW scripts, edit the python_parameters file.
// It is build as a JSON file but it doesn't support tables / lists or object / dictionaries inside the main one.
// A such, every key / value pair should be in the same object.
// Other restrictions:
//  - keys shouldn't have a ":" inside them,
//  - strings values must be declared between "", not '',
//  - python_path, pyth_to_js_file, js_to_pyth_file and clean_pyth_files must exist.
// Other than that, you're free to add anything to it.

function __get_python_parameters(){
	// Stringifies the file (in UTF-8)
	var pythParameters = Components.classes["@oclcpica.nl/scriptinputfile;1"]
		.createInstance(Components.interfaces.IInputTextFile);
	pythParameters.openSpecial("ProfD", "\\alp_scripts\\python-winibw\\python_parameters");
	var doc = ""
	while (pythParameters.isEOF() === false){
		doc += pythParameters.readLine();
	}
	pythParameters.close();
	// Splits the stringified object in different pairs
	doc = doc.substr(doc.indexOf("{")+1, doc.indexOf("}")-1);
	doc = doc.split(",");
	// Creates the object
	var param = {};
	for(ii=0; ii<doc.length; ii++){
		var key = doc[ii].substr(0, doc[ii].indexOf(":"));
		key = key.substr(key.indexOf("\"")+1, key.indexOf("\"", key.indexOf("\"")+1)-1);
		var value = doc[ii].substr(doc[ii].indexOf(":")+1);
		if (value.indexOf("\"") != -1){ // If the value is a string
			value = value.substr(value.indexOf("\"")+1, value.indexOf("\"", value.indexOf("\"")+1)-1);
		}else{ // If the value is a bool / number
			value = value.replace(/\s/gm, "");
		}
		param[key] = value;
	}
	return param;
}

function __execute_python(pyModPath){
	application.shellExecute(pythPar["python_path"], 5, "open", pyModPath);
}

function __clean_python_temp_file(){
	__execute_python(pythPar["clean_pyth_files"]);	
}

function __is_valid_for_Python_WinIBW(key){
	if(!(key in pythPar)){
		missingParams.push(key + " is not defined");
	}else if(key!="pyth_to_js_file" && key != "js_to_pyth_file"){
		var file = Components.classes["@oclcpica.nl/scriptinputfile;1"]
		.createInstance(Components.interfaces.IInputTextFile);
		exists = file.open(pythPar[key]);
		file.close()
		if (!exists){
			missingParams.push(key + " does not point to an existing file");	
		}
	}
}

const pythPar = __get_python_parameters(); // Defines pythPar as a global variable
// Checks on WinIBW's initialisation if Python-WinIBW can be used.
// /!\ Does not check if pyth_to_js_file and js_to_pyth_file are valid paths.
missingParams = []
__is_valid_for_Python_WinIBW("python_path");
__is_valid_for_Python_WinIBW("pyth_to_js_file");
__is_valid_for_Python_WinIBW("js_to_pyth_file");
__is_valid_for_Python_WinIBW("clean_pyth_files");
// if(!("python_path" in pythPar)){missingParams.push("python_path");};
// if(!("pyth_to_js_file" in pythPar)){missingParams.push("pyth_to_js_file");};
// if(!("js_to_pyth_file" in pythPar)){missingParams.push("js_to_pyth_file");};
// if(!("clean_pyth_files" in pythPar)){missingParams.push("clean_pyth_files");};
if(missingParams.length > 0){
	var missingParamsErr = "The following parameters are not / wrongfully defined inside python_parameters:"
	for(ii=0;ii<missingParams.length;ii++){
		missingParamsErr += "\n - "+ missingParams[ii];
	}
	missingParamsErr += "\n\nPlease refrain from using any script using Python-WinIBW."
	missingParamsErr += "\n\n\nIf you can not find the problem, please refer yourself to the documentation on my GitHub:"
	missingParamsErr += "\nhttps://github.com/Alban-Peyrat/WinIBW/blob/main/python-winibw.md"
	missingParamsErr += "\nOr contact me directly."
	application.messageBox("Failed to initialize Python-WinIBW", missingParamsErr, "error-icon")
}

// How to use:
// [First time] Open SCOOP\scripts\python_parameters inside the WinIBW files and edit the files path (and add optionnal parameters if you wish to).
// [First time] Verify in get_python_parameters.py if the absolute path to python_parameters is correct.
// First, execute a JS WinIBW script.
// If you need to transfer data to the python script, write it in the temp_js_to_pyth file using WinIBW's FileOutput object.
// Then, execute the __execute_python() script with the absolute path to the python module as a parameter.
// In your python module, import get_python_parameters and store its main() returned value in a variable.
// Then, write any data you wish to transfer to WinIBW to the temp_pyth_to_js file.
// Back in the JS script (you might need to pause WinIBW to let the python script fully execute [untested]), retrieve the data using WinIBW's FileInput object.
// At the end, run __clean_python_temp_file() to deleted the two temporary files.