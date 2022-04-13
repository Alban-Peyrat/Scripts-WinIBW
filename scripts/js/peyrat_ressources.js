// Scripts Ressources

function __AbesDelTitleCreated(){
//See Abes' standart_copy.js
// But reworked to only delete the top line of P3CLIP var
  //suptag("Cr\u00E9"); // If the file is not encoded in Wetsern Windows 1252
  suptag("Cré");
}

function __AbesDelItemData(){
//See Abes' standart_copy.js
//But reworked to delete the every information about items
//According to this, it should delete every information about items : http://documentation.abes.fr/sudoc//formats/loc/index.htm
  suptag("A");
  suptag("9");
  suptag("E");
  suptag("e");
}

function __addTextToVar(vari, text, sep){
// Returns the variable appended with the separator and the text
// Unless if the variable is empty, which only returns only text
//sep = the wanetd separator
  if (vari == ""){
    return text
  }else{
    return "".concat(vari, sep, text)
  }
}

function __connectBaseProd() {
// Connects to Sudoc production database
  application.connect("nacarat.sudoc.abes.fr", "1040-1055")
}

function __connectBaseTest() {
// Connects to Sudoc test database
  application.connect("cramoisi.sudoc.abes.fr", "1100")
}

function __createWindow(){
//Normally the new window becomes the active window
  application.newWindow();
}

function __dateToYYYYMMDD_HHMM(date){
// Returns the date as a string in YYYYMMDD_HHMM format
  var day = date.getDate();
  var month = date.getMonth()+1;
  var year = date.getFullYear();
  var hour = date.getHours();
  var min = date.getMinutes();
  if (month<10){
    month = "0"+month;
  }
  if (day<10){
    day = "0"+day;
  }
  if (hour<10){
    hour = "0"+hour;
  }
  if (min<10){
    min = "0"+min;
  }
  return "".concat(year, month, day, "_", hour, min)
}

function __deconnect(){
//Closes all windows to be safe
  var nbWin = application.windows.count;
  for(var ii = 0; ii < nbWin; ii++){
    //Quand une fenêtre se ferme, l'index descend automatiquement
    application.windows.item(0).close();
  }
}

//voir Abes parce que sinon ça dysfonctionne
function __findExactText(txt){
// A TESTER
  application.activeWindow.title.startOfBuffer(false);
  application.activeWindow.title.find(txt, true, false, false);
}

function __getEnvVar(varName){
// Returns the variable value or returns false if the variable doesn't exists
  if (theEnv.exists(varName)){
    return theEnv.get(varName)
  } else {
    return false
  }
}

function __getNoticeType(){
// Returns 0 if it's an authority record, 1 a bibliographic record, 2 not a record
  var isAut = application.activeWindow.getVariable("P3VMC");

  if (isAut == ""){
    var scrCode = application.activeWindow.getVariable("scr");
    if (scrCode == "II"){ // Invoer Ingang
      // Creating an Authority record
      return 0
    }else if (scrCode == "IT"){ // Invoer Titel
      // Creating a bibliographic record
      return 1
    }else {
      // Supposedly, every other option is covered
      return 2
    }
  }else if (isAut.charAt(0) == "T"){
    return 0
  }else {
    return 1
  }
}

function __hasWarningMsg(){
// Returns all warnings messages as a string separated by a semi-colon.
// Returns an empty string if there are no warning messages
  var output = "";
  for (ii = 0; ii < application.activeWindow.messages.count; ii++) {
    msg = application.activeWindow.messages.item(ii);
    if (msg.type == 2){
      output =  __addTextToVar(output, msg.text, ";");
    }
  }
  return output
}

//voir Abes parce que sinon ça dysfonctionne
function __insertText(txt){
// A TESTER
  application.activeWindow.title.endOfBuffer(false);
  application.activeWindow.title.insertText(txt);
}

function __isTitle(){
// Returns true or false depending on if the active window has a title
  if(!application.activeWindow.title){
    return false
  }else{
    return true
  }
}

function __logIn(identifiants){
// Logs in
// identifiants = "identifiant motDePasse"
  application.activeWindow.command("\\log "+identifiants, false)
}

function __parseDocLine(line){
// Returns an array of line splitted by horizontal tabulations
  return line.split("\u0009")
}

function __removeAccents(str){
// A revoir
//rip normalize
  str = str.replace(/[\u00EA\u00E9\u00EB\u00E8]/g, "e");
  str = str.replace(/[\u00C8\u00C9\u00CA\u00CB]/g, "E");
  str = str.replace(/[\u00E0\u00E1\u00E2\u00E3\u00E4\u00E5]/g, "a");
  str = str.replace(/[\u00C0\u00C1\u00C2\u00C3\u00C4\u00C5]/g, "A");
  str = str.replace(/[\u00EC\u00ED\u00EE\u00EF]/g, "i");
  str = str.replace(/[\u00CC\u00CD\u00CE\u00CF]/g, "I");
  str = str.replace(/[\u00F9\u00FA\u00FB\u00FC]/g, "u");
  str = str.replace(/[\u00D9\u00DA\u00DB\u00DC]/g, "U");
  str = str.replace(/[\u00F2\u00F3\u00F4\u00F5\u00F6]/g, "o");
  str = str.replace(/[\u00D2\u00D3\u00D4\u00D5\u00D6]/g, "O");
  str = str.replace(/\u00DD/g, "y");
  str = str.replace(/\u00FD/g, "Y");
  str = str.replace(/\u00E7/g, "c");
  str = str.replace(/\u00C7/g, "C");
  str = str.replace(/\u0153/g, "oe");
  str = str.replace(/\u0152/g, "OE");
  str = str.replace(/\u00E6/g, "ae");
  str = str.replace(/\u00C6/g, "AE");
  return str;
};

function __serializeArray(vari, sep){
// Returns the array as a string separated by sep
  output = "";
  for(ii=0;ii<vari.length;ii++){
    output += vari[ii] + sep;
  }
  return output.substr(0, (output.length - sep.length))
}

//Réponse de BeNdErR à : https://stackoverflow.com/questions/16873323/javascript-sleep-wait-before-continuing
function __sleep(milliseconds) {
// Sleeps the script execution for X milliseconds
  var start = new Date().getTime();
  for (var i = 0; i < 1e7; i++) {
    if ((new Date().getTime() - start) > milliseconds){
      break;
    }
  }
}

function __timerToReal(start, end){
// Return the difference bewteen start and end as "X minute(s) X seconde(s)"
  var sec = (end-start)/1000;
  var min = Math.floor(sec/60);
  sec = sec%60;
  return min+" minute(s) "+sec + " seconde(s)"
}