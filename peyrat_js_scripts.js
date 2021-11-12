function AlP_js_removeAccents(str){
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
}

function AlP_js_triPeb(){
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
