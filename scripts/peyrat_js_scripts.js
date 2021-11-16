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
