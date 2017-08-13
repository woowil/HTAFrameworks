// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

var oFso = new ActiveXObject("Scripting.FileSystemObject");
var oWsh = new ActiveXObject("WScript.Shell");
var oWno = new ActiveXObject("WScript.Network");
var oShl = new ActiveXObject("Shell.Application");
var oXml = new ActiveXObject("Microsoft.XMLDOM");
	oXml.async = false;
	oXml.preserveWhiteSpace = false;
	oXml.setProperty("SelectionLanguage","XPath");

function HUILoader(){
	try{
		var base = []		
		base.push("source/library/HUI/Base/HUIPrototypes.js")
		base.push("source/library/HUI/Base/HUICore.js")
		var script = document.createElement("script")
		script.language = "JScript.Encode"
		script.type = "text/javascript"
		script.charset = "UTF-8" //"ISO-8859-1"
		var oHead = document.getElementsByTagName("HEAD")[0]
		for(var i = 0; i < base.length; i++){
			element = script.cloneNode(true)
			element.id = "JSLoad-" + document.scripts.length
			element.src = base[i]
			oHead.appendChild(element)
		}
		
		return true
	}
	catch(e){
		alert("HUILoader(): " + e.description)
		return false
	}
}