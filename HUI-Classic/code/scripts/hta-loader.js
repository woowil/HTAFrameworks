// nOsliw Solutions HUI, HTML Application Library https://github.com/woowil/HTAFrameworks
//**Start Encode**

/*

File:     hta-loader.js
Purpose:
Author:   Woody Wilson
Created:  2006-04
Version:  nope


Disclaimer:
This sample code is provided AS IS WITHOUT WARRANTY
OF ANY KIND AND IS PROVIDED WITHOUT ANY IMPLIED WARRANTY
OF FITNESS FOR PURPOSES OF MERCHANTABILITY. Use this code
is to be undertaken entirely at your risk, and the
results that may be obtained from it are dependent on the user.
Please note to fully back up files and system(s) on a regular
basis. A failure to do so can result in loss of data or damage
to systems.

*/

try{
	if(!oFso || !oWsh || !oWno){
		var oFso = new ActiveXObject("Scripting.FileSystemObject");
		var oWsh = new ActiveXObject("WScript.Shell");
		var oWno = new ActiveXObject("WScript.Network");
	}
}
catch(ee){
	var oFso = new ActiveXObject("Scripting.FileSystemObject");
	var oWsh = new ActiveXObject("WScript.Shell");
	var oWno = new ActiveXObject("WScript.Network");
}

// AJAX
// http://www.c-point.com/javascript_tutorial/Editor/ajax_tutorial.htm
// 

var loader_current_file

function hta_loader_onerror(m,u,l){
	//if(!Window.prototype.closeonerror)
	var s = "Script error\n"
	s = s + "\nMessage:\t" + m
	s = s + "\nUrl:\t" + u
	s = s + "\nLine:\t" + l
	try{
		s = s + "\nFile:\t" + loader_current_file
	}
	catch(e){}
	s = s + "\n\n" + hta_loader_onerror.caller
	s = s + "\n\nClose Window?"
	if(confirm(s)){
		window.setTimeout("window.close()",100)
	}
	return true
}

function hta_loader_load(bLibOnly){
	try{
		String.prototype.isSearch = function(oRe){
			return this.valueOf().search(oRe) > -1
		}

		hta_loader_loadscripts(bLibOnly)
	}
	catch(e){
		//js_log_error(2,e);
		return false;
	}
}

function hta_loader_loadscript(sFile){
	try{ // This function dynamically loads an external script file
		if(typeof sFile != "string" && sFile.length > 3) return;
		loader_current_file = sFile = hta_loader_trim(sFile) // must do this!!
		if(sFile.search(/.+\.(?:js|jse)$|.+\.(?:vbs|vbe)/ig) == -1) return;
		var oScript = document.createElement("script");
		oScript.id = "AJAX-" + document.scripts.length
		//, oScript.type = "text/javascript" // "application/x-javascript"
		if(sFile.search(/\.(?:js|jse)$/ig) > -1) oScript.language = "JScript.Encode"
		//, oScript.type = "text/vbscript"
		else oScript.language = "VBScript.Encode"
		oScript.charset = "ISO-8859-1"
		oScript.src = sFile;
		var head = document.getElementsByTagName('head')[0]
		//document.body.appendChild(oScript);
		head.appendChild(oScript);
		delete oScript
	}
	catch(e){
		alert("hta_loader_loadscript() - " + e.description + " " + sFile)
	}
}

function hta_loader_loadscripts(bLibOnly){
	window.onerror = hta_loader_onerror;
	try{
		var sFile = oFso.GetSpecialFolder(2) + "\\hta_loader_loadscripts-" + Math.ceil(Math.random()*1000 + 1) + ".log"
		var oRe = new RegExp("\.svn|depicted|debug","ig") // folder or files to ignore
		var oReMain = /.*(library-js_[1-2]\.js)/ig
		var sDir = oFso.GetAbsolutePathName(".\\").replace(/(.+)\\$/,"$1")
		var sPath = bLibOnly ? "code\\HUI-1.1.X" : "code"
		var sCmd = "%comspec% /c dir \"" + sDir + "\\" + sPath + "\\*.js*\" \"" + sDir + "\\" + sPath + "\\*.vb*\" /b /ON /s | find /i /v \"depicted\" | find /i /v \"-mof\" | find /i /v \"-customer\" | find /i /v \"ajax\">" + sFile
		oWsh.Run(sCmd,0,true)

		var oFile = oFso.OpenTextFile(sFile,1,false,-2)
		var a = (oFile.ReadAll()).split("\n")
		oFile.close()

		var b = new Array()
		for(var i = 0, len = a.length; i < len; i++){
			a[i] = (a[i]).replace(/\n|\t|\r/ig,"")
			if((a[i]).search(oRe) == -1){ // if not on ignore list
				b.push("." + ((a[i]).replace(sDir,"")).replace(/\\/g,"/"))
				var idx = b.length-1
				if((b[idx]).search(oReMain) > -1){
					hta_loader_loadscript(b[idx]) // The main file should run first
				}
			}
			delete a[i]
		}
		for(var i = 0, len = b.length; i < len; i++){
			if((b[i]).search(oReMain) == -1) hta_loader_loadscript(b[i])
			delete b[i]
		}
		a.length = b.length = 0
		delete a, delete b
	}
	catch(e){
		alert("hta_loader_loadscripts() - " + e.description)
	}
	finally{
		oFso.DeleteFile(sFile,true)
		loader_current_file = ""
		window.onerror = null
	}
}

function hta_loader_refresh(){
	try{
		var bOK = false
		//alert(document.scripts.length)
		hideContextMenus()
		hta_refresh()
		return;
		for(var i = document.scripts.length-1; i >= 0; i--){
			var a = (document.scripts)[i]
			if(typeof(a.src) == "string" && (a.id).search(/_loader_/ig) > -1){
				hta_loader_removescript(a.id)
				bOK = true
			}
		}
		if(bOK) hta_loader_loadscripts()

	}
	catch(e){
		alert("hta_loader_refresh() - " + e.description)
		return false;
	}
}

function hta_loader_removescript(sScriptID){
	if(!sScriptID || typeof(sScriptID) != "string") return;
	try{
		var head = document.getElementsByTagName('head')[0]
		head.removeChild(head.getElementById(sScriptID));
	}
	catch(e){
		//alert("hst_loader_removescript() - " + e.description)
	}
	try{
		document.body.removeChild(document.getElementById(sScriptID));
	}
	catch(e){
		//alert("hst_loader_removescript() - " + e.description)
	}
}

function hta_loader_trim(s){
	try{
		var ii
		s = s.replace(/([ \t\n]*)(.*)/g,"$2"); // Removes space at beginning
		if((ii = s.search(/([ \t\n]*$)/m)) > 0) s = s.substring(0,ii);
		return s.replace(/[ \t]+/ig," "); // Replaces tabs, double spaces
	}
	catch(ee){
		return "";
	}
}
