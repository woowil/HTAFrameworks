// nOsliw Solutions HUI, HTML Application Library https://github.com/woowil/HTAFrameworks
//**Start Encode**

/*

File:     library-js-class-file-compression.js
Purpose:  Development script
Author:   Woody Wilson
Created:  2006-04
Version:
Dependency:  library-js.js, library-js__prototype.js

Description:
JScript library scripting functions used for WSH files or HTA Applications.

Usage:

Description:
JScript library scripting functions used for WSH files or HTA Applications.

Revisions: to many

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


// Global Scripting Objects
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


// Function _Class() must be loaded
class_file_encryption.prototype = new _Class("Encryption","Class for encrypting javascript/jscript files")

function class_file_encryption(bStatus,bDebug){
	try{
		this.isStatus = bStatus ? true : false
		this.isDebug = bDebug ? true : false
		var files = new ActiveXObject("Scripting.Dictionary")		
		
		this.isScript = function(sFile){
			var s
			if(oFso.FileExists(sFile) && (s = oFso.GetExtensionName(sFile))){
				return s.match(/js|jse/i)
			}
			return false
		}
		
		this.encryptFile = function(sFile){
			if(this.isScript(sFile)){
				var oFile = oFso.OpenTextFile(sFile,this.fso.read,false,this.fso.TristateUseDefault)
				var sStream = oFile.ReadAll()
				oFile.close()
				sStream = this.encryptStream(sStream)				
				return this.setStreamOnFile(sFile,sStream)
			}
			return false
		}
		
		this.encryptStream = function(sStream){
			//TODO: 
			return sStream
		}
		
		this.setStreamOnFile = function(sFile,sStream){
			return io_file_create(sFile,sStream)
		}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}
