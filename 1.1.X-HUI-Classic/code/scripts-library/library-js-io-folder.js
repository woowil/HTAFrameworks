// nOsliw Solutions - HTML Aplication Framework.
//**Start Encode**

/*

File:     library-js-io-folder.js
Purpose:  Development script
Author:   Woody Wilson
Created:  08.07.2002
Version:  
Dependency:  library-js.js, library-js-io-file.js

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
	if(!oFso || !oWsh || !oWno || !oXml){
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

////////////////////////////////////////////////////////////////////////////////////////////////
/////// FOLDER SYSTEM FUNCTIONS
////////////////////////////////////////////////////////////////////////////////////////////////



function io_folder_deletesub(sFolder,bNotForce){
	try{
		if(io_folder_path(sFolder)){
			var oFolder = oFso.GetFolder(sFolder), bOk = true, sCmd;		
			var bForce = bNotForce ? false : true
			var oSubFolders = new Enumerator(oFolder.SubFolders);
			for(; !oSubFolders.atEnd(); oSubFolders.moveNext()){
				bOk = io_folder_deletesub(oSubFolders.item());
			}
			try{
				sCmd = "%comspec% /c attrib -r -h /s \"" + sFolder + "\\*.*";
				oWsh.Run(sCmd,oReg.hide,true);
				oFolder.Delete(bForce);
				return bOk;
			}
			catch(ee){
				return false;	
			}			
		}
		return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_folder_delete(sFolder,bNotForce){
	try{
		if(io_folder_path(sFolder)){
			var sCmd = "%comspec% /c attrib -r -h \"" + sFolder + "\\*.*";
			oWsh.Run(sCmd,oReg.hide,true);			
			try{
				oFso.DeleteFolder(sFolder,true);
				return true;
			}
			catch(ee){
				return false;	
			}			
		}
		return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_folder_readable(sFolder){
	try{
		if(oFso.FolderExists(sFolder)){
			return (oWsh.Run("%compspec% /c dir /b \"" + sFolder + "\" >nul",oReg.hide,true) == 0)	
		}
		return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_folder_list_(sParentFolder,sRegFolder){ // Use this function on remote folder list
	var sFile = oReg.temp + "\\io_folder_list.log"
	try{
		if(sParentFolder.match(/[a-g]:.*$/i)) return io_folder_list2(sParentFolder,sRegFolder) // io_folder_list2 function is faster locally
		var aFolders = new Array();		
		if(oFso.FolderExists(sParentFolder)){
			if(oWsh.Run("%comspec% /c dir /b /ad \"" + sParentFolder + "\" | sort>" + sFile,oReg.hide,true) == 0){
				var s, oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault)
				s = oFile.ReadAll(), oFile.close()
				aFolders = (s.replace(/(.+)$/gm,sParentFolder + "\\" + "$1\n")).split(/[\n]+/igm)
				if(typeof(sRegFolder) == "string"){
					var oRe = new RegExp(sRegFolder,"ig"), a = new Array()
					for(var i = 0, len = aFolders.length; i < len; i++){
						var f = new String(aFolders[i]); // Bug
						if((f).match(oRe)) a.push(f)
					}
					aFolders.length = 0
					return a
				}
			}
		}
		return aFolders;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		io_file_delete(sFile)
	}
}

function io_folder_list(sParentFolder,sRegFolder){ // Use this function on system folders
	try{
		var aSubFolders = new Array();
		if(oFso.FolderExists(sParentFolder)){
			var oRe = sRegFolder ? new RegExp(sRegFolder,"ig") : false;
			var oParentFolder = oFso.GetFolder(sParentFolder);
			var oFolders = new Enumerator(oParentFolder.SubFolders);
		   	for(var i = 0; !oFolders.atEnd(); oFolders.moveNext(), i++){
				var sFolder = new String(oFolders.item());
				if(oRe && sFolder.match(oRe)) aSubFolders.push(sFolder);
				else aSubFolders.push(sFolder);
			}
		}
		return aSubFolders.sort();
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_folder_createsub(sSubFolder,bError){
	try{
		if(oFso.FolderExists(sSubFolder)) return true;
		if(oWsh.Run("%comspec% /c md \"" + sSubFolder + "\"",oReg.hide,true) == 0) return true
		var oRe = /(\\\\[a-z0-9_\-]+\\[a-z]\$\\)(.*)/ig // if folder is UNC admin share
		sUNC = sSubFolder.match(oRe) ? RegExp.$1 : ""		
		var aSubFolder = (sUNC == "" ? sSubFolder.split(/\\/g) : (RegExp.$2).split(/\\/g));
		for(var i = 0, sFolder = "", len = aSubFolder.length; i < len; i++){
			sFolder = sFolder + aSubFolder[i] + "\\";
			if(!oFso.FolderExists(sFolder)){
				oFso.CreateFolder(sUNC + sFolder);
			}
		}
		return true;
	}
	catch(e){
		if(bError) js_log_error(2,e);
		return false;
	}
}

function io_folder_create(sFolderStream){
	try{
		var bOk = true
		for(var i = 0, len = arguments.length; i < len; i++){
			sFolder = arguments[i];
			if(!oFso.FolderExists(sFolder)){
				//sCmd = "%comspec% /c mkdir \"" + sFolder + "\""
				//if(oWsh.Run(sCmd,oReg.hide,true) != 0) return false
				if(!io_folder_createsub(sFolder)) bOk = false
			}
		}
		return bOk
	}
	catch(e){
		js_log_error(2,e); 
		return false;
	}
}
