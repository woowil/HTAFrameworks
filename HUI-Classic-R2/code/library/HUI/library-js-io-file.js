// nOsliw Solutions HUI, HTML Application Library http://hui.nosliw.info, License http://hui.codeplex.com/license
//**Start Encode**

/*

File:     library-js-io-file.js
Purpose:  Development script
Author:   Woody Wilson
Created:  08.07.2002
Version:  
Dependency:  library-js.js, library-js-io-folder.js

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

function io_file_shutcut(sUrl,sName,sFolder){
	try{
		sFolder = sFolder ? sFolder : "Desktop"
	    var sTemp, oUrlLink, sDesktop
	    sDesktop = oWsh.SpecialFolders(sFolder);
	    sTemp = sDesktop + "\\" + sName + ".url"
	    var oUrlLink = oWsh.CreateShortcut(strTemp)
	    oUrlLink.TargetPath = strUrl
	    oUrlLink.Save
	    //oWsh.Exec(sTemp, false)
    }
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_file_html(sFolder){
	try{
		sFolder = sFolder ? sFolder : io_folder_path();
		var oFile, sSupport = "";
		var aFiles = io_file_list(sFolder,new RegExp("js|vbs|css|pl|hta","ig"));		
		var sHd = '<table border="0" cellspacing="0" cellpadding="0" class="cSupport">', sFt = '\n</table>';
		for(var oUtc, i = 0, len = aFiles.length; i < len; i++){
			var oFile = io_file_info(aFiles[i].file);			
			sSupport += '\n<tr><td width="80">File name </td><td><b>' + oFile.name + '</b></td></tr>';
			sSupport += '\n<tr valign="top"><td>File path </td><td>' + oFile.path + '</td></tr>';
			//oUtc = new js_tme_utc(oFile.created,1);
			sSupport += '\n<tr><td>Created Date </td><td>' + oFile.created + '</td></tr>';
			//oUtc = new js_tme_utc(oFile.lastmod,1);
			sSupport += '\n<tr><td>Modified Date </td><td>' + oFile.lastmod + '</td></tr>';
			sSupport += '\n<tr><td>File Size </td><td>' + oFile.size + ' bytes</td></tr>';
			sSupport += '\n<tr><td colspan="2">&nbsp;</td></tr>';
		}		
		sSupport = sHd + sSupport + sFt;
		return sSupport;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_file_read(sFile,bError,bNotArray,sRegExp){
	try{
		var aStream = false;
		if(oFso.FileExists(sFile)){
			aStream = new Array()
			var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
			if(!bNotArray){
				var a = (oFile.ReadAll()).split("\n");
				for(var i = 0; i < a.length; i++){
					a[i] = (a[i]).replace(/\n|\t|\r/ig,"")
					if((a[i]).search(/[a-z0-9_]+/ig) == -1) continue
					aStream.push(js_str_trim(a[i]))
					//if((aStream[i]).length == 0) (aStream[i]).splice(i,1)
				}
				js_str_kill(a)
			}
			else {
				var oRe = new RegExp("("+ sRegExp + ")","ig")
				while(!oFile.AtEndOfStream){
					var sLine = oFile.ReadLine();
					if(sRegExp){
						if(sLine.search(oRe) > -1) aStream.push(RegExp.$1);
					}
        			else aStream.push(sLine);
				}
			}
			oFile.Close();
		}
		return aStream;
	}
	catch(e){
		if(bError) js_log_error(2,e);
		return false;
	}
}

function io_file_append(sFile,sStream,bDelete,bAppend,sSep,bNotIter){
	var oFile = null
	try{
		if(typeof(sFile) != "string") return false
		if(bDelete) io_file_delete(sFile);
		if(!oFso.FolderExists(oFso.GetParentFolderName(sFile))) io_folder_create(oFso.GetParentFolderName(sFile))				
		var sSep = !sSep ? "\n" : sSep;
		var IOmode = bAppend ? oReg.append : oReg.write;
		oFile = oFso.OpenTextFile(sFile,IOmode,true,oReg.TristateUseDefault);
		if(typeof(oFile) == "object"){
			var aStream = sStream.split(sSep); // Just remember to separate your stream with a Newline character
			for(var i = 0, len = aStream.length; i < len; i++){				
				if((aStream[i]).search(/^\[ \t]*$/g) == -1) oFile.WriteLine(aStream[i]); // If not blank line
				else oFile.WriteBlankLines(1)
			}
			js_str_kill(aStream)
			return sFile;
		}
		return false
	}
	catch(e){
		var s = (e.description).toString(), n = (e.number & 0xFFFF)
		if(s == "Path Not found" && n == 76){ // Folder maybe does not exist
			if(!bNotIter){
				if(io_folder_create(oFso.GetParentFolderName(sFile))){
					return io_file_append(sFile,sStream,false,false,sSep,true)
				}
			}
		}
		else if(s == "Invalid procedure call or argument"){ 
			// I got this once. The problem was that there were UNICode characters in sStream.
			// Note: TristateUseDefault is for ASCII and TristateUse is for UNICODE
			
		}
		if(s == "Unknown name" && n == 10){
			// The file is probably using UNICODE charactors.. maybe should use 'oReg.TristateUse'
		}
		js_log_error(2,e);
		return false;
	}
	finally{
		try{
			if(typeof(oFile) == "object"){
				oFile.Close();
			}			
		}
		catch(ee){}
		js_str_kill(oFile);
	}
}

function io_file_create(sFile,sStream,bNotIter){
	var oFile;
	try{
		if(typeof(sFile) != "string") return false
		io_file_delete(sFile);
		oFso.CreateTextFile(sFile,true,oReg.TristateUseDefault);
		oFile = oFso.OpenTextFile(sFile,oReg.write,true,oReg.TristateUseDefault);
		oFile.WriteLine(sStream);
		return true;
	}
	catch(e){
		var s = (e.description).toString(), n = (e.number & 0xFFFF)
		if(s == "Path Not found" && n == 76){ // Folder maybe doesnæt exist
			if(!bNotIter){
				if(io_folder_create(oFso.GetParentFolderName(sFile))){
					return io_file_append(sFile,sStream,false,false,sSep,true);
				}
			}
		}
		else if(s == "Invalid procedure call or argument"){ 
			// I got this once. The problem was that there were UNICode characters in sStream.
			// Note: TristateUseDefault is for ASCII and TristateUse is for UNICODE
			
		}
		if(s == "Unknown name" && n == 10){
			// The file is probably using UNICODE charactors.. maybe should use 'oReg.TristateUse'
		}
		js_log_error(2,e);
		return false;
	}
	finally{
		try{
			oFile.Close();
		}
		catch(ee){}
		js_str_kill(oFile);
	}
}

function io_file_merge(sFile,sFile1,sFile2,sPar){
	try{
		if(!js_str_isdefined(sFile,sFile1)) return false
		if(io_file_exists(sFile1)){
			if(typeof(sFile2) != "string") sFile2 = sFile1, sFile1 = sFile
			var sPar = typeof(sPar) == "string" ? sPar : ""
			var sCmd = "%comspec% /u /s /c copy /y \"" + sFile1 + "\" + \"" + sFile2 + "\" \"" + sFile + "\" | find /i \"copied\""
			return (oWsh.Run(sCmd,oReg.hide,true) == 0)
		}
		return false
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_file_readable(sFile){
	try{
		if(oFso.FileExists(sFile)){
			return (oWsh.Run("%comspec% /c type \"" + sFile + "\" >nul",oReg.hide,true) == 0)	
		}
		return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_file_random(sFile,hval,lval){
	if(typeof(sFile) != "string") return undefined
	return oFso.GetParentFolderName(sFile)  + "\\" + oFso.GetBaseName(sFile) + "-" + js_str_randomize(lval,hval)	+ "." + oFso.GetExtensionName(sFile)
}

function io_file_attrset(sFile,iBit){
	try{
		if(oFso.FileExists(sFile) && js_str_isnumber(iBit)){
			var oFile = oFso.GetFile(sFile);
			oFile.Attributes = oFile.Attributes + iBit; // iBit is positive/negative integer: +-1(readonly), +-2(hidden), +-4(system) or +-32(archive)
		}
		else return false;
		return oFile.Attributes;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_file_attrget(sFile){
	try{
		var Obj = new Object();
		if(oFso.FileExists(sFile)){
			var oFile = oFso.GetFile(sFile);
			Obj.normal = (oFile.attributes && 0); // Normal file. No attributes are set
			Obj.readonly = (oFile.attributes && 1); // Read-only file. Attribute is read/write.
			Obj.hidden = (oFile.attributes && 2); // Hidden file. Attribute is read/write.
			Obj.system = (oFile.attributes && 4); // System file. Attribute is read/write.
			Obj.volume = (oFile.attributes && 8); // Disk drive volume label. Attribute is read-only.
			Obj.directory = (oFile.attributes && 16); // Folder or directory. Attribute is read-only.
			Obj.archive = (oFile.attributes && 32); // File has changed since last backup. Attribute is read/write.
			Obj.alias = (oFile.attributes && 64); // Link or shortcut. Attribute is read-only.
			Obj.compress = (oFile.attributes && 128); // Compressed file. Attribute is read-only.			
			Obj.attr = oFile.attributes;
			
			// Definitions
			Obj.readonly_on = -1, Obj.readonly_off = 1;
			Obj.hidden_on = -2, Obj.hidden_off = 2;
			Obj.system_on = -4, Obj.system_off = 4;
			Obj.archive_on = -32, Obj.archive_off =-32;
			Obj.enableall = 0, Obj.disableall = 39;
			// Reset to zero
			
		}
		else return false;
		return Obj;
	}
	catch(e){
		return false;
	}
}

function io_file_info(sFile){
	try{
		var oFile = false;
		if(oFso.FileExists(sFile)){
			oFile = new Object();			
   			oFile.name = oFso.GetFileName(sFile); // file.txt
   			oFile.base = oFso.GetBaseName(sFile); //file
   			oFile.ext = (oFso.GetExtensionName(sFile)).toLowerCase(); // txt
   			oFile.path = oFso.GetAbsolutePathName(sFile); // c:\path\file.txt
   			oFile.parent = oFso.GetParentFolderName(sFile); // c:\path
   			oFile.version = oFso.GetFileVersion(sFile);
   			oFile.attr = io_file_attrget(sFile);
   			var f = oFso.GetFile(sFile);
   			oFile.created = f.DateCreated;
   			oFile.lastmod = f.DateLastModified;
   			oFile.lastacc = f.DateLastAccessed;
   			oFile.size = f.Size // In bytes (If KB, divide this by 1024)
   			oFile.type = f.Type // Ex: Text Document (*.txt)
   			oFile.shortname = f.ShortName
   			oFile.drive = f.Drive
   			oFile.absname = sFile;
   		}
   		return oFile;
	}
	catch(e){
		return false;
	}
}

function io_file_list(sFolder,sFileExt){
	try{
	 	var aFiles = new Array();
	 	if(oFso.FolderExists(sFolder = oFso.GetAbsolutePathName(sFolder))){
			var oFolder = oFso.GetFolder(sFolder); // c:\path or c:\path\ result in c:\path
			var oFiles = new Enumerator(oFolder.Files);
			for(var i = 0; !oFiles.atEnd(); oFiles.moveNext(), i++){
			 	var oFile = oFso.GetFile(oFiles.item()), sFile = oFile.Path;
			 	if(sFileExt && (oFso.GetExtensionName(sFile)).toLowerCase() == sFileExt.toLowerCase()){
			 		aFiles.push(sFile);
			 	}
			 	else aFiles.push(sFile);
			}
		}
		return aFiles.sort();
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_file_listrec(sFolder,sFileExt,sRegFolder,iSub,iSubLimit){
	try{
		var aFiles = new Array(), oFile;
		iSub = js_str_isnumber(iSub) ? iSub : 0;
		iSubLimit = iSubLimit ? iSubLimit : js.fls.SUBLIMIT;
		if(iSub < iSubLimit && oFso.FolderExists(sFolder)){
			var aSubFiles = new Array(), aSubFolders;
			aFiles = aFiles.concat(io_file_list(sFolder,sFileExt));
			aSubFolders = io_folder_list(sFolder,sRegFolder);				
			for(var i = 0, len = aSubFolders.length; i < len; i++){
				aFiles = aFiles.concat(io_file_listrec(aSubFolders[i],sFileExt,sRegFolder,iSub++,iSubLimit));
			}
		}
		return aFiles;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_file_suffix(sPath,sExtension,sSuffix){
	try{
		var aFiles, obj = new Object();
		obj.count = obj.changed = 0;
		if(aFiles = io_file_list(sPath)){ // Folder
			var oFile;
			for(var i = 0, len = aFiles.length; i < len; i++){
				if(oFso.FileExists(aFiles[i]) && (oFile = io_file_info(aFiles[i],sSuffix))){
					if(oFile.ext == sExtension){
						var sCmd = "%comspec% /c cd /d " + oFile.parent + " && ren " + oFile.name + " " + oFile.namesuffix
						if(oWsh.Run(sCmd,oReg.hide,true) == 0) obj.changed++;
						obj.count++;
					}
				}
			}
		}
		return obj;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_file_rename(OrgFile,NewFile){
	try{
		if(io_file_exists(OrgFile,NewFile)){
			var sCmd = "%comspec% /c ren " + OrgFile + " " + NewFile;
			return (oWsh.Run(sCmd,oReg.hide,true) == 0)
		}
		return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

/**
    Author:     Woody Wilson
	Function:   io_file_overdelete()
	Parameters: OrgFile,OverFile
	            OrgFile (string) = Absolute path name of the original file that will be overwritten
	            OverFile (string) = Absolute path name of the file that will overwrite
	Purpose:    Overwrites an original file
	Returns:    Boolean
	Exception:  On Error
*/
function io_file_overdelete(OrgFile,OverFile){
	try{
		if(io_file_exists(OrgFile,OverFile)){
			var oFile = oFso.GetFile(OrgFile);
			oFile.Copy(OverFile,true);
			oFile.Delete(true);
			return true;
		}
		return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_file_exists(filestream){	
	try{
		for(var i = 0, len = arguments.length; i < len; i++){
			if(!oFso.FileExists(arguments[i])){
				return false;
			}
		}
		return true;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_file_delete(filestream){
	try{
		for(var sFile, i = 0, len = arguments.length; i < len; i++){
			try{
				if(oFso.FileExists(sFile = arguments[i])){
					oWsh.Run("%comspec% /c attrib -h -r \"" + sFile + "\">nul & del /f /q \"" + sFile + "\">nul",oReg.hide,true);
					if(oFso.FileExists(sFile)){
						var oFile = oFso.GetFile(sFile); // deletes also empty files
						oFile.Delete();
					}
				}
			}
			catch(ee){}
		}
		return true;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_file_replace(sFile,sRe,sReplace,sChangeFile,sText,sTextErr){
	try{
		var oResult = false;
		if(oFso.FileExists(sFile) && !sChangeFile){ // Requires change.exe
      			var oRe = new RegExp(sRe,"i"), sLine;
      			var oFileTmp = oFso.OpenTextFile(sFile + ".tmp",oReg.write,true,oReg.TristateUseDefault);
      			var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
        		while(!oFile.AtEndOfStream){
        			sLine = oFile.ReadLine();
          			if(sLine = sLine.replace(oRe,sReplace)){
          				oResult = true;
          			}
        			oFileTmp.WriteLine(sLine);
        		}
        		oFileTmp.Close(); oFile.Close();
        		
        		if(oResult){
					io_file_delete(sFile);
        			oFso.CopyFile(sFile + ".tmp",sFile,true);
    			}
      			io_file_delete(sFile + ".tmp");
       	}
		else if(io_file_exists(sFile,sChangeFile)){
			var sCommand = "%comspec% /c " + sChangeFile + " /I " + sFile + " \"" + sRe + "\" \"" + sReplace + "\"";
			oResult = js_io_msdos(sCommand);	
		}
		return oResult;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_file_attrib(iOpt,sFile,sParameter){
	try{
		if(oFso.FileExists(sFile)){
			var sAttr, Obj = new Object(), sCommand, oResult = true;
			sParameter = sParameter ? sParameter : "";
   			if(iOpt == 1){ // Run with or without parameter
   				sCommand = "%comspec% /c attrib " + sParameter + " " + sFile;
   				oWsh.Run(sCommand,oReg.hide,true);
   			}
   			else if(iOpt == 2){ // Gets attributes
   				oResult = js_io_msdos("%comspec% /c attrib " + sFile);
   				sAttr = oResult.message.substring(0,6); // A  SHR
   				sAttr = js_str_rem(sAttr," "); // removes spaces
   				Obj.archive = (!(sAttr.indexOf("A"))) ? true : false;
   				Obj.hidden = (!(sAttr.indexOf("H"))) ? true : false;
   				Obj.readonly = (!(sAttr.indexOf("R"))) ? true : false;
   				Obj.system = (!(sAttr.indexOf("S"))) ? true : false;
   			}
   			else return false;
   		}
   		else return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		return oResult;
	}
}

function io_file_robocopy(sRoboFile,sSrcPath,sDstPath,sOptions,iShow,bShowReturn){
	try{
		sRoboFile = sRoboFile.search(/\\/g) > -1 ? "\"" + sRoboFile + "\"" : sRoboFile // Bug!! do not remove. %comspec% can't have "robocopy.exe" but it can have "C:\\path\\robocopy.exe"
		var sCmd = "%comspec% /u /q /s /c " + sRoboFile + " \"" + oFso.GetAbsolutePathName(sSrcPath) + "\"\\ \"" + oFso.GetAbsolutePathName(sDstPath) + "\"\\ " + sOptions;
		iShow = js_str_isnumber(iShow) ? iShow : oReg.showmin
		//js_log_print("log_result",sCmd)
		//io_file_append(pmt.fls.log_extra,sCmd,false,true);
		js_log_print("log_result","### Copying files: " + sSrcPath + " ==> " + sDstPath);
		var iReturn = oWsh.Run(sCmd,iShow,true);
		var bResult = true;		
		if(bShowReturn){
			var sReturn = "";
			switch(iReturn){
				case 0 : break;
				case 1 : sReturn = "COPY"; break;
				case 2 : sReturn = "XTRA"; break;
				case 3 : sReturn = "XTRA COPY"; break;
				case 4 : sReturn = "MISM"; break;
				case 5 : sReturn = "MISM COPY"; break;
				case 6 : sReturn = "MISM XTRA"; break;
				case 7 : sReturn = "MISM XTRA COPY"; break;
				case 8 : sReturn = "FAIL (ACCESS DENIED)", bResult = false; break;
				case 9 : sReturn = "FAIL COPY", bResult = false; break;
				case 10 : sReturn = "FAIL XTRA", bResult = false; break;
				case 11 : sReturn = "FAIL XTRA COPY", bResult = false; break;
				case 12 : sReturn = "FAIL MISM ", bResult = false; break;
				case 13 : sReturn = "FAIL MISM COPY", bResult = false; break;
				case 14 : sReturn = "FAIL MISM XTRA", bResult = false; break;
				case 15 : sReturn = "FAIL MISM XTRA COPY", bResult = false; break;
				case 16 : sReturn = "FATAL ERROR", bResult = false; break;
				default: break;
			}
			js_log_print("log_result","#### RoboCopy - ExitCode: " + iReturn + ", CopyInfo: " + sReturn + ", CopyResult: " + (bResult ? "SUCCESS" : "FAILED"));
		}
		return bResult;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

