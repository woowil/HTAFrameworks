// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

__H.include("HUI@IO_Folder.js","HUI@Util.js")

//(function(){
__H.register(__H.IO,"File","File Management",function File(){
	
	
	var o_this = this
	var b_initialized = false

	var d_files

	/////////////////////////////////////
	//// DEFAULT

	var o_options = {
		s_file     : false,
		i_sublimit : 10
	}

	var initialize = function initialize(bForce){
		if(b_initialized && !bForce) return;

		d_files = new ActiveXObject("Scripting.Dictionary")

		b_initialized = true
	}

	this.setOptions = function setOptions(oOptions){
		if(__H.isObject(oOptions)) return Object.extend(o_options,oOptions,true)
		return false
	}

	this.getOptions = function(){
		return o_options
	}

	/////////////////////////////////////
	////

	this.isFile = function isFile(sFile){
		try{

		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}

	this.setFile = function setFile(sFile){
		if(!__H.isString(sFile)) return false
		o_options.s_file = sFile
	}

	this.getFile = function getFile(){
		return o_options.s_file
	}

	this.getExtension = function getExtension(){
		try{

		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}

	this.getTemp = function getTemp(){
		try{

		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}

	this.shutcut = function shutcut(sUrl,sName,sFolder){
		try{
			sFolder = typeof(sFolder) == "string" ? sFolder.trim() : "Desktop"
		    var sDesktop = oWsh.SpecialFolders(sFolder);
		    var sTemp = sDesktop + "\\" + sName + ".url"
		    var oUrlLink = oWsh.CreateShortcut(strTemp)
		    oUrlLink.TargetPath = strUrl
		    oUrlLink.Save
		    //oWsh.Exec(sTemp, false)
	    }
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.readall = function readall(sFile){
		try{
			var s = ""
			if(typeof(sFile) == "string" && oFso.FileExists(sFile)){
				var oFile = oFso.OpenTextFile(sFile,__HIO.read,false,__HIO.TristateUseDefault);
				s = (oFile.ReadAll()).trim()
				oFile.Close();
			}
			return s;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.read = function read(sFile,bNotArray,sRegExp){
		try{
			var a = []
			if(oFso.FileExists(sFile)){
				var oFile = oFso.OpenTextFile(sFile,__HIO.read,false,__HIO.TristateUseDefault);
				if(!bNotArray){
					a.push((oFile.ReadAll()).trim().split("\n"))
					var oRe = /[a-z0-9_]+/ig
					for(var i = a.length; i ; i--){
						//a[i] = (a[i]).replace(/\n|\t|\r/g,"")
						a[i-1] = (a[i-1]).trim()
						if(!(a[i-1]).isSearch(oRe)) a.pop(), i--
					}
				}
				else {
					var oRe = new RegExp("("+ sRegExp + ")","ig")
					while(!oFile.AtEndOfStream){
						var sLine = oFile.ReadLine();
						if(sRegExp){
							if(sLine.isSearch(oRe)) a.push(RegExp.$1);
						}
	        			else a.push(sLine);
					}
				}
				oFile.Close();
			}
			return a;
		}
		catch(ee){
			__HLog.error(ee,this)
			return a;
		}
	}

	this.append = function append(sFile,sStream,bDelete,bAppend,sSep,bNotIter){
		var oFile
		try{
			if(!__H.isString(sFile,sStream)) return false
			if(bDelete) this.remove(sFile);
			if(!oFso.FolderExists(oFso.GetParentFolderName(sFile))) __HFolder.create(oFso.GetParentFolderName(sFile))
			sSep = !sSep ? "\n" : sSep;
			oFile = oFso.OpenTextFile(sFile,(bAppend ? __HIO.append : __HIO.write),true,__HIO.TristateUseDefault);
			var a = sStream.split(sSep); // Just remember to separate your stream with a Newline character
			var oRe = /^\[ \t]*$/g
			for(var i = 0, iLen = a.length; i < iLen; i++){
				if(!(a[i]).isSearch(oRe)) oFile.WriteLine(a[i]); // If not blank line
				else oFile.WriteBlankLines(1)
			}
			a.empty()
			return true
		}
		catch(ee){
			var s = (ee.description).toString(), n = (ee.number & 0xFFFF)
			if(s == "Path Not found" && n == 76){ // Folder maybe does not exist
				if(!bNotIter){
					if(__HFolder.create(oFso.GetParentFolderName(sFile))){
						return this.append(sFile,sStream,false,false,sSep,true)
					}
				}
			}
			else if(s == "Invalid procedure call or argument"){
				// I got this once. The problem was that there were UNICode characters in sStream.
				// Note: TristateUseDefault is for ASCII and TristateUse is for UNICODE

			}
			else if(s == "Type Mismatch");
			if(s == "Unknown name" && n == 10){
				// The file is probably using UNICODE charactors.. maybe should use '$Env.TristateUse'
			}
			//__HLog.error(ee,this) // CAUSES A LOOP
			return false;
		}
		finally{
			try{
				oFile.Close()
			}
			catch(ee){}
			delete oFile;
		}
	}

	this.write = function write(sFile,s,bBackup,bOverwite){
		try{
			if(__H.isString(sFile,s) && oFso.FileExists(sFile)){
				if(bBackup) oFso.CopyFile(sFile,sFile + ".bak",bOverwite)
				var oFile = oFso.OpenTextFile(sFile,__HIO.write,true,__HIO.TristateUseDefault);
				oFile.WriteLine(s)
				oFile.Close();
				return true
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return false;
	}

	this.create = function create(sFile,sStream,bNotIter){
		try{
			if(typeof(sFile) != "string") return false
			//this.remove(sFile);
			oFso.CreateTextFile(sFile,true,__HIO.TristateUseDefault);
			var oFile = oFso.OpenTextFile(sFile,__HIO.write,true,__HIO.TristateUseDefault);
			if(typeof(sStream) == "string") oFile.WriteLine(sStream);
			oFile.close()
			return sFile;
		}
		catch(ee){
			var s = (ee.description).toString(), n = (ee.number & 0xFFFF)
			if(s == "Path Not found" && n == 76){ // Folder maybe don't exist
				if(!bNotIter){
					if(__HFolder.create(oFso.GetParentFolderName(sFile))){
						return this.append(sFile,sStream,false,false,sSep,true);
					}
				}
			}
			else if(s == "Invalid procedure call or argument"){
				// I got this once. The problem was that there were UNICode characters in sStream.
				// Note: TristateUseDefault is for ASCII and TristateUse is for UNICODE

			}
			if(s == "Unknown name" && n == 10){
				// The file is probably using UNICODE charactors.. maybe should use '$Env.TristateUse'
			}
			__HLog.error(ee,this)
			return false;
		}
	}

	this.merge = function merge(sFile,sFile1,sFile2){
		try{
			if(__H.isStringEmpty(sFile,sFile1)) return false
			if(this.exists(sFile1)){
				if(typeof(sFile2) != "string") sFile2 = sFile1, sFile1 = sFile
				var sCmd = "%comspec% /u /q /s /c copy /y \"" + sFile1 + "\" + \"" + sFile2 + "\" \"" + sFile + "\" | find /i \"copied\""
				return (oWsh.Run(sCmd,__HIO.hide,true) == 0)
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.isReadable = function isReadable(sFile){
		try{
			if(oFso.FileExists(sFile)){
				return (oWsh.Run("%comspec% /c type \"" + sFile + "\" >nul",__HIO.hide,true) == 0)
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.setAttribute = function setAttribute(sFile,iBit){
		try{
			if(oFso.FileExists(sFile) && __H.isNumber(iBit)){
				var oFile = oFso.GetFile(sFile);
				oFile.Attributes = oFile.Attributes + iBit; // iBit is positive/negative integer: +-1(readonly), +-2(hidden), +-4(system) or +-32(archive)
			}
			else return false;
			return oFile.Attributes;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getAttribute = function getAttribute(sFile){
		try{
			var o = {};
			if(oFso.FileExists(sFile)){
				var oFile = oFso.GetFile(sFile);
				o.normal = (oFile.attributes & 0); // Normal file. No attributes are set
				o.readonly = (oFile.attributes & 1); // Read-only file. Attribute is read/write.
				o.hidden = (oFile.attributes & 2); // Hidden file. Attribute is read/write.
				o.system = (oFile.attributes & 4); // System file. Attribute is read/write.
				o.volume = (oFile.attributes & 8); // Disk drive volume label. Attribute is read-only.
				o.directory = (oFile.attributes & 16); // Folder or directory. Attribute is read-only.
				o.archive = (oFile.attributes & 32); // File has changed since last backup. Attribute is read/write.
				o.alias = (oFile.attributes & 1024); // Link or shortcut. Attribute is read-only.
				o.compress = (oFile.attributes & 2048); // Compressed file. Attribute is read-only.
				o.attr = oFile.attributes;

				// Definitions
				o.readonly_on = -1, o.readonly_off = 1;
				o.hidden_on = -2, o.hidden_off = 2;
				o.system_on = -4, o.system_off = 4;
				o.archive_on = -32, o.archive_off =-32;
				o.enableall = 0, o.disableall = 39;
				// Reset to zero

			}
			else return false;
			return o;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.info = function info(sFile){
		try{
			var o = {};
			if(oFso.FileExists(sFile)){
	   			o.name = oFso.GetFileName(sFile); // file.txt
	   			o.base = oFso.GetBaseName(sFile); //file
	   			o.ext = (oFso.GetExtensionName(sFile)).toLowerCase(); // txt
	   			o.path = oFso.GetAbsolutePathName(sFile); // c:\path\file.txt
	   			o.parent = oFso.GetParentFolderName(sFile); // c:\path
	   			o.version = oFso.GetFileVersion(sFile);
	   			o.attr = this.attrget(sFile);
	   			var f = oFso.GetFile(sFile);
	   			o.created = f.DateCreated;
	   			o.lastmod = f.DateLastModified;
	   			o.lastacc = f.DateLastAccessed;
	   			o.size = f.Size // In bytes (If KB, divide this by 1024)
	   			o.type = f.Type // Ex: Text Document (*.txt)
	   			o.shortname = f.ShortName
	   			o.drive = f.Drive
	   			o.absname = sFile;
	   		}
	   		return o;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.listFiles = function listFiles(sFolder,sFileExt,bRecursivt,sPattern){
		try{
			var sFile = __HIO.temp + "\\" + oFso.GetTempName()

			sFileExt = !__H.isStringEmpty(sFileExt) ? "*." + sFileExt.replace(/.*\./g,"") : "*.*"
			sPattern = !__H.isStringEmpty(sPattern) ? " | find /i \"" + sPattern + "\"" : ""
			bRecursivt = bRecursivt ? " /s " : ""
			sFolder = oFso.GetAbsolutePathName(sFolder)

			var sCmd = "%comspec% /u /q /s /c dir \"" + sFolder + "\\" + sFileExt + "\" /b /OGN /A-D " + bRecursivt + sPattern + " | sort>" + sFile
			if(oWsh.Run(sCmd,__HIO.hide,true) != 0){
				__HExp.IOFileNotFound("Files Not Found In Folder: " + sFolder + "\\" + sFileExt)
			}

			var oFile = oFso.OpenTextFile(sFile,__HIO.read,false,__HIO.TristateUseDefault)
			var a = oFile.ReadAll().trim().split(/\r\n|\n/g)
			oFile.close()

			return a;
		}
		catch(ee){
			if(ee.description.trim() != "Input past end of file"){
				__HLog.error(ee,this)
			}
			else __HLog.debug("Folders Not Found In Folder: " + sFolder + "\\" + sFileExt)
			return [];
		}
		finally{
			__HFile.remove(sFile)
		}
	}

	this.list = function(){
		return this.listFiles.apply(this,arguments)
	}

	this.suffix = function suffix(sPath,sExtension,sSuffix){
		try{
			var aFiles, o = {};
			o.count = o.changed = 0;
			if(aFiles = this.list(sPath)){ // Folder
				var oFile;
				for(var i = 0, iLen = aFiles.length; i < iLen; i++){
					if(oFso.FileExists(aFiles[i]) && (oFile = this.info(aFiles[i],sSuffix))){
						if(oFile.ext == sExtension){
							var sCmd = "%comspec% /c cd /d " + oFile.parent + " && ren " + oFile.name + " " + oFile.namesuffix
							if(oWsh.Run(sCmd,__HIO.hide,true) == 0) o.changed++;
							o.count++;
						}
					}
				}
			}
			return o;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.rename = function rename(OrgFile,NewFile){
		try{
			if(this.exists(OrgFile,NewFile)){
				var sCmd = "%comspec% /c ren " + OrgFile + " " + NewFile;
				return (oWsh.Run(sCmd,__HIO.hide,true) == 0)
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.move = function move(sFileFrom,sFileTo){
		try{
			if(oFso.FileExists(sFileFrom)){
				oFso.MoveFile(sFileFrom,sFileTo)
			}
			return true;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.exists = function exists(args){
		try{
			for(var i = 0, iLen = arguments.length; i < iLen; i++){
				if(!oFso.FileExists(arguments[i])){
					return false;
				}
			}
			return true;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.remove = function remove(){
		try{
			var bRemoved = true
			for(var i = 0, iLen = arguments.length; i < iLen; i++){
				if(oFso.FileExists(arguments[i])){
					bRemoved = (oWsh.Run("%comspec% /c attrib -h -r \"" + arguments[i] + "\">nul && del /f /q \"" + arguments[i] + "\">nul",__HIO.hide,true) == 0)
					if(oFso.FileExists(arguments[i])){
						try{
							oFso.GetFile(arguments[i]).Delete(true); // deletes also empty files
						}
						catch(ee){}
						if(oFso.FileExists(arguments[i])) bRemoved = false
					}
				}
			}
			return bRemoved;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.replace = function replace(sFile,sRe,sReplace){
		try{
			var oResult = false;
			if(oFso.FileExists(sFile)){
	      			var oRe = new RegExp(sRe,"i"), sLine;
	      			var oFileTmp = oFso.OpenTextFile(sFile + ".tmp",__HIO.write,true,__HIO.TristateUseDefault);
	      			var oFile = oFso.OpenTextFile(sFile,__HIO.read,false,__HIO.TristateUseDefault);
	        		while(!oFile.AtEndOfStream){
	        			sLine = oFile.ReadLine();
	          			if(sLine = sLine.replace(oRe,sReplace)){
	          				oResult = true;
	          			}
	        			oFileTmp.WriteLine(sLine);
	        		}
	        		oFileTmp.Close(); oFile.Close();

	        		if(oResult){
						this.remove(sFile);
	        			oFso.CopyFile(sFile + ".tmp",sFile,true);
	    			}
	      			this.remove(sFile + ".tmp");
	       	}
			return oResult;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.setAttrib = function setAttrib(sFile){
		try{
			if(oFso.FileExists(sFile)){
				return (oWsh.Run("%comspec% /c attrib " + sAttrib + " \"" + sFile + "\"",__HIO.hide,true) == 0)
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getAttrib = function getAttrib(sFile,sAttrib){
		try{
			var s, o = {}
			if(oFso.FileExists(sFile)){
				o = __HSys.msdos("attrib \"" + sFile + "\"");
				if(o.message){
					s = o.message.substring(0,6); // A  SHR
					s = s.replace(/[ \t]+/g,""); // removes spaces
					o.archive = (!(s.indexOf("A"))) ? true : false;
					o.hidden = (!(s.indexOf("H"))) ? true : false;
					o.readonly = (!(s.indexOf("R"))) ? true : false;
					o.system = (!(s.indexOf("S"))) ? true : false;
				}
	   		}
			return o
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.robocopy = function robocopy(sFile,sSrcPath,sDstPath,sOptions,iShow,bShowReturn){
		try{
			sFile = sFile.isSearch(/\\/g) ? "\"" + sFile + "\"" : sFile // Bug!! do not remove. %comspec% can't have "robocopy.exe" but it can have "C:\\path\\robocopy.exe"
			var sCmd = "%comspec% /u /q /s /c " + sFile + " \"" + oFso.GetAbsolutePathName(sSrcPath) + "\"\\ \"" + oFso.GetAbsolutePathName(sDstPath) + "\"\\ " + sOptions;
			iShow = __H.isNumber(iShow) ? iShow : __HIO.showmin
			//this.append(pmt.fls.log_extra,sCmd,false,true);
			__HLog.log("### Copying files: " + sSrcPath + " ==> " + sDstPath);
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
				__HLog.log("#### RoboCopy - ExitCode: " + iReturn + ", CopyInfo: " + sReturn + ", CopyResult: " + (bResult ? "SUCCESS" : "FAILED"));
			}
			return bResult;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.createWSH = function createWSH(sFile,sTitle,sStart,sFunctionStream){
		try{
			var s = "///////////\n//// " + sTitle + "\n\n" +
					'var oFso = new ActiveXObject("Scripting.FileSystemObject");\n' +
					'var oWsh = new ActiveXObject("WScript.Shell");\n' +
					'var oWno = new ActiveXObject("WScript.Network");\n' +
					'var __HKeys = {};\n\n' +
					'if(' + sStart + ') WScript.Quit(0);\nelse WScript.Quit(-1);';
			this.append(sFile,s,true);

			var oFile = oFso.OpenTextFile(sFile,__HIO.append,true,__HIO.TristateUseDefault);
			oFile.WriteBlankLines(1);
			oFile.WriteLine(__H.Env);
			for(var i = 3, iLen = arguments.length; i < iLen; i++){
				if(arguments[i] = "string"){
					oFile.WriteBlankLines(2);
					oFile.WriteLine(arguments[i]);
				}
			}
			oFile.Close();
			return true;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

})
//}());

//// Global define
var __HFile = new __H.IO.File()
