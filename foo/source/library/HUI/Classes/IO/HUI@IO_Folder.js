// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions, All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**

__H.include("HUI@IO.js")

__H.register(__H.IO,"Folder","Folder",function Folder(sFolder){
	var o_this = this
	var b_initialized = false
	
	var s_folder
	var d_folders
	var TXT_NOT_DEFINED
	
	/////////////////////////////////////
	//// DEFAULT
	
	var o_options = {
		
	}
	
	var initialize = function initialize(bForce){
		if(b_initialized && !bForce) return;

		d_folders   = new ActiveXObject("Scripting.Dictionary")
		TXT_NOT_DEFINED = "Folder not defined"
		
		b_initialized = true
	}
	
	this.setOptions = function setOptions(oOptions){		
		try{
			Object.extend(o_options,oOptions,true)
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getOptions = function(){
		return o_options
	}
	
	/////////////////////////////////////
	//// 

		
	this.isFolder = function(sFolder){
		return typeof(sFolder) == "string" && oFso.FolderExists(sFolder)
	}
	
	this.setFolder = function setFolder(sFolder){
		if(this.isFolder(sFolder = oFso.GetAbsolutePathName(sFolder))){
			return (s_folder = sFolder)
		}
		return false
	}
	
	this.getFolder = function getFolder(){
		try{
			if(s_folder == null) throw TXT_NOT_DEFINED
			else return s_folder
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	if(sFolder) this.setFolder(sFolder)
	
	this.remove = function remove(){
		try{
			var bRemoved = true
			for(var i = 0, iLen = arguments.length; i < iLen; i++){
				if(this.setFolder((arguments[i]))){
					if(oWsh.Run("%comspec% /c attrib -r -h \"" + this.getFolder() + "\"",__HIO.hide,true) == 0){
						oFso.DeleteFolder(this.getFolder(),true)
						if(oFso.FolderExists(this.getFolder())) bRemoved = false
					}
				}
			}
			return bRemoved
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}		
	}

	this.exists = function exists(){	
		try{
			for(var i = 0, iLen = arguments.length; i < iLen; i++){
				if(!oFso.FolderExists(arguments[i])){
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

	this.isReadable = function isReadable(sFolder){
		try{
			if(oFso.FolderExists(sFolder)){
				return (oWsh.Run("%compspec% /c dir /b \"" + sFolder + "\" | find /I \"File Not Found\"",__HIO.hide,true) != 0)	
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.listFolders = function listFolders(sFolder,bRecursivt,sPattern){
		try{
			var sFile = __HIO.temp + "\\" + oFso.GetTempName()
			
			sPattern = __H.isString(sPattern) ? " | find /i \"" + sPattern + "\"" : ""
			bRecursivt = bRecursivt ? " /s " : ""
			sFolder = oFso.GetAbsolutePathName(sFolder)

			var sCmd = "%comspec% /u /q /s /c dir \"" + sFolder + "\" /b /OGN /AD " + bRecursivt + sPattern + " | sort>" + sFile
			if(oWsh.Run(sCmd,__HIO.hide,true) != 0){
				__HExp.IOFileNotFound("Folders Not Found In Folder: " + sFolder)
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
			else __HLog.debug("Folders Not Found In Folder: " + sFolder)
			return [];
		}
		finally{
			__HFile.remove(sFile)
		}
	}
	
	this.list = function(){
		return this.listFolders.apply(this,arguments)
	}
	
	this.create = function create(){
		try{
			for(var i = 0, iLen = arguments.length; i < iLen; i++){
				if(typeof(arguments[i]) == "string" && !oFso.FolderExists(arguments[i])){
					if(oWsh.Run("%comspec% /c mkdir \"" + arguments[i] + "\"",__HIO.hide,true) != 0) return false
				}
			}
			return true
		}
		catch(ee){
			__HLog.error(ee,this) 
			return false;
		}
	}
	
	this.move = function move(sFolderFrom,sFolderTo){
		try{
			if(oFso.FolderExists(sFolderFrom)){
				oFso.MoveFolder(sFolderFrom,sFolderTo)
				return true;
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.toHtml = function toHtml(sFolder){
		try{
			if(oFso.FolderExists(sFolder = oFso.GetAbsolutePathName(sFolder))){
				var sSupport = "";
				var sHd = '<table border="0" cellspacing="0" cellpadding="0" class="cSupport">', sFt = '\n</table>';
				var aFiles = this.list(sFolder,new RegExp("js|vbs|css|pl|hta","ig"));
				var oFolder = oFso.GetFolder(sFolder); // c:\path or c:\path\ result in c:\path
				var oFiles = new Enumerator(oFolder.Files);
				for(var i = 0; !oFiles.atEnd(); oFiles.moveNext(), i++){
				 	var oFile = oFso.GetFile(oFiles.item()), sFile = oFile.Path;
					sSupport = sSupport.concat('\n<tr><td width="80">File name </td><td><b>' + oFso.GetFileName(sFile) + '</b></td></tr>')
					sSupport = sSupport.concat('\n<tr valign="top"><td>File path </td><td>' + oFso.GetAbsolutePathName(sFile) + '</td></tr>')
					sSupport = sSupport.concat('\n<tr><td>Created Date </td><td>' + oFile.DateCreated + '</td></tr>')
					sSupport = sSupport.concat('\n<tr><td>Modified Date </td><td>' + oFile.DateLastModified + '</td></tr>')
					sSupport = sSupport.concat('\n<tr><td>File Size </td><td>' + oFile.Size + ' bytes</td></tr>')
					sSupport = sSupport.concat('\n<tr><td colspan="2">&nbsp;</td></tr>')
				}		
				sSupport = sHd + sSupport + sFt;
				return sSupport;
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
})

var __HFolder = new __H.IO.Folder()