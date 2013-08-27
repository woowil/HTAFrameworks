// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions, All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**

__H.include("HUI@IO@File.js")

__H.register(__H.IO.File,"INI","INI/INF Files",function INI(){
	var fileini = null
	var oReExt = /ini|inf/ig
	var TXT_NOT_A_FILE = "This is not a file"
	var TXT_FILE_NOT_SET = "The file must be set: setFile(sFile)"
	var files_d = new ActiveXObject("Scripting.Dictionary")
	
	this.isBlankLine = function(sLine){
		return (typeof(sLine) == "string" && sLine.isSearch(/^\[ \t]*$/g) != null);
	}
	
	this.isINI = function isINI(sFile){
		if(typeof(sFile) == "string" && oFso.FileExists(sFile) && (sExt = oFso.GetExtensionName(sFile))){
			if(sExt.isSearch(oReExt)) return true
		}
		return false
	}
	
	this.setFile = function setFile(sFile){
		try{
			if(typeof(sFile) != "string" || !this.isINI(sFile)) throw TXT_NOT_A_FILE			
			fileini = sFile
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.getFile = function getFile(){
		try{
			if(!fileini) throw TXT_FILE_NOT_SET
			return fileini
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.getFileTmp = function getFileTmp(){
		return this.getFile() + ".tmp"
	}
	
	this.hasKeys = function hasKeys(sSection,sKeyName){
		try{
			var sResult = false, iLine
			if(iLine = this.hasSection(sSection)){
				var oFile = oFso.OpenTextFile(this.getFile(),__HIO.read,false,__HIO.TristateUseDefault);
				while(!oFile.AtEndOfStream && oFile.Line < iLine) oFile.SkipLine();
				while(!oFile.AtEndOfStream){
					var sLine = oFile.ReadLine(), oKey;
					if(this.isSection(sLine)) break; // Loop until next section
					else if((oKey = this.isKey(sLine,oFile.Line)) && oKey.name == sKeyName){
						sResult = (oFile.Line-1);
						break;
					}
				}	
				oFile.Close();
			}
			return sResult
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.isKey = function isKey(sLine,iLine){
		try{
			if(__H.isStringEmpty(sLine)) return false;
			var sLineOrg = sLine, sLine = sLine.trim()
			var oRe = new RegExp("[ \t]+","g");
			sLine = sLine.replace(oRe," "); // Removes all tabs and double spaces
			oRe = new RegExp("[ ]{0,1}=[ ]{0,1}","g"), sLine = sLine.replace(oRe,"="); // Removes spaces near '='
			oRe = new RegExp("[\"'%]{0,1}([^=\"\'%]+)[\"'%]{0,1}(={0,1})(.*)","ig"); // This matches: anything=any thing  OR  "any thing"=any thing  OR  anything=any thing  OR  anything
			if(sLine.substring(0,1) == ";");
			else if(oRe.exec(sLine)){
				var o = {};
				o.sline = sLineOrg;
				o.name = RegExp.$1;
				o.value = RegExp.$3 ? RegExp.$3 : "";
				o.key = RegExp.$1 + RegExp.$2 + o.value;
				o.key2 = RegExp.$1 + "=" + o.value;
				o.iline = iLine ? iLine : false;
				return o;
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.isSection = function isSection(sSection){
		try{
			if(__H.isStringEmpty(sSection)) return false;
			var oRe = new RegExp("[\[]([a-z0-9_ \-\.\(\)]+)[\]]$","ig"); // Matches: [paraME TEr_10.A] or [paraME TEr_10.(A)]	
			sSection = sSection.trim()
			if(sSection.substring(0,1) == ";");
			else if(oRe.exec(sSection)){
				return RegExp.$1;
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.hasSection = function hasSection(sSection){
		try{
			var sResult = false
			var oFile = oFso.OpenTextFile(this.getFile(),__HIO.read,false,__HIO.TristateUseDefault);
			while(!oFile.AtEndOfStream){
				if(sSection == this.isSection(oFile.ReadLine())){
					sResult = oFile.Line;
					break;
				}
			}
			oFile.Close();
			return sResult;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.addSection = function addSection(sSection,sKeyName,sKeyValue,sExt){
		try{
			if(!this.hasSection(sSection)){
				var oFile = oFso.OpenTextFile(this.getFile(),__HIO.append,true,__HIO.TristateUseDefault);
				oFile.WriteBlankLines(1);
				oFile.WriteLine("[" + sSection + "]");
				if(sKeyName && sKeyValue) oFile.WriteLine(sKeyName + " = " + sKeyValue);
				oFile.Close();
				return true;
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.getSections = function getSections(){
		try{			
			var a = [], sLine, sSection;
			var oFile = oFso.OpenTextFile(this.getFile(),__HIO.read,false,__HIO.TristateUseDefault);					
			while(!oFile.AtEndOfStream){
				if((sLine = oFile.ReadLine()) && (sSection = this.isSection(sLine))){
					a.push(sSection);
				}
			}
			oFile.Close();
			return a			
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.addKey = function addKey(sSection,sKeyName,sKeyValue,bIgnoreExist){
		try{
			sKeyValue = (typeof(sKeyValue) == "string") ? sKeyValue : ""
			if((iLine = this.hasSection(sSection)) && sKeyName){
				var oFile = oFso.OpenTextFile(this.getFile(),__HIO.read,false,__HIO.TristateUseDefault);
				var oTmpFile = oFso.OpenTextFile(this.getFileTmp(),__HIO.write,true,__HIO.TristateUseDefault);
				while(!oFile.AtEndOfStream){
					var sLine = oFile.ReadLine();
					if(oFile.Line < iLine) oTmpFile.WriteLine(sLine);
					else{
						oTmpFile.WriteLine(sLine); // write the section
						break;
					}
				}
				var v = (sKeyValue === "")  ? sKeyName : sKeyName + "=" + sKeyValue
				var oRe = new RegExp("^$","g"); // Removes empty lines
				while(!oFile.AtEndOfStream){
					var sLine = oFile.ReadLine();				
					if(!bIgnoreExist && (oKey = this.isKey(sLine,oFile.Line)) && oKey.name == sKeyName){
						oTmpFile.WriteLine(v); // If exist, change key value
						break;
					}
					else if(this.isSection(sLine) || oFile.AtEndOfStream){ // Loops until next section
						if(oFile.AtEndOfStream){
							oTmpFile.WriteLine(sLine); // write the key
							oTmpFile.WriteLine(v); // Add key value at end
							oTmpFile.WriteBlankLines(1);								
						}
						else {
							oTmpFile.WriteLine(v); // Add key value at end
							oTmpFile.WriteBlankLines(1);
							oTmpFile.WriteLine(sLine); // write the next section
						}
						break;
					}
					else if(!oRe.exec(sLine)) oTmpFile.WriteLine(sLine);
				}
				while(!oFile.AtEndOfStream){
					oTmpFile.WriteLine(oFile.ReadLine());
				}
				oFile.Close();
				oTmpFile.Close();
				this.move(this.getFileTmp(),this.getFile());
				return true;
			}
			else return this.addSection(sSection,sKeyName,sKeyValue); // Adds section if not exist
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.getKeys = function getKeys(sSection){
		try{
			var sResult = []
			var oFile = oFso.OpenTextFile(this.getFile(),__HIO.read,false,__HIO.TristateUseDefault);
			while(!oFile.AtEndOfStream){
				var sLine = oFile.ReadLine();
				var p = this.isSection(sLine);
				if(p && p.toLowerCase() == sSection.toLowerCase()){
					sResult = [];
					while(!oFile.AtEndOfStream){
						sLine = oFile.ReadLine();							
						if(!this.isSection(sLine)){ // Get all keys to next section
							if(oLine = this.isKey(sLine,oFile.Line)){
								sResult.push(oLine);
							}
						}
						else break;
					}
					break;
				}
			}
			oFile.Close()
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return sResult
	}
	
	this.getKeysDict = function getKeysDict(sSection,bIgnoreEmptyValue){
		try{
			var aKeys, oDict = new ActiveXObject("Scripting.Dictionary")
			if(aKeys = this.getKeys(sSection)){
				for(var i = 0, iLen = aKeys.length; i < iLen; i++){
					if(!oDict.Exists(aKeys[i].name)){
						if(bIgnoreEmptyValue && (aKeys[i].value).length < 1) continue
						oDict.Add(aKeys[i].name,aKeys[i].value)
					}
				}
			}
			return oDict
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
		finally{
			__HUtil.kill(aKeys)
		}
	}
	
	this.delSection = function delSection(sSection){
		try{
			var iLine
			if(iLine = this.hasSection(sSection)){
				var oFile = oFso.OpenTextFile(this.getFile(),__HIO.read,false,__HIO.TristateUseDefault);
				var oTmpFile = oFso.OpenTextFile(this.getFileTmp(),__HIO.write,true);
				while(!oFile.AtEndOfStream){
					var sLine = oFile.ReadLine();
					if(oFile.Line < iLine) oTmpFile.WriteLine(sLine);
					else break;
				}
				while(!oFile.AtEndOfStream){
					var sLine = oFile.ReadLine()
					if(this.isSection(sLine)){
						oTmpFile.WriteLine(sLine);
						break;
					}
				}
				while(!oFile.AtEndOfStream) oTmpFile.WriteLine(oFile.ReadLine());					
				oFile.Close();
				oTmpFile.Close();
				this.move(this.getFileTmp(),this.getFile());
				return true;
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.delKey = function delKey(sSection,sKeyName){
		try{
			var iLine
			if(iLine = this.hasKeys(sSection,sKeyName)){
				var oFile = oFso.OpenTextFile(this.getFile(),__HIO.read,false,__HIO.TristateUseDefault);
				var oTmpFile = oFso.OpenTextFile(this.getFileTmp(),__HIO.write,true,__HIO.TristateUseDefault);
				while(!oFile.AtEndOfStream){
					var sLine = oFile.ReadLine();
					if(oFile.Line != iLine) oTmpFile.WriteLine(sLine);
					else {
						oTmpFile.WriteLine(sLine);
						oFile.SkipLine(); // Delete key
					}
				}							
				oFile.Close();
				oTmpFile.Close();
				this.move(this.getFileTmp(),this.getFile());
				return true;
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.setKey = function setKey(sSection,sKeyName,sKeyValue){
		try{
			var iLine
			if(iLine = this.hasKeys(sSection,sKeyName)){
				var oFile = oFso.OpenTextFile(this.getFile(),__HIO.read,false,__HIO.TristateUseDefault);
				var oTmpFile = oFso.OpenTextFile(this.getFileTmp(),__HIO.write,true,__HIO.TristateUseDefault);
				while(!oFile.AtEndOfStream){
					var sLine = oFile.ReadLine();
					if(oFile.Line != iLine) oTmpFile.WriteLine(sLine);
					else {
						oTmpFile.WriteLine(sLine);
						oFile.SkipLine(); // Delete key
						oTmpFile.WriteLine(sKeyName + "=" + sKeyValue); // Change key value
					}
				}							
				oFile.Close();
				oTmpFile.Close();
				this.move(this.getFileTmp(),this.getFile());
				return true;
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getKeyValue = function getKeyValue(sSection,sKeyName){
		try{
			var aKeys, sValue = false
			if(aKeys = this.getKeys(sSection)){
				for(var i = 0, iLen = aKeys.length; i < iLen; i++){
					if(aKeys[i].name == sKeyName){
						sValue = aKeys[i].value;
						break;
					}
				}
			}
			return sValue
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
		finally{
			__HUtil.kill(aKeys)
		}
	}
	
	this.getKeyLines = function getKeyLines(){
		try{
			var a = [];
			var oFile = oFso.OpenTextFile(this.getFile(),__HIO.read,false,__HIO.TristateUseDefault);
			while(!oFile.AtEndOfStream){
				var sLine = oFile.ReadLine()
				if(!__H.isStringEmpty(sLine) && sLine.substring(0,1) != ";" && sLine.substring(0,1) != "#"){
					if(oLine = this.isKey(sLine,oFile.Line)){
						a.push(oLine);
					}
				}
			}
			oFile.Close();
			return a
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
		
})

var __HINI = new __H.IO.File.INI()