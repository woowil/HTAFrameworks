// nOsliw Solutions - HTML Aplication Framework.
//**Start Encode**

/*

File:     library-js-registry.js
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
/////// REGISTRY FUNCTIONS
////////////////////////////////////////////////////////////////////////////////////////////////


function js_reg_envget(sHive,sItem){
	try{
		var oEnv = oWsh.Environment(sHive); // sHive: SYSTEM, USER, PROCESS
		return oEnv(sItem);
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_reg_envset(sItem,val){
	try{
		return js_reg_add(oReg.env,sItem,val);
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_reg_exists(sKey,sName,sType){
	try{
		var sValue = oWsh.RegRead(sKey + "\\" + sName);
		if(typeof(sValue) == "unknown" || sType == "REG_MULTI_SZ" || sType == "REG_BINARY"){ // A VBArray of strings or integers
			var aValue = ( new VBArray(sValue) ).toArray();
			sValue = aValue;
		}
		return sValue;
	}
	catch(e){
		return false;
	}
}

function js_reg_read(sKey,sName,sType){
	try{
		var sValue = oWsh.RegRead(sKey + "\\" + sName);
		if(typeof(sValue) == "unknown" || sType == "REG_MULTI_SZ" || sType == "REG_BINARY"){ // A VBArray of strings or integers
			var aValue = ( new VBArray(sValue) ).toArray();
			sValue = aValue; // array with decimal values
		}
		return sValue;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_reg_add(sKey,sName,sValue,sType){
	try{
		// type: REG_SZ, REG_EXPAND_SZ, REG_DWORD(integer), REG_BINARY(integer) (REG_MULTI_SZ is not supported)
		sType = sType ? sType : "REG_SZ";
		oWsh.RegWrite(sKey + "\\" + sName,sValue,sType);
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		return true;
	}
}

function js_reg_autologon(sOpt,sComputer,sPass,sUser,sDomain,iCount){
	try{
		sComputer = sComputer ? sComputer : oWno.ComputerName;
		if(sOpt == "change"){ // Change
			js_reg_remote("addkeyname",sComputer,oReg.win,"DefaultUserName",(sUser ? sUser : oWno.UserName));
			js_reg_remote("addkeyname",sComputer,oReg.win,"DefaultPassword",sPass);
			js_reg_remote("addkeyname",sComputer,oReg.win,"DefaultDomainName",(sDomain ? sDomain : oWno.UserDomain));
			js_reg_remote("addkeyname",sComputer,oReg.win,"AutoAdminLogon","1");
			js_reg_remote("addkeyname",sComputer,oReg.win,"AutoLogonCount",(iCount ? iCount : "1"));
		}
		else if(sOpt == "reset"){ // Reset
			js_reg_remote("addkeyname",sComputer,oReg.win,"DefaultUserName","");
			js_reg_remote("delkeyname",sComputer,oReg.win,"DefaultPassword");
			sDomain = js_reg_remote("getvalue",sComputer,oReg.env,"InstBank",null,1,true) ? sDomain : "";
			js_reg_remote("addkeyname",sComputer,oReg.win,"DefaultDomainName",sDomain);
			js_reg_remote("addkeyname",sComputer,oReg.win,"AutoAdminLogon","0");
			js_reg_remote("delkeyname",sComputer,oReg.win,"AutoLogonCount");
		}
		return true;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_reg_regobj(sComputer,oRegObj){
	try{
		this.computer = sComputer ? sComputer : oWno.ComputerName
		this.localhost = js_sys_localhost(sComputer) ? true : false
		this.regobjx = oRegObj ? oRegObj : js.regobjx		
		this.regkey = null
		this.regrem = null
		this.regparent = null
		this.regvalues = null
		this.regpath = null
		this.regtype = null
		this.regname = null
		this.regpathorg = null
		this.keypath = null		
		this.ignorecase = false
		
		this.setRegistry = function(sComputer,sKeyPath,sRegName,sRegValue,iRegType,bIgnoreCase,bCreate){
			this.computer = sComputer ? sComputer : oWno.ComputerName
			this.ignorecase = bIgnoreCase
			this.regname = this.ignorecase ? sRegName.toUpperCase() : sRegName;
			this.regtype = iRegType ? iRegType : oReg.REG_SZ;
			this.keypath = (sKeyPath.substring(0,1) != "\\") ? "\\" + sKeyPath : sKeyPath;			
			this.keypathorg = (sKeyPath.substring(0,1) != "\\") ? sKeyPath : sKeyPath.substring(1,sKeyPath.length);			
			
			if(!this.isRegObj(this.keypath) || !js_sys_ping(this.computer)) return false
			else js_log_print("log_result","## Unable to make a remote registry connection on:" + this.computer);
			
			try{
				if(this.localhost){
					this.regrem = false
					this.regkey = this.regobjx.RegKeyFromString(sKeyPath);
				}
				else {
					this.regrem = this.regobjx.RemoteRegistry(this.computer);
					this.regkey = this.regrem.RegKeyFromString(sKeyPath);
				}
				this.regparent = oRegKey.Parent;
				this.regvalues = oRegKey.Values;
				this.regpath = oRegKey.FullName;
				this.regsubkeys = this.regkey.SubKeys
				return true
			}
			catch(ee){
				//Error: The name is not in use for a subkey or named value ( = does not exist)
				if(bCreate){
					if(this.addKey(sComputer,sKeyPath,sRegName,sRegValue,iRegType)){
						js_log_print("log_result","## REGISTRY: Creating missing key: \\\\" + sComputer + "\\" + sKeyPath);
						return this.addKeyName(sComputer,sKeyPath,sRegName,sRegValue,iRegType,bIgnoreCase);
					}
					return false;
				}
			}
			return false
		};
		
		this.setEnumValues = function(bForce){
			if(!this.enumvalues || bForce) this.enumvalues = new Enumerator(this.regvalues)
			this.enumvalues.moveFirst()
			return true
		};
		
		this.getValue = function(sComputer,sKeyPath,sRegName,sRegValue,iRegType,bIgnoreCase){
			if(!this.setRegistry(sComputer,sKeyPath,sRegName,null,iRegType,bIgnoreCase)) return false
			for(var oKey, sValue = false, oEnum = this.setEnumValues(); !oEnum.atEnd(); oEnum.moveNext()){
				oKey = oEnum.item();
				if(this.isEqual(oKey.Name,this.regname)){
					if(oKey.Type == oReg.REG_BINARY){  // REG_BINARY is a VB Array of Integers
						if(this.localhost){
							try{
								sValue = oWsh.RegRead(this.keypathorg + "\\" + this.regname,"REG_BINARY")
								sValue = ( new VBArray(sValue) ).toArray(); // array with decimal values
							}
							catch(ee){
								sValue = false
							}
						}
						else sValue = oKey.Value//new VBArray(oKey.Value).toArray(); // returns an Array					
					}
					else if(oKey.Type == oReg.REG_MULTI_SZ){ // REG_MULTI_SZ is a VB Array of Strings separated with null (\0) charactors
						sValue = oKey.Value, sValueNew = "";
						for(var i = 0; i < sValue.length; i++){
							if((ch = sValue.substring(i,i+1)) == "\0") sValueNew += "\\0"; 
							else sValueNew += ch;
						}
						sValue = sValueNew.replace("\\0","");
					}
					else if(oKey.Type == oReg.REG_DWORD) sValue = oKey.Value;
					else if(oKey.Type == oReg.REG_SZ) sValue = oKey.Value;
					else sValue = oKey.value;
					break;
				}
			}
			js_str_kill(oEnum,this.enumvalues);
			return sValue
		};
		
		this.setValue = function(sComputer,sKeyPath,sRegName,sRegValue,iRegType,bIgnoreCase){
			if(!this.setRegistry(sComputer,sKeyPath,sRegName,sRegValue,iRegType,bIgnoreCase)) return false
			this.regvalue = sRegValue ? sRegValue : this.regvalue
			for(var oKey, bResult = false, oEnum = this.setEnumValues(); !oEnum.atEnd(); oEnum.moveNext()){
				oKey = oEnum.item();
				if(this.isEqual(oKey.Name,this.regname)){
					if(js_str_isdefined(this.regvalue)){
						js_str_kill(oEnum,this.enumvalues);
						oKey.value = this.regvalue;
						bResult = true
						break;
					}
				}
			}
			js_str_kill(oEnum,this.enumvalues);
			return bResult
		};
		
		this.getKeyType = function(sComputer,sKeyPath,sRegName){
			if(!this.setRegistry(sComputer,sKeyPath,sRegName)) return false
			for(var oKey, oResult = false, oEnum = this.setEnumValues(); !oEnum.atEnd(); oEnum.moveNext()){
				oKey = oEnum.item();
				if(this.isEqual(oKey.Name,this.regname)){
					var o = new Object();
					o.key = this.keypath;
					o.name = oKey.Name;
					o.itype = oKey.type;
					o.stype = (s = oReg.key(oKey.type)) == undefined ? "" : s;
					o.value = this.getValue(this.computer,this.keypath,this.regname,null,oKey.Type,false);
					o.object = typeof(oKey.value);
					oResult = o
					break;
				}
			}
			js_str_kill(oEnum,this.enumvalues);
			return oResult
		};		
		
		this.addKeyName = function(sComputer,sKeyPath,sRegName,sRegValue,iRegType){
			if(!this.setRegistry(sComputer,sKeyPath,sRegName,sRegValue,iRegType,null,true)) return false
			if(js_str_isdefined(this.regvalue)){
				if(js_str_isdefined(this.regname)) this.delKeyName(sComputer,sKeyPath,sRegName);
				else this.regname = "";
				try{
					oValues.Add(this.regname,this.regvalue,this.regtype);
					oValues.Reset();
					return true
				}
				catch(ee){}
			}
			return false
		};
		
		this.delKeyName = function(sComputer,sKeyPath,sRegName){
			if(!this.setRegistry(sComputer,sKeyPath,sRegName)) return false
			if(typeof(this.regname) == "object" && this.regname.length > 0) var aRegName = this.regname;
			else var aRegName = new Array(sRegName);
			for(var i = 0; i < aRegName.length; i++){
				try{
					oValues.Remove(aRegName[i]);
					oValues.Reset();
					return true
				}
				catch(ee){}
			}
			return false
		};
		
		this.addKey = function(sComputer,sKeyPath){
			if(!this.setRegistry(sComputer,sKeyPath)) return false
			try{
				var oRe = /([a-zA-Z_]+)\\(.*)/i, aKey; // HKEY_LOCAL_MACHINE
				if(aKey = sKeyPath.match(oRe)){			
					this.regkey = this.regrem.RegKeyFromString("\\" + aKey[1]);
					this.regkey.SubKeys.Add(aKey[2]);
					return true
				}
			}
			catch(ee){}
			return false
		};
		
		this.delKey = function(sComputer,sKeyPath,sRegName){
			if(!this.setRegistry(sComputer,sKeyPath,sRegName)) return false
			try{
				oRegKey.SubKeys.Remove(this.regname);
				oRegKey.SubKeys.Reset();
				return true
			}
			catch(ee){}
			return false
		};
		
		this.copyValues = function(sComputer,sKeyPath,sRegNameFrom,sRegNameTo){
			if(!this.setRegistry(sComputer,sKeyPath,sRegNameFrom,sRegNameTo)) return false
			var oFrom = this.regkey.SubKeys(this.regname).Values
			var oTo = this.regkey.SubKeys(this.regvalue).Values
			for(var oKey, oEnum = new Enumerator(oFrom); !oEnum.atEnd(); oEnum.moveNext()){
				oKey = oEnum.item()
				try{
					oTo.Add(oKey.Name,oKey.Value,oKey.Type);
				}
				catch(ee){
					return false	
				} // The name is already used for a subkey or named value
			}
			return true
		};
		
		this.getKeys = function(sComputer,sKeyPath){
			if(!this.setRegistry(sComputer,sKeyPath)) return false
			var aResult = new Array();
			for(var oKey, oEnum = new Enumerator(this.regsubkeys); !oEnum.atEnd(); oEnum.moveNext()){
				oKey = oEnum.item();
				aResult.push(oKey.Name);
			}
			return (aResult.length > 0) ? aResult : false
		};
		
		this.getValues = function(){
			if(!this.setRegistry(sComputer,sKeyPath,sRegName)) return false
			var aResult = new Array();
			for(var oKey, o, oEnum = new Enumerator(oValues); !oEnum.atEnd(); oEnum.moveNext()){
				oKey = oEnum.item();
				o = new Object();
				o.name = oKey.Name;
				o.value = oKey.Value;
				o.type = oKey.Type;
				aResult.push(o);			
			}
			return (aResult.length > 0) ? aResult : false			
		};
		
		this.isRegObj = function(){
			if(!js_str_isdefined(this.computer,this.regobx,this.keypath)) return false			
			return true
		};
		
		this.isEqual = function(sName1,sName2){
			if(this.ignorecase) sName2 = sName1.toUpperCase(), sName2 = sName2.toUpperCase()
			return (sName1 === sName2 && sName1 == sName2)
		};
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_reg_remote(sOpt,sComputer,sKeyPath,sRegName,sRegValue,iRegType,bIgnoreCase,bCreate){
	try{
		sComputer = sComputer ? sComputer : oWno.ComputerName
		var sResult = false;
		
		if(!js_str_isdefined(sOpt,sKeyPath)) return false
		else if(js.regobjx){ // RegObj.dll (RegObj.Registry) must be loaded and associated globally as js.regobjx
			
			if(sOpt == 1 || sOpt.match(/getvalue$/i)) sOpt = "getvalue"
			else if(sOpt == 2 || sOpt.match(/setvalue/i)) sOpt = "setvalue"
			else if(sOpt == 3 || sOpt.match(/addkeyname/i)) sOpt = "addkeyname"
			else if(sOpt == 4 || sOpt.match(/delkeyname/i)) sOpt = "delkeyname"
			else if(sOpt == 5 || sOpt.match(/delkey$/i)) sOpt = "delkey"
			else if(sOpt == 6 || sOpt.match(/addkey$/i)) sOpt = "addkey"
			else if(sOpt == 8 || sOpt.match(/getkeytype/i)) sOpt = "getkeytype"
			else if(sOpt == 9 || sOpt.match(/copyvalues/i)) sOpt = "getvalues"
			else if(sOpt == 10 || sOpt.match(/getkeys/i)) sOpt = "getkeys"
			else if(sOpt == 11 || sOpt.match(/getvalues/i)) sOpt = "getvalues"
			//else if(sOpt == 11 || sOpt.match(/|/ig)) sOpt == ""
			else return false
			
			sRegName = bIgnoreCase ? sRegName.toUpperCase() : sRegName;
			iRegType = iRegType ? iRegType : oReg.REG_SZ;
			sKeyPathOrg = (sKeyPath.substring(0,1) != "\\") ? sKeyPath : sKeyPath.substring(1,sKeyPath.length);
			sKeyPath = (sKeyPath.substring(0,1) != "\\") ? "\\" + sKeyPath : sKeyPath;
			
			try{
				if(js_sys_localhost(sComputer)){
					var oRegKey = js.regobjx.RegKeyFromString(sKeyPath);
				}
				else {
					var oRegRem = js.regobjx.RemoteRegistry(sComputer);
					var oRegKey = oRegRem.RegKeyFromString(sKeyPath);
				}
				var oRegParent = oRegKey.Parent;
				var oValues = oRegKey.Values;
				var sRegPath = oRegKey.FullName;
			}
			catch(ee){
				//Error: The name is not in use for a subkey or named value ( = does not exist)
				if(bCreate && sOpt == "addkeyname"){
					if(js_reg_remote("addkey",sComputer,sKeyPath,sRegName,sRegValue,iRegType)){
						js_log_print("log_result","## REGISTRY: Creating missing key: \\\\" + sComputer + "\\" + sKeyPath);
						return js_reg_remote("addkeyname",sComputer,sKeyPath,sRegName,sRegValue,iRegType,bIgnoreCase);
					}
					return false;
				}
				else if(sOpt != "addkey") return false;
			}
			
			if(sOpt == "getvalue" || sOpt == "getkeytype"){
				var sValue = false;
				for(var oKey, oEnum = new Enumerator(oValues); !oEnum.atEnd(); oEnum.moveNext()){
					oKey = oEnum.item();
					if((oKey.Name).toUpperCase() == sRegName.toUpperCase()){						
						if(sOpt == "getvalue"){ // Read Key Name
							if(oKey.Type == oReg.REG_BINARY){  // REG_BINARY is a VB Array of Integers
								if(js_sys_localhost(sComputer)) sValue = js_reg_read(sKeyPathOrg,sRegName,"REG_BINARY");
								else sValue = oKey.Value//new VBArray(oKey.Value).toArray(); // returns an Array					
							}
							else if(oKey.Type == oReg.REG_MULTI_SZ){ // REG_MULTI_SZ is a VB Array of Strings separated with null (\0) charactors
								sValue = oKey.Value, sValueNew = "";
								for(var i = 0, len = sValue.length; i < len; i++){
									if((ch = sValue.substring(i,i+1)) == "\0") sValueNew = sValueNew + "\\0"; 
									else sValueNew = sValueNew + ch;
								}
								sValue = sValueNew.replace("\\0","");
							}
							else if(oKey.Type == oReg.REG_DWORD) sValue = oKey.Value;
							else if(oKey.Type == oReg.REG_SZ) sValue = oKey.Value;
							else sValue = oKey.value;
							sResult = sValue;
						}
						else if(sOpt == "getkeytype"){ // Read Key Type
							sResult = new Object();
							sResult.key = sKeyPath;
							sResult.name = oKey.Name;
							sResult.itype = oKey.type;
							sResult.stype = (s = oReg.key(oKey.type)) == undefined ? "" : s;
							sResult.value = js_reg_remote("getvalue",sComputer,sKeyPath,sRegName,null,oKey.Type,false);;
							sResult.object = typeof(oKey.value);
						}
						break;
					}
				}
				js_str_kill(oEnum);
			}
			else if(sOpt == "setvalue"){ // Add Key Name
				var sValue = false;
				for(var oKey, oEnum = new Enumerator(oValues); !oEnum.atEnd(); oEnum.moveNext()){
					oKey = oEnum.item();
					if((oKey.Name).toUpperCase() == sRegName.toUpperCase()){	
						if(typeof(sRegValue) == "string"){
							if(!sRegValue.match(/\0/g) && oKey.Type == oReg.REG_MULTI_SZ){
								sRegValue = sRegValue.replace(/[;,]/ig,"\0") + "\0\0";// null character. 
							}
							sResult = (oKey.value = sRegValue);
						}
						break;
					}
				}
				js_str_kill(oEnum);
			}
			else if(sOpt == "addkeyname"){ // Add Key Name
				if(typeof(sRegValue) == "string"){
					if(typeof(sRegName) == "string") js_reg_remote("delkeyname",sComputer,sKeyPath,sRegName);
					else sRegName = "";
					try{
						oValues.Add(sRegName,sRegValue,iRegType);
						oValues.Reset();
						sResult = true
					}
					catch(ee){}
				}
			}
			else if(sOpt == "delkeyname"){ // Remove Key Name (also array)
				if(typeof(sRegName) == "object" && sRegName.length > 0) aRegName = sRegName;
				else aRegName = new Array(sRegName);				
				for(var i = 0, len = aRegName.length; i < len; i++){
					try{
						oValues.Remove(aRegName[i]);
						oValues.Reset();
						sResult = true
					}
					catch(ee){}
				}				
			}
			else if(sOpt == "delkey"){ // Remove Key (recursively)
				try{
					oRegKey.SubKeys.Remove(sRegName);
					oRegKey.SubKeys.Reset();
					sResult = true
				}
				catch(ee){}
			}
			else if(sOpt == "addkey"){ // Add Key
				try{
					var oRe = /([a-zA-Z_]+)\\(.*)/i; // HKEY_LOCAL_MACHINE
					if(aKey = sKeyPath.match(oRe)){						
						oRegRem = js.regobjx.RemoteRegistry(sComputer);
						oRegKey = oRegRem.RegKeyFromString("\\" + aKey[1]);
						oRegKey.SubKeys.Add(aKey[2]);
						sResult = true
					}
				}
				catch(ee){}
			}
			else if(sOpt == "renamekey"){ // NOT TESTED!
				// Rename Key
				if(js_reg_remote("addkey",sComputer,sKeyPath,sRegName,null)){
					//oRegKey.SubKeys(sRegName).Value = oValues;
					//oRegParent.Value = ;
					oRegKey.SubKeys.Remove(oRegKey);
					sResult = true
				}
			}			
			else if(sOpt == "copyvalues"){ // Copy Key Values
				var oFrom = oRegKey.SubKeys(sRegName).Values
				var oTo = oRegKey.SubKeys(sRegValue).Values
				for(var oKey, oEnum = new Enumerator(oFrom); !oEnum.atEnd(); oEnum.moveNext()){
					oKey = oEnum.item()
					try{
						oTo.Add(oKey.Name,oKey.Value,oKey.Type);
						sResult = true
					}
					catch(ee){} // The name is already used for a subkey or named value
				}
			}
			else if(sOpt == "getkeys"){ // Enum Keys
				var aResult = new Array();
				var oSubKeys = oRegKey.SubKeys;
				for(var oKey, oEnum = new Enumerator(oSubKeys); !oEnum.atEnd(); oEnum.moveNext()){
					oKey = oEnum.item();
					aResult.push(oKey.Name);
				}
				if(aResult.length > 0) sResult = aResult;
			}
			else if(sOpt == "getvalues"){ // Enum Key Values
				var aResult = new Array();
				for(var oKey, oEnum = new Enumerator(oValues); !oEnum.atEnd(); oEnum.moveNext()){
					oKey = oEnum.item();
					var o = new Object();
					o.name = oKey.Name;
					o.value = oKey.Value;
					o.type = oKey.Type;
					aResult.push(o);			
				}
				if(aResult.length > 0) sResult = aResult;
			}
		}
		return sResult;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_reg_addbinary(sKey,sKeyName,sHexValue){ // This function is to add a REG_BINARY value by creating a registry file
	try{
		var sFile = oReg.temp + "\\reg_binary.reg";
		var bResult = true;
		var sStream = "REGEDIT4\n" + '\n';
		sStream += "[" + sKey + "]" + '\n';
		sStream += "\"" + sKeyName + "\"=hex:" + sHexValue + '\n'; // HexStr: 2E,01,56,A3		
		io_file_append(sFile,sStream);
		var sCmd = "%comspec% /c regedit /s " + sFile;		
		if(oWsh.Run(sCmd,oReg.hide,true) != 0) bResult = false;
		io_file_delete(sFile);
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		return bResult;
	}
}

function js_reg_del(key,Stream){
	try{
		for(var i = 1; i < arguments.length; i++){
			if(js_reg_read(key,arguments[i])){
				oWsh.RegDelete(key + "\\" + arguments[i]);
			}
		}
		return true;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_reg_gettype(iType){ // This is used fron the AddReg section in INF printer files
	try{
		var sType = "REG_SZ";
		switch(iType){
			case 0 : {
				sType = "REG_SZ";
				break;
			}
			case 1 : {
				sType = "REG_BINARY";
				break;
			}
			case 2 : {
				//  Prevent overwrite.. Need to code..
				break;
			}
			case 4 : {
				//  Delete Key
				break;
			}
			case 8 : {
				//  Append to registry (only valid if MULTI_SZ is set)
				break;
			}
			case 16 : {
				//  Only Create the Key
				break;
			}
			case 32 : {
				//  Reset only if the key exists
				break;
			}
			case 4096 : {
				//  Make changes in the 64Bit Registry
				break;
			}
			case 8192 : {
				//  XP Only -> Remove KEY only 
				break;
			}
			case 16384 : {
				//  Make changes in the 32Bit Registry
				break;
			}
			case 65536 : {
				sType = "REG_MULTI_SZ";
				break;
			}
			case 65537 : {
				sType = "REG_DWORD";
				break;
			}
			case 131072 : {
				sType = "REG_EXPAND_SZ";
				break;
			}
			case 131073 : {
				sType = "REG_NONE";
				break;
			}
			default : {
				break;
			}
		}	
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		return sType;
	}
}

function js_reg_pagefile(iOpt,sComputer,iMax,iMin){
	try{
		if(iOpt == 1){ // Local - Get
			var aPagefile = js_reg_read(oReg.mem,"PagingFiles","REG_MULTI_SZ"); // array
			var oPagefile = new Array();
			for(var i = 0; i < aPagefile.length; i++){
				aPage = aPagefile[i].split(" ");
				oPagefile[i] = new Object();
				oPagefile[i].file = aPage[0];
				oPagefile[i].minsize = aPage[1]; // MB
				oPagefile[i].maxsize = aPage[2]; // MB
			}
			return oPagefile;
		}
		else if(iOpt == 2){ // Local - Set NOT TESTED!!
			return;
			sValue = "C:\\pagefile.sys " + iMin + " " + iMax;
			if(!js_reg_add(oReg.mem,"PagingFiles",sValue,"REG_MULTI_SZ")) return false;			
		}
		else if(iOpt == 3){ // Remote - Get (requires RegObj.dll)
			var sPagefile = js_reg_remote("getvalue",sComputer,oReg.mem,"PagingFiles",null,7,false);
			var oPagefile = new Object();
			aPage = sPagefile.split(" ");
			oPagefile.file = aPage[0];
			oPagefile.minsize = aPage[1];
			oPagefile.maxsize = aPage[2].substring(0,aPage[2].length-2);
			return oPagefile;
		}
		else if(iOpt == 4){ // Remote - Set (requires RegObj.dll)
			sValue = "C:\\pagefile.sys " + iMin + " " + iMax;
			js_reg_remote("addkeyname",sComputer,oReg.mem,"PagingFiles",sValue,7,true);
		}
		else return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_reg_regsize(sOpt,sComputer,iLimit,bOptimize){
	try{
		var sRegsize = false, iRegSize;
		if(sOpt == "getlocal"){ // Local - Get
			sRegsize = (js_reg_read(oReg.con,"RegistrySizeLimit")/1024); // KB
		}
		else if(sOpt == "setlocal"){ // Local - Set NOT TESTED!!
			sRegsize = js_reg_add(oReg.con,"RegistrySizeLimit",iLimit,"REG_DWORD");
		}
		else if(sOpt == "getremote"){ // Remote - Get (requires RegObj.dll)
			var oRegsize = new Object(); 
			oRegsize.limit = js_reg_remote("getvalue",sComputer,oReg.con,"RegistrySizeLimit",null,oReg.REG_DWORD,false);
			oRegsize.limit_kb = oRegsize.limit/1024;
			var s1 = oFso.GetFile("\\\\" + sComputer + "\\c$\\winnt\\system32\\config\\software").Size;
			var s2 = oFso.GetFile("\\\\" + sComputer + "\\c$\\winnt\\system32\\config\\SECURITY").Size;
			var s3 = oFso.GetFile("\\\\" + sComputer + "\\c$\\winnt\\system32\\config\\system").Size;
			var s4 = oFso.GetFile("\\\\" + sComputer + "\\c$\\winnt\\system32\\config\\sam").Size;
			var s5 = oFso.GetFile("\\\\" + sComputer + "\\c$\\winnt\\system32\\config\\default").Size;
			oRegsize.current = (s1+s2+s3+s4+s5);
			oRegsize.current_kb = oRegsize.current/1024; // KB
			return oRegsize;
		}
		else if(sOpt == "setremote"){ // Remote - Set (requires RegObj.dll)
			var iRegSize = js_reg_regsize(3,sComputer);
			if(bOptimize) iRegSize += 20000;
			var iLimit = (iRegSize * 1024); // iValue: 30000
			sRegsize = js_reg_remote("addkeyname",sComputer,oReg.con,"RegistrySizeLimit",iLimit,oReg.REG_DWORD,true);
		}
		return sRegsize;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_reg_xload(xName,xFile,bRegistered){
	try{
		var Obj = new ActiveXObject(xName);
		return Obj;
	}
	catch(e){
		if(!bRegistered && oFso.FileExists(xFile)){
			if(js_reg_regsvr("register",xFile,false)){
				return js_reg_xload(xName,xFile,true);
			}
			else return false;
		}
		else {
			js_log_print("log_result","Unable to locate file: " + xFile);
			return false;
		}
	}
}

function js_reg_regsvr(sOpt,sFile){
	try{
		var sMode, sCmd, bResult = false;
		if(oFso.FileExists(sFile = oFso.GetAbsolutePathName(sFile))){			
			if(sOpt == "register") sMode = " "; // Register
			else if(sOpt == "unregister") sMode = " /u "; // Unregister
			sCmd = "%comspec% /c regsvr32 /s" + sMode + "\"" + sFile + "\"";
			if(oWsh.Run(sCmd,oReg.hide,true) == 0) bResult = true;
		}
		return bResult;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

