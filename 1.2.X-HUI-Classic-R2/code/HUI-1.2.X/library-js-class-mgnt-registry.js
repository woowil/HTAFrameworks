// nOsliw Solutions HUI, HTML Application Library http://hui.nosliw.info, License http://hui.codeplex.com/license
//**Start Encode**

/*

File:     library-js-class-mgnt-registry.js
Purpose:  Development script
Author:   Woody Wilson
Created:  2006-03
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
class_mgnt_registry.prototype = new _Class("Registry","Class for accessing remote registry")

function class_mgnt_registry(oService,sComputer,sUser,sPass,bInit,bStatus,bDebug){
	try{
		// http://msdn.microsoft.com/library/en-us/wmisdk/wmi/stdregprov.asp
		// http://www.activexperts.com/activmonitor/windowsmanagement/adminscripts/registry/
		// http://cwashington.netreach.net/depo/view.asp?Index=560&ScriptType=vbscript
		// http://www.serverwatch.com/tutorials/article.php/1476831
		// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/writing_wmi_scripts_in_jscript.asp
		// http://www.serverwatch.com/tutorials/print.php/1476831
		// http://www.codecomments.com/archive298-2004-5-191525.html		
		this.wmi_default = oService
		this.regprov = null
		this.computer = typeof(sComputer) == "string" ? sComputer : oWno.ComputerName
		this.username = typeof(sUser) == "string" ? sUser : null
		this.password = typeof(sPass) == "string" ? sPass : null
		this.isStatus = bStatus ? true : false
		this.isDebug = bDebug ? true : false
		this.oReKey = new RegExp("[\\\\]{0,1}(.+)[\\\\]{0,1}$","ig")
		this.oReMethodNoValue = new RegExp(/EnumKey|EnumValues|DeleteKey|CreateKey/ig)
		this.isLocalhost = this.isHostname()
		this.name = "class_mgnt_registry"
		// For none WMI enabled systems
		this.regobjx = null
		this.hasWMI = true		
		
		this.HKCR = 0x80000000; //HKEY_CLASSES_ROOT
		this.HKCU = 0x80000001; //HKEY_CURRENT_USER
		this.HKLM = 0x80000002; //HKEY_LOCAL_MACHINE
		this.HKU  = 0x80000003; //HKEY_USERS
		this.HKCC = 0x80000005; //HKEY_CURRENT_CONFIG
		this.HKDD = 0x80000006; //HKEY_DYN_DATA (Windows 95/98)

		this.REG_SZ        = 1
		this.REG_EXPAND_SZ = 2
		this.REG_BINARY    = 3
		this.REG_DWORD     = 4
		this.REG_MULTI_SZ  = 7

		this.KEY_QUERY_VALUE        = 0x1
		this.KEY_SET_VALUE          = 0x2
		this.KEY_CREATE_SUB_KEY     = 0x4
		this.KEY_ENUMERATE_SUB_KEYS = 0x8
		this.KEY_NOTIFY             = 0x10
		this.KEY_CREATE_LINK        = 0x20
		this.DELETE                 = 0x10000
		this.READ_CONTROL           = 0x20000
		this.WRITE_DAC              = 0x40000
		this.WRITE_OWNER            = 0x80000

		var typesidx = new Array(
		    "",    // 0
		    "REG_SZ",    // 1
		    "REG_EXPAND_SZ",    // 2
		    "REG_BINARY",    // 3
		    "REG_DWORD",    // 4
		    "REG_DWORD_BIG_ENDIAN",    // 5
		    "REG_LINK",    // 6
		    "REG_MULTI_SZ",    // 7
		    "REG_RESOURCE_LIST",    // 8
		    "REG_FULL_RESOURCE_DESCRIPTOR",    // 9
		    "REG_RESOURCE_REQUIREMENTS_LIST",    // 10
		    "REG_QWORD");    // 11
		
		this.setService = function(oService,sComputer,sUser,sPass,bForce){
			try{
				if(!this.wmi_default || bForce){
					if(oService) this.wmi_default = oService
					else {
						this.computer = typeof(sComputer) == "string" ? sComputer : oWno.ComputerName
						this.username = typeof(sUser) == "string" ? sUser : null
						this.password = typeof(sPass) == "string" ? sPass : null
						if(!this.wmi_locator) this.wmi_locator = new ActiveXObject("WbemScripting.SWbemLocator");
						//this.wmi_cimv2 = this.wmi_locator.ConnectServer(this.computer,"root\\cimv2",this.username,this.password);
						this.wmi_default = this.wmi_locator.ConnectServer(this.computer,"root\\default",this.username,this.password);
					}
				}
				this.isLocalhost = this.isHostname()
				this.setText()
				if(this.wmi_default){
					this.regprov = this.wmi_default.Get("StdRegProv")					
				}
				return this.wmi_default
			}
			catch(ee){
				this.error(ee,"setService()")
				this.hasWMI = false
				if(this.ERR_WMI_ACCESSDENIED) return false;
				else if(this.ERR_WMI_NOTSTARTED){
					if(oWsh.Run("%comspec% /c ping " + sComputer + " -n 1 -l 100 -w 1 -f -i 1 | find /i \"TTL\" >nul",oReg.hide,true) != 0);
					else if(oWsh.Run("%comspec% /c sc \\\\" + sComputer + " config winmgmt start= auto>nul && sc \\\\" + sComputer + " start winmgmt>nul",oReg.hide,true) == 0){
						return this.setService(oService,sComputer,sUser,sPass,bForce)
					}
				}					
				return this.setRegObj()
			}
			return false
		}
		
		this.setRegObj = function(sKeyPath,sFile){
			try{ // RegObj.dll
				if(this.isString(sFile) && oFso.FileExists(sFile)){
					if(oWsh.Run("%comspec% /c regsvr32 /s \"" + sFile + "\"",oReg.hide,true) == 0){
						throw "Unable to register file: " + sFile;
					}
				}
				// Requires that regobj.dll is loaded using regsvr32.exe
				if(!this.regobjx) this.regobjx = new ActiveXObject("RegObj.Registry")
				if(this.isString(sKeyPath)){
					sKeyPath = sKeyPath.replace(this.oReKey,"$1")
					if(this.isLocalhost){
						this.regrem = false
						this.regkey = this.regobjx.RegKeyFromString(sKeyPath);
					}
					else {
						this.regrem = this.regobjx.RemoteRegistry(this.computer);
						this.regkey = this.regrem.RegKeyFromString(sKeyPath);
					}
					this.regparent = this.regkey.Parent;
					this.regvalues = this.regkey.Values;
					this.regpath = this.regkey.FullName;
					this.regsubkeys = this.regkey.SubKeys
				}
				return true
			}
			catch(ee){
				this.error(ee,"setRegObj()")
				this.regobjx = null
			}
			return false
		}
		
		this.setText = function(){
			this._SET_ERR_STRING = "An error occured attempting to set a remote registry value on computer" + this.computer.toUpperCase()
			this._SET_ERR_METHOD = ""
			this._GET_ERR_METHOD = ""
		}
		
		this.getEnvironment = this.getEnv = function(sKeyName,sHive){
			try{
				if(this.isLocalhost){
					// sHive: Win200X/NT/XP - SYSTEM, USER, 
					// sHive: Win95/98 - PROCESS
					sHive = typeof(sHive) == "string" ? sHive : "SYSTEM"
					var oEnv = oWsh.Environment(sHive); 
					return oEnv(sKeyName); // sItem: NUMBER_OF_PROCESSORS, SYSTEMROOT etc
				}
				return this.Exists(this.subkey.env,sKeyName)
			}
			catch(ee){
				this.error(ee,"getEnvironment()");	
			}
			return false;
		}
		
		this.getMethod = function(sKeyPath,sKeyName,sMethod){
			this.reset()
			if(!this.regprov){
				if(!this.setService(this.wmi_default,this.computer,this.username,this.password,true)) return false
			}
			if(!this.hasWMI) return false
			try{
				var o
				if(o = this.getHKEY(sKeyPath)){
					var oMethod = this.regprov.Methods_.Item(sMethod)
					var oInParam = oMethod.InParameters.SpawnInstance_();
					oInParam.hDefKey = o.HK_HEX
					oInParam.sSubKeyName = o.KEY;
					if(!sMethod.match(this.oReMethodNoValue)) oInParam.sValueName = typeof(sKeyName) == "string" ? sKeyName : "";
					// http://msdn.microsoft.com/library/en-us/wmisdk/wmi/swbemobject_execmethod_.asp
					o = this.regprov.ExecMethod_(oMethod.Name,oInParam)
					if(this.isDebug) this.echo(sKeyName + ": " + sMethod + " ReturnValue: " + o.ReturnValue)
					return (o.ReturnValue == 0 ? o : false)
				}
			}
			catch(ee){
				this.error(ee,"setMethod()",this._GET_ERR_METHOD)
			}
			return null
		}
					
		this.setMethod = function(sKeyPath,sKeyName,sKeyValue,sMethod){
			this.reset()
			if(!this.regprov){
				if(!this.setService(this.wmi_default,this.computer,this.username,this.password,true)) return false
			}
			if(!this.hasWMI) return false			
			try{
				var o
				if(o = this.getHKEY(sKeyPath)){
					var oMethod = this.regprov.Methods_.Item(sMethod)
					var oInParam = oMethod.InParameters.SpawnInstance_();
					oInParam.hDefKey = o.HK_HEX
					oInParam.sSubKeyName = o.KEY;
					oInParam.sValueName = typeof(sKeyName) == "string" ? sKeyName : "";;
					
					if(sMethod.match(/StringValue/ig)) oInParam.sValue = sKeyValue
					else if(sMethod.match(/DWORD|Binary/ig)) oInParam.uValue = sKeyValue // DWORD and BINARY
					else return false
					// http://msdn.microsoft.com/library/en-us/wmisdk/wmi/swbemobject_execmethod_.asp
					o = this.regprov.ExecMethod_(oMethod.Name,oInParam)
					return (o.ReturnValue == 0)
				}
			}
			catch(ee){
				this.error(ee,"setMethod()",this._SET_ERR_METHOD)
			}
			return false
		}
		
		//// GET functions		

		this.getHKEY = function(sKeyPath){
			if(typeof(sKeyPath) != "string") return null
			sKeyPath = sKeyPath.replace(this.oReKey,"$1")
			var o = new Object()
			if(sKeyPath.match(/(?:HKEY_LOCAL_MACHINE|HKLM)\\(.+)$/ig)) o.HK_HEX = this.HKLM, o.KEY = RegExp.$1
			else if(sKeyPath.match(/(?:HKEY_CURRENT_USER|HKLM)\\(.+)$/ig)) o.HK_HEX = this.HKCU, o.KEY = RegExp.$1
			else if(sKeyPath.match(/(?:HKEY_CLASSES_ROOT|HKLM)\\(.+)$/ig)) o.HK_HEX = this.HKCR, o.KEY = RegExp.$1
			else if(sKeyPath.match(/(?:HKEY_USERS|HKLM)\\(.+)$/ig)) o.HK_HEX = this.HKU, o.KEY = RegExp.$1
			else if(sKeyPath.match(/(?:HKEY_CURRENT_CONFIG|HKLM)\\(.+)$/ig)) o.HK_HEX = this.HKCC, o.KEY = RegExp.$1
			else if(sKeyPath.match(/(?:HKEY_DYN_DATA|HKDD)\\(.+)$/ig)) o.HK_HEX = this.HKDD, o.KEY = RegExp.$1
			else return null
			return o
		}
		
		this.Exists = function(sKeyPath,sKeyName,bString){
			try{ // only for string
				var o
				if(this.isLocalhost){
					o = oWsh.RegRead(sKeyPath.replace(this.oReKey,"$1") + "\\" + sKeyName)
					if(typeof(o) == "unknown"){
						o = ( new VBArray(o) ).toArray()
						return (bString ? o.toString() : o);
					}
					return o
				}
				else if(o = this.getMethod(sKeyPath,sKeyName,"GetStringValue")){
					return o.sValue;
				}
				else if(o = this.getMethod(sKeyPath,sKeyName,"GetDWORDValue")){
					return o.uValue; // Note that this uses uValue instead of sValue
				}
				else if(o = this.getMethod(sKeyPath,sKeyName,"GetMultiStringValue")){
					o = o.sValue.toArray()
					return (bString ? o.toString() : o);
				}
			}
			catch(ee){
				if(this.isDebug) this.error(ee,"Exists()")
			}
			return false
		}
		
		this.getValueSTRING = this.getValue = function(sKeyPath,sKeyName){
			try{
				var o
				if(!this.isString(sKeyPath,sKeyName)) return false
				if(this.isLocalhost){
					try{
						return oWsh.RegRead(sKeyPath.replace(this.oReKey,"$1") + "\\" + sKeyName)
					}
					catch(ee){
						if(this.isDebug) this.error(ee)
						// Unable to open registry key "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\curComputerADSite" for reading
					}
				}
				else if(o = this.getMethod(sKeyPath,sKeyName,"GetStringValue")){
					if(this.isDebug) this.echo(sKeyName + ": " + o.sValue) 
					return o.sValue;
				}
			}
			catch(ee){
				this.error(ee,"getValue()")
			}
			return false
		}
		
		this.setValueSTRING = this.setValue = function(sKeyPath,sKeyName,sKeyValue){
			try{
				if(!this.isString(sKeyPath,sKeyName)) return false // Important! sKeyValue can be "",do not add here
				if(typeof(sKeyValue) != "string") return false
				
				if(this.isLocalhost){
					oWsh.RegWrite(sKeyPath.replace(this.oReKey,"$1") + "\\" + sKeyName,sKeyValue,"REG_SZ")
					return true
				}
				else return this.setMethod(sKeyPath,sKeyName,sKeyValue,"SetStringValue")
			}
			catch(ee){
				this.error(ee,"setValue()")
			}
			return false
		}		
		
		this.getValueEXPAND = function(sKeyPath,sKeyName){
			try{
				var o
				if(!this.isString(sKeyPath,sKeyName)) return false
				if(this.isLocalhost){
					if(o = oWsh.RegRead(sKeyPath.replace(this.oReKey,"$1") + "\\" + sKeyName)){
						return o.toString()
					}
				}
				else if(o = this.getMethod(sKeyPath,sKeyName,"GetExpandedStringValue")){
					return o.sValue;
				}
			}
			catch(ee){
				this.error(ee,"getValueEXPAND()")
			}
			return false
		}

		this.getValueMULTI = function(sKeyPath,sKeyName,bString){
			var o
			if(!this.isString(sKeyPath,sKeyName)) return false
			if(this.isLocalhost){
				o = oWsh.RegRead(sKeyPath.replace(this.oReKey,"$1") + "\\" + sKeyName)
				if(typeof(o) == "unknown"){
					o = ( new VBArray(o) ).toArray()
					return (bString ? o.toString() : o);
				}
			}
			else if(o = this.getMethod(sKeyPath,sKeyName,"GetMultiStringValue")){
				o = o.sValue.toArray()
				return (bString ? o.toString() : o);
			}
			return false
		}
		
		this.setValueMULTI = function(sKeyPath,sKeyName,aKeyValues){
			try{ // aKeyValues: Either an array or a string
				if(!this.isString(sKeyPath,sKeyName)) return false
				if(this.isString(aKeyValues)) aKeyValues = new Array(aKeyValues)
				var vb = this.setVBArray(aKeyValues,"string")
				if(typeof(vb) != "unknown" || typeof(vb) == "boolean") return false
				return this.setMethod(sKeyPath,sKeyName,vb,"SetMultiStringValue")
			}
			catch(ee){
				this.error(ee,"setValueBINARY()")
			}
			return false
		}
		
		this.getValueDWORD = function(sKeyPath,sKeyName){
			var o
			if(!this.isString(sKeyPath,sKeyName)) return false
			if(this.isLocalhost){
				return oWsh.RegRead(sKeyPath.replace(this.oReKey,"$1") + "\\" + sKeyName)
			}
			else if(o = this.getMethod(sKeyPath,sKeyName,"GetDWORDValue")){
				return o.uValue; // Note that this uses uValue instead of sValue
			}
			return false
		}

		this.setValueDWORD = function(sKeyPath,sKeyName,iKeyValue){
			try{
				if(!this.isString(sKeyPath,sKeyName)) return false
				if(typeof(iKeyValue) != "number") return false
				//if(this.isLocalhost) return oWsh.RegWrite(sKeyPath + "\\" + sKeyName,this.getNum(iKeyValue),"REG_DWORD") // Can't use, get buffer overflow on Hex numbers
				return this.setMethod(sKeyPath,sKeyName,iKeyValue,"SetDWORDValue") // sKeyValue: Either hex or number
			}
			catch(ee){
				this.error(ee,"setValueDWORD()")
			}
			return false
		}
		
		this.getValueBINARY = function(sKeyPath,sKeyName,bString){
			try{
				var o
				if(!this.isString(sKeyPath,sKeyName)) return false
				if(this.isLocalhost){
					o = oWsh.RegRead(sKeyPath.replace(this.oReKey,"$1") + "\\" + sKeyName)
					if(typeof(o) == "unknown"){
						o = ( new VBArray(o) ).toArray()
						return (bString ? o.toString() : o);
					}
				}
				else if(o = this.getMethod(sKeyPath,sKeyName,"GetBinaryValue")){
					return o.uValue.toArray(); // Note that this uses uValue instead of sValue
				}
			}
			catch(ee){
				this.error(ee,"getValueBINARY()")
			}
			return false
		}
		
		this.setValueBINARY = function(sKeyPath,sKeyName,aKeyValues){
			try{ // aKeyValues: Either an array of 0x01 0xf2 or a hex number
				if(!this.isString(sKeyPath,sKeyName)) return false
				if(typeof(aKeyValues) == "number") aKeyValues = new Array(aKeyValues)
				var vb = this.setVBArray(aKeyValues,"number")
				if(typeof(vb) != "unknown" || typeof(vb) == "boolean") return false
				return this.setMethod(sKeyPath,sKeyName,vb,"SetBinaryValue")
			}
			catch(ee){
				this.error(ee,"setValueBINARY()")
			}
			return false
		}
		
		this.getValues = function(sKeyPath,sRegString){
			var o, aOValues = false
			if(!this.isString(sKeyPath)) return false
			if(o = this.getMethod(sKeyPath,null,"EnumValues")){
				aOValues = new Array()
				try{
					var aNames = o.sNames.toArray();
					var aTypes = o.Types.toArray();
				}
				catch(ee){ // Empty values
					var aNames = new Array();
					var aTypes = new Array();
					this.error(ee,"getValues()")
				}
				var oRe = typeof(sRegString) == "string" ? new RegExp(sRegString,ig) : false
				for(var i = 0, len = aNames.length; i < len; i++){
					if(oRe && (aNames[i]).match(oRe)) continue
					var oo = new Object()
					oo.name = aNames[i], oo.typeidx = aTypes[i], oo.typestr = typesidx[oo.typeidx]					
					switch(oo.typeidx){
						case this.REG_SZ        : {oo.value = this.getValue(sKeyPath,oo.name); break;}
						case this.REG_EXPAND_SZ : {oo.value = this.getValueEXPAND(sKeyPath,oo.name); break;}
						case this.REG_BINARY    : {oo.value = this.getValueBINARY(sKeyPath,oo.name); break;}
						case this.REG_DWORD     : {oo.value = this.getValueDWORD(sKeyPath,oo.name); break;}
						case this.REG_MULTI_SZ  : {oo.value = this.getValueMULTI(sKeyPath,oo.name); break;}
						default : {oo.value = ""; break;}
					}
					if(this.isDebug) this.echo("name: " + oo.name + " value: " + oo.value + " str: " + oo.typestr + " idx: " + oo.typeidx)
					aOValues.push(oo)
				}
				this.kill(aNames,aTypes,o)
			}
			return aOValues
		}

		this.getKeys = function(sKeyPath){
			try{
				var o
				if(!this.isString(sKeyPath)) return false
				if(o = this.getMethod(sKeyPath,null,"EnumKey")){					
					try{
						return o.sNames.toArray()
					}
					catch(eee){
						if(this.isDebug) this.echo(sKeyName + ": Subkeys are missing")
						//sNames is not defined if subKeys are missing
					}
				}
			}
			catch(ee){
				this.error(ee,"getKeys()")
			}
			return false
		}

		this.getHex = function(iNum){
			if(isNaN(iNum) || typeof(iNum) != "number") return 0
			else if(iNum > 0) return iNum.toString(16)
			else return (iNum + 0x100000000).toString(16);
		}
		
		this.getNum = function(xHex){
			if(isNaN(xHex) || typeof(xHex) != "number") return 0
			return parseInt(xHex);
		}
		
		//// DEL functions
		
		this.delKey = function(sKeyPath){
			try{
				if(this.isLocalhost){
					return oWsh.RegDelete(sKeyPath.replace(this.oReKey,"$1") + "\\" + sKeyName)
				}
				else return this.getMethod(sKeyPath,null,"DeleteKey")
			}
			catch(ee){
				this.error(ee,"delKey()")
				return false
			}
		}
		
		this.delValue = function(sKeyPath,sKeyName){
			try{
				var o				
				if(this.isLocalhost){
					return oWsh.RegDelete(sKeyPath.replace(this.oReKey,"$1") + "\\" + sKeyName)
				}
				else return (this.getMethod(sKeyPath,sKeyName,"DeleteValue") ? true : false)
			}
			catch(ee){
				this.error(ee,"delValue()")
				return false
			}
		}
		
		// ADD functions
		
		this.createKey = function(sKeyPath){
			try{
				if(this.isLocalhost){
					oWsh.RegWrite(sKeyPath.replace(this.oReKey,"$1") + "\\","")
					return true
				}
				else return (this.getMethod(sKeyPath,null,"CreateKey") ? true : false)
			}
			catch(ee){
				this.error(ee,"createKey()")
				return false
			}
		}
				
		// SECURITY
		/*
		this.KEY_QUERY_VALUE        = 0x1
		this.KEY_SET_VALUE          = 0x2
		this.KEY_CREATE_SUB_KEY     = 0x4
		this.KEY_ENUMERATE_SUB_KEYS = 0x8
		this.KEY_NOTIFY             = 0x10
		this.KEY_CREATE_LINK        = 0x20
		this.DELETE                 = 0x10000
		this.READ_CONTROL           = 0x20000
		this.WRITE_DAC              = 0x40000
		this.WRITE_OWNER            = 0x80000
		*/
		this.getAccess = function(sKeyPath,sKeyName,hSecurity){
			if(!this.regprov){
				if(!this.setService(this.wmi_default,this.computer,this.username,this.password)) return null
			}
			if(!this.hasWMI) return false
			this.err_number = this.err_description = this.err_function = ""
			try{
				var o
				if(o = this.getHKEY(sKeyPath)){
					var oMethod = this.regprov.Methods_.Item("CheckAccess")
					var oInParam = oMethod.InParameters.SpawnInstance_();
					oInParam.hDefKey = o.HK_HEX
					oInParam.sSubKeyName = o.KEY;
					oInParam.lRequired = hSecurity
					
					o = this.regprov.ExecMethod_(oMethod.Name,oInParam)
					if(this.isDebug) this.echo(sMethod + " ReturnValue: " + o)					
					return (o == 0 ? true : false)
				}
			}
			catch(ee){
				this.error(ee,"getAccess()")
			}
			return null
		}
		
		this.hasAccess = function(sKeyPath,hSecurity){
			return this.getAccess(sKeyPath,hSecurity)
		}
		
		this.hasAccessREAD = function(sKeyPath){
			return this.getAccess(sKeyPath,(this.KEY_QUERY_VALUE | this.KEY_SET_VALUE | this.READ_CONTROL))
		}
		
		this.hasAccessCHANGE = function(sKeyPath){
			return this.getAccess(sKeyPath,(this.KEY_QUERY_VALUE | this.KEY_SET_VALUE | this.READ_CONTROL | this.DELETE | this.WRITE_DAC))
		}
		
		this.hasAccessFULL = function(sKeyPath){
			return this.getAccess(sKeyPath,(this.KEY_QUERY_VALUE | this.KEY_SET_VALUE | this.KEY_CREATE_SUB_KEY | this.KEY_ENUMERATE_SUB_KEYS | this.DELETE | this.READ_CONTROL | this.WRITE_OWNER))
		}
		
		//// SPECIAL FUNCTIONS
		
		this.getRegistrySize = function(bKB){
			// TODO: Not working.... should use WMI here!
			try{
				 // http://www.windowsitpro.com/Article/ArticleID/14737/14737.html
				 var o = this.getValueDWORD(this.subkey.con,"RegistrySizeLimit")
				 return ((bKB && o) ? o/1024 : o)
			}
			catch(ee){
				this.error(ee,"getRegistrySize()")
			}
			return false
		}
		
		this.setRegistrySize = function(iSize){
			try{
				// http://www.windowsitpro.com/Article/ArticleID/14737/14737.html
				var REGSIZE_MIN = 4096000000
				if(!this.isNumber(iSize) && iSize < REGSIZE_MIN) return false
				return this.setValueDWORD(this.subkey.con,"RegistrySizeLimit",iSize)
			}
			catch(ee){
				this.error(ee,"setRegistrySize()")
			}
		}
		
		this.getPagefile = function(bString){
			try{
				return this.getValueMULTI(this.subkey.mem,"PagingFiles",bString)
			}
			catch(ee){
				this.error(ee,"getPagefile()")
			}
		}
		
		this.setPagefile = function(sFile,iMax,iMin){
			try{
				if(!this.isString(sFile) || !sFile.match(/[a-z]:\\pagefile\.sys$/ig)) return false
				if(!this.isNumber(iMax,iMin) && iMin > iMax) return false
				return this.setValueMULTI(this.subkey.mem,"PagingFiles",sFile.toLowerCase() + " " + iMax + " " + iMin)
			}
			catch(ee){
				this.error(ee,"setPagefile()")
			}
		}		
		
		if(bInit) this.init(oService)
	}
	catch(e){ alert("Registry error: " + e.description)
		js_log_error(2,e);
		return false;
	}
}
