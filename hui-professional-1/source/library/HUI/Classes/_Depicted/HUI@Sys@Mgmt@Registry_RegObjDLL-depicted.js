// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**


__H.include("HUI@Sys@Mgmt@Registry.js")

// This is for Windows OS that doesn't have WMI enabled or installed, f.ex Windows 95/98/NT
__H.register(__H.Sys.Mgmt.Registry,"RegObjDLL","Local/Remote registry using regobj.dll",function RegObjDLL(){
	var computer = oWno.Computer
	var o_regkey = null
	var o_parent = null
	var o_subkeys = null
	var o_values = null
	var s_keypath = ""
	var s_keypathorg = ""
	var s_regname = ""
	var s_regobjdll
	var i_regtype
	var b_ignorecase = false

	var regobjx = null
	var TXT_NOT_COMPUTER = "Not a computer"
	var TXT_NOT_DEFINED  =  "Computer is not defined"
	var TXT_DLL_NOT_LOADED = "DDL file not entered in argument or unable to load component for regobj.dll"

	this.key = new ActiveXObject("Scripting.Dictionary")
	this.key.add(0,"REG_NONE"); // No value type
	this.key.add(1,"REG_SZ"); // String, A sequence of characters representing human readable text. Unicode null terminated string.
	this.key.add(2,"REG_EXPAND_SZ"); // String, An expandable data string, which is text that contains a variable to be replaced when called by an application (ex : %windir%System\wsock32.dll ). Unicode null terminated string.
	this.key.add(3,"REG_BINARY"); // VBArray Raw binary data. Most hardware component information is stored as binary data, and can be displayed in hexadecimal format
	this.key.add(4,"REG_DWORD"); // = REG_DWORD_LITTLE_ENDIAN, Number, 32 bits number
	this.key.add(5,"REG_DWORD_BIG_ENDIAN"); // 32 bits number but in big endian format
	this.key.add(6,"REG_LINK"); // A symbolic link (unicode)
	this.key.add(7,"REG_MULTI_SZ"); // VBArray, A multiple string. Values that contain lists or multiple values in human readable text are usually this type (unicode). Entries are separated by NULL characters.
	this.key.add(8,"REG_RESOURCE_LIST"); // Resource list in the resource map
	this.key.add(9,"REG_FULL_RESOURCE_DESCRIPTOR"); // Resource list in the hardware description
	this.key.add(10,"REG_RESSOURCE_REQUIREMENT_MAP"); // Resource list in the hardware description

	var regrem_d

	this.initialize = function initialize(sComputer,sUser,sPass,sRegObjDLL){
		try{
			var debug = __H.$debug
			__H.$debug = true
			__HSys.setSystem(sComputer,sUser,sPass)
			i_regtype = __HKeys.REG_SZ
			regrem_d = new ActiveXObject("Scripting.Dictionary")
			if(!__HSys.ping() || !__HSys.netUseAdmin()) return false
			
			// The ActiveX Control of RegObj.dll
			if(!regobjx && !(regobjx = __HSys.regActiveX("RegObj.Registry",sRegObjDLL))){
				throw new Error(__HLog.errorCode("error"),TXT_DLL_NOT_LOADED)
			}
			if(__H.isString(s_regobjdll)) s_regobjdll = sRegObjDLL

			return regobjx
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
		finally{
			__H.$debug = debug
		}
	}

	this.deinitialize = function deinitialize(sComputer,bALL){
		if(bAll){
			regrem_d.RemoveAll()
			regobjx = null
			__HSys.regsvr32(s_regobjdll,true)
		}
		else regrem_d.Remove(sComputer)
	}

	this.getRegObjDLL = function getRegObjDLL(){
		return regobjx || initialize()
	}

	this.remoteRegistry = function remoteRegistry(sComputer){
		try{
			if(!regrem_d.Exists(sComputer)){
				this.initialize(sComputer)
				if(__HSys.isLocalHost()) regrem_d.Add(sComputer,regobjx);
				else regrem_d.Add(sComputer,regobjx.RemoteRegistry(sComputer));
			}
			computer = sComputer
			return regrem_d(sComputer)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getKeyPath = function getKeyPath(){
		try{
			return (this.remoteRegistry(computer)).RegKeyFromString(s_keypath)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.setRegistry = function setRegistry(sKeyPath,sRegName,sRegValue,iRegType,bIgnoreCase,bCreate){
		b_ignorecase = bIgnoreCase
		sRegName = typeof(sRegName) == "string" ? sRegName : ""
		s_regname = b_ignorecase ? sRegName.toUpperCase() : sRegName;
		i_regtype = iRegType ? iRegType : __HKeys.REG_SZ;
		s_keypath = (sKeyPath.substring(0,1) != "\\") ? "\\" + sKeyPath : sKeyPath;
		s_keypathorg = (sKeyPath.substring(0,1) != "\\") ? sKeyPath : sKeyPath.substring(1,sKeyPath.length);

		try{
			if(o_regkey = this.getKeyPath()){
				o_parent  = o_regkey.Parent;
				o_values  = o_regkey.Values;
				s_keypath = o_regkey.FullName;
				o_subkeys = o_regkey.SubKeys
				return true
			}
		}
		catch(ee){
			//Error: The name is not in use for a subkey or named value ( = does not exist)
			if(bCreate){
				if(this.addKey(s_keypath,s_regname,sRegValue,i_regtype)){
					__HLog.debug("Creating missing key: \\\\" + computer + "\\" + s_keypath);
					return this.addKeyName(s_keypath,s_regname,sRegValue,i_regtype);
				}
			}
		}
		return false
	}

	this.setEnumValues = function setEnumValues(bForce){
		if(!this.enumvalues || bForce) this.enumvalues = new Enumerator(o_values)
		this.enumvalues.moveFirst()
		return true
	}

	this.getValue = function getValue(sKeyPath,sRegName,sRegValue,iRegType,bIgnoreCase){
		if(!this.setRegistry(sKeyPath,sRegName,null,iRegType,bIgnoreCase)) return false
		for(var sValue = false, oEnum = this.setEnumValues(); !oEnum.atEnd(); oEnum.moveNext()){
			var oKey = oEnum.item();
			if(this.isEqualKey(oKey.Name,s_regname)){
				if(oKey.Type == __HKeys.REG_BINARY){  // REG_BINARY is a VB Array of Integers
					if(__HSys.isLocalHost()){
						try{
							sValue = oWsh.RegRead(s_keypathorg + "\\" + s_regname,"REG_BINARY")
							sValue = ( new VBArray(sValue) ).toArray(); // array with decimal values
						}
						catch(ee){
							sValue = false
						}
					}
					else sValue = oKey.Value//new VBArray(oKey.Value).toArray(); // returns an Array
				}
				else if(oKey.Type == __HKeys.REG_MULTI_SZ){ // REG_MULTI_SZ is a VB Array of Strings separated with null (\0) charactors
					sValue = oKey.Value, sValueNew = "";
					for(var i = 0, ch, iLen = sValue.length; i < iLen; i++){
						if((ch = sValue.substring(i,i+1)) == "\0") sValueNew = sValueNew.concat("\\0");
						else sValueNew = sValueNew.concat(ch)
					}
					sValue = sValueNew.replace("\\0","");
				}
				else if(oKey.Type == __HKeys.REG_DWORD) sValue = oKey.Value;
				else if(oKey.Type == __HKeys.REG_SZ) sValue = oKey.Value;
				else sValue = oKey.value;
				break;
			}
		}
		__HUtil.kill(oEnum,this.enumvalues);
		return sValue
	};

	this.setValue = function setValue(sKeyPath,sRegName,sRegValue,iRegType,bIgnoreCase){
		if(!this.setRegistry(sKeyPath,sRegName,sRegValue,iRegType,bIgnoreCase)) return false
		this.regvalue = sRegValue ? sRegValue : this.regvalue
		for(var bResult = false, oEnum = this.setEnumValues(); !oEnum.atEnd(); oEnum.moveNext()){
			var oKey = oEnum.item();
			if(this.isEqualKey(oKey.Name,s_regname)){
				if(!__H.isStringEmpty(this.regvalue)){
					oKey.value = this.regvalue;
					bResult = true
					break;
				}
			}
		}
		__HUtil.kill(oEnum,this.enumvalues);
		return bResult
	};

	this.getKeyType = function getKeyType(sKeyPath,sRegName){
		if(!this.setRegistry(sKeyPath,sRegName)) return false
		var oResult = false
		for(var oEnum = this.setEnumValues(); !oEnum.atEnd(); oEnum.moveNext()){
			var oKey = oEnum.item(), s;
			if(this.isEqualKey(oKey.Name,s_regname)){
				var o = {
					key : s_keypath,
					name : oKey.Name,
					itype : oKey.type,
					stype : ((s = this.key(oKey.type)) == undefined ? "" : s),
					value : this.getValue(s_keypath,s_regname,null,oKey.Type,false),
					object : typeof(oKey.value)
				}
				oResult = o
				break;
			}
		}

		return oResult
	};

	this.addKeyName = function addKeyName(sKeyPath,sRegName,sRegValue,iRegType){
		if(!this.setRegistry(sKeyPath,sRegName,sRegValue,iRegType,null,true)) return false
		if(!__H.isStringEmpty(this.regvalue)){
			if(!__H.isStringEmpty(s_regname)) this.delKeyName(s_keypath,s_regname);
			else s_regname = "";
			try{
				o_values.Add(s_regname,this.regvalue,i_regtype);
				o_values.Reset();
				return true
			}
			catch(ee){}
		}
		return false
	};

	this.delKeyName = function delKeyName(sKeyPath,sRegName,aRegName){
		if(!this.setRegistry(sKeyPath,sRegName)) return false
		if(typeof(aRegName) == "object" && aRegName.length > 0) var a = aRegName;
		else var a = [sRegName];
		for(var i = 0, iLen = a.length; i < iLen; i++){
			try{
				o_values.Remove(a[i]);
				o_values.Reset();
				return true
			}
			catch(ee){}
		}
		return false
	};

	this.addKey = function addKey(sKeyPath){
		if(!this.setRegistry(sKeyPath)) return false
		try{
			var oRe = /([a-z_]+)\\(.*)/i, aKey; // HKEY_LOCAL_MACHINE
			if(aKey = sKeyPath.isSearch(oRe)){
				o_regkey = this.getKeyPath("\\" + aKey[1]);
				o_regkey.SubKeys.Add(aKey[2]);
				return true
			}
		}
		catch(ee){}
		return false
	};

	this.delKey = function delKey(sKeyPath,sRegName){
		if(!this.setRegistry(sKeyPath,sRegName)) return false
		try{
			o_regkey.SubKeys.Remove(s_regname);
			o_regkey.SubKeys.Reset();
			return true
		}
		catch(ee){}
		return false
	};

	this.copyValues = function copyValues(sKeyPath,sRegNameFrom,sRegNameTo){
		if(!this.setRegistry(sKeyPath,sRegNameFrom,sRegNameTo)) return false
		var oFrom = o_regkey.SubKeys(s_regname).Values
		var oTo = o_regkey.SubKeys(this.regvalue).Values
		for(var oEnum = new Enumerator(oFrom); !oEnum.atEnd(); oEnum.moveNext()){
			var oKey = oEnum.item()
			try{
				oTo.Add(oKey.Name,oKey.Value,oKey.Type);
			}
			catch(ee){
				return false
			} // The name is already used for a subkey or named value
		}
		return true
	};

	this.getKeys = function getKeys(sKeyPath){
		if(!this.setRegistry(sKeyPath)) return false
		var a = [];
		for(var oEnum = new Enumerator(o_subkeys); !oEnum.atEnd(); oEnum.moveNext()){
			var oKey = oEnum.item();
			a.push(oKey.Name);
		}
		return (a.length > 0) ? a : false
	};

	this.getValues = function getValues(sKeyPath){
		if(!this.setRegistry(sKeyPath)) return false
		var a = [];
		for(var oEnum = new Enumerator(o_values); !oEnum.atEnd(); oEnum.moveNext()){
			var oKey = oEnum.item();
			var o = {
				name  : oKey.Name,
				value : oKey.Value,
				type  : oKey.Type
			}
			a.push(o);
		}
		return a
	}

	this.isEqualKey = function isEqualKey(sName1,sName2){
		if(b_ignorecase) sName2 = sName1.toUpperCase(), sName2 = sName2.toUpperCase()
		return (sName1 === sName2 && sName1 == sName2)
	}
})

var __HRegDLL = new __H.Sys.Mgmt.Registry.RegObjDLL()