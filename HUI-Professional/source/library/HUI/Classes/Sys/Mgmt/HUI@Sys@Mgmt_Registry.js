// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**


__H.include("HUI@Sys@Mgmt.js","HUI@Sys@Mgmt@WMI.js")

__H.register(__H.Sys.Mgmt,"Registry","Accessing/Editing Local/Remote Registry using WMI",function Registry(){
	// http://msdn.microsoft.com/library/en-us/wmisdk/wmi/stdregprov.asp
	// http://www.activexperts.com/activmonitor/windowsmanagement/adminscripts/registry/
	// http://cwashington.netreach.net/depo/view.asp?Index=560&ScriptType=vbscript
	// http://www.serverwatch.com/tutorials/article.php/1476831
	// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/writing_wmi_scripts_in_jscript.asp
	// http://www.serverwatch.com/tutorials/print.php/1476831
	// http://www.codecomments.com/archive298-2004-5-191525.html		
	
	this.regprov = null
	this.oReKey = /[\\]{0,1}(.+)[\\]{0,1}$/ig
	this.oReMethodNoValue = /EnumKey|EnumValues|DeleteKey|CreateKey/ig
	
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
	
	var o_service = null
	var b_loaded = false
	var a_typesidx = [
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
		"REG_QWORD"];    // 11
	
	this.reset = function reset(){} // Don't REMOVE
	
	this.initialize = function initialize(sComputer,sUser,sPass,bForce){
		try{
			if(b_loaded && !bForce) return true;
			if(__HWMI.setServiceWMI(sComputer,"root\\default",sUser,sPass)){
				o_service = __HWMI.getServiceWMI(bForce)
				this.regprov = o_service.Get("StdRegProv")				
			}
			b_loaded = o_service && this.regprov
			return b_loaded
		}
		catch(ee){			
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getEnvironment = function getEnvironment(sKeyName,sHive){
		try{			
			if(__HSys.isLocalHost()){
				if(typeof(sHive) != "string") sHive = ""
				// sHive: Win200X/NT/XP - SYSTEM, USER, 
				// sHive: Win95/98 - PROCESS
				sHive = sHive.isSearch(/system|user|process/ig) ? sHive.toUpperCase() : "SYSTEM"
				var oEnv = oWsh.Environment(sHive); 
				return oEnv(sKeyName); // sItem: NUMBER_OF_PROCESSORS, SYSTEMROOT etc
			}
			return this.Exists(__HKeys.env,sKeyName)
		}
		catch(ee){
			__HLog.error(ee,this);	
		}
		return false;
	}
	
	this.getMethod = function getMethod(sKeyPath,sKeyName,sMethod){
		this.reset()
		if(!this.regprov && !this.initialize()) return null
		try{
			var o
			if(o = this.getHKEY(sKeyPath)){
				var oMethod = this.regprov.Methods_.Item(sMethod)
				var oInParam = oMethod.InParameters.SpawnInstance_();
				oInParam.hDefKey = o.HK_HEX
				oInParam.sSubKeyName = o.KEY;
				if(!sMethod.isSearch(this.oReMethodNoValue)) oInParam.sValueName = typeof(sKeyName) == "string" ? sKeyName : "";
				// http://msdn.microsoft.com/library/en-us/wmisdk/wmi/swbemobject_execmethod_.asp
				o = this.regprov.ExecMethod_(oMethod.Name,oInParam)
				__HLog.debug(sKeyName + ": " + sMethod + " ReturnValue: " + o.ReturnValue)
				return (o.ReturnValue == 0 ? o : null)
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return null
	}
				
	this.setMethod = function setMethod(sKeyPath,sKeyName,sKeyValue,sMethod){
		this.reset()
		if(!this.regprov && !this.initialize()) return false
		try{
			var o
			if(o = this.getHKEY(sKeyPath)){
				var oMethod = this.regprov.Methods_.Item(sMethod)
				var oInParam = oMethod.InParameters.SpawnInstance_();
				oInParam.hDefKey = o.HK_HEX
				oInParam.sSubKeyName = o.KEY;
				oInParam.sValueName = typeof(sKeyName) == "string" ? sKeyName : "";;
				
				if(sMethod.isSearch(/StringValue/ig)) oInParam.sValue = sKeyValue
				else if(sMethod.isSearch(/DWORD|Binary/ig)) oInParam.uValue = sKeyValue // DWORD and BINARY
				else return false
				// http://msdn.microsoft.com/library/en-us/wmisdk/wmi/swbemobject_execmethod_.asp
				o = this.regprov.ExecMethod_(oMethod.Name,oInParam)
				return (o.ReturnValue == 0)
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return false
	}
	
	//// GET functions		

	this.getHKEY = function getHKEY(sKeyPath){
		if(typeof(sKeyPath) != "string") return null
		sKeyPath = sKeyPath.replace(this.oReKey,"$1")
		var o = {}
		if(sKeyPath.match(/(?:HKEY_LOCAL_MACHINE|HKLM)\\(.+)$/ig)) o.HK_HEX = this.HKLM, o.KEY = RegExp.$1
		else if(sKeyPath.match(/(?:HKEY_CURRENT_USER|HKCU)\\(.+)$/ig)) o.HK_HEX = this.HKCU, o.KEY = RegExp.$1
		else if(sKeyPath.match(/(?:HKEY_CLASSES_ROOT|HKCR)\\(.+)$/ig)) o.HK_HEX = this.HKCR, o.KEY = RegExp.$1
		else if(sKeyPath.match(/(?:HKEY_USERS|HKU)\\(.+)$/ig)) o.HK_HEX = this.HKU, o.KEY = RegExp.$1
		else if(sKeyPath.match(/(?:HKEY_CURRENT_CONFIG|HKCC)\\(.+)$/ig)) o.HK_HEX = this.HKCC, o.KEY = RegExp.$1
		else if(sKeyPath.match(/(?:HKEY_DYN_DATA|HKDD)\\(.+)$/ig)) o.HK_HEX = this.HKDD, o.KEY = RegExp.$1
		else return null
		return o
	}
	
	this.Exists = function Exists(sKeyPath,sKeyName,bString){
		try{ // only for string
			var o
			if(__HSys.isLocalHost()){
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
			__HLog.error(ee,this)
		}
		return false
	}
	
	this.getValueSTRING = this.getValue = function getValue(sKeyPath,sKeyName){
		try{
			var o
			if(!__H.isString(sKeyPath,sKeyName)) throw new Error(__HLog.errorCode("error"),"Argument 0 or 1 is undefined")
			if(__HSys.isLocalHost()){
				return oWsh.RegRead(sKeyPath.replace(this.oReKey,"$1") + "\\" + sKeyName)
			}
			else if(o = this.getMethod(sKeyPath,sKeyName,"GetStringValue")){
				return o.sValue;
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return false
	}
	
	this.setValueSTRING = this.setValue = function setValue(sKeyPath,sKeyName,sKeyValue){
		try{
			if(!__H.isString(sKeyPath,sKeyName)) return false // Important! sKeyValue can be "",do not add here
			if(typeof(sKeyValue) != "string") return false
			
			if(__HSys.isLocalHost()){
				oWsh.RegWrite(sKeyPath.replace(this.oReKey,"$1") + "\\" + sKeyName,sKeyValue,"REG_SZ")
				return true
			}
			else return this.setMethod(sKeyPath,sKeyName,sKeyValue,"SetStringValue")
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return false
	}		
	
	this.getValueEXPAND = function getValueEXPAND(sKeyPath,sKeyName){
		try{
			var o
			if(!__H.isString(sKeyPath,sKeyName)) return false
			if(__HSys.isLocalHost()){
				if(o = oWsh.RegRead(sKeyPath.replace(this.oReKey,"$1") + "\\" + sKeyName)){
					return o.toString()
				}
			}
			else if(o = this.getMethod(sKeyPath,sKeyName,"GetExpandedStringValue")){
				return o.sValue;
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return false
	}

	this.getValueMULTI = function getValueMULTI(sKeyPath,sKeyName,bString){
		var o
		if(!__H.isString(sKeyPath,sKeyName)) return false
		if(__HSys.isLocalHost()){
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
	
	this.setValueMULTI = function setValueMULTI(sKeyPath,sKeyName,aKeyValues){
		try{ // aKeyValues: Either an array or a string
			if(typeof(vb) != "unknown" || typeof(vb) == "boolean") return false
			return this.setMethod(sKeyPath,sKeyName,__H.toVBArray(aKeyValues),"SetMultiStringValue")
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return false
	}
	
	this.getValueDWORD = function getValueDWORD(sKeyPath,sKeyName){
		var o
		if(!__H.isString(sKeyPath,sKeyName)) return false
		if(__HSys.isLocalHost()){
			return oWsh.RegRead(sKeyPath.replace(this.oReKey,"$1") + "\\" + sKeyName)
		}
		else if(o = this.getMethod(sKeyPath,sKeyName,"GetDWORDValue")){
			return o.uValue; // Note that this uses uValue instead of sValue
		}
		return false
	}

	this.setValueDWORD = function setValueDWORD(sKeyPath,sKeyName,iKeyValue){
		try{
			if(!__H.isString(sKeyPath,sKeyName)) return false
			if(typeof(iKeyValue) != "number") return false
			//if(__HSys.isLocalHost()) return oWsh.RegWrite(sKeyPath + "\\" + sKeyName,this.getNum(iKeyValue),"REG_DWORD") // Can't use, get buffer overflow on Hex numbers
			return this.setMethod(sKeyPath,sKeyName,iKeyValue,"SetDWORDValue") // sKeyValue: Either hex or number
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return false
	}
	
	this.getValueBINARY = function getValueBINARY(sKeyPath,sKeyName,bString){
		try{
			var o
			if(!__H.isString(sKeyPath,sKeyName)) return false
			if(__HSys.isLocalHost()){
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
			__HLog.error(ee,this)
		}
		return false
	}
	
	this.setValueBINARY = function setValueBINARY(sKeyPath,sKeyName,aKeyValues){
		try{ // aKeyValues: Either an array of 0x01 0xf2 or a hex number
			if(!__H.isString(sKeyPath,sKeyName)) return false
			if(typeof(aKeyValues) == "number") aKeyValues = new Array(aKeyValues)
			var vb = this.setVBArray(aKeyValues,"number")
			if(typeof(vb) != "unknown" || typeof(vb) == "boolean") return false
			return this.setMethod(sKeyPath,sKeyName,vb,"SetBinaryValue")
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return false
	}
	
	this.getValues = function getValues(sKeyPath,sRegString){
		try{
			var o, aOValues = []
			if(!__H.isString(sKeyPath)) return false
			if(o = this.getMethod(sKeyPath,null,"EnumValues")){
				try{
					var aNames = o.sNames.toArray();
					var aTypes = o.Types.toArray();
				}
				catch(ee){ // Empty values
					var aNames = [];
					var aTypes = [];
					//__HLog.error(ee,this)
				}
				var oRe = typeof(sRegString) == "string" ? new RegExp(sRegString,ig) : false
				for(var i = 0, iLen = aNames.length; i < iLen; i++){
					if(oRe && (aNames[i]).isSearch(oRe)) continue
					var oo = {}
					oo.Name = aNames[i], oo.Typeidx = aTypes[i], oo.Typestr = a_typesidx[oo.Typeidx]
					switch(oo.Typeidx){
						case this.REG_SZ        : {oo.Value = this.getValue(sKeyPath,oo.Name); break;}
						case this.REG_EXPAND_SZ : {oo.Value = this.getValueEXPAND(sKeyPath,oo.Name); break;}
						case this.REG_BINARY    : {oo.Value = this.getValueBINARY(sKeyPath,oo.Name); break;}
						case this.REG_DWORD     : {oo.Value = this.getValueDWORD(sKeyPath,oo.Name); break;}
						case this.REG_MULTI_SZ  : {oo.Value = this.getValueMULTI(sKeyPath,oo.Name); break;}
						default : {oo.Value = ""; break;}
					}
					__HLog.debug("name: " + oo.Name + " value: " + oo.Value + " str: " + oo.Typestr + " idx: " + oo.Typeidx)
					aOValues.push(oo)
				}
				//__HUtil.kill(aNames,aTypes,o)
			}
			return aOValues
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getKeys = function getKeys(sKeyPath){
		try{
			var o
			if(__H.isStringEmpty(sKeyPath)) return false
			if(o = this.getMethod(sKeyPath,null,"EnumKey")){					
				try{
					return o.sNames.toArray()
				}
				catch(eee){
					__HLog.debug(eee.description)
					__HLog.debug(sKeyPath + ": Subkey(s) are missing")
					//sNames is not defined if subKeys are missing
				}
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return false
	}

	this.getHex = function getHex(iNum){
		if(isNaN(iNum) || typeof(iNum) != "number") return 0
		else if(iNum > 0) return iNum.toString(16)
		else return (iNum + 0x100000000).toString(16);
	}
	
	this.getNum = function getNum(xHex){
		if(isNaN(xHex) || typeof(xHex) != "number") return 0
		return parseInt(xHex);
	}
	
	//// DEL functions
	
	this.delKey = function delKey(sKeyPath){
		try{
			if(__HSys.isLocalHost()){
				return oWsh.RegDelete(sKeyPath.replace(this.oReKey,"$1") + "\\" + sKeyName)
			}
			else return this.getMethod(sKeyPath,null,"DeleteKey")
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.delValue = function delValue(sKeyPath,sKeyName){
		try{
			var o				
			if(__HSys.isLocalHost()){
				return oWsh.RegDelete(sKeyPath.replace(this.oReKey,"$1") + "\\" + sKeyName)
			}
			else return (this.getMethod(sKeyPath,sKeyName,"DeleteValue") ? true : false)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	// ADD functions
	
	this.createKey = function createKey(sKeyPath){
		try{
			if(__HSys.isLocalHost()){
				oWsh.RegWrite(sKeyPath.replace(this.oReKey,"$1") + "\\","")
				return true
			}
			else return (this.getMethod(sKeyPath,null,"CreateKey") ? true : false)
		}
		catch(ee){
			__HLog.error(ee,this)
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
	this.getAccess = function getAccess(sKeyPath,sKeyName,hSecurity){
		if(!this.regprov && !this.initialize()) return null
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
				__HLog.debug(sMethod + " ReturnValue: " + o)					
				return (o == 0 ? true : false)
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return null
	}
	
	this.hasAccess = function hasAccess(sKeyPath,hSecurity){
		return this.getAccess(sKeyPath,hSecurity)
	}
	
	this.hasAccessREAD = function hasAccessREAD(sKeyPath){
		return this.getAccess(sKeyPath,(this.KEY_QUERY_VALUE | this.KEY_SET_VALUE | this.READ_CONTROL))
	}
	
	this.hasAccessCHANGE = function hasAccessCHANGE(sKeyPath){
		return this.getAccess(sKeyPath,(this.KEY_QUERY_VALUE | this.KEY_SET_VALUE | this.READ_CONTROL | this.DELETE | this.WRITE_DAC))
	}
	
	this.hasAccessFULL = function hasAccessFULL(sKeyPath){
		return this.getAccess(sKeyPath,(this.KEY_QUERY_VALUE | this.KEY_SET_VALUE | this.KEY_CREATE_SUB_KEY | this.KEY_ENUMERATE_SUB_KEYS | this.DELETE | this.READ_CONTROL | this.WRITE_OWNER))
	}
	
	//// SPECIAL FUNCTIONS
	
	this.getRegistrySize = function getRegistrySize(bKB){
		// TODO: Not working.... should use WMI here!
		try{
			 // http://www.windowsitpro.com/Article/ArticleID/14737/14737.html
			 var o = this.getValueDWORD(__HKeys.con,"RegistrySizeLimit")
			 return ((bKB && o) ? o/1024 : o)
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return false
	}
	
	this.setRegistrySize = function setRegistrySize(iSize){
		try{
			// http://www.windowsitpro.com/Article/ArticleID/14737/14737.html__H
			var REGSIZE_MIN = 4096000000
			if(!this.isNumber(iSize) && iSize < REGSIZE_MIN) return false
			return this.setValueDWORD(__HKeys.con,"RegistrySizeLimit",iSize)
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.getPagefile = function getPagefile(bString){
		try{
			return this.getValueMULTI(__HKeys.mem,"PagingFiles",bString)
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.setPagefile = function setPagefile(sFile,iMax,iMin){
		try{
			if(!__H.isString(sFile) || !sFile.isSearch(/[a-z]:\\pagefile\.sys$/ig)) return false
			if(!this.isNumber(iMax,iMin) && iMin > iMax) return false
			return this.setValueMULTI(__HKeys.mem,"PagingFiles",sFile.toLowerCase() + " " + iMax + " " + iMin)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	////
	
	this.getRegistryObject = function getRegistryObject(){
		try{
			if(__HWMI.setServiceWMI(null,"root\\cimv2")){
				oService = __HWMI.setServiceWMI()
			}
			else return false
			var oColItems = oService.ExecQuery("Select * from Win32_Registry","WQL",48);

			for(var oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item(), d
				var o = {
					Caption      : oItem.Caption,
					CurrentSize  : oItem.CurrentSize,
					Description  : oItem.Description,
					InstallDate  : ((d=__HWMI.date(oItem.InstallDate)) ? d : "undefined"),
					MaximumSize  : oItem.MaximumSize,					
					Name         : oItem.Name,
					ProposedSize : oItem.ProposedSize,
					Status       : oItem.Status
				}
				return o
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	/*
	Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7
Const HKLM = &H80000002
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
 


' Driver code.
Dim objReg : Set objReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
ReturnValue = ListSubKeys( HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products", "DisplayName", "KB403742")
WScript.Echo ReturnValue

' Function to retrieve specified match value
Function ListSubKeys(iKey, strKey, sKeyLookup, sKeyContent)

    objReg.EnumKey HKLM, strKey, arrSubkeys
    If IsArray(arrSubkeys) Then
        For Each strSubkey In arrSubkeys
            objReg.EnumValues HKLM,strKey & "\" & strSubkey,arrValues,arrTypes
            If isArray(arrValues) Then
                For i = 0 To Ubound(arrValues)
                    Select Case arrTypes(i)
                        Case REG_SZ
                            If arrValues(i) = sKeyLookup Then
                                objReg.GetStringValue HKLM, strKey & "\" & strSubkey, arrValues(i), strValue1
                                If InStr(strValue1, sKeyContent) Then
                                    'WScript.Echo "REG_SZ[" & arrValues(i)  & "]" & "[" & strValue1 & "]"
                                    ListSubKeys = strValue1
                                    Exit Function ' we got teh match so return with it
                                End If
                            End If
                        Case Else ' only interested in strings right now
                        'WScript.Echo  arrTypes(i)
                    End Select
                Next
            End If
            ListSubKeys = ListSubKeys( iKey, strkey & "\" & strSubKey, sKeyLookup, sKeyContent)
            If Not IsEmpty( ListSubKeys ) Then
                Exit Function
            End If
        Next
       
    End If

End Function
*/
	this.find = function find(sFile,iMax,iMin){
		try{
			
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
})

///////////////////////////

var __HReg = __HRegistry = new __H.Sys.Mgmt.Registry()
