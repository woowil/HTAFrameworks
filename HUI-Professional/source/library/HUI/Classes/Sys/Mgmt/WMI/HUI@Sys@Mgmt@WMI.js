// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**

__H.include("HUI@Sys@Mgmt.js")

__H.register(__H.Sys.Mgmt,"WMI","Windows Management Intrumentation",function WMI(){
	var d_computers = new ActiveXObject("Scripting.Dictionary")
	var d_services  = new ActiveXObject("Scripting.Dictionary")
	var o_locator	= new ActiveXObject("WbemScripting.SWbemLocator")
	var o_refresher = new ActiveXObject("WbemScripting.SWbemRefresher")

	var o_service_tmp  = null
	var a_namespaces   = []
	this.class_win32   = true
	this.class_system  = false
	this.class_cim     = false

	var d_qualifiers   = new ActiveXObject("Scripting.Dictionary");
	var d_refresher    = new ActiveXObject("Scripting.Dictionary");
	var d_colitems     = new ActiveXObject("Scripting.Dictionary");

	// WMI service
	this.wmi = {
		namespace		: new RegExp("root\\\\(cimv2|default|microsofthis)","ig"),
		ns_cimv2		: null,
		ns_default		: null,
		ns_microsofthis	: null,

		ERR_WMI_ACCESSDENIED : -2147217405,
		ERR_WMI_NOTSTARTED   : -2146828218,
		ERR_WMI_NORPC        : -2146828218,

		FlagReturnImmediately    : 0x10, // Used in ExecQuery
		FlagForwardOnly          : 0x20, // Used in ExecQuery
		FlagUseAmendedQualifiers : 0x20000 // Enabling this flag tells WMI to return the entire class definition rather than just the local
	}

	this.initialize = function initialize(){
		try{
			this.wmi.computer = oWno.ComputerName
			this.wmi.namespace = "root\\cimv2"
			this.wmi.username = null
			this.wmi.password = null

			this.wmi.impersonation = 3
			this.wmi.authentication = 6
			this.wmi.privilege = -1
			this.wmi.locale = "MS_409"
			this.wmi.authority = null
			this.wmi.securityflags = 128

			d_computers.RemoveAll()
			d_services.RemoveAll()
		}
		catch(ee){
			__HLog.error(ee,this)
			return null
		}
	}
	this.initialize()

	this.setServiceWMI = function setServiceWMI(sComputer,sNameSpace,sUser,sPass,sImpersonation,sAuthentication,iPrivilege,sLocale,sAuthority,iSecurityFlags){
		try{
			this.wmi.computer = (!__H.isStringEmpty(sComputer) && sComputer.isSearch(/[a-z0-9\-_]+/ig)) ? sComputer : this.wmi.computer
			this.wmi.namespace = !__H.isStringEmpty(sNameSpace) ? sNameSpace : this.wmi.namespace
			this.wmi.username = !__H.isStringEmpty(sUser) ? sUser : this.wmi.username
			this.wmi.password = !__H.isStringEmpty(sPass) ? sPass : this.wmi.password

			__HSys.setSystem(this.wmi.computer,this.wmi.username,this.wmi.password)

			this.wmi.impersonation = typeof(sImpersonation) == "number" ? sImpersonation : 3;
			this.wmi.authentication = typeof(sAuthentication) == "number" ? sAuthentication : 6; // Pkt(4): Authenticates that all data received is from the expected client.
			this.wmi.privilege = typeof(iPrivilege) == "number" ? iPrivilege : -1; // ???: All Privileges, 7: Security
			this.wmi.locale = !__H.isStringEmpty(sLocale) ? sLocale : "MS_409"; //MS_409 is American English
			// Either NTLM {authority=ntlmdomain:DomainName} or Kerberos :{authority=kerberos:DomainName\ServerName}
			// If sUser=DOMAIN\username then sAuthority must be null
			this.wmi.authority = typeof(sAuthority) == "string" ? sAuthority : null;
			this.wmi.securityflags = typeof(iSecurityFlags) == "number" ? iSecurityFlags : 128; // Connection timeout: 0 is indefinitely, 128 sec is max

			if(!d_computers.Exists(this.wmi.computer.toLowerCase())){
				__HLog.log("# Setting WMI Information for system: " + this.wmi.computer)
				if(!__HSys.isadmin || !__HSys.netUseAdmin()){ // WMI crashes otherwise!
					__HLog.logPopup("Either local user is not an Administrator or unable to make a WMI connection using current credentials on system " + this.wmi.computer)
					return false
				}
				d_computers.Add(this.wmi.computer.toLowerCase(),"")
			}
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getServiceWMI = function getServiceWMI(bForce){
		try{
			// http://www.microsoft.com/technet/scriptcenter/scrguide/sas_wmi_vzbp.asp
			//http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/swbemservices.asp

			var sKey = (this.wmi.computer + "::" + this.wmi.namespace).toLowerCase()
			if(bForce && d_computers.Exists(sKey)){
				d_computers.Remove(sKey)
				d_services.Remove(sKey)
			}

			if(!d_services.Exists(sKey)){
				var o
				if(o = o_locator.ConnectServer(this.wmi.computer,this.wmi.namespace,this.wmi.username,this.wmi.password,this.wmi.locale,this.wmi.authority,this.wmi.securityflags,null)){
					o.Security_.ImpersonationLevel = this.wmi.impersonation;
					o.Security_.AuthenticationLevel = this.wmi.authentication;
					if(this.wmi.privilege == -1){
						for(var i = 1; i < 27; i++){
							try{ // Adds all privileges
								o.Security_.Privileges.Add(i); // http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/wbemprivilegeenum.asp
							}
							catch(ee){}
						}
					}
					else o.Security_.Privileges.Add(this.wmi.privilege);
					d_services.Add(sKey,o)
				}
			}
			return (o_service_tmp = d_services(sKey))
		}
		catch(ee){
			var s = "", n = (ee.number & 0xFFFF)
			if(n == 4196) {
				// User credentials cannot be used for local connections
				return this.getServiceWMI(bForce)
			}
			else if(ee.number == this.ERR_WMI_ACCESSDENIED) s = ". Using user: " + sUser + ", pass: ";
			else if(ee.number == this.ERR_WMI_NORPC) s = ". RPC uses port 135";
			else if(ee.number == this.ERR_WMI_NOTSTARTED){
				if(oWsh.Run("%comspec% /c sc \\\\" + this.wmi.computer + " config winmgmt start= auto>nul && sc \\\\" + this.wmi.computer + " start winmgmt>nul",__HIO.hide,true) == 0){
					return this.getServiceWMI(bForce)
				}
			}
			else __HLog.debug("ee.number = " + ee.number + " :: code = " + n + " :: description = " + ee.description)
			__HLog.error(ee,this,s)
			return false
		}
	}

	this.wmiDate = function wmiDate(d){
		try{
			if(__H.isDate(d)){
				d = d.replace(/[\*]/g,0);
				return d.substring(0,4) + "-" + d.substring(4,6) + "-" + d.substring(6,8) + " " + d.substring(8,10) + ":" + d.substring(10,12) + ":" + d.substring(12,14) + "." + d.substring(15,18);
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return "";
	}

	this.wmiArray = function wmiArray(aVBarray,sBreak){
		try{
			if(aVBarray != null){
				var a = new VBArray(aVBarray).toArray();
				sBreak = typeof(sBreak) == "string" ? sBreak : "\n  "
				return sBreak + (a.toString()).replace(/,/g,sBreak)

				/*
				for(var j = 0, sVB = "", iLen = aVB.length; j < iLen; j++){
					 //sVB = sVB + (j == 0 ? "" : sBreak) + sVBName + "-" + (j+1) + ":  " + aVB[j];
					 sVB = sVB + (j == 0 ? "" : sBreak) + aVB[j];
					 a.push(aVB[j])
				}
				aVB.stream = (j == 1) ? sVB.replace(/.+-[0-9]{1,2}:  (.+)$/ig,"$1") : sVB;
				//aVB.array = a;

				return aVB;
				*/
			}
			//throw new Error(8882,"Not a WMI vbArray.")
		}
		catch(ee){
			__HLog.error(ee,this)			
		}
		return false;
	}

	this.isEnabled = function isEnabled(sComputer){
		try{
			return (typeof(GetObject("winmgmts:{impersonationLevel=3,authenticationLevel=6,(Security)}!\\\\\\\\" + sComputer)) == "object")
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.namespaces = function namespaces(sComputer,sNameSpace,sUser,sPass){
		try{
			sNameSpace = typeof(sNameSpace) == "string" ? sNameSpace : "root"
			if(sNameSpace == "root") a_namespaces.length = 0

			if(!this.setServiceWMI(sComputer,sNameSpace,sUser,sPass)) return a_namespaces.sort()
			var o = (this.getServiceWMI()).InstancesOf("__NAMESPACE");

			for(var oEnum = new Enumerator(o); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				if((oItem.Name).isSearch(/ms_[0-9]{3,5}/ig)) continue				
				a_namespaces.push(sNameSpace + "\\" + oItem.Name);
				a_namespaces.concat(this.namespaces(sComputer,sNameSpace + "\\" + oItem.Name,sUser,sPass))
			}
			return a_namespaces; // do not sort here... takes time
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getObjectText = function getObjectText(){
		try{ // Obtains the textual rendition of the object in Managed Object Format (MOF) syntax
			return (o_service_tmp.Get(sClass,this.wmi.FlagUseAmendedQualifiers)).GetObjectText_();
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return false;
	}

	this.setRefresher = function setRefresher(sClass,sComputer){
		try{
			var sKey = (sComputer + "::" + sClass).toLowerCase()
			if(!d_refresher.Exists(sKey)){
				d_refresher.Add(sKey,o_refresher)
				d_colitems.Add(sKey,o_refresher.AddEnum(o_service_tmp,sClass).objectSet)
			}
			d_refresher(sKey).Refresh()
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getRefresher = function getRefresher(sClass,sComputer){
		try{
			var sKey = (sComputer + "::" + sClass).toLowerCase()
			if(!d_colitems.Exists(sKey) && !this.setRefresher(sClass,sComputer)){
				return false
			}
			return d_colitems(sKey)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.setClasses = function setClasses(bWin32,bSystem,bCim){
		try{
			this.class_win32 = (bWin32 || this.class_win32)
			this.class_system = (bSystem || this.class_system);
			this.class_cim = (bCim || this.class_cim);
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getClassesSpecific = function getClassesSpecific(sRegExp){
		try{
			if(!sRegExp) return []
			var a = []
			var oRe = __H.isRegExp(sRegExp) ? sRegExp : new RegExp(sRegExp,"ig")

			for(var oEnum = new Enumerator(o_service_tmp.SubclassesOf()); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				if(!(this.getQualifierName(oItem)) || !d_qualifiers.Exists("dynamic")) continue
				if(!(oItem.Path_.Class).isSearch(oRe)) continue
				a.push(oItem.Path_.Class)
			}

			return a.sort();
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getClassesAll = function getClassesAll(sComputer,sNameSpace,sUser,sPass){
		try{
			var a = []
			if(!this.setServiceWMI(sComputer,sNameSpace,sUser,sPass)) return false

			for(var oEnum = new Enumerator((this.getServiceWMI()).SubclassesOf()); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				if(!(this.getQualifierName(oItem)) || !d_qualifiers.Exists("dynamic")) continue
				a.push(oItem.Path_.Class)
			}

			return a.sort();
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getClassesAllObject = function getClassesAllObject(sComputer,sNameSpace,sUser,sPass){
		try{
			var aClasses = []
			if(!this.setServiceWMI(sComputer,sNameSpace,sUser,sPass)) return false

			for(var oEnum = new Enumerator((this.getServiceWMI()).SubclassesOf()); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item(), s;
				if(!(this.getQualifierName(oItem)) || !d_qualifiers.Exists("dynamic")) continue
				if((s = oItem.Path_.Class)){
					var a = s.split("_");
					if(a[0] == "" || a[0] == null) a[0] = "__";
					if(!this.class_system && a[0] == "__");
					else if(!this.class_cim && a[0].isSearch(/cim/ig));
					else if(!this.class_win32 && a[0].isSearch(/win32/ig));
					else {
						if(!aClasses[a[0]]) aClasses[a[0]] = [];
						aClasses[a[0]].push(s);
					}
				}
			}

			for(var o in aClasses){
				if(aClasses.hasOwnProperty(o)){
					aClasses[o].sort();
				}
			}
			return aClasses;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getClassQualifiers = function getClassQualifiers(sClass){
		try{
			var a = [];
			var oClass = o_service_tmp.Get(sClass);
			for(var oEnum = new Enumerator(oClass.Qualifiers_); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				a.push({
					name  : oItem.Name,
					value : ((oItem.Name.toLowerCase() == "valuemap") ? oItem.Value.replace(/,+/g,", ") : oItem.Value),
					text  : oItem.Name + " = " + sValue
				})
			}
			return a
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getQualifier = function getQualifier(oQual,sStr){
		try{ // http://msdn.microsoft.com/library/en-us/wmisdk/wmi/standard_qualifiers.asp
			var sValue = "";
			for(var i = 1, oEnum = new Enumerator(oQual.Qualifiers_); !oEnum.atEnd(); oEnum.moveNext(), i++){
				var oItem = oEnum.item();
				var sStr2 = sStr ? sStr : i + ". ";
				var a = this.getVBArray(oItem.Value)
				if(oItem.Name.toLowerCase() == "valuemap"){
					a = (a.toString()).replace(/,+/ig,", ");
				}
				sValue = sValue + sStr2 + oItem.Name + " = " + a + "\n";
			}
			return sValue.replace(/(.*)\n$/g,"$1");
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getQualifierName = function getQualifierName(oQual){
		try{ // http://msdn.microsoft.com/library/en-us/wmisdk/wmi/standard_qualifiers.asp
			d_qualifiers.RemoveAll()
			for(var oEnum = new Enumerator(oQual.Qualifiers_); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				d_qualifiers.Add(oItem.Name,"")
			}
			// Association, dynamic, Locale, provider, UUID
			return d_qualifiers;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getQualifierDescription = function getQualifierDescription(oQual){
		try{
			for(var oEnum = new Enumerator(oQual.Qualifiers_); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				if(oItem.Name.toLowerCase() != "description") continue
				try{
					return oItem.Value;
				}
				catch(ee){};
			}
			return ""
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getMethods = function getMethods(sClass){
		try{
			var aMethod = [];
			var oClass = o_service_tmp.Get(sClass);

			for(var oParam, oEnum = new Enumerator(oClass.Methods_); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				var a, o = {}
				// Name
				o.name = oItem.Name;
				// In Parameter
				o.inparam = ""
				if(a = this.getMethodInParameter(oItem)){
					for(var j = 0, iLen = a.length; j < iLen; j++){
						o.inparam = o.inparam + a[j].name + " As " + a[j].typestr + " <" + a[j].cimtype + ">\n";
						delete a[j]
					}
				}
				o.inparam = o.inparam == "" ? o.inparam : (o.inparam).replace(/(.*)\n$/g,"$1")
				// Out Parameter
				oParam = this.getMethodOutParameter(oItem);
				if(oParam.name) o.outparam = oParam.name + " As " + oParam.text + " <" + oParam.cimtype + ">"
				else o.outparam = ""
				o.description = a ? a.desc : ""
				o.qualifiers = this.getQualifier(oItem)
				aMethod.push(o)
			}
			return aMethod;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getMethodInParameter = function getMethodInParameter(oMethod){
		try{
			var oInpars = oMethod.InParameters;
			if(typeof(oInpars) != null){ // Some method parameters are null
				try{
					var a = [], i = 0
					for(var oEnum = new Enumerator(oInpars.Properties_); !oEnum.atEnd(); oEnum.moveNext()){
						a[i++] = this.parameterType(oEnum.item());
					}
				}
				catch(ee){ } // Properties may not exist
				a.name = oMethod.Name;
				a.desc = this.getQualifierDescription(oMethod);
			}
			return a;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getMethodOutParameter = function getMethodOutParameter(oMethod){
		try{
			var oOutpars = oMethod.OutParameters;
			if(typeof(oOutpars) != null){ // Some method parameters are null
				for(var oEnum = new Enumerator(oOutpars.Properties_); !oEnum.atEnd(); oEnum.moveNext()){
					var oItem = oEnum.item();
					return this.parameterType(oItem);
				}
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getVBArray = function getVBArray(aVB){
		try{
			if(typeof(aVB) == "unknown"){
				return (new VBArray(aVB).toArray());
			}
			return [];
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getDerivation = function getDerivation(sClass){
		try{
			var oClass = o_service_tmp.Get(sClass);

			var a = (this.getVBArray(oClass.Derivation_)).reverse();
			var aODerived = []

			for(var i = 0, iLen = a.length; i < iLen; i++){
				var o = {}
				o.Derived = a[i]
				if(i == 0) o.type = "Dynasty";
				else if(i == (aDerived.length-1)) o.type = "SuperClass";
				else o.type = "Class";
				aODerived.push(o)
			}
			return aODerived
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getPropertyNames = function getPropertyNames(sClass){
		try{ // http://www.microsoft.com/technet/scriptcenter/scrguide/sas_wmi_yzik.asp
			var oClass = o_service_tmp.Get(sClass);
			var a = [];
			for(var oEnum = new Enumerator(oClass.Properties_); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				a.push(oItem.Name)
			}
			return a;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getProperty = function getProperty(sClass){
		try{ // http://www.microsoft.com/technet/scriptcenter/scrguide/sas_wmi_yzik.asp
			var oClass = o_service_tmp.Get(sClass);
			var a = [];
			for(var oEnum = new Enumerator(oClass.Properties_); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				a.push(this.parameterType(oItem))
			}
			return a;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getPropertyExtended = function getPropertyExtended(sClass){
		try{ // http://www.microsoft.com/technet/scriptcenter/scrguide/sas_wmi_yzik.asp__H
			var oClass = o_service_tmp.Get(sClass);
			var a = [];

			for(var oEnum = new Enumerator(oClass.Properties_); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				var o = {}, n1
				o.name = n1 = oItem.name
				var n2 = n1.substring(1,n1.length), n2 = n2.replace(/[_]/g,"")
				o.name2       = n1.substring(0,1) + n2.replace(/([A-Z_])/g," $1") // Create spaces for every big letters except the first one
				o.description = this.getQualifierDescription(oItem)
				o.isarray     = oItem.IsArray ? true : false;
				o.isdatetime  = oItem.CIMType == 101 ? true : false;
				o.isnumber    = (typeof(oItem.Value) == "number")
				o.cimtype     = oItem.CIMType
				a.push(o)
			}
			return a;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.parameterType = function parameterType(oProperty){
		try{
			var o = {}
			switch(oProperty.CIMType){
				case 0 : o.typestr = "Empty (NULL) Value", o.type = "CIM_EMPTY"; break;
				case 2 : o.typestr = "16-bit Signed Integer", o.type = "CIM_SINT16"; break;
				case 3 : o.typestr = "32-bit Signed Integer", o.type = "CIM_SINT32"; break;
				case 4 : o.typestr = "32-bit Real Number", o.type = "CIM_REAL32"; break;
				case 5 : o.typestr = "64-bit Real Number", o.type = "CIM_REAL64"; break;
				case 8 : o.typestr = "String Value", o.type = "CIM_STRING"; break;
				case 11 : o.typestr = "Boolean Value", o.type = "CIM_BOOLEAN"; break;
				case 13 : o.typestr = "Object Value", o.type = "CIM_OBJECT"; break;
				case 16 : o.typestr = "8-bit Signed Integer", o.type = "CIM_SINT8"; break;
				case 17 : o.typestr = "8-bit UnSigned Integer", o.type = "CIM_USINT8"; break;
				case 18 : o.typestr = "16-bit UnSigned Integer", o.type = "CIM_USINT16"; break;
				case 19 : o.typestr = "32-bit UnSigned Integer", o.type = "CIM_USINT32"; break;
				case 20 : o.typestr = "64-bit Signed Integer", o.type = "CIM_SINT64";	break;
				case 21 : o.typestr = "64-bit UnSigned Integer", o.type = "CIM_USINT64"; break;
				case 101 : o.typestr = "Datetime Value", o.type = "CIM_DATETIME"; break;
				case 102 : o.typestr = "Reference of a CIM Object", o.type = "CIM_REFERENCE"; break;
				case 103 : o.typestr = "16-bit Character Value", o.type = "CIM_CHAR16"; break;
				case 4095 : o.typestr = "Illegal Value", o.type = "CIM_ILLEGAL"; break;
				case 8192 : o.typestr = "Array Value", o.type = "CIM_FLAG_ARRAY"; break;
				default : o.typestr = "Unknown", o.type = "CIM_UNKNOWN"; break;
			}

			if(oProperty.IsArray) o.typestr += "[]", o.isarray = true;
			o.text = o.typestr + (o.type == "" ? "" : " (" + o.type + ")")
			o.cimtype = oProperty.CIMtype
			o.value = oProperty.Value
			o.name = oProperty.Name
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return o;
	}
})

var __HWMI = new __H.Sys.Mgmt.WMI()