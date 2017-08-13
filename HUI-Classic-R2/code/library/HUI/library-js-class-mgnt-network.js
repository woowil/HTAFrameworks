// nOsliw Solutions HUI, HTML Application Library https://github.com/woowil/HTAFrameworks
//**Start Encode**

/*

File:     library-js-class-mgnt-network.js
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
class_mgnt_network.prototype = new _Class("Network","Class for managing network configuration")

function class_mgnt_network(oService,sComputer,sUser,sPass,bInit,bStatus,bDebug){
	try{
		this.wmi_cimv2 = oService
		this.macaddress = null
		this.ipaddress = null
		this.properties = null
		this.computer = typeof(sComputer) == "string" ? sComputer : oWno.ComputerName
		this.username = typeof(sUser) == "string" ? sUser : null
		this.password = typeof(sPass) == "string" ? sPass : null
		this.isStatus = bStatus ? true : false
		this.isDebug = bDebug ? true : false	
		this.name = "class_mgnt_network"
		
		this.setService = function(oService,sComputer,sUser,sPass,bForce){
			if(!this.wmi_cimv2 || bForce){
				if(oService) this.wmi_cimv2 = oService
				else {
					this.computer = typeof(sComputer) == "string" ? sComputer : oWno.ComputerName
					this.username = typeof(sUser) == "string" ? sUser : null
					this.password = typeof(sPass) == "string" ? sPass : null
					var oLoc = new ActiveXObject("WbemScripting.SWbemLocator");
					this.wmi_cimv2 = oLoc.ConnectServer(this.computer,"root\\cimv2",this.username,this.password);
				}
			}
			this.isLocalhost = this.isHostname()
			return this.wmi_cimv2
		}
		
		this.IsIPAddress = function(sIPaddress){
			try{
				var aNum, oRe = new RegExp("([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{1,3})","ig"); // IP-address
				if(typeof(sIPaddress) != "string");
				else if(aNum = sIPaddress.match(oRe)){
					if(sIPaddress.length == aNum[0].length){
						for(var i = 1; i < aNum[0].length; i++){
							if(parseInt(aNum[i]) > 255) return false;
							delete aNum[i]
						}
						delete aNum						
						return true
					}
				}				
			}
			catch(ee){
				this.error(ee,"IsIPAddress()");
			}
			return false
		}
		
		this.getQuery = function(WQL){
			if(!this.isString(WQL) || !this.setService()) return null
			return this.wmi_cimv2.ExecQuery(WQL,null,48)
		}
		
		this.getProperties = function(sClass,oClassItem,o,WQL){
			try{
				if(!this.isString(sClass,WQL)) return null
				o = typeof(o) == "object" ? o : {}
				var oClass = this.wmi_cimv2.Get(sClass);				
				
				for(var oEnum = new Enumerator(oClass.Properties_); !oEnum.atEnd(); oEnum.moveNext()){
					var oItem = oEnum.item();
					var oRe = new RegExp(".+" + new String(oItem.name) + ".+","g")
					if(WQL.match(oRe)){
						var oo = oClassItem[oItem.name], v
						if(oItem.CIMType == 101) oo = this.wmiDate(oo)
						else if(oItem.IsArray) oo = ((v=this.wmiArray(oo,oItem.name,"\n")) ? v.stream : " ").replace(/.+:[ \t]+(.+)/gm,"$1")
						o[oItem.name] = oo
						this.kill(v)
					}
				}
			}
			catch(ee){
				this.error(ee,"getProperties()");
			}
			finally{
				return o
			}
		}
		
		this.getMAC = function(){
			if(this.macaddress) return this.macaddress
			var sFile = this.fso.temp + "\\getMAC.log"
			var sCmd = "%comspec% /c nbtstat -R>nul & nbtstat -a " + this.computer + " | find /i \"MAC ad\" > " + sFile;
			if(oWsh.Run(sCmd,this.fso.hide,true) == 0){
				var oFile = oFso.OpenTextFile(sFile,this.fso.read,false,this.fso.TristateUseDefault)
				var s = oFile.ReadAll() // MAC Address = 00-0D-56-FE-C1-DE
				oFile.close()
				s = s.replace(/\n|\t\r/ig,"")
				if(s.match(/[a-z \t]+=[ ]([a-z0-9\-]+)[ ]*/ig)){
					this.macaddress = (RegExp.$1).replace(/[\-]/ig,":")
				}
			}
			
			if(oFso.FileExists(sFile)){
				(oFso.GetFile(sFile)).Delete(); // deletes also empty files
			}
			return this.macaddress
		}

		this.setWINS = function(sPrimary,sSecondary){
			if(!this.IsIPAddress(sPrimary)) return false
			if(sSecondary && !this.IsIPAddress(sSecondary)) return false
			if(!this.getMAC()) return false
			var WQL = "select MACAddress from Win32_NetworkAdapterConfiguration where IPEnabled='true' and MACAddress='" + this.macaddress + "'"
			for(var oEnum = new Enumerator(this.getQuery(WQL)); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item()
				oItem.SetWINSServer(sPrimary,(sSecondary ? sSecondary : null))
				return true
			}
			return false
		}

		this.setDNS = function(sPrimary,sSecondary){
			if(!this.IsIPAddress(sPrimary)) return false
			if(sSecondary && !this.IsIPAddress(sSecondary)) return false
			if(!this.getMAC()) return false
			var WQL = "select MACAddress from Win32_NetworkAdapterConfiguration where IPEnabled='true' and MACAddress='" + this.macaddress + "'"
			var aVB = this.setVBArray(sPrimary,(sSecondary ? sSecondary : null))
			for(var oEnum = new Enumerator(this.getQuery(WQL)); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item()
				oItem.SetDNSServerSearchOrder(aVB)
				return true
			}
			return false
		}
		
		this.setDNSSearch = function(sSearchList){
			if(!this.isString(sSearchList)) return false
			if(!this.getMAC()) return false
			var WQL = "select MACAddress from Win32_NetworkAdapterConfiguration where IPEnabled='true' and MACAddress='" + this.macaddress + "'"
			var aVB = this.setVBArray(sSearchList.split(/[,;]*/g))
			for(var oEnum = new Enumerator(this.getQuery(WQL)); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item()
				oItem.SetDNSSuffixSearchOrder(aVB)
				return true
			}
			return false
		}
		
		this.getIPConfig = function(){
			var oEnum, oItem
			try{
				if(!this.getMAC()) return null
				var WQL = "select IPAddress,IPSubnet,DefaultIPGateway,DNSDomain,IPEnabled,MACAddress from Win32_NetworkAdapterConfiguration where IPEnabled='true' and MACAddress='" + this.macaddress + "'"
				var o = {}
				o.SystemName = (this.computer).toUpperCase()
				for(var oEnum = new Enumerator(this.getQuery(WQL)); !oEnum.atEnd(); oEnum.moveNext()){
					var oItem = oEnum.item()
					this.getProperties("Win32_NetworkAdapterConfiguration",oItem,o,WQL)
				}
				delete o.MACAddress
				delete o.IPEnabled
				return o
			}
			catch(ee){
				this.error(ee,"getIPConfig()");
			}
			finally{
				this.kill(oEnum,oItem)
			}
			return false
		}

		this.getIPConfigAll = function(){
			var oEnum, oItem
			try{
				if(!this.getMAC()) return null
				// Bug: Must have IPEnabled and MACAddress
				var WQL = "select IPAddress,MACAddress,IPSubnet,IPEnabled,DefaultIPGateway,DNSDomain,DNSDomainSuffixSearchOrder,DNSEnabledForWINSResolution,DNSServerSearchOrder,DomainDNSRegistrationEnabled,Index,WINSEnableLMHostsLookup,WINSPrimaryServer,WINSSecondaryServer,FullDNSRegistrationEnabled,ServiceName,DHCPEnabled,DHCPLeaseExpires,DHCPLeaseObtained from Win32_NetworkAdapterConfiguration where IPEnabled='true' and MACAddress='" + this.macaddress + "'"
				var o = {}
				o.SystemName = (this.computer).toUpperCase()
				for(var oEnum = new Enumerator(this.getQuery(WQL)); !oEnum.atEnd(); oEnum.moveNext()){
					var oItem = oEnum.item()
					this.getProperties("Win32_NetworkAdapterConfiguration",oItem,o,WQL)
				}
				// NetConnectionID
				WQL = "select AdapterType,Index,MACAddress,Manufacturer,ProductName,Speed,TimeOfLastReset from Win32_NetworkAdapter where Index='" + o.Index + "'"
				for(var oEnum = new Enumerator(this.getQuery(WQL)); !oEnum.atEnd(); oEnum.moveNext()){
					var oItem = oEnum.item()
					this.getProperties("Win32_NetworkAdapter",oItem,o,WQL)
				}
				delete o.Name // Bug
				if(o.Speed == null) o.Speed = "undefined"
				return o
			}
			catch(ee){
				this.error(ee,"getIPConfigAll()");
			}
			finally{
				this.kill(oEnum,oItem)
			}
			return false
		}
		
		this.getDNSDomain = function(){
			var oEnum, oItem
			try{
				if(!this.getMAC()) return null
				var WQL = "select DNSDomain,IPEnabled,MACAddress from Win32_NetworkAdapterConfiguration where IPEnabled='true' and MACAddress='" + this.macaddress + "'"
				var o = {}
				o.SystemName = (this.computer).toUpperCase()
				for(var oEnum = new Enumerator(this.getQuery(WQL)); !oEnum.atEnd(); oEnum.moveNext()){
					var oItem = oEnum.item()
					this.getProperties("Win32_NetworkAdapterConfiguration",oItem,o,WQL)
				}
				delete o.MACAddress
				delete o.IPEnabled
				return o
			}
			catch(ee){
				this.error(ee,"getDNSDomain()");
			}
			finally{
				this.kill(oEnum,oItem)
			}
			return false
		}
		
		this.objectToHTML = function(o,bHead){
			var h = ""
			if(bHead) h = h + '<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0">'
			//for(oo in o) h = h + '<th>' + oo + '</th>'
			//h = h + '</tr>'
			for(var oo in o){
				h = h + '\n<tr><td>' + oo + '</td><td>' + o[oo] + '</td></tr>'
			}
			//alert(h)
			if(bHead) h = h + '\n</table>'
			return h
		}

		this.objectToCSV = function(o,bHead){
			var c = "\n"
			if(bHead); //TODO
			for(var oo in o){
				c = c + o[oo] + ';'
			}
			return c.replace(/(.+);$/g,"$1")
		}

		this.reboot = function(){
			for(var oEnum = new Enumerator(this.getQuery("Select * from Win32_OperatingSystem where Primary='true'")); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item()
				oItem.Reboot()
			}
			return true
		}
		
		if(bInit) this.init(oService)
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

