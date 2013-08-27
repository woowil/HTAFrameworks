// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**

__H.include("HUI@Sys.js","HUI@Sys@Mgmt@WMI.js")

__H.register(__H.Sys,"Net","Network",function Net(){
	this.macaddress = null
	this.ipaddress = null
	this.properties = null
	var o_service = null
	
	this.initialize = function initialize(sComputer,sUser,sPass,bForce){
		try{
			if(__HWMI.setServiceWMI(sComputer,"root\\cimv2",sUser,sPass)){
				o_service = __HWMI.getServiceWMI(bForce)
			}
			return o_service
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	var getQuery = function(WQL){
		try{
			if(!o_service) this.initialize()
			if(!__H.isString(WQL) || !o_service) return null
			return o_service.ExecQuery(WQL,null,48)
		}
		catch(ee){
			__HLog.error(ee,this);
			return null
		}
	}
	
	var getProperties = function(sClass,oClassItem,o,WQL){
		try{
			if(!__H.isString(sClass,WQL)) return null
			o = typeof(o) == "object" ? o : {}
			var oClass = o_service.Get(sClass);
			
			for(var oEnum = new Enumerator(oClass.Properties_); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				var oRe = new RegExp(".+" + new String(oItem.name) + ".+","g")
				if(WQL.isSearch(oRe)){
					var oo = oClassItem[oItem.name], v
					if(oItem.CIMType == 101) oo = this.wmiDate(oo)
					else if(oItem.IsArray) oo = ((v=this.wmiArray(oo,oItem.name,"\n")) ? v.stream : " ").replace(/.+:[ \t]+(.+)/gm,"$1")
					o[oItem.name] = oo
					__HUtil.kill(v)
				}
			}
			return o
		}
		catch(ee){
			__HLog.error(ee,this);
			return null
		}
	}
	
	this.getMAC = function getMAC(){
		try{
			if(this.macaddress) return this.macaddress
			var sFile = __HIO.temp + "\\getMAC.log"
			var sCmd = "%comspec% /c nbtstat -R>nul & nbtstat -a " + this.computer + " | find /i \"MAC ad\" > " + sFile;
			if(oWsh.Run(sCmd,__HIO.hide,true) == 0){
				var oFile = oFso.OpenTextFile(sFile,__HIO.read,false,__HIO.TristateUseDefault)
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
		catch(ee){
			__HLog.error(ee,this);
		}
	}

	this.setWINS = function setWINS(sPrimary,sSecondary){
		try{
			if(!__H.isString(sPrimary)) return false
			if(!this.getMAC()) return false
			var WQL = "select MACAddress from Win32_NetworkAdapterConfiguration where IPEnabled='true' and MACAddress='" + this.macaddress + "'"
			for(var oEnum = new Enumerator(getQuery(WQL)); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item()
				oItem.SetWINSServer(sPrimary,(sSecondary ? sSecondary : null))
				return true
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this);
		}
	}

	this.setDNS = function setDNS(sPrimary,sSecondary){
		try{
			if(!__H.isString(sPrimary)) return false
			if(!this.getMAC()) return false
			var WQL = "select MACAddress from Win32_NetworkAdapterConfiguration where IPEnabled='true' and MACAddress='" + this.macaddress + "'"
			var aVB = this.setVBArray(sPrimary,(__H.isString(sSecondary) ? sSecondary : null))
			for(var oEnum = new Enumerator(getQuery(WQL)); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item()
				oItem.SetDNSServerSearchOrder(aVB)
				return true
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this);
		}
	}
	
	this.setDNSSearch = function setDNSSearch(sSearchList){
		try{
			if(!__H.isString(sSearchList)) return false
			if(!this.getMAC()) return false
			var WQL = "select MACAddress from Win32_NetworkAdapterConfiguration where IPEnabled='true' and MACAddress='" + this.macaddress + "'"
			var aVB = this.setVBArray(sSearchList.split(/[,;]*/g))
			for(var oEnum = new Enumerator(getQuery(WQL)); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item()
				oItem.SetDNSSuffixSearchOrder(aVB)
				return true
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this);
		}
	}
	
	this.getIPConfig = function getIPConfig(){
		var oEnum, oItem
		try{
			if(!this.getMAC()) return null
			var WQL = "select IPAddress,IPSubnet,DefaultIPGateway,DNSDomain,IPEnabled,MACAddress from Win32_NetworkAdapterConfiguration where IPEnabled='true' and MACAddress='" + this.macaddress + "'"
			var o = {}
			o.SystemName = (this.computer).toUpperCase()
			for(var oEnum = new Enumerator(getQuery(WQL)); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item()
				getProperties("Win32_NetworkAdapterConfiguration",oItem,o,WQL)
			}
			delete o.MACAddress
			delete o.IPEnabled
			return o
		}
		catch(ee){
			__HLog.error(ee,this);
		}
		return false
	}

	this.getIPConfigAll = function getIPConfigAll(){
		var oEnum, oItem
		try{
			if(!this.getMAC()) return null
			// Bug: Must have IPEnabled and MACAddress
			var WQL = "select IPAddress,MACAddress,IPSubnet,IPEnabled,DefaultIPGateway,DNSDomain,DNSDomainSuffixSearchOrder,DNSEnabledForWINSResolution,DNSServerSearchOrder,DomainDNSRegistrationEnabled,Index,WINSEnableLMHostsLookup,WINSPrimaryServer,WINSSecondaryServer,FullDNSRegistrationEnabled,ServiceName,DHCPEnabled,DHCPLeaseExpires,DHCPLeaseObtained from Win32_NetworkAdapterConfiguration where IPEnabled='true' and MACAddress='" + this.macaddress + "'"
			var o = {}
			o.SystemName = (this.computer).toUpperCase()
			for(var oEnum = new Enumerator(getQuery(WQL)); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item()
				getProperties("Win32_NetworkAdapterConfiguration",oItem,o,WQL)
			}
			// NetConnectionID
			WQL = "select AdapterType,Index,MACAddress,Manufacturer,ProductName,Speed,TimeOfLastReset from Win32_NetworkAdapter where Index='" + o.Index + "'"
			for(var oEnum = new Enumerator(getQuery(WQL)); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item()
				getProperties("Win32_NetworkAdapter",oItem,o,WQL)
			}
			delete o.Name // Bug
			if(o.Speed == null) o.Speed = "undefined"
			return o
		}
		catch(ee){
			__HLog.error(ee,this);
		}
		return false
	}
	
	this.getDNSDomain = function getDNSDomain(){
		var oEnum, oItem
		try{
			if(!this.getMAC()) return null
			var WQL = "select DNSDomain,IPEnabled,MACAddress from Win32_NetworkAdapterConfiguration where IPEnabled='true' and MACAddress='" + this.macaddress + "'"
			var o = {}
			o.SystemName = (this.computer).toUpperCase()
			for(var oEnum = new Enumerator(getQuery(WQL)); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item()
				getProperties("Win32_NetworkAdapterConfiguration",oItem,o,WQL)
			}
			delete o.MACAddress
			delete o.IPEnabled
			return o
		}
		catch(ee){
			__HLog.error(ee,this);
		}
		return false
	}
	
	this.objectToHTML = function objectToHTML(o,bHead){
		var h = ""
		if(bHead) h = h + '<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0">'
		//for(oo in o) h = h + '<th>' + oo + '</th>'
		//h = h + '</tr>'
		for(var oo in o){
			if(o.hasOwnProperty(oo)){
				h = h + '\n<tr><td>' + oo + '</td><td>' + o[oo] + '</td></tr>'
			}
		}
		//alert(h)
		if(bHead) h = h + '\n</table>'
		return h
	}

	this.objectToCSV = function objectToCSV(o,bHead){
		var c = "\n"
		if(bHead); //TODO
		for(var oo in o){
			if(o.hasOwnProperty(oo)){
				c = c + o[oo] + ';'
			}
		}
		return c.replace(/(.+);$/g,"$1")
	}

	this.reboot = function reboot(sComputer,sUser,sPass){
		try{
			if(!this.initialize(sComputer,sUser,sPass)) return false
			var oColItems = o_service.ExecQuery("Select * from Win32_OperatingSystem where Primary='true'")
			for(var oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item()
				oItem.Reboot()
			}
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.repair = function repair(){
		try{
			oWsh.Run("%comspec% /c arp -d & nbtstat -R & nbtstat -RR & ipconfig /flushdns",__HIO.hide,true)
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
})

var __HNet	= new __H.Sys.Net()