// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

__H.include("HUI@Sys@Net.js")

__H.register(__H.Sys.Net,"Lanman","Network Lanman files and sessions",function Lanman(){
	this.lanmanserver = null
	this.fileservice = null
	this.shares = null
	var o_service = null
	var computer = oWno.ComputerName
	
	this.initialize = function initialize(sComputer,sUser,sPass,bForce){
		try{
			if(__HWMI.setServiceWMI(sComputer,"root\\cimv2",sUser,sPass)){
				o_service = __HWMI.setServiceWMI(bForce)
			}
			this.setLanmanServer(bForce)
			return o_service
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	// SHARES
	
	this.setLanmanServer = function setLanmanServer(bForce){
		if(!this.lanmanserver || bForce){
			try{
				this.lanmanserver = GetObject("WinNT://" + computer + "/LanmanServer,FileService")
			}
			catch(ee){}
		}
		__HLog.debug("LanManServer: typeof-" + typeof this.lanmanserver)
		return this.lanmanserver
	}
	
	this.setFileService = function setFileService(){
		if(!this.fileservice){
			try{
				this.fileservice = GetObject("WinNT://" + computer + "/LanmanServer,FileService")
			}
			catch(ee){}
		}
		return this.fileservice
	}
	
	this.setShares = function setShares(sComputer){
		if(!this.shares){ // WinNT://DOMAIN/SERVER/SHARE
			this.shares = GetObject("WinNT://" + (sComputer ? sComputer : computer) + "/LanmanServer/SHARE")
		}
		return this.shares
	}
	
	this.getShares = function getShares(sComputer,bExtended,bShow){
		var aShares = []
		if(!this.setShares(sComputer)) return aShares
		for(var oEnum = new Enumerator(this.shares); !oEnum.atEnd(); oEnum.moveNext()){
			var oItem = oEnum.item();
			try{
				var o = {}
				o.Name = oItem.Name
				__HLog.log(oItem.Name + " " +oItem.Path)
				o.Class = oItem.Class
				o.ADsPath = oItem.ADsPath
				o.Computer = oItem.HostComputer
				o.Path = oItem.Path
				if(bExtended){
					o.MaxUserCount = oItem.MaxUserCount
					o.Class = oItem.Class
					o.CurrentUserCount = oItem.CurrentUserCount
					o.description = oItem.description
					o.GUID = oItem.GUID
					o.HostComputer = oItem.HostComputer
					o.Parent = oItem.Parent
					o.Schema = oItem.Schema
				}
				o.sort()
				if(bShow){
					__HLog.log("#")
					for(var oo in o){
						if(o.hasOwnProperty(oo)){
							__HLog.log("## " + o + " = " + o[oo])
						}
					}
				}
				else aShares.push(o)
			}
			catch(ee){}
		}
		return aShares
	}
	
	this.delShare = function delShare(sShareName){
		try{
			if(!this.setFileService()) return false
			this.fileservice.Delete("FileShare",sShareName)
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.setShare = function setShare(sName,sLocalPath,sDescription){
		try{
			if(!__H.isString(sShareName,sSharePath)) return false
			if(!this.setFileService()) return false
			var o = this.fileservice.Create("FileShare",sName)
			o.Path = sLocalPath
			o.MaxUserCount = -1 // Unlimited connections
			o.SetInfo()
			// TODO: share rights
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.setShareWMI = function setShareWMI(sName,sLocalPath,sDescription){
		try{
			if(!__H.isString(sName,sLocalPath,sDescription)) return false
			if(!o_service && !this.initialize()) return false
			var FILE_SHARE = 0
			var MAXIMUM_CONNECTIONS = null //255
			var oShare = o_service.Get("Win32_Share") 
			var iReturn = oShare.Create(sLocalPath,sName,FILE_SHARE,MAXIMUM_CONNECTIONS,sDescription)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	// SESSIONS
	
	this.getSessions = function getSessions(bShow){
		try{
			if(!this.setLanmanServer()) return []
			var a = [], i = 0, s
			for(var oEnum = new Enumerator(this.lanmanserver.Sessions()); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				a.push({
					Computer    : oItem.Computer,
					Name        : oItem.Name,
					ConnectTime : oItem.ConnectTime,
					User        : oItem.User,
					IdleTime    : oItem.IdleTime
				})
				if(bShow){
					s = "\nComputer: " + oItem.Computer
					s = s.concat("\nName: " + oItem.Name)
					s = s.concat("\nConnectTime: " + oItem.ConnectTime)
					s = s.concat("\nUser: " + oItem.User)
					s = s.concat("\nIdleTime: " + oItem.IdleTime)
					__HLog.log(s)
				}
			}
			return a;
		}
		catch(ee){
			__HLog.error(ee,this)
			return a
		}
	}
	
	this.delSessions = function delSessions(){
		try{
			if(!this.setLanmanServer()) return false
			var oRe = new RegExp(new String(oWno.UserName),"ig")
			for(var oEnum = new Enumerator(this.lanmanserver.Sessions()); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				try{					
					if(!(new String(oItem.User)).isSearch(oRe)){
						oSessions.Remove(oItem.Name)
					}
				}
				catch(ee){}
			}
			return true;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	// OPEN FILES
	
	this.getFiles = function getFiles(bShow){
		if(!this.setLanmanServer()) return false
		var a = [], s
		for(var oEnum = new Enumerator(this.lanmanserver.Resources()); !oEnum.atEnd(); oEnum.moveNext()){
			try{
				var oItem = oEnum.item();
				if((oItem.Path).isSearch(/PIPE/ig)) continue
				if(bShow){ // ignore pipes
					s = "\nPath: " + oItem.Path
					s = s.concat("\nUser: " + oItem.User)
					s = s.concat("\nName: " + oItem.Name)
					__HLog.log(s)
				}
				a.push({
					Path : oItem.Path,
					User : oItem.User,
					Name : oItem.Name
				})
			}
			catch(ee){}
		}
		return a;
	}
	
	this.delFiles = function delFiles(){
		if(!this.setLanmanServer()) return false
		var oRe = new RegExp(new String(oWno.UserName),"ig")
		for(var oEnum = new Enumerator(this.lanmanserver.Resources()); !oEnum.atEnd(); oEnum.moveNext()){				
			try{
				var oItem = oEnum.item();
				if((oItem.Path).isSearch(/PIPE/ig)) continue
				if(!(new String(oItem.User)).isSearch(oRe)){
					oResources.Remove(oItem.Name)
				}
			}
			catch(ee){}
		}
		return true;
	}
})

var __HLanman = new __H.Sys.Net.Lanman()