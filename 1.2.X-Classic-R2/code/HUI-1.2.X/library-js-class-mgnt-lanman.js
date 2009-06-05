// nOsliw Solutions HUI, HTML Application Library http://hui.nosliw.info, License http://hui.codeplex.com/license
//**Start Encode**

/*

File:     library-js-class-mgnt-storage.js
Purpose:  Development script
Author:   Woody Wilson
Created:  2006-01
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
class_mgnt_lanman.prototype = new _Class("Lanman","Class for managing lanman files and processes")

function class_mgnt_lanman(oService,sComputer,sUser,sPass,bInit,bStatus,bDebug){
	try{
		this.wmi_cimv2 = oService
		this.regprov = null
		this.computer = typeof(sComputer) == "string" ? sComputer : oWno.ComputerName
		this.username = typeof(sUser) == "string" ? sUser : null
		this.password = typeof(sPass) == "string" ? sPass : null
		this.isStatus = bStatus ? true : false
		this.isDebug = bDebug ? true : false
		this.lanmanserver = null
		this.fileservice = null
		this.shares = null
		
		this.setService = function(oService,sComputer,sUser,sPass,bForce){
			if(!this.wmi_cimv2 || bForce){
				if(oService) this.wmi_cimv2 = oService
				else {
					this.computer = typeof(sComputer) == "string" ? sComputer : oWno.ComputerName
					this.username = typeof(sUser) == "string" ? sUser : null
					this.password = typeof(sPass) == "string" ? sPass : null
					var oLoc = new ActiveXObject("WbemScripting.SWbemLocator");
					this.wmi_cimv2 = oLoc.ConnectServer(this.computer,"root\\cimv2",this.username,this.password);
					this.lanmanItems = false
				}				
			}
			this.setLanmanServer(bForce)
			this.isLocalhost = this.isHostname()
			return this.wmi_cimv2
		}
		
		// FILE AND PRINT
		
		this.enableFileAndPrint = function(){
			try{
				this.lanmanItems = this.wmi_cimv2.ExecQuery("Select * from Win32_Service where Name = 'lanmanserver'","WQL",48)
				for(var oEnum = new Enumerator(this.lanmanItems); !oEnum.atEnd(); oEnum.moveNext()){
					var oItem = oEnum.item();
					var oStartup = this.wmi_cimv2.Get("Win32_ProcessStartup")
					var oConfig = oStartup.SpawnInstance_()
					oConfig.ShowWindow = 0

					var oProcess = this.wmi_cimv2.Get("Win32_Process")
					if(oProcess.Create("%comspec% /c echo %time%>C:\\Temp\\EnableLanManServer.flg", null, oConfig) != 0){
						js_log_print("log_result","## Process could not be created.")
					}
					if(!oItem.Started) oItem.StartService()
				}
			}
			catch(ee){
				this.error(ee,"enableFileAndPrint()")
				return false
			}
		}
		
		this.disableFileAndPrint = function(){
			try{
				this.lanmanItems = this.wmi_cimv2.ExecQuery("Select * from Win32_Service where Name = 'lanmanserver'","WQL",48)
				for(var oEnum = new Enumerator(this.lanmanItems); !oEnum.atEnd(); oEnum.moveNext()){
					var oItem = oEnum.item();
					var oColFiles = this.wmi_cimv2.ExecQuery("ASSOCIATORS OF {Win32_Directory.Name='C:\\Temp'} where ResultClass=CIM_DataFile")
					for(var oEnum2 = new Enumerator(oColFiles); !oEnum2.atEnd(); oEnum2.moveNext()){
						var oItem2 = oEnum2.item();							
						if((oItem2.Extension).toLowerCase() == "flg" && (oItem2.FileName).toLowerCase() == "enablelanmanserver"){
							oItem2.Delete()							
							/*
							oItem2.Drive
							oItem2.Path
							oItem2.FileName
							oItem2.Rename(strNewName)
							oItem2.Compress()
							oItem2.Copy()
							*/
						}
					}
					
					if(oItem.Started){
						var oColServiceList = this.wmi_cimv2.ExecQuery("Associators of {Win32_Service.Name='" + oItem.Name + "'} Where " 
							+ "AssocClass=Win32_DependentService Role=Antecedent")
						for(var oEnum3 = new Enumerator(oColServiceList); !oEnum3.atEnd(); oEnum3.moveNext()){
							var oItem3 = oEnum3.item();
							oItem3.StopService()
						}
						oItem.StopService()
					}
				}
			}
			catch(ee){
				this.error(ee,"disableFileAndPrint()")
				return false
			}
		}
		
		// SHARES
		
		this.setLanmanServer = function(bForce){
			if(!this.lanmanserver || bForce){
				try{
					this.lanmanserver = GetObject("WinNT://" + this.computer + "/LanmanServer,FileService")
				}
				catch(ee){}
			}
			if(this.isDebug) this.echo("LanManServer: typeof-" + typeof this.lanmanserver)
			return this.lanmanserver
		}
		
		this.setFileService = function(){
			if(!this.fileservice){
				try{
					this.fileservice = GetObject("WinNT://" + this.computer + "/LanmanServer,FileService")
				}
				catch(ee){}
			}
			return this.fileservice
		}
		
		this.setShares = function(){			
			if(!this.shares){ // WinNT://DOMAIN/SERVER/SHARE
				this.shares = GetObject("WinNT://" + this.computer + "/LanmanServer/SHARE")
			}
			return this.shares
		}
		
		this.getShares = function(bExtended,bShow){
			if(!this.setShares()) return false
			var aShares = new Array()
			for(var oEnum = new Enumerator(this.shares); !oEnum.atEnd(); oEnum.moveNext()){
				try{
					var oItem = oEnum.item();
					var o = new Object()
					o.Name = oItem.Name
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
						this.echo("#")
						for(var oo in o) this.echo("## " + o + " = " + o[oo])
					}
					else aShares.push(o)
				}
				catch(ee){}
			}
			this.kill(oEnum)
			return aShares
		}
		
		this.delShare = function(sShareName){
			try{
				if(!this.setFileService()) return false
				this.fileservice.Delete("FileShare",sShareName)
				return true
			}
			catch(ee){
				this.error(ee,"delShare()")
				return false
			}
		}
		
		this.setShare = function (sName,sLocalPath,sDescription){
			try{
				if(!this.isString(sShareName,sSharePath)) return false
				if(!this.setFileService()) return false
				var o = this.fileservice.Create("FileShare",sName)
				o.Path = sLocalPath
				o.MaxUserCount = -1 // Unlimited connections
				o.SetInfo()
				// TODO: share rights
				return true
			}
			catch(ee){
				this.error(ee,"setShare()")
				return false
			}
		}
		
		this.setShareWMI = function(sName,sLocalPath,sDescription){
			try{
				if(!this.isString(sName,sLocalPath,sDescription)) return false
				if(!this.setService()) return false
				var FILE_SHARE = 0
				var MAXIMUM_CONNECTIONS = null //255
				var oShare = this.wmi_cimv2.Get("Win32_Share") 
				var iReturn = oShare.Create(sLocalPath,sName,FILE_SHARE,MAXIMUM_CONNECTIONS,sDescription)
			}
			catch(ee){
				this.error(ee,"setShareWMI()")
				return false
			}
		}
		
		// SESSIONS
		
		this.getSessions = function(bShow){
			try{
				if(!this.setLanmanServer()) return false
				var aSessions = new Array()
				for(var oEnum = new Enumerator(this.lanmanserver.Sessions()); !oEnum.atEnd(); oEnum.moveNext()){
					try{
						var oItem = oEnum.item();
						var o = new Object()
						o.Computer = oItem.Computer
						o.Name = oItem.Name
						o.ConnectTime = oItem.ConnectTime
						o.User = oItem.User
						o.IdleTime = oItem.IdleTime
						if(bShow){
							var s = "\nComputer: " + oItem.Computer
							s = s + "\nName: " + oItem.Name
							s = s + "\nConnectTime: " + oItem.ConnectTime
							s = s + "\nUser: " + oItem.User
							s = s + "\nIdleTime: " + oItem.IdleTime
							this.echo(s)
						}
						aSessions.push(o)
						this.kill(o)
					}
					catch(eee){}
				}
				return aSessions;
			}
			catch(ee){
				this.error(ee,"getSessions()")
				return false
			}
		}
		
		this.delSessions = function(){
			try{
				if(!this.setLanmanServer()) return false
				var oRe = new RegExp(new String(oWno.UserName),"ig")
				for(var oEnum = new Enumerator(this.lanmanserver.Sessions()); !oEnum.atEnd(); oEnum.moveNext()){
					try{
						var oItem = oEnum.item();
						if(!(new String(oItem.User)).match(oRe)){
							oSessions.Remove(oItem.Name)
						}
					}
					catch(ee){}
				}
				return true;
			}
			catch(ee){
				this.error(ee,"delSessions()")
				return false
			}
		}
		
		// OPEN FILES
		
		this.getFiles = function(bShow){
			if(!this.setLanmanServer()) return false
			var aResources = new Array()
			for(var oEnum = new Enumerator(this.lanmanserver.Resources()); !oEnum.atEnd(); oEnum.moveNext()){
				try{
					var oItem = oEnum.item();
					if((oItem.Path).match(/PIPE/ig)) continue
					var o = new Object()
					o.Path = oItem.Path
					o.User = oItem.User
					o.Name = oItem.Name					
					if(bShow){ // ignore pipes
						var s = "\nPath: " + oItem.Path
						s = s + "\nUser: " + oItem.User
						s = s + "\nName: " + oItem.Name
						this.echo(s)
					}
					aResources.push(o)
					this.kill(o)
				}
				catch(ee){}
			}
			return aResources;
		}
		
		this.delFiles = function(){
			if(!this.setLanmanServer()) return false
			var oRe = new RegExp(new String(oWno.UserName),"ig")
			for(var oEnum = new Enumerator(this.lanmanserver.Resources()); !oEnum.atEnd(); oEnum.moveNext()){				
				try{
					var oItem = oEnum.item();
					if((oItem.Path).match(/PIPE/ig)) continue
					if(!(new String(oItem.User)).match(oRe)){
						oResources.Remove(oItem.Name)
					}
				}
				catch(ee){}
			}
			return true;
		}
		
		if(bInit) this.init(oService)
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}
