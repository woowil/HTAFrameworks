// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**


__H.include("HUI@IO@File@ACL.js")

__H.register(__H.IO.File.ACL,"Share","Share",function Share(){
	var TXT_COMPUTER_UNACCESSABLE = "Unable to access computer"
	var Network = null
	var computer = null
	this.lansrv = null

	// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/win32_securitydescriptor.asp__H

	var initialize = function initialize(sComputer){
		try{
			if(computer != sComputer || this.lansrv == null){
				if(!this.ping(sComputer)) throw TXT_COMPUTER_UNACCESSABLE
				this.lansrv = GetObject("WinNT://" + sComputer + "/LanmanServer")
				computer = sComputer
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.setShare = function setShare(sComputer,sShareName,sLocSharePath,sDescription,sTrustee,sPerm){
		try{
			if(!initialize(sComputer)) return false
			var oShare = this.lansrv.Create("FileShare",sShareName)
			oShare.Path = sLocSharePath
			oShare.MaxUserCount = -1 // Unlimited: Note that workstations has max 10
			oShare.Description = sDescription ? sDescription : ""
			oShare.SetInfo();
			sTrustee = sTrustee ? sTrustee : "EVERYONE"
			sPerm = sPerm ? sPerm : oAcl.SHARE_RIGHT_FULL
			var sUNCShare = "\\\\" + sComputer + "\\" + sShareName

			return this.setACEShare(null,sUNCShare,sTrustee,sPerm)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getShareObject = function getShareObject(sComputer,sShareName){
		try{
			if(!initialize(sComputer)) return false
			var oShare = false
			for(var oEnum = new Enumerator(this.lansrv); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				if((oItem.Name).isSearch(sShareName,"i")){
					oShare = {
						Name			: oItem.Name,
						MaxUserCount	: oItem.MaxUserCount,
						AdsPath			: oItem.AdsPath,
						Path			: oItem.Path,
						Class			: oItem.Class,
						CurrentUserCount : oItem.CurrentUserCount,
						Description		: oItem.description,
						GUID			: oItem.GUID,
						HostComputer	: oItem.HostComputer,
						Parent			: oItem.Parent,
						Schema			: oItem.Schema
					}
					break;
				}
			}
			return oShare
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getShareString = function getShareString(sComputer,sShareName){
		try{
			if(oo = this.getShareObject(sComputer,sShareName)){
				var s = ""
				for(var o in oo){
					if(oo.hasOwnProperty(o)){
						s = s + "\n" + o + " = " + oo[o]
					}
				}
				return s
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.delShare = function delShare(sComputer,sShareName){
		try{
			if(!initialize(sComputer)) return false
			for(var oEnum = new Enumerator(oLan); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item()
				if(oItem != null && (oItem.Name).isSearch(sShareName,"i")){
					this.lansrv.Delete("fileshare",sShareName);
					break;
				}
			}
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
})

__H.register(__H.IO.File.ACL,"FileACL","file/folder ACL security using FileACL.exe",function FileACL(){

	this.setCHANGE = function setCHANGE(sFolder,sTrustee,bNotThread){
		try{
			bNotThread = !!bNotThread
			return (oWsh.Run("%comspec% /c fileacl.exe \"" + sFolder + "\" /S " + sTrustee + ":RWXD | find /i \"error\">nul",__HIO.hide,bNotThread) != 0)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.setFULL = function setFULL(sFolder,sTrustee,bNotThread){
		try{
			bNotThread = bNotThread ? true : false
			return (oWsh.Run("%comspec% /c fileacl.exe \"" + sFolder + "\" /S " + sTrustee + ":F | find /i \"error\">nul",__HIO.hide,bNotThread) != 0)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
})
