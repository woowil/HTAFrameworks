// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**


__H.include("HUI@Sys@LDAP@ADSI.js")

__H.register(__H.Sys.LDAP.ADSI,"IIS","Internet Information Service",function IIS(){
	/////// IIS ADSI
	// IIsHTMServer: http://feurerpc.oet.udel.edu/iishelp/iis/htm/asp/aore7zn6.htm
	// IIsFilters: http://feurerpc.oet.udel.edu/iishelp/iis/htm/asp/aore4h6b.htm
	// IIsCertMapper: http://feurerpc.oet.udel.edu/iishelp/iis/htm/asp/aore4qia.htm__H
	var service = null
	var domain = oWno.ComputerName
	
	this.getIIS = function getIIS(bForce){
		try{
			if(!service || bForce){
				service = GetObject("IIS://" + domain + "/W3SVC");
			}
			return service
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.createVirtual = function createVirtual(sHTMSite,sVirPath,sVirName){
		try{
			if(!this.getIIS()) return false
			// Check if the HTM site directory exists
			for(var bHTMID = false, oEnum = new Enumerator(service); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				if(oItem.Class == "IIsHTMServer"){
					if(oItem.ServerComment == sHTMSite){
						bHTMID = oItem.Name; // 1..n
						break;
					}
				}
			}
			if(!bHTMID){
				var sMsg = "TERMINATING CreateHTMVirtualDir, Error accessing HTM Site '" + sHTMSite + "'. Server may not exist.";
				__HLog.err_description = sMsg;
				return false;
			}
			
			oHTMSite = service.GetObject("IIsHTMServer",bHTMID);
			// Get the web site's virtual root
			var oVRoot = oHTMSite.GetObject("IIsHTMVirtualDir","Root");
			try{
				// Create the new virtual directory if not exist
				var oVDir = oVRoot.Create("IIsHTMVirtualDir",sVirName);
			}
			catch(e1){}
			// Sets and saves the new virtual directory path
			oVDir.AccessRead = true;
			oVDir.Path = sVirPath;
			oVDir.SetInfo();
			
			return true;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.deleteVirtual = function deleteVirtual(sHTMSite,sVirName){
		try{
			if(!this.getIIS()) return false
			// Check if the HTM site and virtual exists
			for(var bHTMID = false, oEnum = new Enumerator(service); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				if(oItem.Class == "IIsHTMServer"){
					if(oItem.ServerComment == sHTMSite){						
						bHTMID = oItem.Name; // 1..n
						break;
					}
				}
			}
			if(!bHTMID) return false
			oHTMSite = service.GetObject("IIsHTMServer",bHTMID);
			var oVRoot = oHTMSite.GetObject("IIsHTMVirtualDir","Root");
			try{
				oVRoot.Delete("IIsHTMVirtualDir",sVirName);
			}
			catch(e1){}
			return true;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}	
})
