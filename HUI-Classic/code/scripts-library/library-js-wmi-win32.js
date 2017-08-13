// nOsliw Solutions - HTML Aplication Framework.
//**Start Encode**

// GLOBAL/EXTERNAL DECLARATIONS
var SUPPORT_WMI_WIN32 = true;
var SUPPORT_WMI_WIN32_VER = "1.0";

///////////


//////////////////


function wmi_win32_networkconnection(oService,sComputer){
	try{
		oService = oService ? oService : wmi_wbem_service();
		var sHtml, sHtmlAll = "";
		var oColItems = oService.ExecQuery("Select * from Win32_NetworkConnection",null,48);

		for(var oItem, i = 1, oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext(), i++){
			oItem = oEnum.item(), sHtml = "\n\n [INSTANCE-" + i + "]";
			sHtml += "\n  AccessMask: " + oItem.AccessMask;
			sHtml += "\n  Caption: " + oItem.Caption;
			sHtml += "\n  Comment: " + oItem.Comment;
			sHtml += "\n  ConnectionState: " + oItem.ConnectionState;
			sHtml += "\n  ConnectionType: " + oItem.ConnectionType;
			sHtml += "\n  Description: " + oItem.Description;
			sHtml += "\n  DisplayType: " + oItem.DisplayType;
			sHtml += "\n  InstallDate: " + ((d=wmi_class_date(oItem.InstallDate)) ? d : "undefined");
			sHtml += "\n  LocalName: " + oItem.LocalName;
			sHtml += "\n  Name: " + oItem.Name;
			sHtml += "\n  Persistent: " + oItem.Persistent;
			sHtml += "\n  ProviderName: " + oItem.ProviderName;
			sHtml += "\n  RemoteName: " + oItem.RemoteName;
			sHtml += "\n  RemotePath: " + oItem.RemotePath;
			sHtml += "\n  ResourceType: " + oItem.ResourceType;
			sHtml += "\n  Status: " + oItem.Status;
			sHtml += "\n  UserName: " + oItem.UserName;
			sHtmlAll += sHtml;
		}
		delete oColItems;
	}
	catch(e){
		wmi_win32_err_show(2,e);
		return false;
	}
	finally{
		return sHtmlAll;
	}
}

function wmi_win32_share(oService,sComputer){
	try{
		oService = oService ? oService : wmi_wbem_service();
		var sHtml, sHtmlAll = "";
		var oColItems = oService.ExecQuery("Select * from Win32_Share where Status='OK'",null,48);

		for(var oItem, i = 1, oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext(), i++){
			oItem = oEnum.item(), sHtml = "\n\n [INSTANCE-" + i + "]";
			sHtml += "\n  AccessMask: " + oItem.AccessMask;
			sHtml += "\n  AllowMaximum: " + oItem.AllowMaximum;
			sHtml += "\n  Caption: " + oItem.Caption;
			sHtml += "\n  Description: " + oItem.Description;
			sHtml += "\n  InstallDate: " + ((d=wmi_class_date(oItem.InstallDate)) ? d : "undefined");
			sHtml += "\n  MaximumAllowed: " + oItem.MaximumAllowed;
			sHtml += "\n  Name: " + oItem.Name;
			sHtml += "\n  Path: " + oItem.Path;
			sHtml += "\n  Status: " + oItem.Status;
			sHtml += "\n  Type: " + oItem.Type;
			sHtmlAll += sHtml;
		}
		delete oColItems;
	}
	catch(e){
		wmi_win32_err_show(2,e);
		return false;
	}
	finally{
		return sHtmlAll;
	}
}


////////// ERROR

function wmi_win32_err_show(iOpt,oErr){
	try{
		js_err_show(iOpt,oErr,false,false,true);
	}
	catch(e){
		try{
			WScript.Echo(oErr.description);
		}
		catch(e){
			alert(oErr.description);
		}
	}
}
