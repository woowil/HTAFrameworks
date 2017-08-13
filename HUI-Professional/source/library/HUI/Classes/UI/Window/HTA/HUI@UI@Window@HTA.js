// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

__H.include("HUI@UI@Window.js")

__H.register(__H.UI.Window,"HTA","HTA Management",function HTA(){
	var o_this = this
	var b_initialized = false
	var buf_error = null
	
	this.config = {}
	
	/////////////////////////////////////
	//// DEFAULT

	var o_options = {
		
	}
	
	var initialize = function initialize(){
		if(b_initialized) return;
		
		if(typeof(document) != "object" || document.mimeType != "HTML Application"){
			throw new Error(__HLog.errorCode("error"),"__H.UI.Window.HTA.initialize(): Not an HTML Application")
		}		
		
		b_initialized = true
	}	
	initialize()
	
	/////////////////////////////////////
	////

	this.setOptions = function setOptions(oOptions){
		if(__H.isObject(oOptions)) return Object.extend(o_options,oOptions,true)
		return false
	}

	this.getOptions = function(){
		return o_options
	}
	
	this.onError = function onError(message,url,line){
		if(!buf_error) buf_error = new __H.Common.StringBuffer()
		
		buf_error.empty()
		buf_error.append("\nError\t: " + message,false,true)
		buf_error.append("\nUrl\t: " + url)
		buf_error.append("\nLine\t: " + line)
		buf_error.append("\nFile\t: " + __HLoad.getLastFile()) // MUST BE GLOBAL DEFINED
		buf_error.append("\n\nAbort\t: Closes HTA")
		buf_error.append("\nRetry\t: Reloads HTA")
		buf_error.append("\nIgnore\t: Continues")

		var r = oWsh.Popup(buf_error,60,"HTA Error",32 + 2)
		if(r == 3) setTimeout("window.close()",5)
		else if(r == 4) setTimeout("location.reload()",1000) // MUST HAVE A TIME DELAY
	}
	
	this.security = function security(bSet){
		if(bSet){
			if(oWsh.RegWrite(__HKeys.ie_styles + "\\","")){ //Ignore that script stops on heavy loop
				// http://www.codecomments.com/archive298-2004-6-206767.html
				// http://support.microsoft.com/default.aspx?scid=kb;en-us;Q175500
				oWsh.RegWrite(__HKeys.ie_styles + "\\MaxScriptStatements",4294967295,"REG_DWORD") // 0xffffffff === 4294967295
			}
			// Prevents ADO message like: Safety settings on this computer prohibit accessing a data source on another domain.
			// => IE Security->Custom Level->Internet|Local Intranet: Access Data Sources Across Domains
			// http://www.jsifaq.com/SF/Tips/Tip.aspx?id=5130
			// This must be set, or it generates error 2716
			oWsh.RegWrite(__HKeys.ie_cu + "\\Zones\\1\\1406",0,"REG_DWORD") // Intranet zone
			oWsh.RegWrite(__HKeys.ie_cu + "\\Zones\\3\\1406",0,"REG_DWORD") // Internet zone
		}
		else {
			try{
				oWsh.RegDelete(__HKeys.ie_styles + "\\MaxScriptStatements")
				oWsh.RegWrite(__HKeys.ie_cu + "\\Zones\\1\\1406",1,"REG_DWORD") // Intranet zone
				oWsh.RegWrite(__HKeys.ie_cu + "\\Zones\\3\\1406",3,"REG_DWORD") // Internet zone
			}
			catch(ee){}
		}
	}
	
	this.onScroll = function onScroll(oElement){
		/*
		Equivalent to: object.doScroll( [sScrollAction])
		http://msdn.microsoft.com/en-us/library/ms536414(VS.85).aspx
		*/
		if(typeof(oElement) == "object" && oElement.nodeType == 1 && oElement.scrollHeight > oElement.clientHeight){
			oElement.scrollTop = (oElement.scrollHeight-oElement.clientHeight);
		}
	}
	
	this.position = function position(iWidth,iHeight,oWindow){
		//if(document.body) document.body.style.visibility = "hidden"
		if(!__H.isNumber(iWidth,iHeight)) return;
		var iLeft = (window.screen.width-iWidth)/2;
		var iTop = (window.screen.height-iHeight)/2;
		oWindow = oWindow ? oWindow : window
		oWindow.moveTo(iLeft,iTop);
		//oWindow.resizeTo(1,1);
		oWindow.resizeTo(iWidth,iHeight);
		
		try{
			oWindow.self.focus()
		}
		catch(ee){}
	}
	
	var exit = function(){
		try{
			window.close();
		}
		catch(ee){
			WScript.Quit();
		}
	}
	
	this.closeHTA = function closeHTA(bAll){
		this.security(false)
		if(bAll){
			var c = oFso.GetAbsolutePathName(".")
			var n = c + "\\" + oFso.GetFileName(document.URL)
			if(__HWMI.setServiceWMI(oWno.ComputerName,"root\\cimv2")){ // Excellence!!
				var oService = __HWMI.getServiceWMI()
				var oColItems = oService.ExecQuery("SELECT Name,CommandLine FROM Win32_Process WHERE Name = 'mshta.exe'","WQL",48)
				var oRe1 = new RegExp(c.replace(/\\/g,"\\\\"),"ig")
				var oRe2 = new RegExp(oFso.GetBaseName(n),"ig")
				for(var i = 0, oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext()){
					var oItem = oEnum.item()
					var c = new String(oItem.CommandLine)
					if(!c.isSearch(oRe1)) continue // Skip opened HTA running from a different parent directory
					if(!c.isSearch(oRe2) || i > 0){ // Terminate all HTA in same root directory. Leave 
						try{
							oItem.Terminate()
						}
						catch(ee){}
					}
					else i++ // in case there are duplicate mshta that's needs to be killed
				}
			}
		}		
		exit();	
	}
})


/////////
var __HHTA	= new __H.UI.Window.HTA()