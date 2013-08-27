// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**

__H.include("HUI@UI.js")

__H.register(__H.UI,"Automation","Windows automation",function Automation(sTitle){
	var title = document.title
	var delay = 200
	var timeout = 30000
	var txt_no_title = "Title is not specified"
		
	this.setTitle = function setTitle(sTitle){
		try{
			if(typeof(sTitle) != "string" || sTitle.length == 0) throw txt_no_title;
			title = sTitle
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}	
	
	this.getTitle = function getTitle(){
		return title
	}
	
	if(sTitle) this.setTitle(sTitle)
	
	this.sendKeys = function sendKeys(sKey,iDelay){
		try{
			iDelay = typeof(iDelay) == "number" ? iDelay : delay
			__HUtil.sleep(iDelay);
			oWsh.SendKeys(sKey);
			return true;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.appActivate = function appActivate(iDelay,iTime){
		try{
			iTime = isNaN(iTime) ? timeout : iTime;
			iDelay = isNaN(iDelay) ? delay : iDelay;
			for(var idle = 0; !oWsh.AppActivate(this.getTitle()); idle += iDelay){
				if(idle > iTime){
					__HExp.TimeOutOfLimit("Unable to activate window '" + this.getTitle() + "' after " + idle + " milliseconds")
				}
				__HUtil.sleep(iDelay);
			}
			oWsh.AppActivate(this.getTitle());
			return true;
		}
		catch(ee){alert(ee.description + " " +dd)
			__HLog.error(ee,this)
			return false;
		}
	}

	this.appClose = function appClose(iDelay,iTime){
		try{
			iTime = isNaN(iTime) ? timeout : iTime;
			iDelay = isNaN(iDelay) ? delay : iDelay;
			__HUtil.sleep(iDelay);
			for(var idle = 0; !oWsh.AppActivate(this.getTitle()); idle += iDelay){
				if(idle > iTime){
					break;
				}
				__HUtil.sleep(iDelay)
			}
			if(oWsh.AppActivate(this.getTitle())){
				this.sendKeys("%{F4}");
				if(oWsh.AppActivate(this.getTitle())){ // DOS windows
					oWsh.SendKeys("% "); // ALT SPACE
					if(oWsh.AppActivate(this.getTitle())){ // DOS windows running tail
						oWsh.SendKeys("^C"); // CTRL-C
					}
				}
			}
			return true;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.minimize = function minimize(sTitle){
		try{
			this.appActivate()
			oWsh.SendKeys("% ")
			oWsh.SendKeys("n")
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	
	this.maximize = function maximize(){
		try{
			this.appActivate()
			oWsh.SendKeys("% ")
			oWsh.SendKeys("x")
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.refresh = function refresh(){
		try{			
			window.blur()
			if(oWsh.AppActivate(document.title)){
				oWsh.SendKeys("% ")
				oWsh.SendKeys("n")
				window.blur()
				__HUtil.sleep(50);
				if(oWsh.AppActivate(document.title)){
					oWsh.SendKeys("% ")
					oWsh.SendKeys("r")					
				}
			}
			window.focus()
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.openRegistry = function openRegistry(sComputer){
		try{
			sComputer = sComputer ? sComputer : oWno.ComputerName
			oWsh.Run("%comspec% /c regedit.exe",__HIO.hide,false);
			this.appActivate("Registry Editor",1200);
			this.sendKeys("%F");
			this.sendKeys("C");
			this.sendKeys(sComputer);
			this.sendKeys("{ENTER}")
			return true;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.logCommand = function logCommand(bLog,sMessage,noreturn,sLogFile){
		try{
			sLogFile = sLogFile ? sLogFile : __HIO.temp + "\\" + oFso.GetTempName()
			var sTitle = "Status log"
			var sCopycon = "copy con " + sLogFile
			if(bLog){
				if(!oWsh.AppActivate(sTitle)){
					oWsh.Run("%comspec% /k title " + sTitle,__HIO.show,false);
					oWsh.AppActivate(sTitle);
					oWsh.SendKeys("~@echo off~");
					oWsh.SendKeys("color 87~");
					oWsh.SendKeys("mode con cols=75 lines=40~cls~");
					oWsh.SendKeys( + "~");
					oWsh.SendKeys("~");
					oWsh.SendKeys(sMessage + "~");
				}
				else {
					oWsh.SendKeys(sMessage);
					if(!noreturn) oWsh.SendKeys("~");
				}
			}
			else {
				if(oWsh.AppActivate(sTitle + " - " + sCopycon)){
					oWsh.SendKeys("^z~exit~")
				}
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

})

var __HAuto = new __H.UI.Automation()
