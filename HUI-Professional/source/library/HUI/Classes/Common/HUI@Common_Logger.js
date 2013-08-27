// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**

__H.include("HUI@Common.js")

//(function(){
	
	__H.register(__H.Common,"Logger","Logger",function Logger(){
		
		if (!(this instanceof arguments.callee) ){
			return new Logger();
		}
		
		var o_this = this
		var b_initialized = false
		
		this.log_event    = false
		this.log_registry = false
		this.file_log     = false
		this.file_error   = false
		
		var ERROR_TYPE_ERROR
		var ERROR_TYPE_WARNING
		var ERROR_TYPE_INFORMATON
		var ERROR_TYPE_SUCCESS
		var ERROR_TYPE_FAILURE
		var ERROR_CODE_DEFAULT
		var ERROR_CODE_ERROR
		var ERROR_CODE_WARNING
		var ERROR_CODE_INFORMATON
		var ERROR_CODE_SUCCESS
		var ERROR_CODE_FAILURE
		
		/////////////////////////////////////
		//// DEFAULT
		
		var o_options = {
			err_number		: 0,
			err_description : "",
			err_function	: "",
			s_lastlog       : "",
			i_max_wait      : 15000,
			b_isactive      : false
		}
		
		
		var initialize = function initialize(){
			if(b_initialized) return;
			
			var d = new Date()
			var f = oFso.GetAbsolutePathName(".") + "\\logs\\" + d.formatYYYYMM()
			oWsh.Run("%comspec% /c md \"" + f + "\"",0,false)
			f = f.concat("\\" + d.formatYYYYMMDD() + "-")
			o_this.file_log   = __H.o_options.file_log ? f + __H.o_options.file_log : o_this.file_log
			o_this.file_error = __H.o_options.file_error ? f + __H.o_options.file_error : o_this.file_error
			
			ERROR_TYPE_ERROR      = "error"
			ERROR_TYPE_WARNING    = "warning"
			ERROR_TYPE_INFORMATON = "information"
			ERROR_TYPE_SUCCESS    = "success"
			ERROR_TYPE_FAILURE    = "failure"
			ERROR_CODE_DEFAULT    = 8880
			ERROR_CODE_ERROR      = 8881
			ERROR_CODE_WARNING    = 8882
			ERROR_CODE_INFORMATON = 8884
			ERROR_CODE_SUCCESS    = 8888
			ERROR_CODE_FAILURE    = 8896
			
			b_initialized = true
		}
		
		this.setOptions = function setOptions(oOptions){
			return Object.extend(o_options,oOptions,true)
		}
		
		this.getOptions = function(){
			return o_options
		}
		
		/////////////////////////////////////
		////
		
		this.logstring = function logstring(s){
			switch(typeof(s)){
				case "string"   :
				case "number"   : return s;
				case "function" : s = "[function]"; break;
				case "object"   : s = "[object]"; break;
				case "boolean"  : s = "[boolean]"; break;
				default         : s = ""; break;
			}
			return s
		}
		
		this.popup = function popup(s,iIdleSec,sTitle,nType){
			s = o_this.logstring(s)
			if(typeof(sTitle) != "string") sTitle = "Log Popup";
			iIdleSec = typeof(iIdleSec) == "number" && iIdleSec > 0 ? iIdleSec : 20;
			
			/* nType = Button Types + Icon Types
				Button Types
				0 Show OK button.
				1 Show OK and Cancel buttons.
				2 Show Abort, Retry, and Ignore buttons.
				3 Show Yes, No, and Cancel buttons.
				4 Show Yes and No buttons.
				5 Show Retry and Cancel buttons.
				
				Icon Types
				16 Show "Stop Mark" icon.
				32 Show "Question Mark" icon.
				48 Show "Exclamation Mark" icon.
				64 Show "Information Mark" icon.
			*/
			nType = typeof(nType) == "number" ? nType : 48;
			
			try{ // Fix for HTA applications, since idle ain't working
				sTitle = sTitle + " - #" + (10).random(100)
				var ID = window.setTimeout(function(){__HLog.popupClose(sTitle)},(iIdleSec*1000))
			}
			catch(ee){
				sTitle = sTitle.replace(/(.+)-#$/g,"$1")
			}
			
			/* IntButton returns
				1 OK button
				2 Cancel button
				3 Abort button
				4 Retry button
				5 Ignore button
				6 Yes button
				7 No button
			*/
			return oWsh.Popup(s,iIdleSec,sTitle,nType);
		}
		
		this.popupClose = function popupClose(sTitle){
			if(oWsh.AppActivate(sTitle)){
				oWsh.SendKeys('%{F4}')
			}
		}
		
		this.log = function log(s){
			initialize()
			try{
				s = __H.caller() + ": " + o_this.logstring(s)
				
				o_this.reset()
				
				for(var iWait = 0, iSleep; o_options.b_isactive; ){
					if(o_options.i_max_wait < iWait) break;
					iSleep = (5).random(300), iWait += iSleep;
					if(typeof(__HUtil) == "object") __HUtil.sleep(iSleep);
				}
				o_options.b_isactive = true;
				
				var d = new Date(), a
				d = d.formatDateTime ? d.formatDateTime() : d.toLocaleTimeString()
				if((a = s.split(/[\n\r]/g)) && a.length > 1){
					for(var i = a.length; i; i--){
						o_options.s_lastlog = o_options.s_lastlog.concat("\n[" + d + "] " + a.shift())
					}
				}
				else o_options.s_lastlog = "[" + d + "] ".concat(s)
				a.empty()
				
				o_this.append(o_this.file_log,o_options.s_lastlog);
				if(__H.o_options.textarea_result) o_this.logTag(__H.o_options.textarea_result,o_options.s_lastlog)
				else {
					try{
						o_this.popup(o_options.s_lastlog)
					}
					catch(e){
						WScript.Echo(o_options.s_lastlog);
					}
				}
			}
			catch(ee){
				o_this.error(ee,o_this)
			}
			finally{
				logRegistry()
			}
		}
		
		this.logFile = function logFile(sFile,s){
			s = o_this.logstring(s)
			o_this.append(sFile,s,false,true);
		}
		
		this.logTag = function logTag(oElement,s){
			s = o_this.logstring(s)
			if(__H.isElement(oElement) && oElement.contentEditable){
				if(typeof(oElement.innerText) == "string") oElement.innerText = (oElement.innerText).concat("\r\n" + s)
				else if(typeof(oElement.value) == "string") oElement.value = (oElement.value).concat(s)
				oElement.doScroll("pagedown")
			}
			else alert("no Tag")
		}
		
		this.logInfo = function(s){
			o_this.info(s)
			o_this.log(s)
		}
		
		this.logPopup = function(s,iIdleMilli,sTitle,nType){
			o_this.log(s)
			o_this.popup(s,iIdleMilli,sTitle,nType)
		}
		
		this.logEvent = function(s,iType){
			o_this.event(s,iType)
			o_this.log(s)
		}
		
		this.event = function(s,iType){
			s = o_this.logstring(s)
			if(o_this.log_event){
				/* iType
					0 SUCCESS
					1 ERROR
					2 WARNING
					4 INFORMATION
					8 AUDIT_SUCCESS
					16 AUDIT_FAILURE
				*/
				iType = !isNaN(iType) ? iType : 0
				o_this.logEvent(iType,s)
			}
		}
		
		var logRegistry = function logRegistry(){
			try{
				if(o_this.log_registry){
					oWsh.RegWrite(__H.o_options.path_reg_hkcu + "LastLog",options.s_lastlog)
					oWsh.RegWrite(__H.o_options.path_reg_base + "LastLog",options.s_lastlog)
					oWsh.RegWrite(__H.o_options.path_reg_hkcu + "LastError",options.lasterror)
				}
			}
			catch(e){}
		}
		
		this.info = function(s){
			
		}
		
		this.reset = function(){
			o_options.err_number = o_options.err_description = o_options.err_function = ""
			o_options.b_isactive = false
			o_options.s_lastlog = ""
		}
		
		this.clear = function(){
			o_this.reset()
			if(oFso.FileExists(o_this.file_log)) oFso.DeleteFile(o_this.file_log,true)
			if(oFso.FileExists(o_this.file_error)) oFso.DeleteFile(o_this.file_error,true)
		}
		
		this.error = function error(oErr,oFunc,sDesc){
			initialize()
			if(!__H.isError(oErr)) return;
			try{
				o_this.reset()
				oErr.setCallers(arguments,oFunc)
				o_options.lasterror = oErr.getErrorText(oFunc,sDesc)
				if(__H.$debug) o_this.popup(o_options.lasterror)
				
				o_this.logTag(__H.o_options.textarea_error,o_options.lasterror)
				o_this.append(o_this.file_error,"\n\n" + o_options.lasterror);
			}
			catch(e){
				o_this.popup("__HLog.error(): " + oErr.getErrorText());
			}
			finally{
				logRegistry()
			}
		}
		
		this.errorXML = function errorXML(oXml,oFunc,sDesc){
			try{
				if(!__H.isXML(oXml)) return;
				var oErr = oXml.parseError
				if(oErr.errorCode == 0) return;
				o_this.reset()
				
				var s = new __H.Common.StringBuffer("\nXML Text:      " + (o_options.err_source = (oErr.srcText).trim()))
				s.append("\nXML Code:     " + (o_options.err_errorCode = (oErr.errorCode & 0xFFFF)))
				s.append("\nXML Url:      " + (o_options.err_line = oErr.url))
				s.append("\nXML Line:     " + (o_options.err_line = oErr.line))
				s.append("\nXML Char:     " + (o_options.err_linepos = oErr.linepos))
				s.append("\nXML Position: " + (o_options.err_fileepos = oErr.filepos))
				s.append("\nXML Reason:\t" + (o_options.err_reason = (oErr.reason).trim()))
				
				oErr = new Error(o_this.errorCode("error"),s)
				oErr.setCallers(arguments,oFunc)
				o_options.lasterror = oErr.getErrorText(oFunc,sDesc)
				o_this.logTag(__H.o_options.textarea_error,o_options.lasterror)
				o_this.append(null,o_options.lasterror);
				s.empty()
			}
			catch(e){
				o_this.popup("__HLog.errorXML(): " + e.description);
			}
		}
		
		this.errorCode = function errorCode(sType){
			if(typeof(sType) != "string") return ERROR_CODE_DEFAULT
			var ec
			switch(sType.toLowerCase()){
				case ERROR_TYPE_ERROR      : ec = ERROR_CODE_ERROR; break;
				case ERROR_TYPE_WARNING    : ec = ERROR_CODE_WARNING; break;
				case ERROR_TYPE_INFORMATON : ec = ERROR_CODE_INFORMATON; break;
				case ERROR_TYPE_SUCCESS    : ec = ERROR_CODE_SUCCESS; break;
				case ERROR_TYPE_FAILURE    : ec = ERROR_CODE_FAILURE; break;
				default                    : ec = ERROR_CODE_DEFAULT; break;
			}
			return ec;
		}
		
		this.append = function append(sFile,s,bString){
			initialize()
			try{
				sFile = sFile ? sFile : o_this.file_log
				if(typeof(sFile) != "string") return false
				
				var oFile = oFso.OpenTextFile(sFile,8,true,-2);
				if(!bString){
					var a = s.split(/\r\n|\n/g); // Just remember to separate your stream with a Newline character
					var oRe = /.+/g
					for(var i = 0, iLen = a.length; i < iLen; i++){
						if(!(a[i]).search() > -1) oFile.WriteLine(a[i]); // If not blank line
						else oFile.WriteBlankLines(1)
					}
					a.empty()
				}
				else {
					oFile.WriteLine(s)
				}			
				oFile.Close()
				return true;
			}
			catch(e){
				o_this.popup("__HLog.append(): " + e.description);
				return false;
			}
		}
		
		this.debug = function debug(s){
			if(__H.$debug){
				o_this.log("@ " + s)
			}
		}
		
		this.Exception = {
			ArgumentIllegal : function(s,arg){
				if(typeof(s) != "string") s = ""
				var a = [].addArguments(arg)
				
				throw new Error(8881,"ArgumentIllegal: " + s + " " + a.toString())
			},
			
			ArgumentNull : function(m){
				
			},
			
			ArgumentOutOfRange : function(s){
				if(typeof(s) != "string") s = ""
				throw new Error(8881,"ArgumentOutOfRange: " + s)
			},
			
			ArgumentUndefined : function(m){
				
			},
			
			ArgumentType : function(m){
				
			},
			
			InvalidOperation : function(m){
				
			},
			
			ScriptLoader : function(s){
				if(typeof(s) != "string") s = ""
				throw new Error(8881,"ArgumentOutOfRange: " + s)
			},
			
			TimeOutOfLimit : function(s){
				if(typeof(s) != "string") s = ""
				throw new Error(8881,"TimeOutOfLimit: " + s)
			},
			
			IOFileNotFound : function(s){
				if(typeof(s) != "string") s = ""
				throw new Error(8881,"IOFileNotFound: " + s)
			},
			
			XMLLoadingError : function(s){
				if(typeof(s) != "string") s = ""
				throw new Error(8881,"XMLLoadingError: " + s)
			},
			
			XMLNotInitialized : function(s){
				if(typeof(s) != "string") s = ""
				throw new Error(8881,"XMLNotInitialized: " + s)
			},
			
			SYSRemoteAccessDenied: function(s){
				if(typeof(s) != "string") s = ""
				throw new Error(8881,"SYSRemoteAccessDenied: " + s)
			}
		}
		
		__H.log = this.log;
		__H.popup = this.popup;
		__H.error = this.error;
		__H.debug = this.debug;
		__H.errorCode = this.errorCode;
	
	})
//}())


var __HLog = new __H.Common.Logger()
var __HExp = __HLog.Exception
