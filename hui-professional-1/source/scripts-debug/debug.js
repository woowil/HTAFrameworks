// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**


//__H.include("HUI@UI@Window.HTA.js")
//__H.include("HUI@UI@Window@Widgets_Panel.js","HUI@UI@Window@Elements.js","HUI@UI@Window_Style.js","HUI@IO@File_INI.js","HUI@IO@File@XML.js")

__H.register(__H.UI.Window.HTA,"Debug","nOsliw Debug",function Debug(){
			
	if (!(this instanceof arguments.callee) ){
		return new Debug();
	}
	
	/////////////////////////////////////
	//// CLASS DEFAULT
	
	var o_this = this
	var b_initialized = false
	
	var o_options = {
		screen : null
	}
	
	var initialize = function initialize(bForce){
		if(b_initialized && !bForce) return;
		
		o_options.screen = {
			width	: oHUIDebug.width,
			height	: window.screen.height*0.9,
			width2	: window.screen.width-50,
			height2	: window.screen.height*0.9,
			width3	: 750,
			height3	: 330,
			xpos	: window.screenLeft,
			ypos	: window.screenTop
		}
		
		b_initialized = true
	}
	initialize()
	
	this.setOptions = function setOptions(oOptions){		
		try{
			Object.extend(oOptions,true)
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getOptions = function(){
		return o_options
	}
	
	/////////////////////////////////////
	//// HTA DEFAULT
	
	this.onload = function onload(){
		try{		
			oFormDebug.reset()
			
			__HXML.getXMLAJAX("https://github.com/woowil/HTAFrameworks",this.ajaxUrl)
			var oElements = __H.byTag("INPUT");
			for(var i = oElements.length-1; i >= 0; i--){
				oElement = oElements[i]
				var c = oElement.className, n = oElement.name
				if(oElement.type == "password"){
					oElement.onpaste = new Function("oWsh.Popup('Paste events are prohibited in password fields',5,__HMBA.title_sub,48);return false;")
				}
				else if(oElement.type == "button"){
					oElement.unselectable = "on"
					if(c.isSearch(/debug-input-bb|debug-input-bg|debug-input-br/ig)){ // Other action buttons
						oElement.title = (c.isSearch(/debug-input-bb/ig) && !n.isSearch(/help|action/ig)) ? oElement.title : oElement.value
						oElement.value = ":: " + oElement.value
					}
					if(oElement.onclick == null) oElement.onclick = new Function("__HDebug.formactivate(this);")
					if(!c) oElement.className = "debug-input-bg"
				}
				else if(oElement.type == "checkbox") oElement.className = "debug-input-checkbox"
				else if(oElement.type == "radio") oElement.className = "debug-input-radio"
			}
			var oElements = __H.byTag("TEXTAREA");
			for(var i = oElements.length-1; i >= 0; i--){
				oElement = oElements[i]
				var c = oElement.className, n = oElement.name
				if(n.match(/readme|release|license/ig)){
					n = n.replace("t_","file_")
					oElement.innerText = __HFile.readall(__H.o_options[n])
				}
			}
			
			__HElem.show(oDivDebug,oDivBody,oDivDebugPanel)
			__HHTA.position(o_options.screen.width2,o_options.screen.height2)
			
		}
		catch(e){
			__HLog.error(e,this);
			return false;
		}
	}
	
	this.ajaxUrl = function ajaxUrl(response){
		//oFrameHUI.src = 
	
	}
	
	this.formactivate = function formactivate(oElement){
		try{
			var oForm = oElement.form;
			var n = oElement.name, a, t
			if(oForm.name == "oFormDebug"){
				if(n == "navigatehui"){
					var elem = __H.byId("oFrameHUI")
					oElement.removeNode()
					elem.setAttribute("src",elem.getAttribute("nosrc"))
					__HElem.show(elem)
					//elem.location.reload()
				}
				else if(n == "dbg_evaluate"){
					eval(oForm.myeval.value)
				}
				else if(n == "dbg_functions"){
					//oForm.t_log.innerText = __HDebug.getJSFunctions()
					__H.log(this.getJSFunctions())
					oDivBody.doScroll('down')
				}
				else if(n == "dbg_encode"){
					//oForm.t_log.innerText += __HDebug.encode(oForm.encode.value)
					var s = oForm.encode.value
					oForm.encode.value = s.encode()
				}
				else if(n == "dbg_script_reload1"){
					if(__HLoad.reloadScripts(oForm.dbg_script_text.value)) __HLog.popup('loaded!')
				}
				else if(n == "dbg_script_reload2"){
					if(oForm.dbg_script_list.options.length > 0){
						if(__HLoad.reloadScripts(oForm.dbg_script_list.value)) __HLog.popup('loaded!')
					}
				}
				else if((n) == "dbg_script_update"){
					if(a = __HLoad.getLoadedScripts()){
						__HSelect.setSelect(oForm.dbg_script_list)
						__HSelect.addArrayObject(a,0,"id","src");
						a.empty()
						__HLog.popup('updated!')
					}
				}
				else if(n == "dbg_style_reload1"){
					if(__HLoad.reloadStyle(oForm.dbg_style_text.value)) __HLog.popup('loaded!')
				}
				else if(n == "dbg_style_reload2"){
					if(oForm.dbg_style_list.options.length > 0){
						if(__HLoad.reloadStyle(oForm.dbg_style_list.value)) __HLog.popup('loaded!')
					}
				}
				else if((n) == "dbg_style_update"){
					if(a = __HLoad.getLoadedStyles()){
						__HSelect.setSelect(oForm.dbg_style_list)
						__HSelect.addArray(a,0);
						a.empty()
						__HLog.popup('updated!')
					}
				}
			}
		
		}
		catch(e){
			__HLog.error(e,this);
			return false;
		}
		finally{
			document.recalc(true)
		}
	}
	
	this.testHUI = function testHUI(){
		__H.initialize()
		__H.popup("hello")
		__H.register(__H.UI,"Web2","Web/HTML",function Web2(){
			this.sayHello = function sayHello(){
				__H.popup("__H.UI.Web2() said hello")
			}
		})
		__H.register(__H.UI.Web2,"Toolbar","Toolbar",function Toolbar(){
			this.sayHi = function sayHi(){
				__H.popup("__H.UI.Web2.Toolbar() Hello ")
				//this.sayHello()
			}
		})
		new __H.UI.Web2().sayHello()
		new __H.UI.Web2.Toolbar().sayHi()
		new __H.UI.Web2.Toolbar().sayHello()
	}
	
	this.getJSFunctions = function getJSFunctions(){
		try{
			var sFiles, sLine, oFile
			//var sFolder = prompt("Enter script folder.",oFso.GetAbsolutePathName(".\\"));
			var sFolder = __HShell.browseFolder(oFso.GetAbsolutePathName(".\\"),"Choose folder path for recursive scan")
			if(!sFolder) return;
			if(sFiles = __HFile.listFiles(sFolder,"js",true)){
				var oRe1 = new RegExp("function[ \t]+([a-z0-9_]+)[ \t]*\(([a-z0-9_, ]*)\)","ig")
				var oRe2 = new RegExp("this\.([a-z0-9_]+)[ \t]*=[ \t]*function[ \t]*\(([a-z0-9_, ]*)\)","ig")
				var oRe3 = new RegExp("function[ \t]*\(([a-z0-9_, ]*)\)","ig")
				var s = new __H.Common.StringBuffer("Count\tJavaScript Files\n------------------------------------------")
				
				__H.log("Getting function count in folder: " + sFolder)
				for(var j = 0, k, i = sFiles.length-1; i >= 0; i--){
					oFile = oFso.OpenTextFile(sFiles[i],__HIO.read,false,__HIO.TristateUseDefault);
					k = 0
					while(!oFile.AtEndOfStream){
						if((sLine = oFile.ReadLine()).isSearch(oRe1) || sLine.isSearch(oRe2) || sLine.isSearch(oRe3)) k++
					}
					oFile.Close();
					s.append("\n" + k + "\t" + sFiles[i])
					j += k 
				}				
				s.append("\n\nTotal\n------------------------------------------\n")
				s.append(j + "\t" + (sFiles.length-1))
				
				__HLog.popup("Found " + sFiles.length + " functions in folder structure " + sFolder)
				
				return s
			}
			return ""
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.encode = function(s){
		if(typeof(s) == "number") s = s.toString();
		else if(typeof(s) != "string") return "";
		return s.encode()
	}
	
	this.getInnerHTML = function getInnerHTML(obj){
		var sFile = oFso.GetSpecialFolder(2) + "\\debug_inner.html";
		var oFile = oFso.CreateTextFile(sFile,true);
		var a = (obj.innerHTML).split("\n");
		for(var i = 0, iLen = a.length; i < iLen; i++){
			oFile.WriteLine(a[i]);
			delete a[i]
		}
		oFile.close()
		__HShell.open(sFile)
	}
	
	this.getDOM = function getDOM(oObj){ // THIS FUNCTION IS USED FOR VIEWING COM OBJECTS
		try{
			if(typeof(oObj) == "string") oObj = __H.$[oObj];
			if(!__H.isElement(oObj)){
				__HLog.popup("Not an HTML DOM object")
				return;
			}
			var sFile = oFso.GetSpecialFolder(2) + "\\debug.html";
			var oFout = oFso.CreateTextFile(sFile,true);
			var sHead = "<html><head>\n<style>td{font-family:Arial;font-size:11px;padding:1px;}</style>";
			sHead = sHead + '<script language="javascript" src="debug2.js"></script>\n</head><body>';
			var sFoot = "\n</table></body></html>";
			var sBody = '\n<table bgcolor="#EEEEEE" border="1" cellspacing="1" cellpadding="2" align="center" bordercolor="#000000">';
			var oRe = new RegExp("[object]","ig")
			
			for(var o in oObj){
				//if(!oObj.hasOwnProperty(o)) continue
				sBody = sBody.concat("\n<tr><td> " + o + " </td>")
				if(oObj[o] == null || oObj[o] == "") sBody = sBody.concat("<td>&nbsp;</td></tr>")
				else if(typeof(oObj[o]) == "object"){
					var s = '<a href="#" nclick="mba_debug(' + oObj[o].nodeName + ')">' + oObj[o] + '</a>';
					sBody = sBody.concat("<td>" + s + " </td></tr>")
				}
				else sBody = sBody.concat("<td>" + oObj[o] + " </td></tr>")
			}
			var a = (sHead + sBody + sFoot).split("\n");
			for(var i = 0, iLen = a.length; i < iLen; i++){
				oFout.WriteLine(a[i]);
				delete a[i]
			}
			oFout.Close();
			
			if(typeof(__HShell) == "undefined") __HShell = new ActiveXObject("Shell.Application")
			__HShell.open(sFile)
			return true;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.getEvent = function(e){
		e = e ? e : window.event
		var result = '';
		try{
			result += "altKey = " + e.altKey + "<br/>";
			result += "button = " + e.button + "<br/>";
			result += "keyCode = " + e.keyCode + "<br/>";
			result += "clientX = " + e.clientX + "<br/>";
			result += "clientY = " + e.clientY + "<br/>";
			result += "ctrlKey = " + e.ctrlKey + "<br/>";
			result += "offsetX = " + e.offsetX + "<br/>";
			result += "offsetY = " + e.offsetY + "<br/>";
			result += "screenX = " + e.screenX + "<br/>";
			result += "screenY = " + e.screenY + "<br/>";
			result += "shiftKey = " + e.shiftKey + "<br/>";
			result += "target.id = " + e.target.id + "<br/>";
			result += "type = " + e.type + "<br/>";
			
			return result
		}
		catch(ee) {
			__HLog.error(ee,this)
		}
	}
	
	this.testPanel = function testPanel(){
		
	
	}
})

var __HDebug = new __H.UI.Window.HTA.Debug()			



