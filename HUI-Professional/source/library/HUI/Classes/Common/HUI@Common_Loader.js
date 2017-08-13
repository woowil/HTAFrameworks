// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

// http://www.javascriptkit.com/javatutors/loadjavascriptcss2.shtml

__H.include("HUI@Common.js","HUI@Common_Logger.js")

//(function(){
	
	__H.register(__H.Common,"Loader","Script/Style File Loader",function Loader(){
		
		if (!(this instanceof arguments.callee) ){
			return new Loader();
		}
		
		var o_this = this
		var b_initialized = false
		
		o_this.d_load_scripts = null
		o_this.d_load_styles  = null
		
		var oReIgnoreFolder, sReIgnoreFiles, oReExt, oReExtJS, oReExtCSS, oReExtVBS
		
		/////////////////////////////////////
		//// DEFAULT
		
		var o_cur_load
		var o_options = {}
		
		var initialize = function initialize(){
			if(b_initialized) {return;}
			
			o_this.d_load_scripts	 = __H.d_lib_includes.Count > 0 ? __H.d_lib_includes : new ActiveXObject("Scripting.Dictionary")
			o_this.d_load_styles	 = new ActiveXObject("Scripting.Dictionary")
			
			sReIgnoreFiles	= "template example depicted"
			oReIgnoreFolder	= /template|example|depicted/ig // Private
			oReExt			= /.+\.(?:js|jse|vbs|vbe)$/ig
			oReExtJS		= /.+\.(?:js|jse)$/ig
			oReExtCSS		= /.+\.(?:css)$/ig
			oReExtVBS		= /.+\.(?:vbs|vbe)$/ig
			
			o_cur_load = {
				s_file      : "[unknown]",
				o_script    : null,
				b_loading   : false
			}
			
			b_initialized = true
		}
		initialize()
		
		this.setOptions = function setOptions(oOptions){
			if(__H.isObject(oOptions)) {
				return Object.extend(o_options,oOptions,true)
			}
			return false
		}
		
		this.getOptions = function(){
			return o_options
		}
		
		this.isLoading = function(){
			return o_cur_load.b_loading
		}
		
		this.getLastFile = function(){
			return o_cur_load.s_file
		}
		
		this.loadScript = function loadScript(sFile,bNotAppend){
			try{ // This function dynamically loads an external script file
				if(typeof(sFile) != "string") {
					return false;
				}
				o_cur_load.s_file = sFile.trim().replace(/\\|\\\\/g,"/") // must do this!!
				__H.s_load_file = o_cur_load.s_file
				
				if(o_this.d_load_scripts.Exists(o_cur_load.s_file)) {
					return true
				}
				else if(!oFso.FileExists(o_cur_load.s_file)) {
					throw new Error(__H.errorCode("error"),"Unable to load/locate script file: " + sFile)
				}
				else if(!o_cur_load.s_file.isSearch(oReExt)) {
					
					return false;
				}
				o_cur_load.o_script = __H.byClone("script") // Clone is faster than create
				o_cur_load.o_script.id = __H.o_options.s_prefix_script + document.scripts.length
				if(o_cur_load.s_file.isSearch(oReExtJS)){
					o_cur_load.o_script.language = "JScript.Encode"
					o_cur_load.o_script.type = "text/javascript" // "application/x-javascript"
				}
				else if(o_cur_load.s_file.isSearch(oReExtVBS)){
					o_cur_load.o_script.language = "VBScript.Encode"
					o_cur_load.o_script.type = "text/vbscript"
				}
				o_cur_load.onLoad = scriptLoaded;
				o_cur_load.o_script.charset = "UTF-8" //"ISO-8859-1"
				o_cur_load.o_script.src = o_cur_load.s_file;
				
				// DON*T CHANGE! MUST BE BEFORE appendChild..
				o_this.d_load_scripts.Add(o_cur_load.s_file,o_cur_load.o_script.id)
				o_cur_load.b_loading = true			
				if(!bNotAppend) __H.o_dom_head.appendChild(o_cur_load.o_script)
				o_cur_load.b_loading = false
				
				return true
			}
			catch(ee){
				__H.popup("__HLoad.loadScript(): " + ee.description)
				__H.error(ee,this)
				return false
			}
		}
		
		var scriptLoaded = function scriptLoaded(){
			o_cur_load.b_loading = false
			return false
		}
		
		this.loadScripts = function loadScripts(){
			try{
				for(var i = 0, iLen = arguments.length; i < iLen; i++){
					this.loadScript(arguments[i])
				}
				return true
			}
			catch(ee){
				__HLog.error(ee,this)
			}
			return false
		}
		
		this.loadScriptFolder = function loadScriptFolder(sFolder){
			try{
				if(typeof(sFolder) != "string") return false
				
				var sFile = oFso.GetSpecialFolder(2) + "\\" + oFso.GetTempName()
				var bReturn = true
				
				var f = oFso.GetAbsolutePathName(".") + "\\" + sFolder
				var sCmd = "%comspec% /c dir \"" + f + "\\*.js*\" \"" + f + "\*.vb*\" /b /ON /A-D /s | findstr /i /v \"" + sReIgnoreFiles + "\" >" + sFile
				if(oWsh.Run(sCmd,0,true) != 0){
					__HLog.logPopup("WARNING - No valid JS script files were found recursively in folder: " + sFolder)
					return
				}
				
				var oFile = oFso.OpenTextFile(sFile,1,false,-2)
				var a = oFile.ReadAll().split(/\r\n|\n/g)
				oFile.close()
				var s = oFso.GetAbsolutePathName(".\\")
				var bb = []
				for(var i = a.length; i; i--){ // DON'T CHANGE THE FOR LOOP ORDER
					var ss = (a.shift()).trim()
					if(!ss.isSearch(oReIgnoreFolder)){ // if not on ignore list
						if((ss = "." + ss.replace(s,"").replace(/\\/g,"/")) == ".") continue
						if(!this.loadScript(ss)) bReturn = false;					
					}
				}
				return  bReturn
			}
			catch(ee){
				__HLog.popup("__HLoad.loadScriptFolder(): " + ee.description)
				if(ee.description == "Input past end of file"){
					__HLog.debug("No Script Files Found in Folder " + sFolder)
				}
				else __HLog.error(ee,this)
				return false
			}
			finally{			
				o_cur_load.s_file = ""
				if(typeof(__HFile) != "undefined") __HFile.remove(sFile)
			}
		}
		
		this.loadFolders = function loadFolders(){
			var bLoaded = true
			for(var i = arguments.length-1; i >= 0; i--){
				if(!this.loadScriptFolder(arguments[i])){
					bLoaded = false
				}
			}
			return bLoaded
		}
		
		this.removeScript = function removeScript(sScriptID){
			try{
				if(__H.isStringEmpty(sScriptID)) return false;
				
				var o
				if(o = __H.byIds(sScriptID)){
					if(o_this.d_load_scripts.Exists(o.src)) o_this.d_load_scripts.Remove(o.src)
					//__H.byTag('head').removeChild(o);
					o_cur_load.o_head.removeChild(o);
					return true
				}
				return false
			}
			catch(e){
				__HLog.error(e,this,"removeScript()-1")
			}
			try{
				document.removeChild(o);
				return true
			}
			catch(e){
				__HLog.error(e,this,"removeScript()-2")
			}
			try{
				o.removeNode(true);
				return true
			}
			catch(e){
				__HLog.error(e,this,"removeScript()-1")
			}
			return false
		}
		
		this.replaceScript = function replaceScript(sRegExp,sNewFile){
			try{
				if(this.loadScript(sNewFile,true)){
					var a = this.getLoadedScripts(sRegExp)
					if(a.length){
						a[0].obj.parentNode.replaceChild(o_cur_load.o_script,a[0].obj)
						return true
					}				
				}
			}
			catch(ee){
				__HLog.error(ee,this)
			}
			return false
		}
		
		this.removeScripts = function removeScripts(sRegExp){
			try{
				var a = (this.getLoadedScripts(sRegExp)).reverse()
				for(var i = 0, iLen = a.length; i < iLen; i++){
					this.removeScript(a[i].id)
				}
				return true
			}
			catch(ee){
				__HLog.error(ee,this)
			}
			return false
		}
		
		this.getLoadedScripts = function getLoadedScripts(sRegExp){
			try{
				var oElements = __H.byTag('SCRIPT'), a = []
				var oRe = typeof(sRegExp) == "string" ? new RegExp(sRegExp,"ig") : oReExt
				for(var i = 0, iLen = oElements.length; i < iLen; i++){
					if(oElements[i].src && (new String(oElements[i].src)).isSearch(oRe)){
						a.push({
							obj   : oElements[i],
							src   : oElements[i].src,
							index : i,
							id    : oElements[i].id
						})
					}
				}			
			}
			catch(ee){
				__HLog.error(ee,this)
			}
			return a
		}
		
		this.reloadScripts = function reloadScripts(sRegExp){
			try{
				var oElements = __H.byTag('script')
				var oRe = typeof(sRegExp) == "string" ? new RegExp(sRegExp,"ig") : oReExt
				var bLoad = false
				//for(var j = 0, iLen = arguments.length; j < iLen; j++){
				var bFound = false
				for(var i = 0, s, iLen = oElements.length; i < iLen; i++){
					if(oElements[i].src && (new String(oElements[i].src)).isSearch(oRe)){
						oElements[i].id = !oElements[i].id ? "oScrRemove" : oElements[i].id
						var s = oElements[i].src
						if(this.removeScript(oElements[i].id)){
							i++
							if(this.loadScript(s)){
								--i, bLoad = true
							}
							bFound = true
						}
					}
				}
				if(!bFound){
					if(oFso.FileExists(sRegExp)){
						return this.loadScript(sRegExp)
					}
				}
			//}
				return bLoad
			}
			catch(ee){
				__HLog.error(ee,this)
				return false
			}
		}
		
		// STYLESHEET FUNCTIONS
		this.loadStyle = function loadStyle(sFile){
			try{
				if(__H.isStringEmpty(sFile)) return false;
				o_cur_load.s_file = sFile.trim().replace(/\\|\\\\/g,"/")
				if(o_this.d_load_styles.Exists(o_cur_load.s_file)) return true
				var len = document.styleSheets(0).imports.length
				document.styleSheets(0).addImport(o_cur_load.s_file,len)
				this.d_load_styles.Add(o_cur_load.s_file,len+"")
				return true
			}
			catch(ee){
				__HLog.error(ee,this)
				return false
			}
		}
		
		this.loadStyles = function loadStyles(){
			try{
				for(var i = 0, iLen = arguments.length; i < iLen; i++){
					this.loadScript(a[i])
				}
				return true
			}
			catch(ee){
				__HLog.error(ee,this)
			}
			return false
		}
		
		this.loadStyleFolder = function loadStyleFolder(sFolder){
			if(__H.isStringEmpty(sFolder)) return false
			try{			
				var sFile = oFso.GetSpecialFolder(2) + "\\" + oFso.GetTempName()
				var bReturn = true
				var sCmd = "%comspec% /c dir \"" + __H.s_path_hta + "\\" + sFolder + "\\*.css\" /b /ON /A-D /s | findstr /i /v \"" + sReIgnoreFiles + "\" | sort>" + sFile
				if(oWsh.Run(sCmd,0,true) != 0){
					__HLog.logPopup("WARNING - No valid CSS script files were found recursively in folder:" + sFolder)
					return false
				}
				var oFile = oFso.OpenTextFile(sFile,1,false,-2)
				var a = (oFile.ReadAll()).split(/\r\n|\n/g)
				oFile.close()			
				
				for(var i = a.length-1; i >= 0; i--){
					a[i] = ((a[i]).replace(__H.s_path_hta + "\\","")).replace(/\\/g,"/")
					if(!this.loadStyle(a[i])) bReturn = false
				}
				a.empty()
				return bReturn
			}
			catch(ee){
				if(ee.description == "Input past end of file"){
					__HLog.debug("No CCS files Found in folder " + sFolder)
				}
				else __HLog.error(ee,this)
				return false
			}
			finally{
				if(typeof(__HFile) != "undefined") __HFile.remove(sFile)
			}
		}
		
		this.getLoadedStyles = function getLoadedStyles(sRegExp){
			try{
				var a = []
				var oRe = typeof(sRegExp) == "string" ? new RegExp(sRegExp,"ig") : false
				var oStyles = document.styleSheets
				for(var i = 0, iLen = oStyles.length; i < iLen; i++){
					if(oStyles(i).owningElement.tagName == "STYLE"){
						for(var j = 0, iLen2 = oStyles(i).imports.length; j < iLen2; j++){
							var sFile = oStyles(i).imports(j).href
							if(oRe && sFile.isSearch(oRe)){
								a.push(sFile)
								continue
							}
							a.push(sFile)
						}
					}
				}
			}
			catch(ee){
				__HLog.error(ee,this)
			}
			return a
		}
		
		this.reloadStyle = function reloadStyle(sFile){
			try{
				var oRe1 = typeof(sFile) == "string" ? new RegExp(sFile,"ig") : new RegExp(".+\.css$","ig")
				var oStyles = document.styleSheets
				for(var i = 0, iLen1 = oStyles.length; i < iLen1; i++){
					if(oStyles(i).owningElement.tagName == "STYLE"){
						var iLen2 = oStyles(i).imports.length
						for(var j = 0; j < iLen2; j++){
							sFile = oStyles(i).imports(j).href
							if(sFile.isSearch(oRe1)){
								oStyles(i).addImport(sFile,j)
								//document.styleSheets(i).imports(j).href = obj.css_href  // Won't work!!
								if(!o_this.d_load_styles.Exists(sFile)){
									o_this.d_load_styles.Add(sFile,iLen2+"")
								}
								return true
							}
						}
					}
				}
				return false
			}
			catch(ee){
				__HLog.error(ee,this)
			}
		}
		
		this.removeStyle = function removeStyle(sFile){
			try{
				if(typeof(sFile) != "string") return false
				var oRe = new RegExp(sFile,"ig")
				var oStyles = document.styleSheets
				for(var i = 0, iLen = oStyles.length; i < iLen; i++){
					if(oStyles(i).owningElement.tagName == "STYLE"){
						for(var j = 0, iLen2 = oStyles(i).imports.length; j < iLen2; j++){
							sFile = oStyles(i).imports(j).href
							if(sFile.isSearch(oRe)){
								oStyles(i).removeImport(j)
								if(o_this.d_load_styles.Exists(sFile)){
									o_this.d_load_styles.Remove(sFile)
								}
								return true
							}
						}
					}
				}
				
				return false
			}
			catch(ee){
				__HLog.error(ee,this)
				return false
			}
		}
		
		__H.loadScript = this.loadScript;
	})
	
//}())

var __HLoad  = new __H.Common.Loader()
