// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

var __H = {
	__name			: "HUI",
	__alias			: "__H",
	__description	: "nOsliw Solutions HUI - HTA/HTML Library Framework",
	__type			: "namespace",
	__version		: "2.5.0",
	
	$  : document.getElementById,
	$C : document.createElement,
	$N : document.getElementsByName,
	$T : document.getElementsByTagName,
	$S : document.styleSheets,
	
	o_dom_head      : document.getElementsByTagName("HEAD")[0],
	
	$date           : new Date(),
	$resource		: oWno.ComputerName,
	$debug			: false,
	$stopit			: false,
	$load_exception : false,
	$config         : {}, // DO NOT CHANGE!! Used by HTA/MBA Application
	
	s_path_hta       : oFso.GetAbsolutePathName("."),
	s_pth_absolute   : "",
	b_lib_loaded	 : false,
	b_lib_loaded_def : false,
	a_lib_common     : ["HUI@Common_Logger.js","HUI@Common_Loader.js","HUI@Common_StringBuffer.js","HUI@UI@Window@HTA.js","HUI@Lang.js"],
	d_lib_includes   : new ActiveXObject("Scripting.Dictionary"),
	d_byclones       : new ActiveXObject("Scripting.Dictionary"),
	d_registered     : new ActiveXObject("Scripting.Dictionary"),
	
	o_options		 : {
		path_library    : "source/library/HUI",
		path_binary     : "source/library/HUI/Binary",
		path_classes    : "source/library/HUI/Classes",
		path_styles     : "source/library/HUI/Styles",
		path_language   : "source/library/HUI/Language",
		path_plugins    : "source/library/HUI/Plugins",
		path_scripts_js : "source/scripts",
		file_readme     : "source/library/HUI/Readme.txt",
		file_releases   : "source/library/HUI/Release.txt",
		file_license    : "source/library/HUI/License.txt",
		lib_separator   : "@",
		lib_sep_regex   : /@/ig,
		include_all     : false,
		include_plugins : false,
		include_css     : true,
		textarea_result : null,
		textarea_error  : null,
		tbody_result	: null,
		tbody_error		: null,
		file_log        : false,
		file_error      : false,
		b_onerror       : true,
		path_reg_base   : "HKCU\\Software\\nOsliw Solutions\\Framework\\HUI Library\\",
		path_reg_hkcu   : "HKCU\\Software\\nOsliw Solutions\\Framework\\HUI Library\\",
		path_reg_hklm   : "HKLM\\Software\\nOsliw Solutions\\Framework\\HUI Library\\",
		s_prefix_script : "JSLoad-"
	},
	
	toString : function(){
		return this.__description + ", v" + this.__version
	},
	
	valueOf : function(){
		// DO NOT REMOVE OR CHANGE !!
		// 1. valueOf is used by call and apply built-in functions
		// 2. valueOf is used when concatenating using the "+, +=" operators
		// 3. valueOf is used when comparing using the "<, <=, >=, ==" operators
		return this.toString()
	},
	
	initialize : function initialize(oOptions){
		try{
			if(this.b_lib_loaded || this.b_lib_loading) {return true};
			this.b_lib_loading = true
			
			if(typeof(oOptions) === "object"){
				for(var property in oOptions){
					if(!oOptions.hasOwnProperty(property)) continue
					if(property in this.o_options){
						this.o_options[property] = oOptions[property]
					}
				}
			}
			this.setOnError(this.o_options.b_onerror)
			
			this.o_options.path_reg_hkcu += this.__version + "\\"
			this.o_options.path_reg_hklm += this.__version + "\\"
			oWsh.RegWrite(this.o_options.path_reg_hkcu,"")
			oWsh.RegWrite(this.o_options.path_reg_hkcu + "LoadDate",this.$date.toLocaleDateString() + " " + this.$date.toLocaleTimeString())
			oWsh.RegWrite(this.o_options.path_reg_hkcu + "LoadUser",oWno.UserDomain + "\\" + oWno.UserName)
			oWsh.RegWrite(this.o_options.path_reg_base + "LoadStatus","Library Loading..")
			
			this.debug("Loading common library Classes")
			this.o_options.path_library = this.o_options.path_library.replace(/^[.\/\\]*|[.\/\\]*$/g,"")
			this.o_options.path_classes = this.o_options.path_classes.replace(/^[.\/\\]*|[.\/\\]*$/g,"")			
			this.o_options.path_styles = this.o_options.path_styles.replace(/^[.\/\\]*|[.\/\\]*$/g,"")
			this.s_pth_absolute = (this.s_path_hta + "/" + this.o_options.path_library).replace(/\\/g,"/")
			
			this.include.apply(this,this.a_lib_common)
			this.b_lib_loaded = true
			
			if(this.o_options.include_plugins){
				this.debug("Loading Plugins")
				this.o_options.path_plugins = this.o_options.path_plugins.replace(/^[.\/\\]*|[.\/\\]*$/g,"")
				__HLoad.loadScriptFolder(this.o_options.path_plugins)
			}
			
			this.b_lib_loaded_def = true
			
			if(this.o_options.include_all){
				this.includeAll()
			}
			
			if(this.o_options.include_css){
				this.debug("Loading library Styles")
				__HLoad.loadStyleFolder(this.o_options.path_styles)
			}
			return true
		}
		catch(ee){
			this.reload(this.__alias + ".initialize(): " + ee.description)
			return false
		}
		finally{
			this.b_lib_loading = false
		}
	},
	
	register : function register(parent,name,description,func,c,v){
		try{
			var sep = this.o_options.lib_separator
			if(this.$load_exception) return false;
			if(this.isStringEmpty(name) || !this.isFunction(func)){
				throw new Error(this.errorCode("error"),"Invalid Arguments [1]: " + name + " or [3]: " + func.getName())
			}
			else if(!this.isFunction(parent) && !this.isObject(parent)){
				throw new Error(this.errorCode("error"),"Argument [0] must be a function or object for: " + name)
			}
			func.__name = parent.__name + sep + name
			
			var ore = new RegExp(sep,"ig")
			var name2 = this.__alias + "." + func.__name.replace(this.__name + sep,"").replace(ore,".")
			if(this.d_registered.Exists(name2)){
				return true
			}
			if(parent.__name != this.__name){
				var sFile = parent.__name.toString()
				sFile = this.o_options.path_classes.concat("/" + sFile.replace(ore,"/") + "/" + sFile + ".js")
				if(oFso.FileExists(sFile)){
					this.include(parent.__name + ".js")
				}
			}
			
			func.__description = typeof(description) == "string" ? description : "[unknown description]"
			func.__type = "Class"
			
			this.log("# Registering Class: " + name2)
			this.d_registered.Add(name2,func.__name)
			parent[name] = func
			return true
		}
		catch(ee){
			//this.$load_exception = true
			this.reload(this.__name + ".register(): " + ee.description)
			return false
		}
	},
	
	include : function include(){
		try{
			if(this.$load_exception) return false;
			
			var dFilesToLoad = new ActiveXObject("Scripting.Dictionary")
			var aFiles = [], i, j, sBaseName, iLen, sFileName, aBaseName, sFileNameInner, sExtension
			for(i = 0, iLen = arguments.length; i < iLen; i++){
				aFiles.push(arguments[i])
			}
			
			for(i = iLen; i; i--){
				sFileName = aFiles.shift()
				if(!this.isClass(sFileName)){
					throw new Error(this.errorCode("error"),"Invalid " + this.__name + " Library type or file name syntax for: '" + sFileName + "'")
				}
				sBaseName = RegExp.$1, sExtension = RegExp.$2
				aBaseName = sBaseName.split(/@|_|\./g) // Ex: HUI@Common_Loader => ["HUI","Common","Loader"]
				if(sBaseName.isSearch(/_/g)) aBaseName.pop() // Ex: HUI@Common_Loader => HUI@Common => ["HUI","Common"]
				
				sFileName = this.o_options.path_classes.concat("/" + aBaseName.join("/") + "/" + sFileName).replace(/Classes\/HUI/ig,"Classes")
				if((this.d_lib_includes).Exists(sFileName)) continue
				
				for(j = 0, sSubFolder = "", iLen = aBaseName.length; j < iLen; j++){
					sSubFolder = sSubFolder.concat(aBaseName[j] + "/")
					sFileNameInner = this.o_options.path_classes.concat("/" + sSubFolder + aBaseName.slice(0,j+1).join(this.o_options.lib_separator) + "." + sExtension)
					sFileNameInner = sFileNameInner.replace(/Classes\/HUI/ig,"Classes") // Must do this because of folder structure
					if(sFileNameInner.match(/HUI\.js/ig)) continue // Must do this because of folder structure
					if(!(this.d_lib_includes.Exists(sFileNameInner)) && !dFilesToLoad.Exists(sFileNameInner)){
						dFilesToLoad.Add(sFileNameInner,"")
					}
				}
				if(!dFilesToLoad.Exists(sFileName)) dFilesToLoad.Add(sFileName,"")
			}
			
			if(dFilesToLoad.Count > 0){
				var sFileToLoad
				var aFilesToLoad = new VBArray(dFilesToLoad.Keys()).toArray()
				dFilesToLoad.RemoveAll()
				for(var i = aFilesToLoad.length; i; i--){
					sFileToLoad = aFilesToLoad.shift()
					
					if(!this.loadScript(sFileToLoad)){
						this.popup("Unable to " + this.__alias + ".include(): " + sFileToLoad)
					}
				}
			}
			
			return true
		}
		catch(ee){
			//this.$load_exception = true
			this.reload(this.__alias + ".include(): " + ee.description)
			return false
		}
	},
	
	includeAll : function includeAll(){
		try{
			this.debug("Loading all library classes")
			var sFile = oFso.GetSpecialFolder(2) + "\\" + oFso.GetTempName()
			var sFolder = oFso.GetAbsolutePathName(this.o_options.path_classes);
			var sIgnore = " | find /i /v \"debug\" | find /i /v \"example\" | find /i /v \"depicted\" "
				
			var sCmd = "%comspec% /u /q /s /c dir \"" + sFolder + "\\*.js*\" /b /OGN /A-D /s " + sIgnore + ">" + sFile
			if(oWsh.Run(sCmd,0,true) != 0){
				throw new Error(this.errorCode("error"),"Library files not found in: " + sFolder)
			}
			//oWsh.Run("%comspec% /c notepad " + sFile,1,true)
			var oFile = oFso.OpenTextFile(sFile,1,false,-2)
			var aFiles = (oFile.ReadAll()).split(/\r\n|\n/g)
			oFile.close()
			var sCommonFiles = this.a_lib_common.toString()
			
			for(var i = aFiles.length-1; i >= 0; i--){
				aFiles[i] = oFso.GetFileName(aFiles[i])
				if(sCommonFiles.isSearch(aFiles[i]) || !this.isClass(aFiles[i])){
					aFiles.splice(i,1) // Removes default and unwanted files from the array
					continue
				}
			}
			
			return this.include.apply(this,aFiles) // EXCELLENCE!!
		}
		catch(ee){
			this.reload(this.__alias + ".includeAll(): " + ee.description)
			return false
		}
		finally{
			if(oFso.FileExists(sFile)) oFso.DeleteFile(sFile,true)
		}
	},
	
	reload : function reload(s){
		if(this.$load_exception) return
		s = typeof(s) == "string" ? s : "Reloading Library.."
		if(oWsh.Popup(s + "\nCurrent/last file: " + this.s_load_file + "\n\nReload?",60,this.__name + " Library",32 + 4) == 6){
			window.setTimeout("window.location.reload()",1)
		}
		else window.setTimeout("window.close()",0)
		this.$load_exception = true
	},
	
	setOnError : function(b){
		window.onerror = !b ? null : function(message,url,line){
			var s = "\nError\t: " + message
			s = s.concat("\nUrl\t: " + url)
			s = s.concat("\nLine\t: " + line)
			s = s.concat("\nFile\t: " + this.s_load_file) // MUST BE GLOBAL DEFINED
			s = s.concat("\n\nAbort\t: Closes Windows")
			s = s.concat("\nRetry\t: Reloads Window")
			s = s.concat("\nIgnore\t: Continues")
			
			var r = oWsh.Popup(s,60,__H.__name + " Loader",32 + 2)
			if(r == 3) window.setTimeout("window.close()",5)
			else if(r == 4) window.setTimeout("location.reload()",1000) // MUST HAVE A TIME DELAY
		}
	},	
	
	getOnError : function(){
		return window.onerror
	},
	
	////////////////////////////////////////////////////////////////////////////////
	//////// METHODS BELOW WILL BE REPLACED BY THE CLASS __H.Log (alias __HLog) ////
	////////////////////////////////////////////////////////////////////////////////
	
	log : function log(s){
		try{
			var dt = (new Date()).format("#YYYY##MM##DD# #hh#:#mm#:#ss#")
			s = "[" + dt + "] Initialize(): ".concat(s)
			var o = this.o_options.textarea_result
			if(o && o.contentEditable){
				o.innerText = (o.innerText).concat("\n" + s)
			}
			else this.popup(s)
		}
		catch(e){
			this.error(e,this)
		}
		finally{
			oWsh.RegWrite(this.o_options.path_reg_hkcu + "LastLog",s)
			oWsh.RegWrite(this.o_options.path_reg_base + "LastLog",s)
		}
	},
	
	popup : function popup(s,t){
		s = typeof(s) == "string" ? s : ""
		t = typeof(t) == "string" ? t : "Popup"
		oWsh.Popup(s,20,t,48);
	},
	
	debug : function debug(s,f){
		if(this.$debug) this.log("DEBUG " + s,f)
	},
	
	error : function error(oErr,f,s){
		if(!(oErr instanceof Error)) return;
		s =  "Initialize(): " + oErr.description + (typeof(s) == "string" || "")
		this.popup(s)
		oWsh.RegWrite(this.o_options.path_reg_hkcu + "LastError",s)
	},
	
	errorCode : function errorCode(s){
		return 8881
	},
	
	s_load_file : "n/a",
	
	loadScript : function loadScript(file){
		try{
			var script = this.byClone("script")
			//file = file.trim().replace(/\\|\\\\/g,"/") // must do this!!
			script.id = this.o_options.s_prefix_script + document.scripts.length
			script.language = "JScript.Encode"
			script.type = "text/javascript"
			script.charset = "UTF-8" //"ISO-8859-1"
			script.src = file
			this.d_lib_includes.Add(script.src,script.id)
			this.s_load_file = file
			this.o_dom_head.appendChild(script)
			
			return true
		}
		catch(e){
			this.error(e,this)
			return false
		}
	},
	
	////////////////////////////////////////////////////////////////////////////////
	////////  ////
	////////////////////////////////////////////////////////////////////////////////
	
	emptyFn : function(){},
	
	extendFn : function extendFn(subc,superc,overrides){
		try{
			var F = function(){};
			F.prototype = superc.prototype;
			subc.prototype = new F();
			
			subc.prototype.constructor = subc;
			subc.parent = superc.prototype;
			
			if(superc.prototype.constructor == Object.prototype.constructor){
				superc.prototype.constructor = superc;
			}
			if(overrides){
				for(var i in overrides){
					//__HLog.debug("### "+i)
					subc.prototype[i] = overrides[i];
				}
			}
		}
		catch(ee){
			this.error(ee,this)
		}
	},
	
	byId    : document.getElementById,
	
	byIds    : function byIds(){
		var a = [], iLen = arguments.length
		if(iLen <= 1) return document.getElementById(arguments[0])
		
		for(var i = iLen-1, el; i >= 0; i--){
			if(typeof arguments[i] != "string") continue
			el = document.getElementById(arguments[i])
			if(el) a.push(el)
		}
		return a;
	},
	
	byName  : document.getElementsByName,
	byTag   : document.getElementsByTagName,
	
	byClone : function byClone(sElement){
		try{
			if(!this.d_byclones.Exists(sElement = sElement.toLowerCase())){
				this.d_byclones.Add(sElement,document.createElement(sElement))
			}
			return this.d_byclones(sElement).cloneNode(true)
		}
		catch(ee){
			this.error(ee,this)
			return null
		}
	},
	
	byClass : function byClass(sClass,oNode,sTag) {
		var a = [];
		var aNodes = (oNode || document).getElementsByTagName(sTag || "*");
		var oRe = new RegExp("(^|\\\\s)" + sClass + "(\\\\s|$)");
		for(var i = aNodes.length-1; i >= 0; i--) {
			if(oRe.test(aNodes[i].className)){
				a.push(aNodes[i])
			}
		}
		return a;
	},
	
	tree : function tree(oo){
		try{
			if(!oo) oo = __H
			var buf = new __H.Common.StringBuffer(oo.__name), a = []
			
			for(var m in oo){
				if(oo.hasOwnProperty(m)){
					a.push(m)
				}
			}
			for(var i = 0, iLen = a.length; i < iLen; i++){
				if(typeof(oo[a[i]]) == "function") buf.append("[F] " + a[i] + "()" + "\n")
				else if(this.isArray(oo[a[i]])){
					buf.append("[A] " + a[i] + "[]" + "\n")
					if(a[i] == "__DOES NOT EXIST ANYMORE subclasses"){
						var ooo = oo[a[i]]
						for(var j = 0, iLen3 = ooo.length; j < iLen3; j++){
							buf.append(this.tree((new ooo[j]())))
						}
						//alert(j)
					}
				}
				else if(this.isDictionary(oo[a[i]])) buf.append("[D] " + a[i] + "\n")
				else if(typeof(oo[a[i]]) == "object") buf.append("[O] " + a[i] + "{}" + "\n")
				else if(typeof(oo[a[i]]) == "boolean") buf.append("[B] " + a[i] + "\n")
				else if(typeof(oo[a[i]]) == "number") buf.append("[N] " + a[i] + "\n")
				else if(typeof(oo[a[i]]) == "string") buf.append("[S] " + a[i] + " = " + oo[a[i]] + "\n")
				else buf.append("[N/A] " + a[i] + "\n")
			}
			return buf
			//__HLog.append(null,s)
		}
		catch(ee){
			this.error(ee,this)
		}
		finally{
			buf.empty()
		}
	},
	
	caller : function caller(bArgs){
		try{
			var o = arguments.callee
			if(typeof(o.caller) == "function") o = o.caller
			
			var oRe1 = /[ \t]*function[ \t]+([a-z0-9_]+)[ \t]*([a-z0-9_, \t\(\)]{2,})[ \t]*.*$/ig
			var oRe2 = /[ \t]*function[ \t]*([a-z0-9_, \t\(\)]{2,})[ \t]*.*$/ig
			var oRe3 = /debug|log|caller/g
			
			if((o.toString().split("\n")[0]).isSearch(oRe1)){
				for(var f = RegExp.$1; oRe3.test(f); ){
					o = o.caller
					if(o){
						o.toString().split("\n")[0].isSearch(oRe1);
						f = RegExp.$1
					}
					else break;
				}
				return RegExp.$1 + (bArgs ? "" + (RegExp.$2).replace(/ /g,"") + "" : "()")
			}
			else if((o.toString().split("\n")[0]).isSearch(oRe2)){
				return "Anonymous" + (bArgs ? "" + (RegExp.$1).replace(/ /g,"") + "" : "()")
			}
		}
		catch(ee){
			//__HLog.error(ee,this) // don't run from here.. Will loop
		}
		return "Unknown()"
	},
	
	toVBArray : function(a,bEmpty){
		var d = new ActiveXObject("Scripting.Dictionary")
		if(!this.isArray(a)) return d.Items()
		for(var i = a.length-1; i >= 0; i--){
			d.Add("" + i,a[i])
		}
		if(bEmpty) a.empty()
		return d.Items()
	},
	
	toJSON : function(object){
		switch(typeof(object)){
			case "undefined" :
			case "function"  :
			case "unknown"   : return;
		}
		
		if(object === null) return "null";
		else if(object.toJSON) return object.toJSON();
		else if(this.isElement(object)) return;
		
		var a = [], v;
		for(var property in object){
			v = this.toJSON(object[property]);
			if(this.isDefined(v)){
				a.push('"' + property + '":' + v);
			}
		}
		
		return '{' + a.join(', ') + '}';
	},
	
	typeOf : function(v){
		var type = typeof v
		if(type == "object"){
			if(this.isDictionary(v)) type = "dictionary"
			else if(this.isEnumerator(v)) type = "enumerator"
			else if(this.isArray(v)) type = "array"
			else if(this.isElement(v)) type = "element"
			else if(this.isElementArray(v)) type = "elementarray"
			else if(this.isError(v)) type = "error"
			else if(this.isNull(v)) type = "null"
			else if(this.isObject(v) && typeof(v.length) != "number") type = "object"
			else if(this.isRegExp(v)) type = "regexp"
			else if(this.isDate(v)) type = "date"
			else if(this.isArguments(v)) type = "argument"
		}
		return type
	},
	
	forEach : function(object,block){
		if(object){ // http://dean.edwards.name/weblog/2006/07/enum/
			var resolve = Object; // default
			if(object instanceof Function){
				// functions have a "length" property
				resolve = Function;
			}
			else if(object.forEach instanceof Function) {
				// the object implements a custom forEach method so use that
				object.forEach(block);
				return;
			}
			else if(typeof object == "string") {
				resolve = String;
			}
			else if(typeof object == "number") {
				resolve = Number;
			}
			else if(typeof object.length == "number") {
				resolve = Array;
			}
			return resolve.forEach(object,block);
		}
		return false
	},
	
	
	////////////////////////////////////////////////////////////////////////////////
	//////// IS METHODS ////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////
	
	
	isHTA : function(doc){
		var doc = doc || document
		return doc
		&& typeof(doc) == "object"
		&& doc.nodeType == 9
		&& doc.mimeType == "HTML Application"
	},
	
	isClass : function(s){
		var oRe = new RegExp("(HUI@[a-z@0-9_]*)\.((?:js|jse|vb|vbe))$","ig") // Matches HUI$..+$.js
		return typeof(s) == "string" && s.isSearch(oRe);
	},
	
	isWScript : function(){
		return typeof(WScript) != "undefined" && typeof(WScript) == "object"
	},
	
	isElement2 : function(){
		var i, o, l;
		for(i = 0, l = arguments.length; i < l; i++){
			o = arguments[i];
			if(!o || typeof o !== "object" || o.nodeType !== 1) return false
		}
		return !!i
	},
	
	isElement : function(object,sElem){
		if(!object || typeof(object) != "object" || object.nodeType != 1) return false
		else if(typeof(object.hasOwnProperty) == "function") return false
		else if(typeof(sElem) == "string") return object.tagName == sElem.toUpperCase();
		return true
	},
	
	isTextNode : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(!arguments[i] || typeof(arguments[i]) != "object" || arguments[i].nodeName !== "#text") return false
		}
		return !!i
	},
	
	isElementArray : function(object,sElem){
		if(!object || typeof(object) != "object" || typeof(object.length) != "number") return false
		else if(typeof(object.hasOwnProperty) == "function") return false
		else if(typeof(sElem) == "string") return object[0].tagName == sElem.toUpperCase();
		return true
	},
	
	isElementEditable : function(object){
		return this.isElement.apply(this,arguments) && arguments[0].contentEditable
	},
	
	isFrame : function(){
		var i, o, l
		for(i = 0, l = arguments.length; i < l; i++){
			o = arguments[i]
			if(!o || typeof(o) !== "object" || o.nodeType !== 9);
			else if(typeof((arguments[i]).hasOwnProperty) == "function");
			else if(o.tagName == "iframe") continue
			return false
		}
		return !!i
	},
	
	isVBArray : function(){
		var i, o, l
		for(i = 0, l = arguments.length; i < l; i++){
			o = arguments[i]
			if(!o || !(typeof o === "unknown" && this.toArray(o))) return false;
		}
		return !!i
	},
	
	isString : function(){
		var i, s, l
		for(i = 0, l = arguments.length; i < l; i++){
			s = arguments[i]
			if(!(typeof s == "string" || (s instanceof String))) return false;
		}
		return !!i
	},
	
	isStringEmpty : function(){
		var i, s, l, oRe = /.+/g
		l = arguments.length
		if(l == 0) return true
		for(i = 0; i < l; i++){
			s = arguments[i]
			if(typeof s !== "string" || !s.trim().isSearch(oRe)) return true
		}
		return false
	},
	
	isStringNumber : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(!arguments[i]);
			else if(typeof(arguments[i]) == "string" ||
			(arguments[i]).constructor == String || arguments[i] instanceof String) continue
			else if(typeof(arguments[i]) == "number" ||
			(arguments[i]).constructor == Number || arguments[i] instanceof Number) continue
			return false
		}
		return !!i
	},
	
	isFunction : function(){
		var i, f, l
		for(i = 0, l = arguments.length; i < l; i++){
			f = arguments[i]
			if(typeof f !== "function" || !(f instanceof Function)) return false
		}
		return !!i
	},
	
	// isInstance : function(){
	// var i, f, l
	// for(i = 0, l = arguments.length; i < l; i++){
	// f = arguments[i]
	// if(typeof f !== "function" || !(f instanceof Function)) return false
	// }
	// return !!i
	// },
	
	isUndefined : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(typeof(arguments[i]) === "undefined") continue
			return false
		}
		return !!i
	},
	
	isDefined : function(){
		// http://www.webreference.com/js/column26/apply.html__H
		return !this.isUndefined.apply(this,arguments)
	},
	
	isNot : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(typeof(arguments[i]) === "undefined" || !arguments[i]) continue
			return false
		}
		return !!i
	},
	
	isNull : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(typeof(arguments[i]) === "object" && arguments[i] == null) continue
			return false
		}
		return !!i
	},
	
	isRegExp : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(arguments[i] == null);
			else if((arguments[i]).constructor == RegExp || arguments[i] instanceof RegExp) continue
			return false
		}
		return !!i
	},
	
	isEnumerator : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(arguments[i] == null);
			else if((arguments[i]).constructor == Enumerator || arguments[i] instanceof Enumerator) continue
			return false
		}
		return !!i
	},
	
	isDictionary : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(arguments[i] == null);
			else if(typeof(arguments[i]) == "object" &&
			typeof(arguments[i].Exists) == "unknown" && typeof(arguments[i].Remove) == "unknown") continue
			//this.isUnknown(arguments[i].Exists,arguments[i].Add,arguments[i].Remove)) continue // BUG!! don't work
			return false
		}
		return !!i
	},
	
	isDate : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(arguments[i] == null);
			else if(arguments[i] instanceof Date || (arguments[i]).constructor == Date) continue
			return false
		}
		return !!i
	},
	
	isArray : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(arguments[i] == null);
			else if(arguments[i] instanceof Array || (arguments[i]).constructor == Array) continue
			return false
		}
		return !!i
	},
	
	isArrayLike : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(arguments[i] == null);
			else if(typeof(arguments[i]) == "object" &&
			!(arguments[i] instanceof Array) && arguments[i].length >= 0) continue
			return false
		}
		return !!i
	},
	
	isArguments : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(arguments[i] == null);
			else if((arguments[i] instanceof Object)
			&& typeof(arguments[i].callee) == "function"
			&& arguments[i].length >= 0) continue
			return false
		}
		return !!i
	},
	
	isObject : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(!arguments[i]);
			else if((arguments[i]).constructor == Object
			|| arguments[i] instanceof Object
			|| typeof(arguments[i]) == "object") continue
			return false
		}
		return !!i
	},
	
	isError : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(!arguments[i] || typeof(arguments[i]) != "object");
			else if(arguments[i] instanceof Error) continue
			return false
		}
		return !!i
	},
	
	isUnknown : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(!arguments[i]);
			else if(typeof(arguments[i]) == "unknown") continue
			return false
		}
		return !!i
	},
	
	isNumber : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(arguments[i] == null);
			else if((arguments[i].constructor == Number || arguments[i] instanceof Number) && isFinite(arguments[i])) continue
			return false
		}
		return !!i
	},
	
	isNumeric : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(arguments[i] == null);
			else if(!isNaN(parseFloat(arguments[i])) && isFinite(arguments[i])) continue
			return false
		}
		return !!i
	},
	
	isBoolean : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(arguments[i] == null);
			else if(arguments[i].constructor == Boolean || arguments[i] instanceof Boolean) continue
			return false
		}
		return !!i
	},
		
	isXML : function(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(arguments[i] == null);
			else if(typeof(arguments[i]) == "object" &&
			typeof(arguments[i].async) == "boolean" &&
			typeof(arguments[i].parseError) == "object") continue
			return false
		}
		return !!i
	}
}

		
		
				