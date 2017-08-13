// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions, All Rights Reserved.

//**Start Encode**

__H.include("HUI@IO@File.js")

__H.register(__H.IO.File,"INI","INI, INF, REG and URL files",function INI(){
	var o_this = this;
	var b_initialized = false

	var oReExt
	var oReSec
	var oReKey;
	var oReNL

	var s_return
	var d_files
	var d_sections
	var o_cur_file
	var o_cur_sec
	var o_cur_key
	var b_sections_loaded = false

	/////////////////////////////////////
	//// DEFAULT
	
	var o_options = {
		b_backup          : true,
		b_backup_overwite : true
	}
	
	var initialize = function initialize(){
		if(b_initialized) return;

		__HLog.debug("Initializing class: __H.IO.File.INI")
		
		// => {C:\windows\win.ini} OR  {C:\windows\inf\wmp.inf}
		oReExt = /ini|inf|conf|reg|url/ig
		// => {[ 1section.nam� with any ch@ractor except equal & semi-colon signs]}
		oReSec = /^[ \t]*[\[]([^=;]+)[\]][ \t]*$/ig
		// => { name = value}  OR {'name% = value }
		oReKey1 = /^["'%]{0,1}[ \t]*([^=;]+)[ \t]*["'%]{0,1}[ \t]*=[ \t]*([^;]+)[ \t]*$/ig
		// => {name}  OR  { %name%}
		oReKey2 = /^[ \t]*["'%]{0,1}[ \t]*([^=;]+)[ \t]*["'%]{0,1}[ \t]*$/ig 
		// => { name ; comment}  OR  { "name" ; }
		oReKey3 = /^[ \t]*["'%]{0,1}[ \t]*([^=;]+)[ \t]*["'%]{0,1}[ \t]*;[ \t]*([^;]*)[ \t]*$/ig 
		 // => { name = value ; comment}  OR  {"name% = value ;}
		oReKey4 = /^[ \t]*["'%]{0,1}[ \t]*([^=;]+)[ \t]*["'%]{0,1}[ \t]*=[ \t]*([^;]+)[ \t]*;[ \t]*(.*)[ \t]*$/ig
		// => { %na;me;jj%}  OR  { "na;me;jj'}
		oReKey5 = /^[ \t]*["'%]([ \t]*[^=]+)["'%][ \t]*$/
		// => 
		oReNL   = /\r\n|\r|\n/g
		// => {; comment}
		oReCom1 = /^[ \t]*[;][ \t]*(.*)[ \t]*$/g
		// => { name = value ; comment}
		oReCom2 = /^[ \t]*([^;]+)[ \t]*[;][ \t]*(.*)[ \t]*$/g

		s_return   = new String("\r\n")
		d_files    = new ActiveXObject("Scripting.Dictionary")
		d_sections = new ActiveXObject("Scripting.Dictionary")

		o_cur_file = {}
		o_cur_sec  = {}
		o_cur_key  = {}

		b_initialized = true
	}

	this.isBlankLine = function(s){
		return (typeof(s) == "string" && s.isSearch(/^\[ \t]*$/g));
	}
	
	this.setOptions = function setOptions(oOptions){		
		if(__H.isObject(oOptions)) return Object.extend(o_options,oOptions,true)
		return false
	}
	
	this.getOptions = function(){
		return o_options
	}
	
	this.toString = function(){
		return this.getSections().toString()
	}
	
	//this.valueOf = function(){
		//return this.toJSON()
	//}
	
	/////////////////////////////////////
	//// INI/INF FILE

	this.isFile = function isFile(args){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(arguments[i]
				&& typeof(arguments[i]) == "string"
				&& oFso.FileExists(arguments[i])
				&& (sExt = oFso.GetExtensionName(arguments[i])).isSearch(oReExt)){
				continue
			}
			__HLog.debug("Either invalid or access unavailable for INI file: " + arguments[i])
			return false
		}
		return !!i
	}

	var load = function load(sFile){
		try{
			 initialize()
			 if(o_this.isFile(sFile)){
				if(!d_files.Exists(o_cur_file.s_file_tmp = sFile.toLowerCase())){
					o_cur_file.o_file = oFso.GetFile(o_cur_file.s_file_tmp)
					d_files.Add(o_cur_file.s_file_tmp,o_cur_file.o_file.OpenAsTextStream(__HIO.read,__HIO.TristateUseDefault).ReadAll().split(oReNL))
				}
				return true
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return false
	}

	this.setFile = function setFile(sFile){
		try{
			this.reset()
			if(!load(sFile)){
				__HExp.ArgumentIllegal("Argument[0] is not a valid INI/INF file: " + sFile)
			}
			o_cur_file.s_name  = o_cur_file.s_file_tmp
			o_cur_file.a_lines = d_files(o_cur_file.s_name)
			o_cur_file.i_len   = o_cur_file.a_lines.length
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getFile = function getFile(){
		return o_cur_file.s_name
	}

	this.write = function write(){
		try{
			if(__HFile.write(o_cur_file.s_name,o_cur_file.a_lines.join(s_return),o_options.b_backup,o_options.b_backup_overwite)){
				o_cur_file.i_len = o_cur_file.a_lines.length
				return true
			}
		return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.tidy = function tidy(){
		if(!this.loadSections()) return false
		try{			
			var iCurNameLength = 0
			var maxNameLength = function(){
				for(var i = 0; i < o_cur_file.i_len; i++){
					if(!o_this.isSection(o_cur_file.a_lines[i])){
						if(o_this.setKey(o_cur_file.a_lines[i],i)){
							if(o_cur_key.b_hasvalue){
								iCurNameLength = Math.max(iCurNameLength,o_cur_key.s_name.length)
							}
						}
					}
				}
			}
			maxNameLength()
			
			var addNameSpace = function(n){	
				var s = "", j = n.length, i = 0
				while(i++ <= (iCurNameLength+2-j)) s = s.concat(" ")
				return n.concat(s)
			}
			
			for(var i = 0, s; i < o_cur_file.i_len; i++){
				if(this.isSection(o_cur_file.a_lines[i])){					
					o_cur_file.a_lines.insert(i++,"")
					o_cur_file.i_len++
				}
				else if(this.setKey(o_cur_file.a_lines[i],i)){
					s = o_cur_key.s_comment ? "; " + o_cur_key.s_comment : ""
					if(o_cur_key.b_hasvalue){
						o_cur_file.a_lines[i] = addNameSpace(o_cur_key.s_name) + " = " + o_cur_key.s_value + s
					}
					else o_cur_file.a_lines[i] = o_cur_key.s_key + s					
				}
				else if(o_cur_file.a_lines[i].isSearch(oReCom1)){
					o_cur_file.a_lines[i] = "; " + RegExp.$1
				}
				else if(this.isBlankLine(o_cur_file.a_lines[i])){
					//o_cur_file.a_lines.splice(i,0)
				}
			}
			return this.write() && this.reset()
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.info = function info(){
		try{
			/*__HLog.log("File: " + o_cur_file.s_name)
			for(var o in o_cur_file){
				if(!o_cur_file.hasOwnProperty(o)) continue
				__HLog.log("# " + o + ": " + o_cur_file[o])
			}*/

			__HLog.log("Section: " + o_cur_sec.s_name)
			for(var o in o_cur_sec){
				if(!o_cur_sec.hasOwnProperty(o)) continue
				__HLog.log("# " + o + ": " + o_cur_sec[o])
			}
			return;
			__HLog.log("Key: " + o_cur_key.s_name)
			for(var o in o_cur_key){
				if(!o_cur_key.hasOwnProperty(o)) continue
				__HLog.log("# " + o + ": " + o_cur_key[o])
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.toXML = function toXML(){
		try{
			
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.toObject = function toObject(){
		try{
			var o = {}
			var a = this.getSections()
			var oRe = /^\[|\]$/ig
			for(var i = 0, iLen = a.length; i < iLen; i++){
				if(this.setSection(a[i])){
					o[a[i].replace(oRe,"")] = this.getKeys(true)
				}
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		a.empty()
		return o
	}
	
	this.toJSON = function toJSON(bObject){		
		return __H.toJSON(this.toObject())		
	}
	
	this.reset = function reset(bAll){
		if(!b_initialized) return;
		
		if(o_cur_file.a_lines){
			o_cur_file.o_file = null
			o_cur_file.s_name = undefined	
		
			o_cur_file.a_lines.empty()
			o_cur_file.i_len = 0
		}
		
		d_sections.RemoveAll()
		o_cur_sec.empty()
		b_sections_loaded = false

		o_cur_key.empty()

		if(bAll) d_files.RemoveAll()		
		return true
	}

	/////////////////////////////////////
	//// SECTION

	this.loadSections = function loadSections(bForce){
		try{
			if(b_sections_loaded && !bForce) return true			
			for(var i = 0; i < o_cur_file.i_len; i++){
				if(this.isSection(o_cur_file.a_lines[i])){
					d_sections.Add(o_cur_file.a_lines[i],i)
					__HLog.debug("Loading index: " + i + ", section: " + o_cur_file.a_lines[i])
				}
			}
			b_sections_loaded = true
			return true
		}
		catch(ee){
			__HLog.debug("index => " + i + ", line => " + o_cur_file.a_lines[i])
			__HLog.error(ee,this)
			return false
		}
	}

	this.isSection = function isSection(sSection){
		return typeof(sSection) == "string" && sSection.trim().isSearch(oReSec)
	}

	this.hasSection = function hasSection(sSection){
		return this.loadSections() && typeof(sSection) == "string" && d_sections.Exists(this.toSection(sSection))
	}

	this.toSection = function toSection(sSection){
		o_cur_sec.s_name_tmp = ""
		if(typeof(sSection) == "string"){
			o_cur_sec.s_name_tmp = "[" + sSection.trim().replace(/^[\[]{1,}[ \t]*(.+)[ \t]*[\]]{1,}$/g,"$1") + "]"			
		}
		return o_cur_sec.s_name_tmp
	}

	this.setSection = function setSection(sSection){
		try{
			if(sSection === o_cur_sec.s_name_org) return true
			if(this.hasSection(sSection)){
				__HLog.debug("Section: " + o_cur_sec.s_name_tmp)

				o_cur_sec.s_name_org = sSection
				o_cur_sec.s_name  = o_cur_sec.s_name_tmp
				o_cur_sec.i_start = d_sections(o_cur_sec.s_name)

				for(var i = o_cur_sec.i_start + 1; i < o_cur_file.i_len; i++){
					if(this.isSection(o_cur_file.a_lines[i])) break;					
				}

				o_cur_sec.i_end = i
				o_cur_sec.i_len = i - o_cur_sec.i_start-1
				o_cur_sec.s_range = "0 < index <= " + o_cur_sec.i_len
				//this.info()
				__HLog.debug(" Range: " + o_cur_sec.s_range)

				return true
			}
			__HExp.ArgumentIllegal("Argument[0] is not a valid section name for file " + o_cur_file.s_name)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getSection = function getSection(){
		return o_cur_sec.s_name
	}
	
	this.getSections = function getSection(){
		this.loadSections()
		return new VBArray(d_sections.Keys()).toArray()
	}
	
	this.addSection = function addSection(sSection,aKeys){
		try{
			var bWrite = false
			if(!d_sections.Exists(this.toSection(sSection))){
				d_sections.Add(o_cur_sec.s_name_tmp,o_cur_file.a_lines.length) // Must Add() before push()
				o_cur_file.a_lines.push(o_cur_sec.s_name_tmp)
				bWrite = true
			}
			else __HLog.log("Ignoring section add for: " + o_cur_sec.s_name_tmp + ". Already exists in file: " + o_cur_file.s_name)

			if(this.setSection(sSection)){
				if(__H.isArray(aKeys)){
					for(var i = 0, iLen = aKeys.length; i < iLen; i++){
						o_cur_file.a_lines.push(aKeys[i])
					}
					o_cur_sec.i_len += i
					o_cur_sec.i_end += i
					bWrite = i > 0
				}
			}
			if(bWrite) return this.write()
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.delSections = function delSections(){
		try{
			for(var i = 0, iLen = arguments.length; i < iLen; i++){
				if(d_sections.Exists(this.toSection(arguments[i]))){
					if(this.setSection(arguments[i])){
						o_cur_file.a_lines.splice(o_cur_sec.i_start,o_cur_sec.i_len)
						d_sections.Remove(o_cur_sec.s_name)
					}
				}
			}
			return this.write()
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}		
	}

	/////////////////////////////////////
	//// KEY

	this.isKeys = function isKeys(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(typeof(arguments[i]) != "string"
				|| (arguments[i] = arguments[i].trim()).length == 0
				|| arguments[i].isSearch(oReCom1)){
				return false
			}
		}
		return i > 0
	}


	this.hasKeys = function hasKeys(){
		var iLen = arguments.length, iHas = 0
		for(var j = 1, i = o_cur_sec.i_start+1; i < o_cur_file.i_len; i++){
			if(j++ > o_cur_sec.i_len) break;
			if(this.setKey(o_cur_file.a_lines[i],i)){
				for(var k = 0; k < iLen; k++){
					if(o_cur_key.name === arguments[k]){
						iHas++
						break;
					}
				}
				if(iHas == iLen) return true
			}
		}
		return false
	}

	this.toKey = function toKey(sKeyName,sKeyValue){
		if(!this.isComment(sKeyName)){
			return sKeyName + (typeof(sKeyValue) == "string" ? " = " + sKeyValue : "")
		}
		__HExp.ArgumentIllegal("Argument [0] is not a valid key name")
	}

	this.setKey = function setKey(sKey,iIndex,bSimple){
		try{
			if(!this.isKeys(sKey)) return false			
			var n, v, c
			if(sKey.isSearch(oReKey1))      n = RegExp.$1, v = RegExp.$2, c = undefined
			else if(sKey.isSearch(oReKey2)) n = RegExp.$1, v = undefined, c = undefined
			else if(sKey.isSearch(oReKey3)) n = RegExp.$1, v = undefined, c = RegExp.$2
			else if(sKey.isSearch(oReKey4)) n = RegExp.$1, v = RegExp.$2, c = RegExp.$3
			else if(sKey.isSearch(oReKey5)) n = RegExp.$1, v = undefined, c = undefined
			else {
				n = RegExp.$1, v = undefined, c = undefined
				__HLog.debug("Strange Key " + sKey + " in section: " + o_cur_sec.s_name + " file: " + o_cur_file.s_name)
				// Strange keys in section [Optional Components] in file c:\windows\inf\accessor.inf
				//__HExp.ArgumentIllegal("Argument [0] " + sKey + " is not a valid key for section: " + o_cur_sec.s_name)
			}
			
			o_cur_key.empty()
			o_cur_key.s_name = n.replace(/[ \t]+/g," ").trim()
			o_cur_key.b_hasvalue = !!v
			o_cur_key.s_value = v ? v.replace(/[ \t]+/g," ") : v
			if(!bSimple){	
				o_cur_key.s_key = o_cur_key.s_name + (v ? " = " + o_cur_key.s_value : "")
				o_cur_key.s_key_org = sKey
				o_cur_key.s_comment = c ? c.trim() : c			
				o_cur_key.i_index = !isNaN(iIndex) ? iIndex : 0
			}
			else delete o_cur_key.b_hasvalue
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getKey = function getKey(){
		return o_cur_key.s_name
	}

	this.getKeys = function getKeys(bSimple){
		try{
			var a = []
			for(var j = 1, i = o_cur_sec.i_start+1; i < o_cur_file.i_len; i++){
				if(j++ > o_cur_sec.i_len) break;
				if(this.setKey(o_cur_file.a_lines[i],i,bSimple)){
					a.push(o_cur_key)
				}
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return a
	}

	var addKeySafe = function addKeySafe(sKeyName,sKeyValue){
		try{
			for(var j = 1, i = o_cur_sec.i_start+1; i < o_cur_file.i_len; i++){
				if(j++ > o_cur_sec.i_len) break;
				if(o_this.setKey(o_cur_file.a_lines[i],i)){
					if(o_cur_key.s_name == sKeyName){
						// replaces value and prevents duplicates
						o_cur_file.a_lines[i] = o_this.toKey(sKeyName,sKeyValue)
						__HLog.debug("Updated index: " + i + " to " + o_cur_file.a_lines[i])
						return true
					}
				}
			}
			o_cur_file.a_lines.insert(o_cur_sec.i_end,o_this.toKey(sKeyName,sKeyValue))
			o_cur_file.i_len = o_cur_file.a_lines.length
			o_cur_sec.i_len++
			o_cur_sec.i_end++
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.addKey = function addKey(sKeyName,sKeyValue){
		try{
			if(__H.isStringEmpty(sKeyName)){
				__HExp.ArgumentIllegal("Argument [0] has invalid type (should be string) or is empty")
			}
			return addKeySafe(sKeyName,sKeyValue) && this.write()
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.addKeys = function addKeys(oKeys){
		try{
			if(!__H.isObject(oKeys)) __HExp.ArgumentIllegal("Argument [1] has invalid type (should be an Object)")
			var j = 0
			for(var o in oKeys){
				if(!oKeys.hasOwnProperty(o)) continue
				if(!addKeySafe(o,oKeys[o])){
					__HLog.logPopup("Unable to add key: " + o + " in section: " + o_cur_sec.s_name + " for file: " + o_cur_file.s_name)
				}
				else j++
			}
			if(j == 0){
				__HLog.debug("No keys were added in section: " + o_cur_sec.s_name + " for file: " + o_cur_file.s_name)
				return true
			}
			return this.write()
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.delKey = function delKey(sKeyName){
		if(this.indexOfKey(sKeyName) > 0){
			o_cur_file.a_lines.remove(o_cur_key.i_index)
			o_cur_sec.i_len--
			o_cur_sec.i_end--
			return this.write()
		}
		return false
	}

	this.setKeyValue = function setKeyValue(sKeyName,sKeyValue){
		return this.addKey.apply(this,arguments)
	}

	this.getKeyValue = function getKeyValue(sKeyName){
		if(this.indexOfKey(sKeyName) > 0) return o_cur_key.s_value
		return undefined
	}

	this.getValue = function getValue(sKeyName){
		return this.getKeyValue.call(this,sKeyName)
	}

	this.indexOfKey = function indexOfKey(sKeyName){
		try{
			for(var j = 1, i = o_cur_sec.i_start+1; i < o_cur_file.i_len; i++){
				if(j++ > o_cur_sec.i_len) break
				if(this.setKey(o_cur_file.a_lines[i],i)){					
					if(o_cur_key.s_name == sKeyName){
						__HLog.debug("name => " + sKeyName + "==" + o_cur_key.s_name 
							+ ", value => " + o_cur_key.s_value 
							+ ", comment => " + o_cur_key.s_comment)
						return i
					}
				}
			}
			return -1;
		}
		catch(ee){
			__HLog.error(ee,this)
			return -1;
		}
	}

	/////////////////////////////////////
	//// COMMENT

	this.isComment = function isComment(sComment){
		if(typeof(sComment) != "string") return false
		return sComment.isSearch(/;/g) && (sComment.isSearch(oReCom1) || sComment.isSearch(oReCom2))
	}

	this.addComment = function addComment(sComment,iIndex){
		try{
			if(!__H.isStringNumber(sComment,iIndex)){
				__HExp.ArgumentIllegal("Arguments [0] or [1] has invalid types (should be string,number)")
			}
			if(iIndex <= o_cur_sec.i_len && iIndex > 0){
				o_cur_file.a_lines.insert(iIndex,sComment)
				return this.write()
			}
			__HExp.ArgumentOutOfRange("Argument [1] is out of range for section: " + o_cur_sec.s_name + ". Should be: " + o_cur_sec.s_range)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.delComment = function delComment(iIndex){
		try{
			if(typeof(iIndex) != "number"){
				__HExp.ArgumentIllegal("Arguments [0] has invalid type (should be number)")
			}
			if(iIndex <= o_cur_sec.i_len && iIndex > 0){
				if(this.isComment(o_cur_file.a_lines[iIndex])){
					if(o_cur_file.a_lines[iIndex].isSearch(oReCom2)){
						o_cur_file.a_lines[iIndex] = RegExp.$1
					}
					else o_cur_file.a_lines.remove(iIndex)
					return this.write()
				}
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
})

var __HINI = new __H.IO.File.INI()