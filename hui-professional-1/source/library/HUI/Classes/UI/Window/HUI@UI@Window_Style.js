// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**


__H.include("HUI@UI@Window.js")

__H.register(__H.UI.Window,"Style","CSS Stylesheet",function Style(){
	var o_this = this
	var b_initialized = false
	
	var d_styles
	var d_rules
	var o_style
	var o_style_tmp
	var o_rule
	var o_rule_tmp
	var o_sheets
	
	// IE		| 	Other browser
	// ---------------------------------
	// rules		cssRules
	// addRule		insertRule
	
	/////////////////////////////////////
	//// DEFAULT
	
	var o_cur_load	
	var o_options = {
		
	}
	
	var initialize = function initialize(bForce){
		if(b_initialized && !bForce) return;
		
		d_styles   = new ActiveXObject("Scripting.Dictionary")
		d_rules    = new ActiveXObject("Scripting.Dictionary")

		b_initialized = true		
	}
	initialize()
	
	/////////////////////////////////////
	//// 
	
	this.setOptions = function setOptions(oOptions){		
		try{
			Object.extend(o_options,oOptions,true)
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
	
	this.refresh = function refresh(){
		try{
			initialize(true)
			d_styles.RemoveAll()
			d_rules.RemoveAll()
			o_sheets = document.styleSheets
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.setStyleCSS = function setStyleCSS(oElement,sValues){
		try{ // http://userjs.org/help/tutorials/efficient-code__HWin
			oElement.setAttribute("style",sValues)
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getStyleCSS = function getStyleCSS(oElement){
		try{
			return oElement.getAttribute("style")
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.setStyle = function setStyle(sClass,sElement,sValue,bForce){
		try{
			if(!isNaN(sValue)) sValue = new String(sValue)
			if(__H.isStringEmpty(sClass,sElement,sValue)) __HExp.ArgumentIllegal(null,arguments)
			if(d_styles.Exists(sClass));
			else{
				if((o_style_tmp = this.getRuleByClass(sClass))){
					d_styles.Add(sClass,o_style_tmp)
				}
			}
			if((o_style = d_styles(sClass)) && (o_style[sElement] || bForce)){
				o_style[sElement] = sValue
				return true
			}
			if(!o_style) __HLog.debug("Unable to locate class: " + sClass)
			return false
		}
		catch(ee){alert(ee.description)
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getStyle = function getStyle(sClass){
		try{
			if(__H.isStringEmpty(sClass)) __HExp.ArgumentIllegal(null,arguments)
			
			if(d_styles.Exists(sClass));
			else if((o_style_tmp = this.getRuleByClass(sClass))){
				d_styles.Add(sClass,o_style_tmp)
			}
			
			return d_styles(sClass)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.setClass = function setClass(sClass){
		try{
			if(__H.isStringEmpty(sClass)) __HExp.ArgumentIllegal(null,arguments)
			
			if(d_styles.Exists(sClass));
			else{
				if((o_style_tmp = this.getRuleByClass(sClass))){
					d_styles.Add(sClass,o_style_tmp)
				}			
			}
			if((o_style = d_styles(sClass)) && (o_style[sElement] || bForce)){
				o_style[sElement] = sValue
				return true
			}
			
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.importStyleURL = function getStyle(sStyleId,sStyleUrl){
		try{
			if(__H.isStringEmpty(sClass)) __HExp.ArgumentIllegal(null,arguments)
			
			var oStyleTag = document.getElementById(sStyleId);
            var oSheet = oStyleTag.sheet ? oStyleTag.sheet : oStyleTag.styleSheet;
			var s = oFso.GetAbsolutePathName(".")
            sStyleUrl = sStyleUrl.replace(s,"").replace(/\\/g,"/")
            if (oSheet.insertRule) { // all browsers, except IE before version 9
                oSheet.insertRule("@import url('" + sStyleUrl + "');",0);
            }
            else {  // Internet Explorer  before version 9
                if (oSheet.addImport) {
                    oSheet.addImport(sStyleUrl);
                }
            }
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.addRule = function addRule(sHref,sClass,sElement,sValue){
		try{
			if(o_sheet = this.getRulesByHref(sHref)){
				o_sheet.rules.addRule(sClass,sElement + ": " + sValue + ";");
				return true
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getRuleByClass = function getRuleByClass(sClass){
		try{
			if(__H.isStringEmpty(sClass)) __HExp.ArgumentIllegal(null,arguments)
			o_sheets = document.styleSheets
			for(var i = 0, iLen = o_sheets.length; i < iLen; i++){
				if(o_sheets(i).owningElement.tagName == "STYLE"){
					for(var oRules, j = 0, iLen2 = o_sheets(i).imports.length; j < iLen2; j++){
						oRules = o_sheets(i).imports(j).rules
						for(var k = 0, iLen3 = oRules.length; k < iLen3; k++){
							if(oRules(k).selectorText == sClass || oRules(k).selectorText == "." + sClass){
								return oRules(k).style
							}
						}
					}
					oRules = o_sheets(i).rules
					for(j = 0, iLen2 = oRules.length; j < iLen2; j++){
						if(oRules(j).selectorText == sClass || oRules(j).selectorText == "." + sClass){
							return oRules(j).style
						}
					}
				}
				else {
					//alert("getRuleByClass " +o_sheets(i).owningElement.tagName)
				}
			}				
			return null
		}
		catch(ee){
			__HLog.error(ee,this)
			return null
		}
	}
	
	this.getRulesByHref = function getRulesByHref(sHref){
		try{
			if(__H.isStringEmpty(sHref)) __HExp.ArgumentIllegal(null,arguments)
			if(d_rules.Exists(sHref)) return d_rules(sHref)
			o_sheets = document.styleSheets
			for(var i = 0, iLen = o_sheets.length; i < iLen; i++){
				if(o_sheets(i).owningElement.tagName == "STYLE"){
					for(var j = 0, iLen2 = o_sheets(i).imports.length; j < iLen2; j++){
						if((o_sheets(i).imports(j).href).isSearch(sHref)){
							d_rules.Add(sHref,o_sheets(i).imports(j).rules)
							return d_rules(sHref,{
								index : i,
								index_import : j,
								rules : o_sheets(i).imports(j).rules,
								href  : sHref
							})
						}						
					}
					for(j = 0, iLen2 = o_sheets(i).rules.length; j < iLen2; j++){
						if((o_sheets(i).href).isSearch(sHref)){
							d_rules.Add(sHref,{
								index : j,
								rules : o_sheets(i).rules,
								href  : sHref
							})
							return o_sheets(i).rules
						}
					}
				}
				else {
					
				}
			}				
			return null
		}
		catch(ee){
			__HLog.error(ee,this)
			return null
		}
	}
	
	this.removeImport = function removeImport(sHref){
		try{
			if(o_sheet = this.getRulesByHref(sHref)){			
				if(d_rules.Exists(sHref)){
					d_rules.Remove(sHref)
				}
				document.styleSheets(o_sheet.index).removeImport(o_sheet.index_import)
				return true
			}
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
})

var __HStyle = new __H.UI.Window.Style()
var __HCSS = __HStyle.setStyle