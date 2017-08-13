// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

__H.include("HUI@UI@Window@Elements.js")

__H.register(__H.UI.Window.Elements,"Input","HTML Input",function Input(){
	var form = null
	var forms = new ActiveXObject("Scripting.Dictionary")
	
	this.isInput = function isInput(oInput){
		return this.isElement(oInput) && oInput.tagName == "INPUT"
	}
	
	this.setForm = function setForm(){
		
	}
	
	this.getRadioValue = function getRadioValue(oInput){
		if(!this.isInput(oInput)) return false
		if(oInput.type == "radio"){
			var oRadio = oInput.form.elements[oInput.name]
			for(var i = 0, iLen = oRadio.length; i < iLen; i++){
				if(oRadio[i].checked){
					return oRadio[i].value
				}
			}
		}
	}

	this.isChecked = function isChecked(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(!this.isInput(arguments[i]) || !arguments[i].checked) return false
		}
		return true
	}	
	
	this.setCheckbox = function setCheckbox(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			var sName = arguments[i];
			if(!oForm[sName]) continue
			for(var j = k = 0, iLen2 = oElem.length; j < iLen2; j++){
				var n = oElem[j].name
				if(n == sName && k == 0) k++,oElem[j].onclick = new Function("__HInput.setCheckboxSup(this.form,'" + sName + "');");
				else if(n.substring(0,sName.length) == sName){
					oElem[j].onclick = new Function("__HInput.setCheckboxSub(this.form,'" + sName + "');");
				}
			}
		}
	}
	
	this.setCheckboxSup = function setCheckboxSup(oForm){
		var bChecked = (oForm[sElements][1]) ? oForm[sElements][0].checked : oForm[sElements].checked;
		for(var i = 0, iLen = oElem.length; i < iLen; i++){
			if(!oElem[i].disabled){ // ignore disabled tags
				var sNamePrefix = (oElem[i].name).substring(0,sElements.length);
				if(oForm[sElements][1] && oElem[i].name == sElements) oElem[i].checked = bChecked
				else if(sNamePrefix == sElements) oElem[i].checked = bChecked;
			}
		}
	}
	
	this.setCheckboxSub = function setCheckboxSub(){
			var oElements = sElements
			var oElements2 = (oForm[sElements2][1]) ? oForm[sElements2][0] : oForm[sElements2]
			if(oElements.checked) oElements2.checked = true
			else { // This removes check on mainbox if last subbox
				var oRe = new RegExp(sElements2,"ig")
				var bSubChecked = false
				for(var j = k = 0, iLen = oElem.length; j < iLen; j++){
					var n = oElem[j].name
					if(n.isSearch(oRe) && n > sElements2){ // subbox
						if(oElem[j].checked){ // at least one is checked... don't bother
							bSubChecked = true // Execellence!!
							break;
						}
					}/*
					else if(n.isSearch(oRe) && n >= sElements2){ // subbox same
						if(oElem[j].checked){ // at least one is checked... don't bother
							alert(2 + " "+ n + " "+ sElements2)
							bSubChecked = (k == 0) ? false : true // Execellence!!
							break;
						}
						k++
					}*/
				}
				if(!bSubChecked){ // Am real good!!
					oElements2.checked = false
					oElements2.onclick()
				}
			}
			//__HUtil.kill(oElements,oElements2)
		}
		
	this.isValueHasHtml = function isValueHasHtml(sValue){
		if(typeof(oInput) != "string") return false
		var oRe = /([\<])([^\>]{1,})*([\>])/i
		return sValue.isSearch(oRe)
	}
	
	this.isValueEmpty = function isValueEmpty(sValue){
		if(typeof(oInput) != "string") return false
		var oRe = /([\s]+)/i
		return sValue.isSearch(oRe)
	}
	
	this.isValuePhome = function isValuePhome(sValue,lang){
		if(typeof(oInput) != "string") return false
		// /\(?\d{3}\)?([-\/\.])\d{3}\1\d{4}/g; // US phone syntax
		var oRe = /(\+\d{1,4})?\s?(\(\d{1,2}\))?(\d{1,3})?\s?\d{1,8}/; // +46 (0)8 51850962 or 71545846
		return sValue.isSearch(oRe)
	}
	
	this.isValueUrl = function isValueUrl(sValue){
		if(typeof(oInput) != "string") return false
		var oRe = /^http:\/\/.+\..+|^ftp:\/\/.+\..+/ig
		return sValue.isSearch(oRe)
	}	
	
	this.isValueName = function isValueName(sValue){
		if(typeof(oInput) != "string") return false
		var oRe = /[a-z_\xC5\xE5\xC4\xE4\xD6\xF6]+\s*[A-Za-z_\xC5\xE5\xC4\xE4\xD6\xF6]*/ig
		return sValue.isSearch(oRe)
	}
	
	this.isValueEmail = function isValueEmail(sValue){
		// /(@.*@)|(\.\.)|(@\.)|(\.@)|(^\.)/;
		var oRe = /(^[_a-z0-9\-_]+(\.[_a-z0-9\-_]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$)/ig
		return sValue.isSearch(oRe)
	}
})

var __HInput = new __H.UI.Window.Elements.Input()