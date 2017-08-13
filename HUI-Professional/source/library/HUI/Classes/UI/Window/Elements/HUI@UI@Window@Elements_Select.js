// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

__H.include("HUI@UI@Window@Elements.js")

__H.register(__H.UI.Window.Elements,"Select","Select",function Select(oSelect){
	var select = null
	var index  = 0
	var d_selects = new ActiveXObject("Scripting.Dictionary")

	this.isSelect = function isSelect(oSel){
		return oSel != null && typeof(oSel) == "object" && oSel.tagName == "SELECT"
	}

	this.isSelectRange = function isSelectRange(iIndex){
		return typeof(iIndex) == "number" && iIndex >= 0 && iIndex <= select.options.length
	}

	this.isSelected = function isSelected(){
		for(var i = select.options.length-1; i >= 0; i--){
			if(select.options[i].selected){
				return true
			}
		}
		return false
	}

	this.setSelect = function setSelect(oSel){
		if(!(select = this.getSelect(oSel))){
			select = __H.byClone("select")
		}
		return select
	}

	this.getSelect = function getSelect(oSel,bForce){
		if(!this.isSelect(oSel)) return false
		var sKey = (oSel.form.name + "::" + oSel.name + "::" + oSel.sourceIndex).toLowerCase()
		if(bForce && d_selects.Exists(sKey)) d_selects.Remove(sKey)
		if(!d_selects.Exists(sKey)){
			d_selects.Add(sKey,oSel)
		}
		return d_selects(sKey)
	}

	// DON*T MOVE!! MUST BE AFTER getSelect and setSelect
	this.setSelect(oSelect)

	this.reset = function reset(bDisabled){
		select.options.selectedIndex = 0
		select.disabled = bDisabled ? true : false;
		if(select.onchange) select.onchange();
		return true
	}

	this.addIndex = function addIndex(iIndex,sValue,sText){
		iIndex = typeof(iIndex) == "number" && iIndex >= 0 ? iIndex : select.options.length;
		select.options[iIndex] = new Option(sText,sValue);
		select.options[iIndex].selected = true
	}

	this.clear = function clear(iIndex,bDefault){
		bDefault = bDefault ? true : false
		iIndex = this.isSelectRange(iIndex) ? iIndex : 0;
		for(var i = select.options.length; i > iIndex; i--){
			select.options[iIndex].removeNode()// = null; // Also possible to use options.removeNode()
		}
		if(!bDefault) select.options[iIndex] = new Option(" ","",false,false);
	}

	this.setIndex = function setIndex(iIndex){
		iIndex = this.isSelectRange(iIndex) ? iIndex : 0
		select.options.selectedIndex = iIndex
		if(select.onchange) select.onchange();
	}

	this.getIndex = function getIndex(){
		bResult = -1;
		iIndex = __H.isNumber(iIndex) ? iIndex : 0;
		for(var len = select.options.length; iIndex < len; iIndex++){
			if(select.options[iIndex].value == sValue){
				bResult = iIndex;
				break;
			}
		}
		return iIndex
	}

	this.disable = function disable(bDisabled){
		select.disabled = bDisabled ? true : false;
	}

	this.setValue = function setValue(sValue){
		if(typeof(sValue) == "string") return false
		select.options.value = sValue
		if(select.onchange) select.onchange();
	}

	this.getValue = function getValue(){
		return select.value
	}

	this.setText = function setText(iIndex){
		iIndex = this.isSelectRange(iIndex) ? iIndex : select.options.selectedIndex
		select.options[iIndex].text
	}

	this.getTextIndex = function getTextIndex(sText,iIndex){
		iIndex = this.isSelectRange(iIndex) ? iIndex : 0;
		sText = (new String(sText)).toLowerCase()
		for(var sText1, sText2, len = select.options.length; iIndex < len; iIndex++){
			var sText2 = select.options[iIndex].text
			var sText1 = (select.options[iIndex].text).replace(/([0-9]{3} )(.*)/g,"$2")
			if(sText1.toLowerCase() == sText || sText2.toLowerCase() == sText){
				return iIndex
			}
		}
		return iIndex
	}

	this.addArraySplit = function addArraySplit(aObject,sSplit,bNoCount){
		bNoCount = bNoCount ? true : false
		iIndex = this.isSelectRange(iIndex) ? iIndex : 1;
		if(__H.isObject(sSplit)){
			this.clear(iIndex)
			for(var i = 0, iLen = aObject.length; i < iLen; i++, iIndex++){
				var n = !bNoCount ? (iIndex+1).toNumberZero() : ""
				var a = (aObject[i]).split(sSplit), sValue = a[1], sText = n + " " + a[0];
				this.addIndex(iIndex,sValue,sText);
			}
		}
		if(select.options.length > 0) select.options[0].selected = true
		else bResult = false
	}

	this.addArrayObject = function addArrayObject(aObject,bNoCount,iIndex,sValue,sText){
		bNoCount = bNoCount ? true : false
		var bText = (typeof(sText) == "string")
		iIndex = this.isSelectRange(iIndex) ? iIndex : 1;

		if(typeof(aObject) == "object"){
			//this.clear(iIndex);
			for(var i = 0, iLen = aObject.length; i < iLen; i++, iIndex++){
				var n = !bNoCount ? (iIndex+1).toNumberZero() : ""
				var sText2 = n + " " + aObject[i][sValue]
				if(bText) sText2 = sText2 + " (" + aObject[i][sText] + ")"
				var sValue2 = aObject[i][sValue]
				this.addIndex(iIndex,sValue2,sText2);
			}
		}
		if(select.options.length > 0){
			return select.options[0].selected = true
		}
		return false
	}

	this.addArray = function addArray(a,iIndex,bNoCount,sExtra){
		bNoCount = bNoCount ? true : false
		var sExtra = typeof(sExtra) == "string" ? sExtra : ""
		iIndex = this.isSelectRange(iIndex) ? iIndex : 1;
		if(typeof(a) == "object"){
			for(var i = 0, iLen = a.length; i < iLen; i++, iIndex++){
				var n = !bNoCount ? (iIndex+1).toNumberZero() : ""
				var sValue = a[i], sText = n + " " + sExtra + a[i];
				this.addIndex(iIndex,sValue,sText)
			}
		}
		if(select.options.length > 0){
			return (select.options[0].selected = true)
		}
		return false
	}

	this.addArrayItem = function addArrayItem(aObject,bNoCount,sItem){
		bNoCount = bNoCount ? true : false
		if(__H.isObject(aObject)){
			for(var i = 0, iLen = aObject.length; i < iLen; i++, iIndex++){
				var n = !bNoCount ? (iIndex+1).toNumberZero() : ""
				sValue = aObject[i][sItem], sText = n + " " + sValue;
				this.addIndex(iIndex,sValue,sText)
			}
		}
		if(select.options.length > 0){
			return (select.options[0].selected = true)
		}
		return false
	}

	this.getArray = function getArray(bSelected){
		var a = []
		var bSelected = bSelected ? true : false
		iIndex = this.isSelectRange(iIndex) ? iIndex : 1;
		var o = sValue == "text" ? sValue : "value"
		for(var i = iIndex, iLen = select.length; i < iLen; i++){
			if(bSelected && !select.options[i].selected) continue
			a.push(select.options[i][o])
		}
		return a
	}

	////

	this.setSearch = function setSearch(oSel,iIndex,bValue,sRegPrefix){
		bValue = bValue ? true : false
		sRegPrefix = typeof(sRegPrefix) == "string" ? sRegPrefix : ""
		iIndex = typeof(iIndex) == "number" ? iIndex : 0

		oSel.onkeydown = new Function("__HSelect.searchKeyDown(this," + iIndex + "," + bValue + ",'" + sRegPrefix + "')")
		oSel.onkeyup = new Function("__HSelect.searchKeyUp(this," + iIndex + "," + bValue + ",'" + sRegPrefix + "')")
	}

	this.searchKeyDown = function searchKeyDown(oSel,iIndex,bValue,sRegPrefix){
		try{ // This is needed for IE. FireFox and Opera has this implemented by default
			var oEvent = window.event
			var sKeyCode = oEvent.keyCode;
			var sToChar = String.fromCharCode(sKeyCode); // excellence
			//if(sKeyCode > 47 && sKeyCode < 91){
			if(!sToChar.isSearch(/[0-9a-z\- ]/ig)) return;
			var iNow = new Date().getTime();
			oSel.setAttribute("active",true)
			if(oSel.getAttribute("finder") == null){
				oSel.setAttribute("finder",sToChar.toUpperCase())
				oSel.setAttribute("timer",iNow)
			}
			else if(iNow > parseInt(oSel.getAttribute("timer")) + 2000) { //Rest all;
				oSel.setAttribute("finder",sToChar.toUpperCase())
				oSel.setAttribute("timer",iNow) //reset timer;
				oSel.setAttribute("active",false)
			}
			else {
				oSel.setAttribute("finder",oSel.getAttribute("finder") + sToChar.toUpperCase())
				oSel.setAttribute("timer",iNow); // update timer;
			}
			var sFinder = oSel.getAttribute("finder");
			iIndex = __H.isNumber(iIndex) ? iIndex : 0;
			var o = bValue ? "value" : "text"
			sRegPrefix = typeof(sRegPrefix) == "string" ? sRegPrefix : false
			for(var i = iIndex; i < oSel.options.length; i++){
				var sTest = oSel.options[i][o];
				if(sRegPrefix){
					var oRe = new RegExp("("+sRegPrefix+")(.*)","ig")
					if(sTest.match(oRe)) sTest = RegExp.$2
				}
				if(sTest.toUpperCase().indexOf(sFinder) == 0){
					oSel.options[i].selected = true;
					break;
				}
			}
			event.returnValue = false;
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}

	this.searchKeyUp = function searchKeyUp(oSel){
		if(!oSel.getAttribute("active")) return;
		var iNow = new Date().getTime();
		if(iNow > parseInt(oSel.getAttribute("timer"))){
			oSel.setAttribute("active",false)
			try{oSel.onchange() }
			catch(ee){}
		}
	}

	this.clone = function clone(oSel){ // NOT 100% doesn't change "name"
		if(!this.isSelect(oSel)) return false
		var oClone = select.cloneNode(true)
		var n = oSel.name
		oClone.onchange = oSel.onchange
		oClone.selectedIndex = oSel.selectedIndex
		oClone.setAttribute("name",n)
		oSel.replaceNode(oClone)
		return oSel
	}
})

var __HSelect = new __H.UI.Window.Elements.Select()
