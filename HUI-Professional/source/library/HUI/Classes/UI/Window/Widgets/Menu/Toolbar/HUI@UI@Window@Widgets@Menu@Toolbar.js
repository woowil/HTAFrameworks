// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

/*
	This dynamic toolbar menu was originally created by Toh Zhiqiang called this.ObjectMenubar Version 1.1.1
	But as you can see,comparing my version,the menubar is totally revised
	
	This file depends on 
*/


var px = "px";
var pl = "paddingLeft"
var btw = "borderTopWidth"
var bbw = "borderBottomWidth"
var blw = "borderLeftWidth"
var brw = "borderRightWidth"
	
var menuCount = 0;
var itemCount = 0;
var sepCount = 0;
var popUpMenuObj = null;
var ContextMenuObj = null // woody
var ContextTag = null
var ContextRange = null
var ContextText = ""
var showValue = 0;
var hideValue = 0;
var staticMenuId = [];
var aMenus = [] // Woody
var MENU_ICON_PATH


__H.include("HUI@UI@Window@Widgets@Menu.js")

__H.register(__H.UI.Window.Widgets.Menu,"Toolbar","Toolbar",function Toolbar(){
	
})

function mnu_get_propIntVal(obj,propertyName){
	try{
		var n = 0
		if(obj.style && (n = obj.style[propertyName]));
		else if(obj.currentStyle) n = obj.currentStyle[propertyName]
		else if(document.defaultView && document.defaultView.getComputedStyle){
			n = document.defaultView.getComputedStyle(obj,null).getPropertyValue(propertyName);
		}
		
		//if(isNaN(n)) return 0	
		return n && typeof(n = parseInt(n)) == "number" ? n : 0;
	}
	catch(e){
		__HLog.error(e,this)
		return 0
	}
}

/*
Get the left position of the pop-up menu.
*/
function mnu_get_mainLeftPos(menuObj,x){
	if(x + menuObj.offsetWidth <= __HWindow.getClientWidth()){
		return x;
	}
	else {
		return x - menuObj.offsetWidth;
	}
}

/*
Get the top position of the pop-up menu.
*/
function mnu_get_mainTopPos(menuObj,y){
	if(y + menuObj.offsetHeight <= __HWindow.getClientHeight()){
		return y;
	}
	else {
		return y - menuObj.offsetHeight;
	}
}

/*
Get the left position of the submenu.
*/
function mnu_get_subLeftPos(menuObj,x,offset){
	if(x + menuObj.offsetWidth - 2 <= __HWindow.getClientWidth()){
		return x - 2;
	}
	else {
		return x - menuObj.offsetWidth - offset;
	}
}

/*
Get the top position of the submenu.
*/
function mnu_get_subTopPos(menuObj,y,offset){
	var top = mnu_get_propIntVal(menuObj,btw);
	var bottom = mnu_get_propIntVal(menuObj,bbw);
	if(y + menuObj.offsetHeight <= __HWindow.getClientHeight()){
		if(__HWindow.bSafari){
			return y - top;
		}
		else {
			return y;
		}
	}
	else {
		if(__HWindow.bSafari){
			return y - menuObj.offsetHeight + offset + bottom;
		}
		else {		
			return y - menuObj.offsetHeight + offset + top + bottom;
		}
	}
}

/*
Pop up the submenu.
*/
function mnu_set_activatePopupSub(obj){
	var parentMenuObj = obj.parent.menuObj;
	var menuObj = obj.subMenu.menuObj;	
	var menuObjLayer = obj.subMenu.menuObjLayer;	
	var x, y;
	if(parentMenuObj.style.position == "fixed"){
		x = parentMenuObj.offsetLeft + parentMenuObj.offsetWidth - mnu_get_propIntVal(parentMenuObj,brw);
		y = parentMenuObj.offsetTop + obj.offsetTop + mnu_get_propIntVal(parentMenuObj,btw) - mnu_get_propIntVal(menuObj,btw);
		menuObj.style.position = "absolute";
		menuObj.style.left = mnu_get_subLeftPos(menuObj,x,obj.offsetWidth) + px;
		menuObj.style.top = mnu_get_subTopPos(menuObj,y,obj.offsetHeight) + px;
		menuObj.style.position = "fixed";
	}
	else {
		if(parentMenuObj.mode == "static" && !__HWindow.bIE50){
			x = obj.offsetLeft + parentMenuObj.offsetWidth - mnu_get_propIntVal(parentMenuObj,blw) - mnu_get_propIntVal(parentMenuObj,brw) - __HWindow.getScrollLeft();
			y = obj.offsetTop - mnu_get_propIntVal(menuObj,btw) - __HWindow.getScrollTop();
			if(__HWindow.bIEnew){
				x += mnu_get_propIntVal(parentMenuObj,blw);
				y += mnu_get_propIntVal(parentMenuObj,btw);
			}
			else if(__HWindow.bSafari){
				x += 8;
				y += mnu_get_propIntVal(menuObj,btw) + 13;
			}
			menuObj.style.left = (mnu_get_subLeftPos(menuObj,x,obj.offsetWidth) + __HWindow.getScrollLeft()) + px;
			menuObj.style.top = (mnu_get_subTopPos(menuObj,y,obj.offsetHeight) + __HWindow.getScrollTop()) + px;
		}
		else {			
			x = parentMenuObj.offsetLeft + parentMenuObj.offsetWidth - mnu_get_propIntVal(parentMenuObj,brw) - __HWindow.getScrollLeft();
			y = parentMenuObj.offsetTop + obj.offsetTop + mnu_get_propIntVal(parentMenuObj,btw) - mnu_get_propIntVal(menuObj,btw) - __HWindow.getScrollTop();
			menuObj.style.left = (mnu_get_subLeftPos(menuObj,x,obj.offsetWidth) + __HWindow.getScrollLeft()) + px;
			menuObj.style.top = (mnu_get_subTopPos(menuObj,y,obj.offsetHeight) + __HWindow.getScrollTop()) + px;
		}
	}
	if(__HWindow.bIE && menuObj.mode == "fixed"){
		menuObj.initialLeft = parseInt(menuObj.style.left) - __HWindow.getScrollLeft();
		menuObj.initialTop = parseInt(menuObj.style.top) - __HWindow.getScrollTop();
	}
	mnu_show_layer(menuObjLayer,menuObj)
	menuObj.style.visibility = "visible";
}

/*
Pop up the main menu.
*/
function mnu_set_activatePopupMain(menuObj,e){
	try{
		menuObj.style.left = (mnu_get_mainLeftPos(menuObj,__HWindow.getClientX(e)) + __HWindow.getScrollLeft()) + px;
		menuObj.style.top = (mnu_get_mainTopPos(menuObj,__HWindow.getClientY(e)) + __HWindow.getScrollTop()) + px;
		var display = popUpMenuObj.menuObj.style.display;
		mnu_show_layer(popUpMenuObj.menuObjLayer,menuObj)
		popUpMenuObj.menuObj.style.display = "none";
		popUpMenuObj.menuObj.style.visibility = "visible";
		popUpMenuObj.menuObj.style.display = display;
	}
	catch(e){}
}

function mnu_set_contextMain(menuObj,e){ // Woody right click in text box
	menuObj.style.left = (mnu_get_mainLeftPos(menuObj,__HWindow.getClientX(e)) + __HWindow.getScrollLeft()) + px;
	menuObj.style.top = (mnu_get_mainTopPos(menuObj,__HWindow.getClientY(e)) + __HWindow.getScrollTop()) + px;
	var display = popUpMenuObj.menuObj.style.display;
	mnu_show_layer(ContextMenuObj.menuObjLayer,menuObj)
	ContextMenuObj.menuObj.style.display = "none";
	ContextMenuObj.menuObj.style.visibility = "visible";
	ContextMenuObj.menuObj.style.display = display;
}

function mnu_set_itemOver(e){
	var previousItem = this.parent.previousItem;	
	if(previousItem){
		HUIToolbarActivate(previousItem,'menuItemOut')
		if(previousItem.subMenu){
			HUIToolbarActivate(previousItem.arrowObj,'menuArrowOut')
			if(previousItem.iconObj){
				HUIToolbarActivate(previousItem.iconObj,'menuIconOut')
			}
		}
		var menuObj = __H.byIds(this.parent.menuObj.id);
		for(var i = 0; i < menuObj.childNodes.length; i++){
			if(menuObj.childNodes[i].enabled && menuObj.childNodes[i].subMenu){
				mnu_hide_menus(menuObj.childNodes[i].subMenu.menuObj,menuObj.childNodes[i].subMenu.menuObjLayer);
			}
		}
	}
	if(this.enabled){
		HUIToolbarActivate(this,'menuItemOver')
		if(this.subMenu && (!this.disabled || this.items)){
			HUIToolbarActivate(this.arrowObj,'menuArrowOver')
			this.subMenu.menuObjLayer.style.visibility = this.subMenu.menuObj.style.visibility = "visible" // Woody
			mnu_set_activatePopupSub(this);
		}
		if(this.iconObj && this.iconClassNameOver){
			HUIToolbarActivate(this.iconObj,'menuIconOver')
		}
	}	
	this.parent.previousItem = this;
}

function mnu_set_itemClick(e){
	try{
		if(this.subMenu || this.isClickable) return; // Woody
		if(this.enabled && this.actionOnClick){
			var action = this.actionOnClick;
			// Woody
			this.parent.menuObj.lastindex = this.index
			
			if(action.indexOf("link:") == 0){
				location.href = action.substr(5);
			}
			else if(action.indexOf("target:") == 0){
				window.open(action.substr(7),"TargetWin","") 
			}
			else if(action.indexOf("open:") == 0){ // Woody
				oShX.open(action.substr(5))
			}
			else {
				if(action.indexOf("code:") == 0){
					eval(action.substr(5));
				}
				else {
					location.href = action;
				}
			}
		}
		
		HUIToolbarCancel(e)
		
		var t0 = mnu_hide_context()
		if(this.parent.menuObj.mode == "cursor"){
			var t1 = window.setTimeout(function(){mnu_hide_visable()},0) // Woody
			var t2 = window.setTimeout(function(){mnu_hide_cursor()},1) // Woody
		}
		else if(this.parent.menuObj.mode == "absolute" || this.parent.menuObj.mode == "fixed"){
			var t1 = window.setTimeout(function(){mnu_hide_visable()},0) // Woody
			if(typeof(mnu_hide_menubars) == "function"){
				var t2 = window.setTimeout(function(){mnu_hide_menubars()},1) // Woody
			}
		}
	}
	catch(ee){
		__HLog.error(ee)
	}
}

/*
Event handler that handles onmouseout event of the menu item.
*/
function mnu_set_itemOut(){
	if(this.enabled){
		if(!(this.subMenu && this.subMenu.menuObj.style.visibility == "visible")){
			HUIToolbarActivate(this.iconObj,'menuIconOut')
		}
		if(this.subMenu){
			if(this.subMenu.menuObj.style.visibility == "visible"){
				HUIToolbarActivate(this.arrowObj,'menuArrowOver')
				if(this.iconObj){
					HUIToolbarActivate(this.iconObj,'menuIconOver')
				}
			}
		}
		else {
			HUIToolbarActivate(this,'menuItemOut')
			if(this.iconObj){
				HUIToolbarActivate(this.iconObj,'menuIconOut')
			}
		}
	}
}

function mnu_set_activatePopup(e){
	e = e ? e : window.event;
	if(!popUpMenuObj){
		return;
	}
	var state = popUpMenuObj.menuObj.style.visibility;
	if(state == "visible"){		
		for(var i = 1; i <= menuCount; i++){
			var menuObj = __H.byIds("DOMenu" + i);
			var menuObjLayer = __H.byIds("DOMenuLayer" + i);
			if(menuObj.mode == "cursor"){
				mnu_hide_layer(menuObjLayer)
				menuObj.style.visibility = "hidden";
				menuObj.style.left = "0px";
				menuObj.style.top = "0px";
				menuObj.initialLeft = 0;
				menuObj.initialTop = 0;
			}
		}
	}
	else {
		var targetElm = (e.target) ? e.target : e.srcElement;
		if(targetElm.disabled || targetElm.tagName == "SELECT") return; // Fix on disabled buttons,=> undefined
		if(targetElm.nodeType == 3){
			ContextTag = targetElm
			targetElm = targetElm.parentNode;
		}
		var oDummy = window.setTimeout(function(){mnu_hide_menubars()},1)
		mnu_set_activatePopupMain(popUpMenuObj.menuObj,e);
	}
}

function mnu_set_activateContext(e){ // Woody
	e = e ? e : window.event;
	if(!ContextMenuObj){
		return;
	}
	var state = ContextMenuObj.menuObj.style.visibility;
	if(state == "visible"){
		for(var i = 1; i <= menuCount; i++){
			var menuObj = __H.byIds("DOMenu" + i);
			var menuObjLayer = __H.byIds("DOMenuLayer" + i);
			if(menuObj.mode == "cursor"){
				mnu_hide_layer(menuObjLayer)
				menuObj.style.visibility = "hidden";
				menuObj.style.left = "0";
				menuObj.style.top = "0";
				menuObj.initialLeft = 0;
				menuObj.initialTop = 0;
			}
		}
	}
	else {
		var targetElm = (e.target) ? e.target : e.srcElement;
		if(targetElm.disabled || targetElm.readonly) return; // Fix on disabled and readonly objects... 
		else if(targetElm.nodeType == 1){ // If text
			ContextTag = targetElm
			targetElm = targetElm.parentNode;		
		}
		mnu_set_contextMain(ContextMenuObj.menuObj,e);
	}
}

/*
Event handler that handles right click event.
*/
function mnu_set_contextMenu(e){
	e = e ? e : window.event
	var oElement = e.srcElement, t
	
	if(__HWindow.getClientX(e) > __HWindow.getClientWidth() || __HWindow.getClientY(e) > __HWindow.getClientHeight()){
		return;
	}
	else if(e.button == 1) mnu_hide_context();
	else if(e.button == 2){ // Right Click		
		mnu_hide_context()
		var t1 = window.setTimeout(function(){mnu_hide_visable()},0) // Woody
		if(typeof(mnu_hide_menubars) == "function"){
			var t2 = window.setTimeout(function(){mnu_hide_menubars()},0) // Woody
		}
		if(ContextMenuObj && (oElement.tagName == "TEXTAREA" || (oElement.tagName == "INPUT" && oElement.type == "text") )){ // Woody
			var state = ContextMenuObj.menuObj.style.visibility;
			if(state == "visible" && (hideValue == 0 || hideValue == 2)){
				mnu_set_activateContext(e);
			}
			if((state == "hidden" || state == "") && (showValue == 0 || showValue == 2)){
				mnu_set_activateContext(e);
			}
		}
		else if(popUpMenuObj){
			var state = popUpMenuObj.menuObj.style.visibility;
			if(state == "visible" && (hideValue == 0 || hideValue == 2)){
				mnu_set_activatePopup(e);
			}
			if((state == "hidden" || state == "") && (showValue == 0 || showValue == 2)){
				mnu_set_activatePopup(e);
			}
		}
	}
}

/*
Show the icon before the display text.
Arguments:
className					: Required. String that specifies the CSS class selector for the icon.
classNameOver			: Optional. String that specifies the CSS class selector for the icon when 
										 the cursor is over the menu item.
*/
function mnu_show_itemIcon(){
	try{
		var element = __H.byClone("span");
		element.id = this.id + "Icon";
		element.className = arguments[0];
		this.insertBefore(element,this.firstChild);
		var height;
		
		if(__HWindow.bIE){
			height = mnu_get_propIntVal(element,"height");
		}
		else {
			height = element.offsetHeight;
		}
		
		element.style.top = Math.floor((this.offsetHeight - height) / 2) + px;
		if(__HWindow.bIE){			
			var left = mnu_get_propIntVal(element,"left");
			if(__HWindow.bIEnew){
				element.style.left = (left - mnu_get_propIntVal(this,"padding-left")) + px;
			}
			else {
				element.style.left = left + px;
			}
		}
		
		this.iconClassName = element.className;		
		HUIToolbarActivate(this,'menuIconOut')
		
		if(arguments.length > 1 && arguments[1].length > 0){
			this.iconClassNameOver = arguments[1];
		}
		this.iconObj = element;
		this.setIconClassName = function setIconClassName(className){
			this.iconObj.className = className
		};
	}
	catch(e){
		__HLog.error(e,this)
	}
}

function mnu_show_itemIcon2(obj,sIcon){
	try{
		if(typeof(sIcon) != "string") return;
		var element = __H.byClone("span");
		element.id = obj.id + "Icon";
		element.innerHTML = '<img src="' + MENU_ICON_PATH + "/" + sIcon + '" height="16" border=0>'
		obj.insertBefore(element,obj.firstChild)
		element.style.position = "absolute"
		element.style.top = Math.floor((obj.offsetHeight - 16) / 2) + px;
		element.style.left = (4 - mnu_get_propIntVal(obj,"paddingLeft")) + px;
	}
	catch(e){
		__HLog.error(e,this)
	}
}

/*
Set the menu object that will show up when the cursor is over the menu item object.
Argument:
menuObj						: Required. Menu object that will show up when the cursor is over the 
										 menu item object.
*/
function mnu_set_sub(menuObj){
	try{
		var element = __H.byClone("div");
		element.id = this.id + "Arrow";
		HUIToolbarActivate(element,"menuArrowOut")
		this.appendChild(element);
		
		var height;
		if(__HWindow.bIE){
			height = mnu_get_propIntVal(element,"height");
		}
		else {
			height = element.offsetHeight;
		}
		element.style.top = Math.floor((this.offsetHeight - height) / 2) + px;
		this.subMenu = menuObj;
		this.subMenu.isClickable = true
		
		this.subMenu.setEnabled = function(){
			this.disabled = false
			this.isClickable = true
		}
		this.subMenu.setDisabled = function(){
			this.disabled = true
			this.isClickable = false
		}
		
		this.arrowObj = element;
		this.setArrowClassName = function setArrowClassName(className){
			this.arrowObj.className = className
		};
		
		//menuObj.menuObj.style.zIndex = this.parent.style.zIndex + this.parent.menuObj.level + 1;
		menuObj.menuObj.style.zIndex = this.parent.menuObj.level + 1;
		menuObj.menuObj.level = this.parent.menuObj.level + 1;
	}
	catch(e){
		__HLog.error(e,this)
	}
}

/*
Add a new menu item to the menu.
Argument:
obj				: Required. Menu item object that is going to be added to the menu object.
*/
function mnu_add_item(obj,sCtrlKey,sIcon){
	try{
		var element, s = "Start"
		if(obj.displayText == "-"){
			element = __H.byClone("div");
			element.innerHTML = '<table width="100%" border="0" cellspacing="0" cellpadding="0"><tr><td><hr size="1" noshade width="100%"></td></tr></table>';
			element.id = obj.id;
			
			s = "#1.1"
			HUIToolbarActivate(element,"menuItemSep")
			
			this.menuObj.appendChild(element);
			element.parent = this;
			
			element.setClassName = function(className){
				this.className = this.className;
			};
			s = "#1.2"
			element.onclick = HUIToolbarCancel;
			element.setEnabled = element.setDisabled = function(){}
			
			if(obj.itemName.length > 0){
				this.items[obj.itemName] = element;
			}
			else {
				this.items[this.items.length] = element;
			}
		}
		else {		
			element = __H.byClone("div");
			element.id = obj.id;
			element.actionOnClick = obj.actionOnClick;
			element.enabled = obj.enabled;
			element.subMenu = null;
			s = "#2.1"
			var textNode = document.createTextNode(obj.displayText);
			element.appendChild(textNode)
			element.innerHTML = '<table width="100%" border="0" cellspacing="0" cellpadding="0"><tr><td>' + element.innerHTML + '</td><td align="right">' + (sCtrlKey ? sCtrlKey : "&nbsp;") + '</td></tr></table>';
			this.menuObj.appendChild(element);
			element.parent = this;
			element.index = this.items.length // Woody
			s = "#2.2"
			// Woody
			this.isClickable = true
			
			element.setEnabled = function(){
				this.disabled = false
				this.isClickable = true
			}
			element.setDisabled = function(){
				this.disabled = true
				this.isClickable = false
			}
			element.setColorOver = function(sBorder,sBackground){
				this.style.borderColor = sBorder
				this.style.backgroundColor = sBackground
			}
			
			element.setClassName = function(className){
				this.className = className;
			};
			
			element.setDisplayText = function(text){
				if(this.childNodes[0].nodeType == 3){
					this.childNodes[0].nodeValue = text;
				}
				else {
					this.childNodes[1].nodeValue = text;
				}
			};
			s = "#2.3"
			element.mnu_set_sub = mnu_set_sub;
			element.showIcon = mnu_show_itemIcon;
			element.onmouseover = mnu_set_itemOver;
			element.onclick = mnu_set_itemClick;
			element.onmouseout = mnu_set_itemOut;
			element.onmouseout()
			s = "#2.4"
			if(!obj.enabled) element.disabled = true;		
			if(obj.itemName.length > 0){
				this.items[obj.itemName] = element;
			}
			else {
				this.items[this.items.length] = element;
			}
			if(sIcon){
				mnu_show_itemIcon2(element,sIcon)
			}
		}
	}
	catch(e){
		__HLog.error(e,this)
		return false
	}
}

function mnu_set_item(sOpt,oMenuItem,sOpt3){
	try{
		if(sOpt == "setactive"){
			if(typeof(sOpt3) == "boolean"){
				oMenuItem.enabled = sOpt3
				oMenuItem.disabled = !sOpt3;
			}
		}
	}
	catch(e){
		__HLog.error(e,this)
		return false
	}
}

/*
Create a new menu item object.
Arguments:
displayText				: Required. String that specifies the text to be displayed on the menu item. If 
										 displayText = "-",a menu separator will be created instead.
itemName					 : Optional. String that specifies the name of the menu item. Defaults to "" (no 
										 name).
actionOnClick			: Optional. String that specifies the action to be done when the menu item is 
										 being clicked. Defaults to "" (no action).
enabled						: Optional. Boolean that specifies whether the menu item is enabled/disabled. 
										 Defaults to true.
className					: Optional. String that specifies the CSS class selector for the menu item. 
										 Defaults to "jsdomenuitem".
*/
function mnu_obj_menuItem(){
	this.displayText = arguments[0];
	if(this.displayText == "-"){
		this.id = "menuSep" + (++sepCount);
		//HUIToolbarActivate(this,"menuItemSep")
	}
	else {
		this.id = "mnu_obj_item" + (++itemCount);
	}
	this.itemName = "";
	this.actionOnClick = "";
	this.enabled = true;	
	
	var len = arguments.length;
	if(len > 1 && arguments[1].length > 0){
		this.itemName = arguments[1];
	}
	if(len > 2 && arguments[2].length > 0){
		this.actionOnClick = arguments[2];
	}
	if(len > 3 && typeof(arguments[3]) == "boolean"){
		this.enabled = arguments[3];
	}
	if(len > 4 && arguments[4].length > 0){
		this.className = arguments[4];
	}
}

/*
Create a new menu object.
Arguments:
width							: Required. Integer that specifies the width of the menu.
mode							 : Optional. String that specifies the mode of the menu. Defaults to "cursor".
id								 : Optional,except when mode = "static". String that specifies the id of 
										 the element that will contain the menu. This argument is required when 
										 mode = "static".
alwaysVisible			: Optional. Boolean that specifies whether the menu is always visible. Defaults 
										 to false.
className					: Optional. String that specifies the CSS class selector for the menu. Defaults 
										 to "jsdomenudiv".
*/
function mnu_obj_menu(){
	this.items = [];
	var element;
	var len = arguments.length;
	if(len > 2 && arguments[2].length > 0 && arguments[1] == "static"){
		element = __H.byIds(arguments[2]);
		if(!element) return;
		staticMenuId[staticMenuId.length] = arguments[2];
	}
	else {
		element = __H.byClone("div");
		element.id = "DOMenu" + (++menuCount);
	}
	element.level = 10;
	element.previousItem = null;
	element.mode = "cursor";
	element.alwaysVisible = false;
	element.initialLeft = 0;
	element.initialTop = 0;	
	element.onselectstart = new Function("return false") // prevents selection
	HUIToolbarActivate(element,"menu")
	
	if(len > 1 && arguments[1].length > 0){
		switch(arguments[1]){
			case "cursor":
				element.style.position = "absolute";
				menuMode.mode = "cursor";
				break;
			case "absolute":
				element.style.position = "absolute";
				element.mode = "absolute";
				break;
			case "fixed":
				if(__HWindow.bIE){
					element.style.position = "absolute";
				}
				else {
					element.style.position = "fixed";
				}
				element.mode = "fixed";
				break;
			case "static":
				element.style.position = "static";
				element.mode = "static";
				break;
		}
	}
	if(len > 3 && typeof(arguments[3]) == "boolean"){
		element.alwaysVisible = arguments[3];
	}
	if(len > 4 && arguments[4].length > 0){
		element.className = arguments[4];
	}
	element.style.width = arguments[0] + px;
	element.style.left = "0";
	element.style.top = "0";
	//element.style.borderWidth = "2px";
	
	if(element.mode != "static"){
		document.body.appendChild(element);
	}
	element.lastindex = undefined
	this.menuObj = element;
	this.menuObjLayer = mnu_add_layer(this.menuObj) 
	this.mnu_add_item = mnu_add_item;
	
	// Woody	
	this.delItemIndex = function delItemIndex(iIndex){
		try{
			iIndex = typeof(iIndex) == "number" ? iIndex : this.menuObj.lastindex
			
			var o = this.menuObj.childNodes(iIndex)
			if(!o.subMenu){
				return o.removeNode(true)
			}
		}
		catch(e){}
		return false
	};
	this.delSubMenu = function delSubMenu(sName){
		try{
			if(typeof(sName) != "string") return;
			
			return this.items[sName].removeNode(true)
		}
		catch(e){}
		return false
	};
	this.setItemsEnabled = function setItemsEnabled(){
		try{
			for(i = 0, iLen = arguments.length; i < iLen; i++){
				var o = this.items[arguments[i]]
				if(typeof(o) != "object") continue
				if(o.subMenu || arguments[i] < this.items.length) o.setEnabled()
			}
			return true
		}
		catch(e){}
		return false
	}
	this.setItemsDisabled = function setItemsDisabled(){
		try{
			for(i = 0, iLen = arguments.length; i < iLen; i++){
				var o = this.items[arguments[i]]
				if(typeof(o) != "object") continue
				if(o.subMenu || arguments[i] < this.items.length) o.setDisabled()
			}
			return true
		}
		catch(e){}
		return false
	}
	
	// Woody	
	this.setClassName = function setClassName(className){
		this.menuObj.className = className;
	};
	this.setMode = function setMode(mode){
		switch(mode){
			case "cursor":
				this.menuObj.style.position = "absolute";
				this.menuObj.mode = "cursor";
				break;
			case "absolute":
				this.menuObj.style.position = "absolute";
				this.menuObj.mode = "absolute";
				this.menuObj.initialLeft = parseInt(this.menuObj.style.left);
				this.menuObj.initialTop = parseInt(this.menuObj.style.top);
				break;
			case "fixed":
				if(__HWindow.bIE){
					this.menuObj.style.position = "absolute";
					this.menuObj.initialLeft = parseInt(this.menuObj.style.left);
					this.menuObj.initialTop = parseInt(this.menuObj.style.top);
				}
				else {
					this.menuObj.style.position = "fixed";
				}
				this.menuObj.mode = "fixed";
				break;
		}
	};
	this.setAlwaysVisible = function setAlwaysVisible(alwaysVisible){
		if(typeof(alwaysVisible) == "boolean"){
			this.menuObj.alwaysVisible = alwaysVisible;
		}
	};
	// Woody
	this.disable = function disable(bDisable){
		this.menuObj.disabled = bDisable ? bDisable : false
	};
	
	this.show = function show(){
		this.menuObj.style.visibility = "visible";
	};
	this.hide = function hide(){
		mnu_hide_layer(this.menuObjLayer)
		this.menuObj.style.visibility = "hidden";
		if(this.menuObj.mode == "cursor"){
			this.menuObj.style.left = "0px";
			this.menuObj.style.top = "0px";
			this.menuObj.initialLeft = 0;
			this.menuObj.initialTop = 0;
		}
	};
	this.setX = function setX(x){
		this.menuObj.initialLeft = x;
		this.menuObj.style.left = x + px;
	};
	this.setY = function setY(y){
		this.menuObj.initialTop = y;
		this.menuObj.style.top = y + px;
	};
	this.moveTo = function moveTo(x,y){
		this.menuObj.initialLeft = x;
		this.menuObj.initialTop = y;
		this.menuObj.style.left = x + px;
		this.menuObj.style.top = y + px;
	};
	this.moveBy = function moveBy(x,y){
		var left = parseInt(this.menuObj.style.left);
		var top = parseInt(this.menuObj.style.top);
		this.menuObj.initialLeft = left + x;
		this.menuObj.initialTop = top + y;
		this.menuObj.style.left = (left + x) + px;
		this.menuObj.style.top = (top + y) + px;
	};
	this.setBorderWidth = function setBorderWidth(width){
		this.menuObj.style.borderWidth = width + px;
	};
}

function mnu_obj_menudelete(obj){
	try{
		var aItems = obj.items.reverse()
		for(var i = 0, iLen = aItems.length; i < iLen; i++){
			if(aItems[i].subMenu) mnu_obj_menudelete(aItems[i].subMenu)
			else aItems[i].removeNode(true)//,obj.delItemIndex(i)
		}
		
		obj.menuObj.removeNode(true)
		obj.menuObjLayer.removeNode(true)
		
		for(var o in obj){
			if(obj.hasOwnProperty(o)){
				delete obj[o]
			}
		}
		obj = null // <=> SubMenu = null
		
		return true
	}
	catch(e){
		__HLog.error(e,this)
		return false
	}
}

function mnu_hide_layer(menuObjLayer){
	if(typeof(menuObjLayer) == "object" && menuObjLayer != null && menuObjLayer.style.visibility != "hidden"){
		//var id = (menuObjLayer.id).replace(/DOMenuLayer([0-9]+)/ig,"$1")
		menuObjLayer.style.position = "absolute"
		menuObjLayer.style.visibility = "hidden";
		menuObjLayer.style.display = "none"
		menuObjLayer.style.left = "0px";		
		menuObjLayer.style.top = "0px";
		menuObjLayer.style.width = "0px"
		menuObjLayer.style.height = "0px"
		menuObjLayer.initialLeft = 0;
		menuObjLayer.initialTop = 0;
		//mdict.Exists(id) mdict.Remove(id)
	}
}
//var mdict = new ActiveXObject("Scripting.Dictionary")
function mnu_show_layer(menuObjLayer,menuObj){
	try{
		menuObjLayer = menuObjLayer ? menuObjLayer : mnu_add_layer(menuObj)
		var id = (menuObjLayer.id).replace(/DOMenuLayer([0-9]+)/ig,"$1")
		var oFrame = document.all["oFrame" + id]
		menuObjLayer.style.left = new String(menuObj.style.left)
		menuObjLayer.style.top = new String(menuObj.style.top)
		menuObjLayer.style.width = new String(menuObj.offsetWidth)
		menuObjLayer.style.height = new String(menuObj.offsetHeight)
		menuObj.style.zIndex = 100
		menuObjLayer.style.zIndex = 90
		//alert(menuObj.id + " "+id + " "+oFrame)
		if(oFrame.style){			
			oFrame.style.width = new String(menuObj.offsetWidth) - 2
			oFrame.style.height = new String(menuObj.offsetHeight) - 4
		}
		menuObjLayer.style.display = "none"	
		menuObjLayer.style.visibility = "visible";
		menuObjLayer.style.display = "block"
		//var o = {}
		//o.menu = menuObj,o.layer = menuObjLayer,mdict.Add(id,o)
		return menuObjLayer
	}
	catch(e){
		__HLog.error(e,this)
		return null
	}
	finally{
		aMenus.push(menuObjLayer,menuObj)		
	}
}

function mnu_add_layer(menuObj){
	var menuObjLayer = __H.byClone("div");
	var id = (menuObj.id).replace(/DOMenu([0-9]+)/ig,"$1")
	menuObjLayer.innerHTML = '<iframe id="oFrame' + id + '" _ALLOWTRANSPARENCY=true APPLICATION=no style="height:100%;width:100%;" frameborder="0" border="0" src="about:"></iframe>'; // This is a fix to 
	menuObjLayer.id = "DOMenuLayer" + id;
 	HUIToolbarActivate(menuObjLayer,"menuLayer")
	mnu_hide_layer(menuObjLayer)
	document.body.appendChild(menuObjLayer);	
	return menuObjLayer
}

/*
Specifies how the pop-up menu shows/hide.
Arguments:
showValue					: Required. Integer that specifies how the menu shows.
hideValue					: Optional. Integer that specifies how the menu hides. If not specified,the 
										 menu shows/hides in the same manner.

0: Shows/Hides the menu by left click only.
1: Shows/Hides the menu by right click only.
2: Shows/Hides the menu by left or right click.
*/
function mnu_set_activatePopupBy(){
	showValue = typeof(arguments[0]) == "number" && (arguments[0] > -1 ? arguments[0] : 0);
	if(arguments.length > 1){
		hideValue = typeof(arguments[1]) == "number" && arguments[1] > -1 ? arguments[1] : 0;
	}
	else {
		hideValue = showValue;
	}
	if(showValue == 1 || showValue == 2 || hideValue == 1 || hideValue == 2){
		document.oncontextmenu = mnu_set_contextMenu;
	}
}

/*
Hide all menus,except those with alwaysVisible = true.
*/
function mnu_hide_all(){
	try{
		var oDummy = window.setTimeout(function(){mnu_hide_context()},0)
		for(var i = menuCount; i; i--){
			var menuObj = __H.byIds("DOMenu" + i);
			var menuObjLayer = __H.byIds("DOMenuLayer" + i);
			if(!menuObj.alwaysVisible){
				mnu_hide_layer(menuObjLayer)
				if(menuObj.style.position == "fixed"){
					menuObj.style.position == "absolute";
					menuObj.style.visibility = "hidden";
					menuObj.style.position == "fixed";
				}
				else {
					menuObj.style.visibility = "hidden";
					if(menuObj.mode == "cursor"){
						menuObj.style.left = "0px";
						menuObj.style.top = "0px";
						menuObj.initialLeft = 0;
						menuObj.initialTop = 0;
					}
				}
			}
		}	
	}
	catch(e){
		__HLog.error(e,this)
	}
}

function mnu_hide_context(bIgnore){ // Woody
	try{		
		mnu_hide_layer(popUpMenuObj.menuObjLayer)
		popUpMenuObj.menuObj.style.visibility = "hidden";
		mnu_hide_layer(ContextMenuObj.menuObjLayer)
		ContextMenuObj.menuObj.style.visibility = "hidden"
		if(bIgnore) return;
		for(var i = 0, iLen = aMenus.length; i < iLen; i++){
			(aMenus[i]).style.visibility = "hidden";
			delete aMenus[i]
		}
		aMenus.length = 0
	}
	catch(e){
		//__HLog.error(e,this)
	}
}

/*
Hide all menus with mode = "cursor",except those with alwaysVisible = true.
*/
function mnu_hide_cursor(){
	try{
		for(var i = menuCount; i; i--){
			var menuObj = __H.byIds("DOMenu" + i);
			var menuObjLayer = __H.byIds("DOMenuLayer" + i);
			if(!menuObjLayer) continue
			if(menuObj.mode == "cursor" && !menuObj.alwaysVisible){
				mnu_hide_layer(menuObjLayer)
				menuObj.style.visibility = "hidden";
				menuObj.style.left = "0px";
				menuObj.style.top = "0px";
				menuObj.initialLeft = 0;
				menuObj.initialTop = 0;
			}
		}
	}
	catch(e){
		__HLog.error(e,this)
	}
}

/*
Hide all menus with mode = "absolute" or mode = "fixed" or mode = "static",except those with 
alwaysVisible = true.
*/
function mnu_hide_visable(){
	try{
		for(var i = menuCount; i; i--){
			var menuObj = __H.byIds("DOMenu" + i);
			var menuObjLayer = __H.byIds("DOMenuLayer" + i);
			if(!menuObjLayer) continue
			if((menuObj.mode == "absolute" || menuObj.mode == "fixed") && !menuObj.alwaysVisible){
				mnu_hide_layer(menuObjLayer)
				if(menuObj.style.position == "fixed"){
					menuObj.style.position = "absolute";
					menuObj.style.visibility = "hidden";
					menuObj.style.position = "fixed";
				}
				else {				
					menuObj.style.visibility = "hidden";
					menuObj.style.left = "0";
					menuObj.style.top = "0";
					menuObj.initialLeft = 0;
					menuObj.initialTop = 0;
				}
			}
		}
		if(typeof(a_menu_static) == "object"){
			if(typeof(__HMenubar) == "undefined"){
				__HMenubar = __H.UI.Window.Menu.Toolbar.Menubar()
			}
			for(var i = a_menu_static.length-1; i >= 0; i--){
				__HMenubar.refreshItems(__H.byIds(a_menu_static[i]));
			}
		}
	}
	catch(e){
		__HLog.error(e,this)
	}
}

/*
Hide the menu and all its submenus.
Argument:
menuObj						: Required. Menu object that specifies the menu and all its submenus to 
										 be hidden.
*/
function mnu_hide_menus(menuObj,menuObjLayer){
	for(var i = menuObj.childNodes.length-1; i >= 0; i--){
		if(menuObj.childNodes[i].enabled && menuObj.childNodes[i].subMenu){
			mnu_hide_menus(menuObj.childNodes[i].subMenu.menuObj,menuObj.childNodes[i].subMenu.menuObjLayer);			
		}
	}
	mnu_hide_layer(menuObjLayer)
	if(menuObj.style.position == "fixed"){
		menuObj.style.position = "absolute";
		menuObj.style.visibility = "hidden";
	}
	else {
		menuObj.style.visibility = "hidden";
		menuObj.style.left = 0;
		menuObj.style.top = 0;
		menuObj.initialLeft = 0;
		menuObj.initialTop = 0;
	}
}

function mnu_set_popup(menuObj){
	popUpMenuObj = menuObj;
}

function mnu_set_context(menuObj){ // Woody
	ContextMenuObj = menuObj;
}

function mnu_hide_menubars(){    // Public method
	if(typeof(__HMenubar) == "undefined"){
		__HMenubar = __H.UI.Window.Widgets.Menu.Toolbar.Menubar()
	}
	__HMenubar.refreshItemsAll()
} 

