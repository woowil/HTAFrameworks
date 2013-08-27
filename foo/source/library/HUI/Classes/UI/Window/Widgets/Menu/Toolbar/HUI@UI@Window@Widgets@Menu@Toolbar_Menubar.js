// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**

/*
	This dynamic toolbar menu was originally created by Toh Zhiqiang called this.addMenubar Version 1.1.1
	But as you can see, comparing my version, the menubar is totally revised
	
	This file depends on 
*/

__H.include("HUI@UI@Window@Widgets@Menu@Toolbar.js")

__H.register(__H.UI.Window.Widgets.Menu.Toolbar,"Menubar","Menu bar",function Menubar(){
	var this_top = this
	var px = "px"
	var pt = "paddingTop"
	var pb = "paddingBottom"
	var btw = "borderTopWidth"
	var bbw = "borderBottomWidth"
	var blw = "borderLeftWidth"
	var brw = "borderRightWidth"
	
	this.count_menu = 0
	this.count_item = 0
	
	var o_dragging = null
	var a_menu_static = []	
	
	var getIntValue = function getIntValue(obj,value){
		return parseInt(obj.style[value])
	}
	
	var getLeftPos = function getLeftPos(menuBarObj,menuBarItemObj,menuObj,x){
		try{
			if(x + menuObj.offsetWidth <= __HWindow.getClientWidth()){
				return x;
			}
			else {
				return x + menuBarItemObj.offsetWidth - menuBarObj.offsetWidth + mnu_get_propIntVal(menuObj, blw) + mnu_get_propIntVal(menuObj, brw);
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return 0
		}
	}
		
	var getTopPos = function getTopPos(menuBarObj,menuBarItemObj,menuObj,y){
		try{
			if(y + menuObj.offsetHeight <= __HWindow.getClientHeight()){
				return y + 1;
			}
			else {
				if(__HWindow.bIEnew && menuBarObj.mode == "static" && __HWindow.pagemode == 0){
					y = menuBarObj.offsetTop + menuBarObj.offsetHeight - __HWindow.getScrollTop();
				}
				if(__HWindow.bIEnew && menuBarObj.mode == "static" && __HWindow.pagemode == 1){
					return menuBarItemObj.offsetTop - menuObj.offsetHeight - mnu_get_propIntVal(menuBarObj, pt) + mnu_get_propIntVal(menuBarObj, pt) + mnu_get_propIntVal(menuBarObj, btw) - __HWindow.getScrollTop();
				}
				else{
					return y - menuObj.offsetHeight - menuBarObj.offsetHeight;
				}
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return 0
		}
	}
	
	var setPopup = function setPopup(menuBarObj,menuBarItemObj,menuObj,menuObjLayer){
		try{
			var x, y;
			if(menuBarObj.style.position == "fixed"){
				x = menuBarObj.offsetLeft + menuBarItemObj.offsetLeft + mnu_get_propIntVal(menuBarObj,blw) - mnu_get_propIntVal(menuObj,blw);
				y = menuBarObj.offsetTop + menuBarObj.offsetHeight;
				
				menuObj.style.position = "absolute";
				menuObj.style.left = getLeftPos(menuBarObj,menuBarItemObj,menuObj,x) + px;
				menuObj.style.top = getTopPos(menuBarObj,menuBarItemObj,menuObj,y) + px;
				menuObj.style.position = "fixed";
			}
			else{
				if(menuBarObj.mode == "static"){
					x = menuBarItemObj.offsetLeft - mnu_get_propIntVal(menuObj,blw) - __HWindow.getScrollLeft();
					y = menuBarObj.offsetTop + menuBarObj.offsetHeight - __HWindow.getScrollTop();					
					if(__HWindow.bIE55 || __HWindow.bIE6){ // DO NOT CHANGE!! IE7 have problem
						x += mnu_get_propIntVal(menuBarObj,blw);
						y = menuBarItemObj.offsetTop + menuBarItemObj.offsetHeight + mnu_get_propIntVal(menuBarObj,bbw) + mnu_get_propIntVal(menuBarObj,pb) - mnu_get_propIntVal(menuBarObj,bbw) - __HWindow.getScrollTop();
					}
					menuObj.style.left = (getLeftPos(menuBarObj,menuBarItemObj,menuObj,x) + __HWindow.getScrollLeft()) + px;
					menuObj.style.top = (getTopPos(menuBarObj,menuBarItemObj,menuObj,y) + __HWindow.getScrollTop()) + px;
				}
				else{
					x = menuBarObj.offsetLeft + menuBarObj.offsetLeft + mnu_get_propIntVal(menuBarObj,blw) - mnu_get_propIntVal(menuObj,blw) - __HWindow.getScrollLeft();
					y = menuBarObj.offsetTop + menuBarObj.offsetHeight - __HWindow.getScrollTop();
					menuObj.style.left = (getLeftPos(menuBarObj,menuBarItemObj,menuObj,x) + __HWindow.getScrollLeft()) + px;
					menuObj.style.top = (getTopPos(menuBarObj,menuBarItemObj,menuObj,y) + __HWindow.getScrollTop()) + px;
				}
			}
			if(__HWindow.bIEnew && menuObj.mode == "fixed"){
				menuObj.initialLeft = parseInt(menuObj.style.left) - __HWindow.getScrollLeft();
				menuObj.initialTop = parseInt(menuObj.style.top) - __HWindow.getScrollTop();
			}
			mnu_show_layer(menuObjLayer,menuObj);
			menuObj.style.visibility = "visible";
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.refreshItems = function refreshItems(menuBarObj){
		for(var i = 0, iLen = menuBarObj.childNodes.length; i < iLen; i++){
			if(menuBarObj.childNodes[i].enabled && menuBarObj.childNodes[i].clicked){
				menuBarObj.childNodes[i].clicked = false;
				if(menuBarObj.childNodes[i].menu){
					mnu_hide_menus(menuBarObj.childNodes[i].menu.menuObj,menuBarObj.childNodes[i].menu.menuObjLayer);
				}
				break;
			}
		}
		menuBarObj.activated = false;
	}
	
	this.refreshItemsAll = function refreshItemsAll(){
		 for(var i = 1; i <= this.menuBarCount; i++){
			this.refreshItems(__H.byIds("DOMenuBar" + i));
		}
	}
	
	var setItemOut = function setItemOut(){
		if(!this.parent.menuBarObj.activated){
			HUIToolbarActivate(this,'menuBarItemOut')
		}
	}
	
	var setItemOver = function setItemOver(e){
		if(this.parent.menuBarObj.activated){
			if(!this.clicked){				
				var menuBarObj = this.parent.menuBarObj;
				for(var i = 0, iLen = menuBarObj.childNodes.length; i < iLen; i++){
					if(menuBarObj.childNodes[i].enabled && menuBarObj.childNodes[i].clicked){
						menuBarObj.childNodes[i].clicked = false;
						if(menuBarObj.childNodes[i].menu){
							mnu_hide_menus(menuBarObj.childNodes[i].menu.menuObj,menuBarObj.childNodes[i].menu.menuObjLayer);
						}
						break;
					}
				}
				if(this.enabled){
					if(this.menu) this.onclick(e);
					else{
						if(this.actionOnClick){
							HUIToolbarActivate(this,"menuBarItemOver")
							this.clicked = true;
						}
					}
				}
			}
		}
		else{
			var ID = window.setTimeout(function(){mnu_hide_context()},0);
			if(this.enabled && (this.menu || this.actionOnClick)){
				HUIToolbarActivate(this,'menuBarItemOver')
			}
		}
	}
	
	var setItemClick = function setItemClick(e){
		try{
			if(this.enabled){
				if(this.menu){
					if(this.clicked){
						HUIToolbarActivate(this,'menuBarItemOver')
						this.clicked = false;
						this.parent.menuBarObj.activated = false;
					}
					else{
						HUIToolbarActivate(this,'menuBarItemClick')
						setPopup(this.parent.menuBarObj,this,this.menu.menuObj,this.menu.menuObjLayer);
						this.clicked = true;
						this.parent.menuBarObj.activated = true;
					}
				}
				else{
					if(this.actionOnClick){
						var action = this.actionOnClick;
						if(action.indexOf("link:") == 0){
							location.href = action.substr(5);
						}
						else{
							if(action.indexOf("code:") == 0){
								eval(action.substr(5));
							}
							else{
								location.href = action;
							}
						}
						HUIToolbarActivate(this,'menuBarItemOut')
						this.clicked = false;
						this.parent.menuBarObj.activated = false;
					}
				}
			}
			HUIToolbarCancel(e)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	var setDown = function setDown(e){
		o_dragging = this.parent.menuBarObj;
		var menuBarObj = this.parent.menuBarObj;
		menuBarObj.differenceLeft = __HWindow.getClientX(e) - menuBarObj.offsetLeft;
		menuBarObj.differenceTop = __HWindow.getClientY(e) - menuBarObj.offsetTop;
		hideMenubars();
		document.onmousemove = function(e){
			if(o_dragging){
				o_dragging.style.left = (__HWindow.getClientX(e) - o_dragging.differenceLeft) + px;
				o_dragging.style.top = (__HWindow.getClientY(e) - o_dragging.differenceTop) + px;
			}
		}
	}
	
	var setUp = function setUp(){
		o_dragging = null;
		var menuBarObj = this.parent.menuBarObj;
		menuBarObj.differenceLeft = 0;
		menuBarObj.differenceTop = 0;
		menuBarObj.initialLeft = menuBarObj.offsetLeft - __HWindow.getScrollLeft();
		menuBarObj.initialTop = menuBarObj.offsetTop - __HWindow.getScrollTop();
		document.onmousemove = null;
	}
	
	var hideMenubars = function hideMenubars(){
		
		for(var i = 1; i <= this_top.count_menu; i++){
			this.refreshItems(__H.byIds("DOMenuBar" + i));
		}
	}
	
	/*
	Arguments:
	displayText		: Required. String that specifies the text to be displayed on the menu bar item.
	menuObj			: Optional. Menu object that is going to be the main menu for the menu bar item. Defaults to null(no menu).
	itemName		: Optional. String that specifies the name of the menu bar item. Defaults to "" (no name).
	enabled			: Optional. Boolean that specifies whether the menu bar item is enabled/disabled. Defaults to true.
	actionOnClick	: Optional. String that specifies the action to be done when the menu item is
					  being clicked if no main menu has been set for the menu bar item. Defaults to "" (no action).
	*/
	var addItem = function addItem(){
		try{
			var sDisplayText = arguments[0];
			var sItemName = "";
			var bEnabled = true;
			var sActionOnClick = "";			
			var len = arguments.length;
			var oMenu = null
			
			if(len > 1 && typeof(arguments[1]) == "object"){
				oMenu = arguments[1];
			}
			if(len > 2 && arguments[2].length > 0){
				sItemName = arguments[2];
			}
			if(len > 3 && typeof(arguments[3]) == "boolean"){
				bEnabled = arguments[3];
			}
			if(len > 4 && arguments[4].length > 0){
				sActionOnClick = arguments[4];
			}
			
			var element = __H.byClone("span");
			element.id = "menuBarItem" +(++this.count_item);
			element.menu = oMenu
			element.enabled = bEnabled
			element.clicked = false;
			element.actionOnClick = sActionOnClick
			
			var textNode = document.createTextNode(sDisplayText);
			element.appendChild(textNode);
			this.menuBarObj.appendChild(element);
			element.parent = this;
			
			element.setClassName = function(className){
				this.className = className
			};
			element.setDisplayText = function(text){
				if(this.childNodes[0].nodeType == 3){
					this.childNodes[0].nodeValue = text;
				}
				else{
					this.childNodes[1].nodeValue = text;
				}
			};
			element.setMenu = function(menu){
				this.menu = menu;
			};
			
			element.onmouseover = setItemOver;
			element.onclick = setItemClick;
			element.onmouseout = setItemOut;
			element.onmouseout()
			
			if(sItemName.length > 0){
				this.items[sItemName] = element;
			}
			else{
				this.items[this.items.length] = element;
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	/*
	Arguments:
	mode		: Optional. String that specifies the mode of the menu bar. Defaults to "absolute".
	id			: Optional,except when mode = "static". String that specifies the id of
				  the element that will contain the menu bar. This argument is required when mode = "static"
	draggable	: Optional. Boolean that specifies whether the menu bar is draggable. Defaults to false.
	className	: Optional. String that specifies the CSS class selector for the menu bar.
	width		: Optional. Integer that specifies the width of the menu bar. Defaults to "auto".
	height		: Optional. Integer that specifies the height of the menu bar. Defaults to "auto".
	*/
	
	this.addMenubar = function addMenubar(){
		try{
			this.items = [];
			var element_drag = __H.byClone("span");			
			var textNode = document.createTextNode("");
			element_drag.appendChild(textNode);
			
			var element;
			var len = arguments.length;
			if(len > 1 && arguments[1].length > 0 && arguments[0] == "static"){
				if(!(element = __H.byIds(arguments[1]))) {
					__HLog.logPopup("The menubar tag " + arguments[1] + " was not found!")
					return this;
				}
				a_menu_static[a_menu_static.length] = arguments[1];
				element.appendChild(element_drag);
			}
			else{
				element = __H.byClone("div");
				element.appendChild(element_drag);
				element.id = "DOMenuBar" +(++this.count_menu);
			}
			
			element.mode = "absolute";
			element.activateMode = "click";
			element.draggable = false;			
			element.activated = false;
			element.initialLeft = 0;
			element.initialTop = 0;
			element.differenceLeft = 0;
			element.differenceTop = 0;			
		 	element.onselectstart = new Function("return false")
			
			HUIToolbarActivate(element,"menuBarItem")
			HUIToolbarActivate(element_drag,"menuBarItemDrag")
			
			if(len > 0 && arguments[0].length > 0){
				switch(arguments[0]){
					case "absolute" :
						element.style.position = "absolute";
						element.mode = "absolute";
						break;
					case "fixed" :
						element.style.position = __HWindow.bIE ? "absolute" : "fixed"
						element.mode = "fixed";
						break;
					case "static" : default :
						element.style.position = "static";
						element.mode = "static";
						break;
				}
			}
			if(len > 2 && typeof(arguments[2]) == "boolean"){
				element.draggable = arguments[2];
				if(element.draggable) element_drag.style.visibility = "visible";
				else element_drag.style.visibility = "hidden";
			}
			if(len > 3 && arguments[3].length > 0){
				element.className = arguments[3];
			}
			if(len > 4 && typeof(arguments[4]) == "number" && arguments[4] > 0){
				element.style.width = arguments[4] + px;
			}
			else if(len > 4 && arguments[4] == null && arguments[4] > 0){
				element.style.width = "100%";
			}
			if(len > 5 && typeof(arguments[5]) == "number" && arguments[5] > 0){
				element.style.height = arguments[5] + px;
			}
			
			if(element.mode != "static") document.body.appendChild(element);
			else{
				element.style.height = "1%"
			}
			
			this.menuBarObj = element;
			this.menuBarObj.onclick = HUIToolbarCancel;
			this.menuBarObjDrag = element_drag;
			this.addItem = addItem;
			
			element_drag.parent = this;
			element_drag.onmousedown = setDown;
			element_drag.onmouseup = setUp;			
			
			this.setMode = function setMode(mode){
				switch(mode){
					case "absolute":
						this.menuBarObj.style.position = "absolute";
						this.menuBarObj.mode = "absolute";
						this.menuBarObj.initialLeft = parseInt(this.menuBarObj.style.left);
						this.menuBarObj.initialTop = parseInt(this.menuBarObj.style.top);
						break;
					case "fixed":
						if(__HWindow.bIE){
							this.menuBarObj.style.position = "absolute";
							this.menuBarObj.initialLeft = parseInt(this.menuBarObj.style.left);
							this.menuBarObj.initialTop = parseInt(this.menuBarObj.style.top);
						}
						else{
							this.menuBarObj.style.position = "fixed";
						}
						this.menuBarObj.mode = "fixed";
						break;
				}
			};
			this.setActivateMode = function setActivateMode(activateMode){
				this.menuBarObj.activateMode = activateMode;
			};
			this.setDraggable = function setDraggable(draggable){
				if(typeof(draggable) == "boolean" && this.menuBarObj.mode != "static"){
					this.menuBarObj.draggable = draggable;
					if(this.menuBarObj.draggable){
						this.menuBarObjDrag.style.visibility = "visible";
					}
					else{
						this.menuBarObjDrag.style.visibility = "hidden";
					}
				}
			};
			this.setClassName = function setClassName(className){
				this.menuBarObj.className = className;
			};
			this.setDragClassName = function setDragClassName(className){
				this.menuBarObjDrag.className = className;
			};
			this.show = function show(){
				this.menuBarObj.style.visibility = "visible";
			};
			this.hide = function hide(){
				this.menuBarObj.style.visibility = "hidden";
			};
			this.setX = function setX(x){
				this.menuBarObj.initialLeft = x;
				this.menuBarObj.style.left = x + px;
			};
			this.setY = function setY(y){
				this.menuBarObj.initialTop = y;
				this.menuBarObj.style.top = y + px;
			};
			this.moveTo = function moveTo(x,y){
				this.menuBarObj.initialLeft = x;
				this.menuBarObj.initialTop = y;
				this.menuBarObj.style.left = x + px;
				this.menuBarObj.style.top = y + px;
			};
			this.moveBy = function moveBy(x,y){
				var left = parseInt(this.menuBarObj.style.left);
				var top = parseInt(this.menuBarObj.style.top);
				this.menuBarObj.initialLeft = left + x;
				this.menuBarObj.initialTop = top + y;
				this.menuBarObj.style.left = (left + x) + px;
				this.menuBarObj.style.top = (top + y) + px;
			};
			this.setBorderWidth = function setBorderWidth(width){
				this.menuBarObj.style.borderWidth = width + px;
			};
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
})

function HUIToolbarActivate(oElement,sOpt){
	try{
		if(typeof(oElement) != "object" || typeof(sOpt) != "string") return;		
		__MStyle[sOpt](oElement)
	}
	catch(e){
		//alert(sOpt + " "+e.description)		
	}
}

function HUIToolbarCancel(e){
	try{
		if(!e){
			var e = window.event;
			e.cancelBubble = true;
		}
		if(e.stopPropagation){
			e.stopPropagation();
		}
	}
	catch(ee){}
}

var __HMenubar = new __H.UI.Window.Widgets.Menu.Toolbar.Menubar()

