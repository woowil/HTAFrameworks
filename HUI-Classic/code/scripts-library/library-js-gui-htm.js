// nOsliw Solutions - HTML Aplication Framework.
//**Start Encode**

/* 

File:     library-js-gui-htm.js
Purpose:  Development script for Dynamic HTML
Author:   Woody Wilson
Created:  2002-18-01
Version:  see LIB_VERSION

Description:
JScript library scripting functions used for WSH files or HTA Applications.
 
Usage:
Need library-js.js

Description:
JScript library scripting functions used for WSH files or HTA Applications.

Revisions: to many

Disclaimer:
This sample code is provided AS IS WITHOUT WARRANTY
OF ANY KIND AND IS PROVIDED WITHOUT ANY IMPLIED WARRANTY
OF FITNESS FOR PURPOSES OF MERCHANTABILITY. Use this code
is to be undertaken entirely at your risk, and the
results that may be obtained from it are dependent on the user.
Please note to fully back up files and system(s) on a regular
basis. A failure to do so can result in loss of data or damage
to systems.		


Woody's 10 bud när du kodar en sajt i dynamisk html:
 1. Du skall koda så att sajten blir så generaliserad och dynamisk som möjligt. Helst utgå från en databas, "en" javascript fil och "en" style-sheet fil
 2. Du skall utgå från en struktur som underlättar administrativ arbete, användarvänlig utformning och utvecklaranpassad filosofi.
 3. Du skall skapa generaliserade funktioner och procedurer så att de är utbyggbara, dynamiska, smarta, genomtänkta och debugutformade.
 4. Du skall ha en genomtänkt och enhetlig namnkonvention, utformad gränssnitt i typsnitt, struktur och kvalitet balanserad med bilder och färger.
 5. Du skall ha en uppfattning om vem du riktar sajten mot och hur den skall användas. Inte för teknisk, inte för lite innehåll och så bra språk som möjligt.
 6. Du skall anpassa sajten till de olika webbläsare så långt som bara möjligt (om krav ställs, annars utgå från IE5.5+ och Netscape7+)
 7. Du skall känna till juridiska webblagar och copyright. Du skall även känna till de "oskrivna" men samtidigt "kända" webbreglerna som t.ex. antal färger, animation, banners, frames etc. 
 8. Du skall utgå från att du ska lämna över sajten till en mindre kvalificerad utvecklare eller ett företag. Glöm inte dokumentation och mallar.
 9. Du skall använda dig av de bästa och "oberoende" webbverktygen som finns. T.ex Macromedia HomeSite5, DreamWeaver MX. PaintShop Pro, TopStyle Pro 
10. Du skall tänka på att det är en sak att koda en sajt men något helt annat att sköta det. Ligg ett steg före hela tiden.
 
*/



var LIB_NAME    = "Library HTM";
var LIB_VERSION = "3.0";
var LIB_FILE    = "library-js-htm.js"

var htm = new htm_object()

function htm_object(){
	try{
		this.version = new js_objectversion(LIB_NAME,LIB_FILE,LIB_VERSION)
	}
	catch(ee){}
	this.id		= false;
	this.img		= new Array();
	try{ // only for IE DOM
		if(window);
		this.br		= new htm_objectbrowser(); // Important that "new" is used
	}
	catch(ee){
		this.br = null;
	}
	this.menu	= new Object();
}


////////////////////////////////////////////////////////////////////////////////
///////// WINDOW FUNCTIONS

function htm_init(){
	try{
		if(htm.br.dyn && !window.saveInnerWidth) {
			if(htm.br.ie){
				window.onresize = htm_div_resize;
				window.saveInnerWidth = window.innerWidth;
				window.saveInnerHeight = window.innerHeight;
			}
			else if(htm.br.nsg){
				
			}
			else if(htm.br.nsc){
				window.onresize = htm_div_resize;
				window.saveInnerWidth = document.body.offsetWidth;
				window.saveInnerHeight = document.body.offsetHeight;
			}
		}
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{
		return true;
	}
}

function htm_nav_menu(iOpt,oMenu,arg3,arg4,arg5,arg6,arg7){
	// <div class="dvIndexMenu" title="Home" OnMouseOver="htm_nav_menu(1,this,1);" OnMouseOut="htm_nav_menu(2,this,1)" OnMouseDown="htm_nav_menu(3,this,'documents/home.html',1)">
	try{
		var bResult = true;
		switch(iOpt){
			case 1 : { // On mouse over
				oMenu.style.cursor = htm.br.nsg ? "pointer" : "hand";
				if(htm.menu.idx != arg3) { // if the div is not activated
					htm_div_color(oMenu,arg7,arg6,arg5,arg4);
				}
				break;
			}
			case 2 : {// On mouse out
				oMenu.style.cursor = "default";
				if(htm.menu.idx == arg3){
					htm_nav_menu(1,oMenu,-1);
					break;
				}
				else htm_div_color(oMenu,arg7,arg6,arg5,arg4);
				break;
			}
			case 3 : { // On mouse down
				if(htm.menu.oMenu) htm_nav_menu(2,htm.menu.oMenu,-1);
				oMenu.style.backgroundColor = arg6;
				self.status = location.href + "?doc=" + arg3;
				htm_nav_url(3,arg3,arg4); // arg3: is the url, arg4: is the window name
				htm.menu.idx = arg5;
				htm.menu.oMenu = oMenu;			
				break;
			}
			default : {
				bResult = false;
				break;
			}
		}
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{
		return bResult;
	}
}

function htm_nav_url(iOpt,url,arg3,arg4){
	try{
		htm.br.exec = true;
		switch(iOpt){
			case 1 : { // Opens new Window
				if(url) window.open(url);
				else htm_log_error(3,"Error opening new window. The url argument is invalid.");
				break;			
			}
			case 2 : { // Opens self Window
				if(url) document.location = url;
				else htm_log_error(3,"Error opening self window. The url argument is invalid.");
				break;
			}
			case 3 : { // Opens an Iframe or known window
				if(url && arg3){ // arg3: iframe id or name
					if(htm.br.ie){
						if(document.all[arg3]) document.all[arg3].src = url;
						else frame[arg3].location.href = url;
					}
					else if(htm.br.nsg || htm.br.w3c) document.getElementById(arg3).src = url;
					else htm_log_error(3,"Sorry your browser does not support this site!");
				}
				else htm_log_error(3,"Error opening url in window. The url argument is invalid.");
				break;
			}
			case 4 : { // Opens a search Window
				if(url && arg4){ // arg4: The html document 
					top.location.href = url + "?doc=" + arg4; // url: http://www.ncr.no/psitis/  page: /document/link.html
				}
				else htm_log_error(3,"Error opening window. The url or search-page argument are invalid.");
				break;
			}
			case 5 : { // Opens a declared window
				if(url && arg3){ // arg3: The window title
					variOptions = arg4 ? arg4 : 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=500,height=500';
					var win = window.open(url,arg3,options);
					win.focus();
				}
				else htm_log_error(3,"Error opening window. The url or title argument are invalid.");
				break;
			}
			case 6 : { // Navigates to an anchor (hash) link in document
				if(url && arg3){ // arg3: The hash link
					hash = arg3 ? '#' + arg3 : '';
					frames[arg4].document.location = url + hash; // arg4: The window name
				}
				else htm_log_error(3,"Error opening window. The url, title or window argument are invalid.");				
				break;
			}
			case 7 : {
				
				break;
			}
			default : {
				htm.br.exec = false;
				htm_log_error(3,"Error calling function htm_nav_url() with argumenti.");
				break;
			}
		}
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{
		return htm.br.exec;
	}
}

function htm_nav_search(win,iOpt){
	try{
		htm.br.exec = true;
		if(str = win.location.search){
			switch(iOpt){
				case 1 : {
					var Re = /page=([a-zA-Z0-9_\\\/\s\.]+)&*/ig;
					if(ary = Re.exec(str)) return ary[1]; // Returns document/main.html from ?page=document/main.html 
					break;
				}
				case 2 : {
					var Re = /page=([a-zA-Z0-9_\\\/\s\.]+)&id=([0-9]{1,2}).*/ig; // %26 is ekvivalent with &
					if(ary = Re.exec(str)){
						return ary[2]; // Returns 3 from ?page=document/consultants.html&id=3 
					}
					break;
				}
				case 3 : {
					
					break;
				}
				default : {
					htm.br.exec = false;
					htm_log_error(3,"Error calling function htm_nav_search() with argumentiOpt.");
					break;
				}
			}
		}
		return htm.br.exec = false;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{
		return htm.br.exec;
	}
}

function htm_nav_remframe(){
	if(self.parent.frames.length != 0){ // Remove parent frames
		self.parent.location = document.location;
	}
}


////////////////////////////////////////////////////////////////////////////////
///////// OBJECT FUNCTIONS

function htm_objectbrowser(){
	try{
		if(!window) return false;
		
		var b = navigator.appName;
		if (b == "Netscape") this.b = "ns";
		else if (b == "Microsoft Internet Explorer") this.b = "ie";
		else this.b = b;
		
		this.version = navigator.appVersion;
		this.agent = navigator.userAgent
		this.v = parseFloat(this.version.substring(0,1));
		this.w3c = (document.getElementById) ? true : false; // Opera6+, Nescape6+ and Internet Explore4+ are identified as w3c 
		this.mac = (navigator.userAgent.indexOf("Mac") != -1);
		
		this.ns3 = (this.b == "ns" && this.v < 4);
		this.ns = (this.b == "ns" && this.v >= 4);
		this.nsc = (this.b == "ns" && this.v == 4);
		this.ns6 = (navigator.vendor == "Netscape6" && this.nsc);
		this.ns7 = (this.w3c && navigator.appName.indexOf("Netscape") >= 0) ? true : false;
		this.nsg = (this.ns6 || this.ns7);
		
		this.ie3 = (this.b == "ie" && this.v <= 3.04);
		this.ie = (this.b == "ie" && this.v >= 4);
		this.ie4 = (this.version.indexOf('MSIE 4') > 0 && this.w3c);
		this.ie5 = (this.version.indexOf('MSIE 5') > 0 && this.w3c);
		this.ie55 = (this.version.indexOf('MSIE 5.5') > 0 && this.w3c);
		this.ie6 = (this.version.indexOf('MSIE 6') > 0 && this.w3c);
		this.opera = this.agent.indexOf("Opera") > -1;
		this.safari = this.agent.indexOf("Safari") > -1;
		
		this.min = (this.ns || this.ie || this.w3c);
		this.img = (document.images);
		this.je = navigator.javaEnabled();
		this.dyn = (this.ie || this.nsc);
		
		this.mse = new Object();
		this.mse.enable = false;
		
		this.cke = new Object();
		this.cke.scrollmenu = "ScrollMenu";
		this.cke.beenhere = "BeenHere";
		this.cke.ns6 = "Netscape6";
		this.cke.drmouse = "DisableMouse";
		
		this.exec = true;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}


function htm_win_screen(iOpt,iWidth,iHeight,iPosX,iPosY){
	try{
		if(iOpt == 1){ // Full screen
			window.moveTo(0,0);
			window.resizeTo(window.screen.width,window.screen.height)
		}
		else if(iOpt == 2){ // restore
			if(iWidth && iHeight && iPosX && iPosY){
				// iPosX: window.screenLeft, iPosY: window.screenTop
				window.moveTo(iPosX,iPosY);
				// iHeight: window.screen.height*0.8
				window.resizeTo(iWidth,iHeight);
			}
			else return false;
		}
		return true;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}	
}


////////////////////////////////////////////////////////////////////////////////
///////// MENU FUNCTIONS




////////////////////////////////////////////////////////////////////////////////
///////// DIV FUNCTIONS

function htm_div_init(aDiv){
	try{
		if(htm.br.ie){ // Internet Explorer (all versions)
			for(var i = 0; i < aDiv.length; i++){
				if(oDiv = document.all[aDiv[i]]){
					aDiv[i].css 	= oDiv.style;
					aDiv[i].doc 	= oDiv;
					aDiv[i].LOAD	= false;
				}
			}
		}
		else if(htm.br.nsg){ // Netscape Navigator 6 and 7
			for(var i = 0; i < aDiv.length; i++){
				var sDivName = ""; // not implemented
				if(oDiv = document.getElementById(sDivName)[aDiv[i]]){
					aDiv[i].css 	= oDiv.style;
					aDiv[i].doc 	= oDiv;
					aDiv[i].LOAD	= false;
				}
			}
		}
		else if(htm.br.nsc){ // Netscape Communicator 4
			for(var i = 70; i < aDiv.length; i++){
				if(oDiv = document.layers[aDiv[i]]){
					aDiv[i].css 	= oDiv;
					aDiv[i].doc 	= oDiv.document;
					aDiv[i].LOAD	= false;
				}
			}
		}
		else return false;
		htm_mse_init(); // Initializes mouse click on site 
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{
		return true;
	}
}

function htm_div_resize(){
	try{
		if(htm.br.ie){
			if(document.body.offsetWidth != window.saveInnerWidth || document.body.offsetHeight != window.saveInnerHeight){
				htm_div_init();
			}
		}
		else if(htm.br.nsg){
		
		}
		else if(htm.br.nsc){
		    if(saveInnerWidth < window.innerWidth || saveInnerWidth > window.innerWidth || saveInnerHeight > window.innerHeight || saveInnerHeight < window.innerHeight ){
			    window.history.go(0);
		    }
		}
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{
		return true;
	}
}

function htm_div_center(oDiv){
   oDiv.style.setExpression("left","document.body.clientWidth/2 - oDiv.offsetWidth/2");
   oDiv.style.setExpression("top","document.body.clientHeight/2 - oDiv.offsetHeight/2");
   document.recalc(true);
}

function htm_div_isvisible(obj){
	try{
		if(typeof(obj) != "object") return false;
		if(obj.style){
			if(htm.br.nsc) sVisibility = 'show';
			else sVisibility = 'visible';
			//alert(sVisibility + " " + obj.style.visibility + " " + obj.style.display) 
			if(obj.style.visibility == sVisibility || obj.style.display == "block"){				
				return true
			}
		}
		return false;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_div_show(obj){
	try{		
		if(typeof(obj) != "object") return false;
		else if(obj.style){
			obj.style.display = "none";
			obj.style.display = "block";
			obj.style.visibility = "visible";
			return true
		}
		return false
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_div_showrec(){
	try{
		var bResult = true
		for(var i = 0, len = arguments.length; i < len; i++){
			if(!htm_div_show(arguments[i])) bResult = false
		}
		return bResult;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_div_ishidden(obj){
	try{
		if(typeof(obj) != "object") return false;
		else if(obj.style){
			if(htm.br.nsc) sVisibility = 'hide';
			else sVisibility = 'hidden';
			if(obj.style.visibility == sVisibility || obj.style.display == "none" || obj.style.display == undefined || obj.style.display == null){
				return true
			}			
		}
		return false;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_div_hide(obj){
	try{
		if(typeof(obj) != "object") return false;
		else if(obj.style){
			obj.style.visibility = "hidden"
			obj.style.display = "none";
		}
		return true;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_div_hiderec(ostream){
	try{
		bResult = true
		for(var i = 0, len = arguments.length; i < len; i++){
			if(!htm_div_hide(arguments[i])) bResult = false
		}
		return bResult;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_div_hideimg(oDiv,oImg,sSrc,bHideDiv){
	try{
		bHideDiv = bHideDiv ? true : false
		if(typeof(oDiv) != "object" || typeof(oImg) != "object") return false;
		else if(oDiv.style && oImg.tagName == "IMG"){
			//oImg.style.backgroundImage = null
			//oImg.style.backgroundImage = 'url(' + sSrcUrl + ')';
			if(bHideDiv) htm_div_hide(oDiv)
			oImg.src = sSrc	
			return true;
		}
		return false;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_div_showimg(oDiv,oImg,sSrc,bShowDiv){
	try{
		bShowDiv = bShowDiv ? true : false
		if(typeof(oDiv) != "object" || typeof(oImg) != "object") return false;
		else if(oDiv.style && oImg.tagName == "IMG"){
			//oImg.style.backgroundImage = null
			//oImg.style.backgroundImage = 'url(' + sSrcUrl + ')';
			if(bShowDiv) htm_div_show(oDiv)
			oImg.src = sSrc	
			return true;
		}
		return false;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_div_toggleimg(oDiv,oImg,sSrcVisible,sSrcHidden,bToggleDiv){
	try{
		bToggleDiv = bToggleDiv ? true : false
		if(typeof(oDiv) != "object" || typeof(oImg) != "object") return false;
		else if(oDiv.style && oImg.tagName == "IMG"){
			//oImg.style.backgroundImage = null
			if(htm_div_isvisible(oDiv)) htm_div_hideimg(oDiv,oImg,sSrcHidden,bToggleDiv)
			else htm_div_showimg(oDiv,oImg,sSrcVisible,bToggleDiv)
			//oImg.style.backgroundImage = 'url(' + sSrcUrl + ')';
			return true;
		}
		return false;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_div_toggle(iOpt,obj,arg3,arg4){
	try{
		var bResult = true;
		switch(iOpt){
			case 1 : { // Block Display switch
				if(obj.style){
					if(htm_div_ishidden(obj)) htm_div_show(obj);
					else htm_div_hide(obj);
				}
				break;
			}
			case 2 : { // Block Display switch, including dublicates
				idx = !isNaN(arg3) ? arg3 : 0
				for(var i = idx; i < obj.length; i++){
					oTag = obj[i]
					htm_div_toggle(1,oTag)
				}
				break;
			}
			case 3 : { // Image switch
				var oRe1 = new RegExp(arg3,"i");
				var oRe2 = new RegExp(".+" + arg3 + ".*","i");
				if(obj.src) obj.src = (obj.src.match(oRe1)) ? arg4 : arg3;
				else {
					if(( obj.style.backgroundImage ).match(oRe2)) obj.style.backgroundImage = 'url(' + arg4 + ')';
					else obj.style.backgroundImage = 'url(' + arg3 + ')';
				}
				break;
			}
			default : {
				bResult = false;
				break;
			}
		}
		return bResult;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_div_getelement(elem){
	// test which method is supported by the browser
	if(document.getElementById)
		return document.getElementById(elem);
	else if(document.all)
		return document.all[elem];
	return null;
}

function htm_div_write(obj,html,nest) {
	if(obj && html){
		if(htm.br.ie || htm.br.nsg)	obj.innerHTML = html;
		else if(htm.br.nsc){
			obj.open();
			obj.writeln(html);
			obj.close();
		}
		else;
	}
}

function htm_div_zoom(obj,init,step,stop,delay,ID,type){
	try{
		if(js_str_isnumber(val = parseFloat(init))){
			if((type == 'in' && val < stop) || (type == 'out' && val > (stop-step))) {
				a6 = ID, htm.id = setTimeout("htm_div_zoom(null,null,null,null,null,a6,null)",delay);
			}
			else{
				a1 = obj, a2 = (val+step), a3 = step, a4 = stop, a5 = delay, a6 = ID, a7 = type;
				if(htm.br.ie || htm.br.nsg){ obj.style.zoom = a2;}
				htm.id = setTimeout("htm_div_zoom(a1,a2,a3,a4,a5,a6,a7)",delay);
			}
		}
		else htm_tme_clear(ID);
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{
		return true;
	}
}

function htm_div_transition(obj,trans,dur){
	obj.css = obj.css ? obj.css : obj.style;
	obj.doc = obj.doc ? obj.doc : obj.document;
	if(obj.css && obj.doc){
		if(htm.br.ie || htm.br.nsg){
			htm_div_hide(obj);
			obj.doc.filters.item(0).apply();
			obj.doc.filters.item(0).duration = (dur ? dur : '3.0');
			obj.doc.filters.item(0).transition = (trans ? trans : '23');
			htm_div_show(obj);
			obj.doc.filters(0).play(1.000);
		}
	}
}

function htm_div_color(obj,txt,style,bg,bor){
	if(obj.style && bor) obj.style.borderColor = bor;
	if(obj.style && bg) obj.style.backgroundColor = bg;
	if(obj.style && style) obj.style.borderStyle = style;
	if(obj.style && txt) obj.style.color = txt;
}

function htm_div_switch(obj1,obj2){
	obj1.css = obj1.css ? obj1.css : obj1.style;
	obj2.css = obj2.css ? obj2.css : obj2.style;
	if(obj1.css && obj2.css){
		xtmp = obj1.css.left, ytmp = obj1.css.top, ztmp = obj1.css.zIndex;
		obj1.css.left = obj2.css.left
		obj1.css.top = obj2.css.top
		obj1.css.zIndex = obj2.css.zIndex
		obj2.css.left = xtmp;
		obj2.css.top = ytmp;
		obj2.css.zIndex = ztmp;	
	}
}

function htm_div_blink(obj,color,delay,iter){
	try{
		obj.css = obj.css ? obj.css : obj.style;
		if(obj && js_str_isnumber(delay) && js_str_isnumber(iter)){
			if(iter <= 0) return true;
			org_color = obj.css.backgroundColor ? obj.css.backgroundColor : "white";
			obj.css.backgroundColor = color;
			objx = obj, colorx = org_color, delayx = delay, iterx = --iter;
			return setTimeout("htm_div_blink(objx,colorx,delayx,iterx)",delay);
		}
		return true;
	}	
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_div_blink2(id,obj,color,delay){
	try{
		if(id){
			org_color = obj.style.backgroundColor ? obj.style.backgroundColor : "white";
			obj.style.backgroundColor = color;
			idx = id, objx = obj, colorx = org_color, delayx = delay;
			id = setTimeout("htm_div_blink2(idx,objx,colorx,delayx)",delay);
		}
	}	
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_div_pos(obj,winname){
	try{
		obj.css = obj.css ? obj.css : obj.style;
		obj.doc = obj.doc ? obj.doc : obj.document;
		if(obj.css && obj.doc && winname){
			if(htm.br.ie){
				this.Bw = winname ? frame[winname].document.body.offsetWidth-4 : document.body.offsetWidth-4;
				this.Bh = winname ? frame[winname].document.body.offsetHeight-2 : document.body.offsetHeight-2;
				this.Ow = obj.doc.offsetWidth-30;
				this.Oh = obj.doc.offsetHeight;
			}
			else if(htm.br.nsg){
				return false; // Not yet implemented
			}
			else if(htm.br.nsc){
				this.Bw = winname ? frame[winname].document.innerWidth : null; // Need to check		
				this.Bh = obj.doc.innerHeight;			
				this.Ow = obj.doc.width;
				this.Oh = obj.doc.height;
			}
			this.xtmp = 0, this.ytmp = 0;
			this.BXc = this.Bw/2;
			this.BYc = this.Bh/2;
			this.OXp = this.BXc - this.Ow/2;
			this.OYp = this.BYc - this.Oh/2;	
		}
	}	
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_div_textzoom(obj,ID,mn,mx,delay){
	obj.css = obj.css ? obj.css : obj.style;
	if(obj.css){
		if(mn < mx){
			obj.css.fontSize = mn += 1, tmpID = ID, tmpObj = obj, tmpMn = mn, tmpMx = mx;
			setTimeout("htm_div_textzoom(tmpID,tmpObj,tmpMn,tmpMx)",50);
		}
		else htm_tme_clear(ID);
	}
}

function htm_div_scroll(oDiv,sPos){
	try{
		if(htm.br.ie || htm.br.nsc){			
			oDiv.style.pixelTop = document.body.scrollTop + sPos;
		}
		else if(htm.br.nsg){
			// need sure
		}
		else return false;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{
		return true;
	}
}

function htm_div_position(oDiv,sPos){
	 try{
	 	var iPos = 0;
		if(htm.br.ie){
			if(sPos == 'y'){
				yPos = oDiv.offsetTop;   
				tmpY = oDiv.offsetParent; 
				while(true){
					yPos += tmpY.offsetTop;
					if((tmpY = tmpY.offsetParent) == null)
			 		break;
				}
		  		iPos = yPos;
		  	}
			else if(sPos == 'x'){
				xPos = oDiv.offsetLeft;
				tmpX = oDiv.offsetParent;
				while(true){
					xPos += tmpX.offsetLeft;
					if((tmpX = tmpX.offsetParent) == null) break;
				}
				iPos = xPos;
			}
		}
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{
		return iPos;
	}
}
function htm_div_html(sOpt,sOpt2,sOpt3,sOpt4,bSort){
	try{
		if(sOpt == "settable"){
			sOpt4 = sOpt4 ? sOpt4 : ";"
			sOpt3 = sOpt3 ? sOpt3 : "\n"
			var r = sOpt2.split(sOpt3), s
			if(bSort){
				s = (r.slice(1,r.length)).sort()
				r = (new Array(r[0])).concat(s)
			}
			var t = '<table width="98%" border="0" cellspacing="0" cellpadding="0" class="cTable2">'
			for(var i = 0, c, t, len = r.length; i < len; i++){
				c = r[i].split(sOpt4)
				t = t + "\n<tr>"
				for(var j = 0, w; j < c.length; j++){
					h = (c[j] == "" || c[j] == null) ? "&nbsp;" : c[j]
					w = j == 0 ? " width=\"1%\" nowrap" : ""
					t = t + (i == 0 ? "<th" + w + " align=\"left\"> " + h + '</th>' : "<td" + w + "> " + h + '</td>')
				}
				t = t + "</tr>"
			}
			t = t + '</table>'
			return t
		}
		return false
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_div_table(sOpt,sTable,iIndex1,iIndex2,bRefresh,bNowrap){
	try{ // http://msdn.microsoft.com/library/default.asp?url=/workshop/author/dhtml/reference/methods/insertadjacenthtml.asp
		var oTR, oTD, sResult = true;
		var oTable = (typeof(sTable) == "string") ? document.getElementById(sTable) : sTable;
		if(!htm_str_isdefined(oTable)) return false // Must be done otherwise.. your system will crash if the object is not found
		else if(sOpt == "cell_insert"){ //Also: oTable.rows[iIndex1].insertCell(iIndex2);
			iIndex1 = js_str_isnumber(iIndex1) ? iIndex1 : -1;  // If iSource is not specified, all the cells are deleted
			oTR = oTable;
			if(oTD = oTR.insertCell(iIndex1)){
				oTD.innerHTML = htm_str_isdefined(iIndex2) ? iIndex2 : "&nbsp;"; //.replace(/[\\]*/g,"\\\\")
				sResult = oTD;
			}
		}
		else if(sOpt == "cell_delete"){ // Delete Cell
			iIndex2 = js_str_isnumber(iIndex2) ? iIndex2 : -1;  // If iSource is not specified, all the cells are deleted
			oTable.rows[iIndex1].deleteCell(iIndex2);
			if(bRefresh) oTable.refresh();
		}
		else if(sOpt == "row_insert"){ // Insert Row
			iIndex1 = js_str_isnumber(iIndex1) ? iIndex1 : -1;  // If iSource is not specified, all the cells are deleted
			if(oTR = oTable.insertRow(iIndex1)){
				sResult = oTR;
				if(bRefresh) oTable.refresh();
			}
		}
		else if(sOpt == "row_move"){ // Move Row
			oRow = oTable.moveRow(iIndex1,iIndex2);
			if(bRefresh) oTable.refresh();
		}
		else if(sOpt == "row_delete"){ // Delete Row
			if(oTable.rows(iIndex1)){
				oTable.deleteRow(iIndex1);
				oTable.refresh();
			}
		}
		else if(sOpt == "rows_clear"){ // Clear table rows
			for(var i = 0, len = oTable.rows.length; i < len; i++){
				htm_div_table(4,oTable,i);
			}
			sResult = oTable;
			if(bRefresh) oTable.refresh();
		}
		else if(sOpt == "row_clone_end"){ // Clone a row and insert at end of table. (THIS WON'T WORK if: <TABLE><FORM>..</FORM></TABLE>)
			iIndex1 = js_str_isnumber(iIndex1) ? iIndex1 : 0;			
			var oNodes = oTable.childNodes;
			for(var i = 0, len = oTable.rows.length; i < len; i++){
				if((oNodes(i).nodeName).toUpperCase() == "TBODY"){
					var oClone = oNodes(i).rows(iIndex1).cloneNode(true);
					iIndex2 = js_str_isnumber(iIndex2) ? iIndex2 : oNodes(i).rows.length-1;
					oNodes(i).rows(iIndex2).insertAdjacentElement("afterEnd",oClone);
					break;
				}
			}
			if(bRefresh) oTable.refresh();
		}
		else if(sOpt == "row_tbody_insert"){
			iIndex1 = js_str_isnumber(iIndex1) ? iIndex1 : -1;  // If iSource is not specified, all the cells are deleted
			var oNodes = oTable.childNodes;
			for(var i = 0, len = oNodes.length; i < len; i++){
				if((oNodes(i).nodeName).toUpperCase() == "TBODY"){
					if(oTR = oNodes(i).insertRow(iIndex1)){
						sResult = oTR;
					}
					break;
				}
			}
			if(bRefresh) oTable.refresh();
		}
		else if(sOpt == "row_tbody_clone"){ // Clone row in TBODY and insert at end of table
			var oBody = oTable;
			if((oBody.nodeName).toUpperCase() == "TBODY" && oBody.rows.length > 0){
				iIndex1 = js_str_isnumber(iIndex1) ? iIndex1 : oBody.rows.length-1;
				var oClone = oBody.rows(iIndex1).cloneNode(true);
				iIndex2 = js_str_isnumber(iIndex2) ? iIndex2 : oBody.rows.length-1;
				oBody.rows(iIndex2).insertAdjacentElement("afterEnd",oClone);
			}
			else sResult = false;
		}
		else if(sOpt == "row_tbody_delete"){ // Delete Row in TBODY
			var oBody = oTable;
			if(!oBody.rows || oBody.rows.length <= 1) return true;
			else if((oBody.nodeName).toUpperCase() == "TBODY"){
				iIndex1 = js_str_isnumber(iIndex1) ? iIndex1 : oBody.rows.length-1;
				oBody.deleteRow(iIndex1);
			}
			else sResult = false;
		}
		else if(sOpt == "rows_tbody_delete"){ // Delete Row in TBODY
			var oBody = oTable;			
			if(!oBody.rows || oBody.rows.length <= 1) return true;
			else if((oBody.nodeName).toUpperCase() == "TBODY"){ // Delte rows from iIndex1 to iIndex2
				iIndex2 = js_str_isnumber(iIndex2) ? iIndex2 : oBody.rows.length-1;
				iIndex1 = js_str_isnumber(iIndex1) ? iIndex1 : 0;
				for(var i = iIndex1; i < iIndex2; i++){
					oBody.deleteRow(iIndex1);
				}
			}
			else sResult = false;
		}
		else return false;
		
		return sResult;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

////////////////////////////////////////////////////////////////////////////////
///////// MENU FUNCTIONS

function htm_doc_getselect(){ // http://www.quirksmode.org/js/selected.html
	try{
		if(window.getSelection){ // Mozilla 1.75*, Safari 1.3 
			return window.getSelection();
		}
		else if(document.getSelection){ // Explorer 5.2 Mac, Mozilla 1.75*, Netscape, Opera
			return document.getSelection();
		}
		else if(document.selection){ // Explorer 5 Win, Explorer 6 Win 
			return document.selection.createRange().text;
		}
		return "";
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}



////////////////////////////////////////////////////////////////////////////////
///////// TIME/DATE FUNCTIONS

function htm_dte_show(lang,date,opt){
	if(!date) var date = new Date();
	if(!lang) return false; // lang: en, sv, no
	date = new Date(date);
	var hour = date.getHours(), min = date.getMinutes(), sec = date.getMinutes();
	hour = (hour<10) ? '0' + hour : hour;
	min = (min<10) ? '0' + min : min;
	sec = (sec<10) ? '0' + sec : sec;
	year = !htm.nsc ? date.getYear() : '';
	var today = htm_dte_day(lang,date.getDay()) + (lang != "en" ? " den " : ", ") + date.getDate() + ' ' + htm_dte_month(lang,date.getMonth())
	if(opt == 2) res = today + ' ' + year;
	else if(opt == 3) res = hour + ':' + min + ":" + sec;
	else res = today + ' ' + year + ', ' + hour + ':' + min + ":" + sec;
	return res;
}

function htm_dte_month(lang,month){
	var m_ary = new Array();
	m_ary['sv'] = new Array('Januari','Februari','Mars','April','Maj','Juni','Juli','Augusti','September','Oktober','November','December');
	m_ary['en'] = new Array('January','Febuary','March','April','May','June','July','August','September','October','November','December');
	m_ary['no'] = new Array('Januar','Februar','Mars','April','Mai','Juni','Juli','Augusti','September','Oktober','November','December');
	return m_ary[lang][month];
}

function htm_dte_day(lang,day){
	var d_ary = new Array();
	d_ary['en'] = new Array('Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday');
	d_ary['sv'] = new Array('Söndag','Måndag','Tisdag','Onsdag','Torsdag','Fredag','Lördag');
	d_ary['no'] = new Array('Søndag','Mandag','Tirsdag','Onsdag','Torsdag','Fridag','Lørdag');	
	return d_ary[lang][day];
}

function htm_dte_get(iopt,lang,date,bNoMille,bNoSep){
	var res = false, m, t;
	var d = date ? new Date(date) : new Date();
	if(iopt == 1){
		m = (d.getMonth()+1), m = (m<10 ? "0" + m : m);
		t = d.getDate(), t = (t<10 ? "0" + t : t);
		if(lang == "no") {
			sSep = !bNoSep ? "." : ""
			res = t + sSep + m + sSep + d.getFullYear();
		}
		else if(lang == "sv"){
			sSep = !bNoSep ? "-" : ""
			res = d.getFullYear() + sSep + m + sSep + t;
		}
		else {
			sSep = !bNoSep ? "/" : ""
			res = t + sSep + m  + sSep + d.getFullYear();
		}
	}
	else if(iopt == 2){ // DataTime
		var sDate = htm_dte_get(1,lang,date,bNoMille,bNoSep);
		h = d.getHours(), h = (h<10 ? "0" + h : h);
		m = d.getMinutes(), m = (m<10 ? "0" + m : m);
		s = d.getSeconds(), s = (s<10 ? "0" + s : s);
		if(!bNoMille){
			ms = "." + d.getMilliseconds(), ms = (ms <= 9) ? "00" + ms : ((ms <= 99) ? "0" + ms : ms); 
		}
		else ms = "";
		res = sDate + " " + h + ":" + m + ":" + s + ms;
	}
	else if(iopt == 3){ // DataTime without milli
		res = htm_dte_get(2,lang,date,true,bNoSep);		
	}
	return res;
}

function htm_dte_utc(iopt,lang,date){
	var utc = false, d, m, t;
	d = date ? new Date(date) : new Date();
	if(iopt == 1){
		m = (d.getUTCMonth()+1), m = (m<10 ? "0" + m : m);
		t = (d.getUTCDate()+1), t = (t<10 ? "0" + t : t);
		if(lang == "no") utc = t + "." + m + "." + d.getUTCFullYear();
		else if(lang == "sv") utc = d.getUTCFullYear() + "-" + m + "-" + t;
		else utc = t + "/" + m  + "/" + d.getUTCFullYear();
	}
	else if(iopt == 2){ // DataTime
		var sDate = htm_dte_utc(1,lang,date);
		h = (d.getHours()+1), h = (h<10 ? "0" + h : h);
		m = (d.getMinutes()+1), m = (m<10 ? "0" + m : m);
		s = (d.getSeconds()+1), s = (s<10 ? "0" + s : s);
		ms = (d.getMilliseconds()+1), ms = (ms <= 9) ? "00" + ms : ((ms <= 99) ? "0" + ms : ms);
		utc = sDate + " " + h + ":" + m + ":" + s + "." + ms;
	}
	else if(iopt == 3){
		
	}
	return utc;
}

function htm_dte_conv(iOpt,sDate,sLang){
	try{
		var aDate = new Array(), sNewDate = "";
		if(!sDate) return false
		else sDate = sDate.toLowerCase();
		if(iOpt == 1){ // sDate = 10-Oct-2003 (Norwegian)
			aDate["jan"] = "01",aDate["feb"] = "02",aDate["mar"] = "03",aDate["apr"] = "04",aDate["may"] = aDate["mai"] = "05",aDate["jun"] = "06";
			aDate["jul"] = "07",aDate["aug"] = "08",aDate["sep"] = "09",aDate["oct"] = aDate["okt"] = "10",aDate["nov"] = "11",aDate["dec"] = "12";
			var oRe = new RegExp("([0-9]{1,2})-([a-z]{3})-([0-9]{4})(.*)","ig");
			if(oRe.exec(sDate)){
				sMounth = aDate[RegExp.$2];
				sDate = (RegExp.$1 < 10) ? "0" + RegExp.$1 : RegExp.$1;
				sTimeStamp = RegExp.$4 ? RegExp.$4 : ""
				if(sLang == "no") sNewDate = sDate + "." + sMounth + "." + RegExp.$3 + sTimeStamp;
				else if(sLang == "sv") sNewDate = RegExp.$3 + "-" + sMounth + "-" + sDate + sTimeStamp;
				else sNewDate = sDate + "/" + sMounth + "/" + RegExp.$3 + sTimeStamp;
			}
		}
		else if(iOpt == 2){
			sNewDate = htm_dte_conv(1,sDate,sLang) + " 00:00:00.000";
		}
		else if(iOpt == 3){ // Return the date object
			sDate = htm_dte_conv(1,sDate,"sv");
			d = sDate.split("-");
			t = sDate.split(" "), t = t[1].split(":");
			return new Date(d[0],(d[1]-1),d[2].substring(0,2),t[0],t[1],t[2])
		}
		return sNewDate;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_dte_lastday(iOpt,oDate){
	try{
		// iDay: 0-6 (Sunday-Saturday)
		var oDateTmp = oDate = (oDate ? oDate : new Date());
		var iDayMilli = 60 * 60 * 24 * 1000;
		var iDayTime = oDate.getTime();
		for(var d = 1; d < 32; d++){ // Count down for a week
			oDateTmp.setTime(iDayTime - (d*iDayMilli));
			if(iOpt == 1){ // Day before
				return oDateTmp;		
			}
			else if(iOpt == 2){ // Last weekday (mon-fre)
				iDay = oDateTmp.getDay();
				if(iDay > 0 && iDay < 6){ // If first monday in the current month
					return oDateTmp;
				}
			}
			else if(iOpt == 3){ // Last month day
				if(oDateTmp.getDate() == 1){ // If first monday in the current month
					return oDateTmp.setTime(iDayTime - ((d+1)*iDayMilli));;
				}
			}
		}
		return false;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_tme_clear(ID){
	if(ID) clearTimeout(ID);
	ID = false;
}

function htm_tme_sleep(id,delay){
	id = setTimeout("htm_tme_clear",delay,id);
}

////////////////////////////////////////////////////////////////////////////////
///////// FUNCTIONS (OTHER)

function htm_randomize(lval,hval){
	hval = new Number(hval);
	lval = new Number(lval);
	if(hval >= lval){
		var now = new Date(), tmp;
		tmp = Math.round(Math.abs(Math.sin(now.getTime())*1000000000));
		return tmp % (hval-lval+1) + lval;
	}
	else return false;
}

function htm_ary_push(ary,obj){
	if(htm.br.nsc) len = ary.push(obj); // IE 5.0 doesn't support the push metod
	else ary[ary.length] = obj, len = ary.length;
	return len;
}

function htm_ary_getitem(opt,aArray,sItemStr,bIgnoreCase){
	try{
		if(typeof(aArray) == "object"){
			sItemStr = bIgnoreCase ? sItemStr : sItemStr.toUpperCase();
			if(!bIgnoreCase){
				sArray = ( aArray.toString() ).toUpperCase();
				aArray = sArray.split(",");
			}
			for(var sItem in aArray){
				sItem = bIgnoreCase ? sItem : sItem.toUpperCase();
				if(sItem == sItemStr){
					return aArray[sItem];
				}
			}
		}
		return false;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_ary_sortobject(o){
	try{
		var o1 = new Array()
		for(i in o){
			o1.push(i)
		}
		o1.sort()
		var o2 = new Object()
		for(var i = 0, len = o1.length; i < len; i++){
			o2[o1[i]] = o[o1[i]]
		}
		return o2
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_ary_sort(opt,aArray,sObject,bIgnoreCase){
	try{
		var bSort = false;
		for(var i = 0, len = aArray.length; i < (len-1); i++){
		    for(var j = i+1; j < len; j++){
		    	if(opt == 1) bSort = (aArray[j] < aArray[i]);
			else if(opt == 2){ // Object
				a = bIgnoreCase ? (aArray[j][sObject]).toUpperCase() : aArray[j][sObject];
				b = bIgnoreCase ? (aArray[i][sObject]).toUpperCase() : aArray[i][sObject];
				bSort = (a < b);
			}
			if(bSort){
				var dummy = aArray[i];
				aArray[i] = aArray[j];
				aArray[j] = dummy;
				bSort = false;
			}
		}
		}
		return aArray;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_ary_duplicate(opt,aArray1,aArray2,bIgnoreCase){
	try{
		var aArray = new Array();
		if(opt == 1){ // Return distinct Array list of arrays
			if(aArray2) aArray1 = aArray1.concat(aArray2)
			aArray1 = aArray1.sort();
			aArray2 = new Array()
			for(var i = 0, len = aArray1.length; i < len; i++){
				aArray2.push(aArray1[i])
				if(!bIgnoreCase) {
					while(aArray1[i] == aArray1[i+1]) i++;
				}
				else{
					try{ while((aArray1[i]).toUpperCase() == (aArray1[i+1]).toUpperCase()) i++; }
					catch(ee){}
				}
			}
			aArray = aArray2;
		}
		else if(opt == 2){ // Return a subtraction array: aArray1 - aArray2
			var aDict = new ActiveXObject("Scripting.Dictionary");
			for(var i = 0, len = aArray1.length; i < len; i++){
				if(!aDict.Exists(aArray1[i])){ // Remember case sentitive!!
					aDict.Add(aArray1[i],i);
				}
			}
			for(var i = c = 0, len = aArray2.length; i < len; i++){
				if(aDict.Exists(aArray2[i])){
					aDict.Remove(aArray2[i]); c++
				}
			}
			aArray = new VBArray(aDict.Keys()).toArray();//alert(c)
		}
		else if(opt == 3){ // Return Array list of items from Array2 that is not included in Array1
			var aDict = new ActiveXObject("Scripting.Dictionary");
			for(var i = 0, len = aArray1.length; i < len; i++){
				if(!aDict.Exists(aArray1[i])){ // Remember case sentitive!!
					aDict.Add(aArray1[i],i);
				}
			}
			aArray1 = new Array()
			for(var i = 0, len = aArray2.length; i < len; i++){				
				if(!aDict.Exists(aArray2[i])){
					aArray1.push(aArray2[i]);
				}
			}
			aDict.RemoveAll();
			aArray = aArray1;
		}
		return aArray;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_msg_pop(iOpt,msg){
	try{
		htm.br.exec = true;
		switch(iOpt){
			case 1 : {
				if(htm.br) alert(msg);
				else if(WScript) WScript.Echo(msg);
				break;
			}
			case 2 : {
				
				break;
			}
			case 3 : {
				
				break;
			}
			default : {
				htm.br.exec = false;
				break;
			}
		}
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{
		return htm.br.exec;
	}
}

function htm_encode_init(form) {
	username = form.user.value, password = form.passwd.value;
	top.main.location = encode ((tmp = username+password),tmp.length % 10) + ".html";
}

function htm_encode(sString,CipherVal){
	try{
		// http://www.htmlhelp.com/reference/charset/iso032-063.html
		// convert to hex to get the charactor
	    var space = "\u0020"
	    var alpha = "0123456789abcdefghijklmnopqrstuvwxyz._~ABCDEFGHIJKLMNOPQRSTUVWXYZ-+$\\";
	    var CipherVal = typeof(CipherVal) == "number" ? parseInt(CipherVal) : (sString.length % 10)
		var sConvert = "", sChar
	    for(var i = 0, len = sString.length; i < len; i++){
	    	sChar = Cipher = sString.substring(i,i+1);
	    	if(sChar != " " && sChar != space){
	    		Conv = alpha.indexOf(sChar), Cipher = Conv^CipherVal;
	        	Cipher = alpha.substring(Cipher,Cipher+1);
	        }
	        sConvert = sConvert + Cipher
	    }
	    return sConvert;
	}
    catch(e){
		htm_log_error(2,e);
		return false;
	}
}


////////////////////////////////////////////////////////////////////////////////
///////// IMAGE FUNCTIONS

function htm_img_init(){
	try{
		
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{
		return true;
	}
}

function htm_img_act(name,action){
	if(htm.br.img && htm.br.dyn){
		if(src = document.images[name].src){
			var re = /(.+)\_\w+\.(.{3,}$)/gi; // A pattern that matches mypic_out.gif or mypic_over.jpg
			if(src.match(re)){
				if(htm.br.ie5){
					idx = src.lastIndexOf('_'), len = src.length;
					newsrc = src.substring(0,idx) + '_' + action +  src.substring(len-4,len);
					document.images[name].src = newsrc;					
				}
				else{
					if((ary = re.exec(src))){
						newsrc = ary[1] + '_' + action +  '.' + ary[2];
						document.images[name].src = newsrc;
					}
				}
			}
			else htm_msg_pop(1,'The image: "' + name + '" either don\'t exist or the file name has wrong naming convention.')	
		}
	}	
}

function htm_img_load(){	
	htm.img[0] = new Image(195,96);
	htm.img[0].src = '/pics/pix_1_195px.gif';
	htm.img[0].large = '/pics/pix_1_400px.gif';
	htm.img[1] = new Image(195,140);
	htm.img[1].src = '/pics/pix_2_195px.gif';
	htm.img[1].large = '/pics/pix_2_400px.gif';
	htm.img.load = true;
}


function htm_img_animate(){
	try{
		// PROTECTED VARIABLES
		var load = false
		var anime_img = new Array()
		var anime_num = 0
		var anime_len = 0
		var anime_ID = null
		var interval = 50
		var delay = 100
		var delay_MAX = 4000
		var image = ""
		
		// PRIVATE FUNCTION
		this.load = function(aImages,ID,sImageID,w,h){
			if(typeof(aImages) != "object" || !aImages.length) return false
			anime.length = 0
			for(var i = 0, len = aImages.length; i < len; i++){   
				anime_img[i] = (w && h) ? new Image(w,h) : new Image();
				anime_img[i].src = aImages[i] // Remember to use front slash: "bin-pics/cartoon-fight" + i + ".gif";
			}
			load = (i > 0)
			anime_ID = ID
			image = sImageID
			anime_len = len
			return load
		}

		// PRIVATE FUNCTION		
		this.start = function(){
			if(!anime_ID && load){
				anime_ID = setTimeout('anime()',interval);
			}
		}

		// PRIVATE FUNCTION
		this.stop = function(){
			if(anime_ID){
				clearTimeout(anime_ID);
			}
			anime_ID = null
		}

		// PUBLIC FUNCTION
		anime = function(){
			if(anime_ID && load){
				// alert(anime_num + " " + anime_img[anime_num].src)
				if(anime_num >= anime_len) anime_num = 0
				document.images[image].src = anime_img[anime_num++].src
				if(anime_ID){
					setTimeout('anime()',delay);
				}
			}
		}

		// PRIVATE FUNCTION		
		this.animeSlower = function(){			
			if(delay < delay_MAX){
				delay += 15
			}			
		}

		// PRIVATE FUNCTION		
		this.animeFaster = function(){			
			if(delay > 0){
				delay -= 15
			}			
		}
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}


////////////////////////////////////////////////////////////////////////////////
///////// FORMS FUNCTIONS

function htm_frm_isvalid(iOpt,frm,arg3,arg4){
	try{
		var val, bOK = false;		
		if(iOpt != 3 && ((val = frm.elements[arg3].value) == "not entered" || val == "" || val == undefined || val == null || !val)) bOK = false;
		
		switch(iOpt){
			case 1 : { // If is an e-mail
				var Re1 = /(@.*@)|(\.\.)|(@\.)|(\.@)|(^\.)/; 
				var Re2 = /(^[_a-zA-Z0-9-]+(\.[_a-zA-Z0-9-]+)*@[a-zA-Z0-9-]+(\.[a-zA-Z0-9-]+)*(\.[a-zA-Z]{2,3})$)/i;
				bOK = Re2.exec(val);
				break;
			}
			case 2 : { // If is a name
				var Re = /[A-Za-z_\xC5\xE5\xC4\xE4\xD6\xF6]+\s*[A-Za-z_\xC5\xE5\xC4\xE4\xD6\xF6]*/;
				bOK = Re.exec(val);
				break;
			}
			case 3 : { // If is an url
				var Re = /^http:\/\/.+\..+|^ftp:\/\/.+\..+/g;
				bOK = Re.exec(val);
				break;
			}
			case 4 : { // If is phone
				var ReUS = /\(?\d{3}\)?([-\/\.])\d{3}\1\d{4}/g; // US phone syntax
				var Re = /(\+\d{1,4})?\s?(\(\d{1,2}\))?(\d{1,3})?\s?\d{1,8}/; // +46 (0)8 51850962 or 71545846
				Re = (arg4 == 'us') ? ReUS : Re;
				bOK = Re.exec(val);
				break;
			}
			case 5 : { // If is a date
				var Re = /([0-9]*)/i;
				if(!(Re.exec(val)) || val.length != 8 || isNaN(val)) return false;
				else if((tmp = parseFloat(val.substring(4,6))) < 1 || tmp > 12 ) bOK = false; // If the months is incorrect
				else if((tmp = parseFloat(val.substring(6,8))) < 1 || tmp > 31 ) bOK = false; // If the date is incorrect
				break;
			}
			case 6 : { // If radio is checked
				bOK = false;
				for(var i = 0; i < arg4; i++){
					if(frm.elements[arg3][i].checked){
						bOK = true;
						break;
					}
				}
				break;
			}
			case 7 : { // If element is not empty
				Re = /([\s]+)/i;
				bOK = !Re.exec(val);
				break;
			}
			case 8 : { //If is contains HTML
				var Re = /([\<])([^\>]{1,})*([\>])/i
				bOK = Re.exec(val);
				break;
			}
			case 9 : { // If is a name
				var Re = /[0-9A-Za-z_\xC5\xE5\xC4\xE4\xD6\xF6]+\s*[A-Za-z_\xC5\xE5\xC4\xE4\xD6\xF6]*/;
				bOK = Re.exec(val);
				break;
			}
			case 10 : { // Array og Email
				var Re2 = /(^[_a-zA-Z0-9-]+(\.[_a-zA-Z0-9-]+)*@[a-zA-Z0-9-]+(\.[a-zA-Z0-9-]+)*(\.[a-zA-Z]{2,3})$)/i;
				if(typeof(val) == "string"){
					bOK = true;
					var aEmail = val.split(",");
					for(var i = 0; i < aEmail.length; i++){
						if(!Re2.exec(aEmail[i])){
							bOK = false; break;
						}
					}
				}
				break;
			}
			default : {
				return false;
			}
		}
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{		
		return bOK;
	}
}

function htm_frm_replace(frm,item,trim){
	if((val = frm.elements[item].value) != '' || val != null){
		//Re = /\.+/;
		var Re = /{trim}+/gi;
		frm.elements[item].value = val.replace(Re, "");
	}
}
function htm_frm_input(iOpt,oInput){
	try{
		var bResult = true;
		if(typeof(oInput) != "object" && oInput.tagName != "INPUT") return false;
		switch(iOpt){
			case 1 : { // UpperCase
				oInput.value = (oInput.value).toUpperCase();
				break;
			}
			case 2 : { // Clears select
				
				break;
			}
			default : {
				bResult = false;
				break;
			}
		}
		return bResult;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_frm_selectkeydown(oForm,oSel,iIndex,bValue,sRegPrefix){ 
	try{ // This is needed for IE. FireFox and Opera has this implemented by default
		var oEvent = window.event
		if(!js_str_isdefined(oForm,oSel)) return false
		else if(typeof(oSel) == "string") oSel = oForm[oSel];
		else if(typeof(oSel) == "object"); 
		else return ;
		var sKeyCode = oEvent.keyCode;
		var sToChar = String.fromCharCode(sKeyCode); // excellence
		var oRe1 = /[0-9a-z\- ]/ig
		//if(sKeyCode > 47 && sKeyCode < 91){
		if(sToChar.match(oRe1)){
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
				//setTimeout("htm_frm_selectkeyup('" + oForm.name + "','" + oSel.name + "')",2000,"JavaScript")
			}
			var sFinder = oSel.getAttribute("finder");
			iIndex = js_str_isnumber(iIndex) ? iIndex : 0;
			var o = bValue ? "value" : "text"
			sRegPrefix = typeof(sRegPrefix) == "string" ? sRegPrefix : false
			for(var i = iIndex, len = oSel.options.length; i < len; i++){
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
		else{
			//Not a digit;
		}
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_frm_selectkeyup(oForm,oSel){
	try{
		if(!js_str_isdefined(oForm,oSel)) return false
		if(typeof(oForm) == "string") oForm = document.forms[oForm]
		if(typeof(oSel) == "string") oSel = oForm[oSel];
		else if(typeof(oSel) == "object"); 
		else return ;

		if(!oSel.getAttribute("active")) return;
		var iNow = new Date().getTime();
		if(iNow > parseInt(oSel.getAttribute("timer")) + 2000){
			oSel.setAttribute("active",false)
			try{ oSel.onchange() }
			catch(ee){}
		}
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}
function htm_frm_select(sOpt,oForm,oSel,iIndex,sValue,sText,sOpt7,sOpt8,sOpt9,sOpt10){
	try{
		if(sOpt != 0 && !js_str_isdefined(sOpt,oForm,oSel)) return false
		else if(typeof(oSel) == "string") oSel = oForm[oSel];
		else if(typeof(oSel) == "object"); 
		else return false;
		var bResult = true;
		sOpt = sOpt.toString()
		if(typeof(oSel) != "object" || oSel.tagName != "SELECT") return false
		else if(sOpt == 0 || sOpt.match(/reset/i)){
			oSel.options.selectedIndex = 0
			oSel.disabled = iIndex ? true : false;
			if(oSel.onchange) oSel.onchange();
		}
		else if(sOpt == 1 || sOpt.match(/addindex/i)){
			iIndex = js_str_isnumber(iIndex) ? iIndex : oSel.options.length;
			oSel.options[iIndex] = new Option(sText,sValue);
			oSel.options[iIndex].selected = true
		}
		else if(sOpt == 2 || sOpt.match(/clear/i)){
			bDefault = sOpt7 ? true : false
			iIndex = js_str_isnumber(iIndex) ? iIndex : 0;
			while(oSel.options.length > iIndex){
				oSel.options[iIndex] = null; // Also possible to use options.removeNode()
			}
			if(!bDefault) oSel.options[iIndex] = new Option(" ","",false,false);
		}
		else if(sOpt == 3 || sOpt.match(/setindex/i)){
			if(iIndex >= 0 && iIndex < oSel.options.length){
				oSel.options.selectedIndex = iIndex ? iIndex : 0;
				if(oSel.onchange) oSel.onchange();
			}
			else bResult = false;
		}
		else if(sOpt == 4 || sOpt.match(/disable/i)){
			oSel.disabled = iIndex = iIndex ? true : false;
		}
		else if(sOpt == 5 || sOpt.match(/setvalue/i)){
			if(oSel.options[iIndex]) oSel.options[iIndex].selected = true;
			else if(sValue) oSel.options.value = sValue;
			else return false
			if(oSel.onchange) oSel.onchange();
		}
		else if(sOpt == 6 || sOpt.match(/getvalue/i)){
			iIndex = js_str_isnumber(iIndex) ? iIndex : oSel.options.selectedIndex
			bResult = oSel.options[iIndex].value;
		}
		else if(sOpt == 7 || sOpt.match(/getindex/i)){
			bResult = -1;
			iIndex = js_str_isnumber(iIndex) ? iIndex : 0;
			for(var len = oSel.options.length; iIndex < len; iIndex++){
				if(oSel.options[iIndex].value == sValue){
					bResult = iIndex;
					break;
				}
			}
		}
		else if(sOpt == 9 || sOpt.match(/gettext$/i)){
			bResult = oSel.options[oSel.options.selectedIndex].text;
		}
		else if(sOpt == 10 || sOpt.match(/gettextindex/i)){
			bResult = iIndex = js_str_isnumber(iIndex) ? iIndex : 0;
			sText = (new String(sText)).toLowerCase()
			for(var sText1, sText2, len = oSel.options.length; iIndex < len; iIndex++){
				var sText2 = oSel.options[iIndex].text
				var sText1 = (oSel.options[iIndex].text).replace(/([0-9]{3} )(.*)/g,"$2")
				if(sText1.toLowerCase() == sText || sText2.toLowerCase() == sText){
					bResult = iIndex;
					break;
				}
			}
		}
		else if(sOpt == 11 || sOpt.match(/addarraysplit/i)){
			var aObject = sOpt7, sSplit = sOpt8
			var bNoCount = sOpt9 ? true : false
			iIndex = js_str_isnumber(iIndex) ? iIndex : 1;
			if(typeof(aObject) == "object" && js_str_isdefined(sSplit)){
				htm_frm_select("clear",oForm,oSel,iIndex);
				for(var i = 0, len = aObject.length; i < len; i++, iIndex++){
					var n = !bNoCount ? js_str_number(iIndex+1) : ""
					var a = (aObject[i]).split(sSplit), sValue = a[1], sText = n + " " + a[0];
					htm_frm_select("addindex",oForm,oSel,iIndex,sValue,sText);
				}
			}
			if(oSel.options.length > 0) oSel.options[0].selected = true
			else bResult = false
		}
		else if(sOpt == 12 || sOpt.match(/addarrayobject/i)){ // htm_frm_select("addarrayobject",oForm,"servers",0,"Name","Type",aServers,false);
			var aObject = sOpt7
			var bNoCount = sOpt8 ? true : false
			var bText = (typeof(sText) == "string")
			iIndex = js_str_isnumber(iIndex) ? iIndex : 1;
			if(typeof(aObject) == "object" && js_str_isdefined(sValue)){
				htm_frm_select("clear",oForm,oSel,iIndex);
				for(var i = 0, len = aObject.length; i < len; i++, iIndex++){
					var n = !bNoCount ? js_str_number(iIndex+1) : ""					
					var sText2 = n + " " + aObject[i][sValue]
					if(bText) sText2 = sText2 + " (" + aObject[i][sText] + ")"
					var sValue2 = aObject[i][sValue]
					htm_frm_select("addindex",oForm,oSel,iIndex,sValue2,sText2);
				}
			}
			if(oSel.options.length > 0) oSel.options[0].selected = true
			else bResult = false
		}
		else if(sOpt == 13 || sOpt.match(/addarray$/i)){ // htm_frm_select("addarray",oForm,"locations",0,null,null,aDomains);
			var aObject = sOpt7
			var bNoCount = sOpt8 ? true : false
			var sTextExtra = sOpt9 ? sOpt9 : ""			
			if(typeof(aObject) == "object"){
				iIndex = js_str_isnumber(iIndex) ? iIndex : 1;
				htm_frm_select("clear",oForm,oSel,iIndex);
				for(var i = 0, len = aObject.length; i < len; i++, iIndex++){
					var n = !bNoCount ? js_str_number(iIndex+1) : ""
					sValue = aObject[i], sText = n + " " + sTextExtra + aObject[i];
					htm_frm_select("addindex",oForm,oSel,iIndex,sValue,sText);
				}
			}
			if(oSel.options.length > 0) oSel.options[0].selected = true
			else bResult = false
		}
		else if(sOpt.match(/addarrayitem/i)){ // htm_frm_select("addarray",oTag,"domains",null,"domain",null,aDomains);
			var aObject = sOpt7, sItem = sValue
			var bNoCount = sOpt8 ? true : false
			var sTextExtra = sOpt9 ? sOpt9 : ""			
			if(typeof(aObject) == "object" && js_str_isdefined(sItem)){
				iIndex = js_str_isnumber(iIndex) ? iIndex : 1;
				htm_frm_select("clear",oForm,oSel,iIndex);
				for(var i = 0, len = aObject.length; i < len; i++, iIndex++){
					var n = !bNoCount ? js_str_number(iIndex+1) : ""
					sValue = aObject[i][sItem], sText = n + " " + sValue;
					htm_frm_select("addindex",oForm,oSel,iIndex,sValue,sText);
				}
			}
			if(oSel.options.length > 0) oSel.options[0].selected = true
			else bResult = false
		}
		else if(sOpt.match(/isselected/i)){
			bResult = false
			for(var i = 0, len = oSel.options.length; i < len; i++){
				if(oSel.options[i].selected){
					bResult = true
					break;
				}
			}			
		}
		else if(sOpt.match(/getarray$/i)){ // htm_frm_select("getarray",oForm,"locations",0,null,null,false);
			var aSelect = new Array()
			var bSelected = sOpt7 ? true : false
			iIndex = js_str_isnumber(iIndex) ? iIndex : 1;
			var o = sValue == "text" ? sValue : "value"
			for(var i = iIndex, len = oSel.length; i < len; i++){
				if(bSelected){
					if(!oSel.options[i].selected) continue
				}
				aSelect.push(oSel.options[i][o])	
			}
			bResult = aSelect
		}
		else if(sOpt.match(/clone/i)){
			
		}
		return bResult;
	}
	catch(e){ //alert(sOpt + " " + oSel.name)
		htm_log_error(2,e);
		return false;
	}
}

function htm_frm_radio(oForm,sRadioName,iOpt){
	try{
		var oRadio, sResult = false;
		if(oRadio = oForm.elements[sRadioName]){
			for(var i = 0, len = oRadio.length; i < len; i++){
				if(oRadio[i].checked){
					sResult = (iOpt == 2) ? i : oRadio[i].value;
					break;
				}
			}
		}
		return sResult;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_frm_checkbox(sOpt,oForm,sElemStream){
	try{
		if(typeof(oForm) != "object") return false
		else if(sOpt == "ischecked"){ // one of them is checked
			for(var i = 2, len = arguments.length; i < len; i++){
				if(oForm[arguments[i]] && oForm[arguments[i]].checked) return true;
			}
		}
		else if(sOpt == "isnotchecked"){ // one of them is checked
			for(var i = 2, len = arguments.length; i < len; i++){
				if(oForm[arguments[i]] && !oForm[arguments[i]].checked) return true;
			}
		}
		return false
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

////////////////////////////////////////////////////////////////////////////////
///////// COOKIE FUNCTIONS

function htm_cke_set(win,name,value,expire){
	try{
		win.document.cookie = name + "=" + escape(value) + ((expire == null) ? "" : ("; expires=" + expire.toGMTString()));
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{
		return true;
	}
}
function htm_cke_get(Name){
   var search = Name + "=";
   if(document.cke.length > 0){ // if there are any cookies
      offset = document.cookie.indexOf(search) 
      if (offset != -1) { // if cookie exists 
         offset = offset + search.length 
         // set index of beginning of value
         end = document.cookie.indexOf(";", offset) 
         // set index of end of cookie value
         if (end == -1) end = document.cookie.length
         return unescape(document.cke.substring(offset, end))
      } 
   }
}

function htm_cke_reg(name) {
   var today = new Date()
   var expires = new Date()
   expires.setTime(today.getTime() + 1000*60*60*24*365)
   htm_cke_set("TheCoolJavaScriptPage", name, expires)
}

function htm_cke_save(name,value,days,iOpt) {
	if (days) {
		var date = new Date();
		date.setTime(date.getTime() + (days*24*60*60*1000))
		var expires = "; expires=" + date.toGMTString()
	}
	else expires = "";
	document.cookie = name + "=" + value + expires + "; path=/";
	if(!iOpt) history.go(0);
}

function htm_cke_read(name) {
	var nameEQ = name + "="
	var ca = document.cookie.split(';')
	for(i = 0; i < ca.length; i++){
		c = ca[i];
		while (c.charAt(0) == ' ') c = c.substring(1,c.length);
		if(c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length,c.length)
	}
	return null;
}

function htm_cke_del(name) {
	htm_cke_save(name,"",-1)
}

///////// STRING 

function htm_str_isdefined(sStream){
	try{
		for(var s, i = 0; i < arguments.length; i++){
			s = arguments[i];
			if(s == "" || s == false || s == null || s == undefined || s == " " || s == "undefined"){
				return false;
			}
		}
		return true;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_str_html(iOpt,sHtml){
	try{
		if(iOpt == 1){
			// Removes HTML tags in a string
			sHtml = (sHtml.replace(/<[0-9a-z="': ;\-()]*>/ig,"")).replace(/<\/[a-z ]*>/ig,"");
		}
		return sHtml;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

////////////////////////////////////////////////////////////////////////////////
///////// MOUSE FUNCTIONS

function htm_mse_init(){
	try{
		if(htm.br.mse.enable){
			if(htm.br.ie) {
				document.onmousedown = htm_mse_disable;
				document.onkeydown = htm_mse_checkkey;
			}
			else if(htm.br.nsg){
				
			}
			else if(htm.br.nsc) {
				document.onmousedown = htm_mse_disable;
				document.captureEvents(Event.MOUSEDOWN);
			}
		}
		else return false;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{
		return true;
	}
}

function htm_mse_disable(e){
	try{
		var no = 0;
		if(htm.br.ie) no = event.button;
		else if(htm.br.nsg);
		else if(htm.br.nsc) no = e.which;
		else {
			return false;
		}		
		if(no == 2 || no == 3){
			var msg = "Nothing wrong with your browser. Right clicking is disabled.";
			if(htm.br.ie){
				htm_log_error(3,msg);
		    	return false;
			}
			else if(htm.br.nsg){
				
			}
			else if(htm.br.ns && !htm_cke_read(htm.br.cke.drmouse)){
				htm_cke_save(htm.br.cke.drmouse,'Yes',1,true);
				htm_log_error(3,msg);
			}
		}
		return false;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{
		return true;
	}
}

function htm_mse_checkkey(){
    try{
		var mouse_key = 93, keycode = event.keyCode;
    	if(htm.br.ie){
			if(keycode == mouse_key) htm_msg_pop(1,"Mouse Key Is Disabled" );
		}
		else if(htm.br.nsg){
		
		}
		else if(htm.br.nsc){
		
		}
		else return false;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
	finally{
		return false;
	}
}

function htm_mse_cursor(iOpt,obj,type){
	if(obj.css){
		switch(iOpt){
			case 1 : {
				if(htm.br.ie) obj.css.cursor = 'hand';
				else if(htm.br.nsg) obj.css.cursor = 'pointer';
				else if(htm.br.nsc);
				break;
			}
			case 2 : {
				obj.css.cursor = 'default';
				break;
			}
			case 3 : {
				
				break;
			}
			default : {
				htm.br.exec = false;
				break;
			}
		}
		
	}	
}


////////////////////////////////////////////////////////////////////////////////
///////// DEBUG/ERROR FUNCTIONS

function htm_dbg_object(Obj){ // THIS FUNCTION IS USED FOR VIEWING COM OBJECTS
	try{
		if(typeof(Obj) == "string") Obj = document.all[Obj];
		var sFile = oReg.temp + "\\debug.html";
		var oFout = oFso.CreateTextFile(sFile,true);
		var sHead = "<html><head>\n<style>td{font-family:Arial;font-size:11px;padding:1px;}</style>";
		sHead = sHead + '<script language="javascript" src="debug2.js"></script>\n</head><body>';
		var sFoot = "\n</table></body></html>";
		var sBody = '\n<table bgcolor="#EEEEEE" border="1" cellspacing="1" cellpadding="2" align="center" bordercolor="#000000">';
		var oRe = new RegExp("[object]","ig")
		
		for(var sItem in Obj){
			sBody = sBody + "\n<tr><td> " + sItem + " </td>";
			if(Obj[sItem] == null || Obj[sItem] == "") sBody = sBody + "<td>&nbsp;</td></tr>";
			else if(typeof(Obj[sItem]) == "object"){
				sAnchor = '<a href="#" nclick="htm_dbg_showobj(' + Obj[sItem].nodeName + ')">' + Obj[sItem] + '</a>';
				sBody = sBody + "<td>" + sAnchor + " </td></tr>";
			}
			else sBody = sBody + "<td>" + Obj[sItem] + " </td></tr>";
		}
		var aStream = (sHead + sBody + sFoot).split("\n");
		for(var i = 0, len = aStream.length; i < len; i++){
			oFout.WriteLine(aStream[i]);
		}		
		oFout.Close();
		js_shl_open(sFile)
		return true;
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_log_error(iOpt,oErr){
	try{
		js_log_error(iOpt,oErr);
	}
	catch(e){
		try{
			WScript.Echo(oErr.description);
		}
		catch(ee){
			alert(oErr.description);
		}
	}
}
