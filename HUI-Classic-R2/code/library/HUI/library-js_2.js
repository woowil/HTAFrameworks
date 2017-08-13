// nOsliw Solutions HUI, HTML Application Library https://github.com/woowil/HTAFrameworks
//**Start Encode**

/*

File:     library-js_2.js
Purpose:  Development script
Author:   Woody Wilson
Created:  2006-04
Version:
Dependency:  library-js_1.js

Description:
JScript library scripting functions used for WSH files or HTA Applications.

REMEMBER THAT PROTOTYPES ONLY WORKS 100% WHEN SCRIPT FILES ARE INCLUDED AS SRC FILES
I.E, NOT DIRECTLY IN THE SCRIPT IN THE HTA FILE

Disclaimer:
This sample code is provided AS IS WITHOUT WARRANTY
OF ANY KIND AND IS PROVIDED WITHOUT ANY IMPLIED WARRANTY
OF FITNESS FOR PURPOSES OF MERCHANTABILITY. Use this code
is to be undertaken entirely at your risk, and the
results that may be obtained from it are dependent on the user.
Please note to fully back up files and system(s) on a regular
basis. A failure to do so can result in loss of data or damage
to systems.

*/


// Global Scripting Objects
try{
	if(!oFso || !oWsh || !oWno){
		var oFso = new ActiveXObject("Scripting.FileSystemObject");
		var oWsh = new ActiveXObject("WScript.Shell");
		var oWno = new ActiveXObject("WScript.Network");
	}
}
catch(ee){
	var oFso = new ActiveXObject("Scripting.FileSystemObject");
	var oWsh = new ActiveXObject("WScript.Shell");
	var oWno = new ActiveXObject("WScript.Network");
}

// http://www.webreference.com/programming/javascript/gr/column9/2.html
// http://www.webreference.com/programming/javascript/gr/column18/index.html
// http://phrogz.net/JS/Classes/ExtendingJavaScriptObjectsAndClasses.html#example4
// http://www.bigbold.com/snippets/tag/javascript/5

var LIB_NAME	= "Main Library for prototypes";
var LIB_VERSION = "1.0";
var LIB_FILE	= oFso.GetAbsolutePathName("library-js_2.js")

var js2 = new js2_object();

function js2_object(){
	try{
		try{
			this.version = new js_objectversion(LIB_NAME,LIB_FILE,LIB_VERSION)
		}
		catch(ee){}
		this.REG_REGIONAL = "HKEY_USERS\\.DEFAULT\\Control Panel\\International"
		this.regional = false
		this.setRegional = function(){
			if(this.regional) return;
			try{
				Date.prototype.firstdayofweek = oWsh.RegRead(this.REG_REGIONAL + "\\iFirstDayOfWeek")
				Date.prototype.longdate = oWsh.RegRead(this.REG_REGIONAL + "\\sLongDate")
				Date.prototype.shortdate = oWsh.RegRead(this.REG_REGIONAL + "\\sShortDate")
				Date.prototype.longtime = oWsh.RegRead(this.REG_REGIONAL + "\\sTimeFormat")
				Date.prototype.shorttime = (Date.prototype.longtime).substring(0,5)
				Date.prototype.septime = oWsh.RegRead(this.REG_REGIONAL + "\\sTime")
				Date.prototype.sepdate = oWsh.RegRead(this.REG_REGIONAL + "\\sDate")
				Date.prototype.weekdays = 7
				Number.prototype.currency = oWsh.RegRead(this.REG_REGIONAL + "\\sCurrency")
				Number.prototype.country = oWsh.RegRead(this.REG_REGIONAL + "\\sCountry")				
				return true
			}
			catch(ee){
				//alert(ee.description)
				return false
			}
			this.regional = true
		}
	}
	catch(e){
		js_log_error(2,e);
		return false
	}
}

//_Window()
function _Window(){
	try{
		//this.show = function(){
			
		//}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

_Function()
function _Function(){

	Function.prototype.closure = function(obj){
		// Init object storage.
		if (!window.__objs){
			window.__objs = [];
			window.__funs = [];
		}

		// For symmetry and clarity.
		var fun = this;

		// Make sure the object has an id and is stored in the object store.
		var objId = obj.__objId;
		if (!objId)
		__objs[objId = obj.__objId = __objs.length] = obj;

		// Make sure the function has an id and is stored in the function store.
		var funId = fun.__funId;
		if (!funId)
		__funs[funId = fun.__funId = __funs.length] = fun;

		// Init closure storage.
		if (!obj.__closures)
		obj.__closures = [];

		// See if we previously created a closure for this object/function pair.
		var closure = obj.__closures[funId];
		if (closure)
		return closure;

		// Clear references to keep them out of the closure scope.
		obj = null;
		fun = null;

		// Create the closure, store in cache and return result.
		return __objs[objId].__closures[funId] = function (){
			return __funs[funId].apply(__objs[objId], arguments);
		};
	};
}

_Boolean()
function _Boolean(){
	try{
		Boolean.prototype.XOR = function(b2){
			//true.XOR(false); //returns a value of true
			var b1 = this.valueOf();
			return (b1 && !b2) || (b2 && !b1);
		}

		Boolean.prototype.AND = function(b2){
			if(typeof b2 != "boolean") return false
			var b1 = this.valueOf();
			return (b1 && b2);
		}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

_Number()
function _Number(){
	try{
		// http://devedge-temp.mozilla.org/library/manuals/2000/javascript/1.5/reference/number.html#1200968
		// http://www.irt.org/xref/Number.htm
		Number.prototype.isPositive = function(n){
			if(isNaN(n)) var n = this
			return (n >= 0 && n <= Number.MAX_VALUE)
		}
		
		Number.prototype.isNegative = function(n){
			if(isNaN(n)) var n = this
			return (n < 0 && n > Number.MIN_VALUE)
		}
		
		Number.prototype.toFixed = function(fractionDigits){
			var m = Math.pow(10,fractionDigits);
			return Math.round(this*m,0)/m;
		}

		Number.prototype.toExponential = function(fractionDigits){
		   var l = Math.floor(Math.log(this)/Math.LN10);
		   var lm = Math.pow(10,l);
		   var n = this / lm;
		   var s = fractionDigits ? n.toFixed(fractionDigits) : n.toString();
		   s += 'e' + ((l > 0) ? '+' : '-') + l;
		   return s;
		}

		Number.prototype.toPrecision = function(precision){
			var l = Math.floor(Math.log(this)/Math.LN10);
			var m = Math.pow(10,l + 1 - precision);
			return Math.round(this/m,0)*m;
		}
		
		Number.prototype.toDecimal = function(iDecimal){
			return Math.round((this-Math.floor(this))*Math.pow(10,iDecimal))/100
		}
		
		Number.prototype.toProcent = function(iDecimal){ // Equivalent to VBScript FormatPercent, NOT 100%
			if(!this.isPositive(iDecimal)) var iDecimal = 2
			var s = this*100
			s = Math.floor(s) + s.toDecimal(iDecimal)
			var s_num = Math.floor(s)
			var s_dec = Math.floor(Math.pow(10,iDecimal)*(s-s_num)) + ""	
			for(s = "";;){
				if(s_num < 1000){
					s = s_num + s
					break;
				}
				var t = s_num/1000				
				s_num = Math.floor(t)
				s = " " + Math.round(1000*(t-s_num)) + s
			}
			while(s_dec.length < iDecimal) s_dec = s_dec + "0"
			return s + "," + s_dec + "%"
		}
		
		Number.prototype.toCurrency = function(n){ // Equivalent to VBScript FormatCurrency, NOT 100%
			if(!this.isPositive(iDecimal)) var iDecimal = 2
			var s = Math.floor(this) + this.toDecimal(iDecimal)
			var s_num = Math.floor(s)
			var s_dec = Math.floor(Math.pow(10,iDecimal)*(s-s_num)) + ""
			for(s = "";;){
				if(s_num < 1000){
					s = s_num + s
					break;
				}
				var t = s_num/1000
				s_num = Math.floor(t)
				s = " " + Math.round(1000*(t-s_num)) + s
			}
			while(s_dec.length < iDecimal) s_dec = s_dec + "0"
			js2.setRegional()
			return s + "," + s_dec + this.currency
		}
		
		Number.prototype.toHex = function(r){
			if(typeof(r) != "number") r = 16
			return this.toString(r)
		}
		
		Number.prototype.toDec = function(r){
			if(typeof(r) != "number") r = 16
			return parseInt(this,r) // Supposing this is a Hex number
		}
		
		Number.prototype.isHex = function(n){

		}

		Number.prototype.isOct = function(n){

		}

		Number.prototype.isDec = function(n){

		}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

_String()
function _String(){
	try{
		String.prototype.isEqual = function(s){
			return (this.valueOf() === s) // Identical
		}

		String.prototype.encode = function(c,s){
			// http://www.htmlhelp.com/reference/charset/iso032-063.html
			// convert to hex to get the charactor
			if(typeof(s) != "string") var s = this
			c = typeof(c) == "number" ? parseInt(c) : (s.length % 10)
			var space = "\u0020", alpha = "0123456789abcdefghijklmnopqrstuvwxyz._~ABCDEFGHIJKLMNOPQRSTUVWXYZ-+$\\";
			// \xE5\xE6\xF8\xC5\xC6\xD8 æ,ø,Å,Æ,Ø (Norwegian characters)
			var sConvert = ""
			for(var i = 0, len = s.length; i < len; i++){
				var sChar = Cipher = s.substring(i,i+1);
				if(sChar != " " && sChar != space){
					Conv = alpha.indexOf(sChar), Cipher = Conv^c;
					Cipher = alpha.substring(Cipher,Cipher+1);
				}
				sConvert = sConvert + Cipher
			}
			return sConvert;
		}

		String.prototype.trim = function(s){
			var ii
			if(typeof(s) != "string") var s = this
			s = s.replace(/([ \t\n]*)(.*)/g,"$2"); // Removes space or newline at beginning and end
			if((ii = s.search(/([ \t\n]*$)/m)) > 0) s = s.substring(0,ii);
			return s.replace(/[ \t]+/ig," "); // Replaces tabs, double spaces
		}

		String.prototype.capatilize = function(s){
			if(typeof(s) != "string") var s = this
			var a = s.split(/[ \t]+/g)
			for(var i = 0, len = a.length; i < len; i++){
				var C = ((a[i]).substring(0,1)).toUpperCase()
				var R = new String(((a[i]).substring(1,(a[i]).length)))
				R = C.search(/[0-9]/) > -1 ? R : R.toLowerCase()
				if(i == 0) s = C + R
				else s = s + " " + C + R
			}
			return s
		}
		
		String.prototype.isSearch = function(oRe){
			return this.valueOf().search(oRe) > -1
		}	
		

		String.prototype.shuffle = function(s){
			if(typeof(s) != "string") var s = this
			var a = new Array()
			for(var i = s.length; i;) a.push(s.charCodeAt(--i)) // Get the UniCode number
			for(var i = a.length; i; ){  // code based on Fisher-Yates algorithm
				var j = parseInt(Math.random() * i)
				var x = a[--i]
				a[i] = a[j]
				a[j] = x
			}
			for(var s = "", i = a.length; i;){
				s = s + String.fromCharCode(a[--i]) // Get the UniCode character
				delete a[i]
			}
			delete a
			return s;
		};
		
		String.prototype.random = function(s,l){
			if(typeof(s) != "string") var s = this
			if(typeof(l) != "number" || l > s.length) var l = s.length-1
			return s.charAt(parseInt(Math.random() * l))
		}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

_Date()
function _Date(){
	try{
		var lang = "sv"
		
		Date.prototype.sepdate = "-"
		Date.prototype.septime = ":"
		
		Date.prototype.dtGeneralDate = 0
		Date.prototype.dtLongDate = 1
		Date.prototype.dtShortDate = 2
		Date.prototype.dtLongTime = 3
		Date.prototype.dtShortTime = 4
				
		Date.prototype.isDate = function(o){
			try{
				if(!o || typeof(o) != "object") return false
				// http://blogs.msdn.com/ericlippert/comments/53352.aspx
				return (o.IsPrototypeOf(Date) || o.constructor == Date || o.constructor == _Date || o instanceof _Date)
			}
			catch(ee){
				return false
			}
		}

		Date.prototype.setLang = function(l){
			lang = l
			if(l == "sv") this.sepdate = "-"
			else if(l == "no") this.sepdate = "."
			else if(l == "en") this.sepdate = "\\"
		}

		Date.prototype.nDigits = function(d,n){
			if(typeof n != "number") return d
			for(var d = "" + d; d.length < n; ) d = "0" + d;
			return d;
		}

		////

		Date.prototype.dayHours = function(d){
			if(typeof d != "number") var d = 1
			return Math.abs(24*d);
		}

		Date.prototype.dayMinutes = function(d){
			return this.dayHours(d)*60;
		}

		Date.prototype.daySeconds = function(d){
			return this.dayMinutes(d)*60;
		}

		Date.prototype.dayMilliseconds = function(d){
			return this.daySeconds(d)*1000;
		}

		////		
		
		Date.prototype.formatDateSV = function(){
			return this.getYear() + "-" + this.nDigits(this.getMonth()+1,2) + "-" + this.nDigits(this.getDate(),2)
		}

		Date.prototype.formatDateNO = function(){
			return this.getDate() + "." + this.nDigits(this.getMonth()+1,2) + "." + this.nDigits(this.getYear(),2)
		}

		Date.prototype.formatDateEN = function(){
			return this.getDate() + "/" + this.nDigits(this.getMonth()+1,2) + "/" + this.nDigits(this.getYear(),2)
		}

		Date.prototype.formatYYYYMMDD = function(d,bSep){
			if(!this.isDate(d)) var d = this
			var s = bSep ? this.sepdate : ""
			return d.getYear() + s + this.nDigits(d.getMonth()+1,2) + s + this.nDigits(d.getDate(),2)
		}

		Date.prototype.formatYYYYMM = function(d,bSep){
			if(!this.isDate(d)) var d = this
			return (d.formatYYYYMMDD(bSep)).substring(0,6)
		}

		Date.prototype.formatHHMMSS = function(d,bSep){
			if(!this.isDate(d)) var d = this
			var s = bSep ? this.septime : ""
			return this.nDigits(d.getHours(),2) + s +  this.nDigits(d.getMinutes(),2) + s + this.nDigits(d.getSeconds(),2)
		}
		
		Date.prototype.formatMMSS = function(d,bSep){
			if(!this.isDate(d)) var d = this
			return (d.formatHHMMSS(bSep)).substr(2)
		}

		Date.prototype.formatHHhMMmSSs = function(d){
			if(!this.isDate(d)) var d = this
			return this.nDigits(d.getHours(),2) + "h " + this.nDigits(d.getMinutes(),2) + "m " + this.nDigits(d.getSeconds(),2) + "s"
		}

		Date.prototype.formatMMmSSs = function(d){
			if(!this.isDate(d)) var d = this
			return (d.formatHHhMMmSSs()).substr(2)
		}

		//// 64Bits date from t.ex logintimestamp
		
		Date.prototype.getTime64 = function(o){
			try{
				var d = new Date(1601,0,1,0,0,0) //				
				// represents number of 100 nanoseconds since 1601-01-01
				var iTime = o.HighPart * Math.pow(2,32) + o.LowPart // oo = 64 bits number
				iTime = iTime / (60*10000000)
				iTime = iTime / 1440			
				var iTimeMilli = Math.ceil(iTime*24*3600*1000)
				d.setTime(d.getTime()+iTimeMilli)
				return d
			}
			catch(e){
				return undefined
			}
		}

		
		////
		
		Date.prototype.formatDateTime = function(d,iFormat,bMilli){
			if(!this.isDate(d)) var d = this
			if(!isNaN(iFormat) && iFormat >= 0 && iFormat <= 4){ // Equivalent to VBScript function FormatDateTime
				js2.setRegional()
				var s = ""
				if(iFormat == this.dtGeneralDate) s = this.shortdate // 21.04.2006
				else if(iFormat == this.dtLongDate) s = this.longdate // 21. april 2006
				else if(iFormat == this.dtShortDate) s = this.shortdate // 21.04.2006
				else if(iFormat == this.dtLongTime) s = this.longtime // 00:00:00
				else if(iFormat == this.dtShortTime) s = this.shorttime // 00:00
				
				s = s.replace(/dd/,this.nDigits(d.getDate(),2))
				s = s.replace(/d/,d.getDate())
				s = s.replace(/MMMM/,this.aMonths("no")[d.getMonth()])
				s = s.replace(/MM/,this.nDigits(d.getMonth(),2))
				s = s.replace(/yyyy/,d.getYear())
				//s = s.replace(/yy/,d.getDate())
				s = s.replace(/HH/,this.nDigits(d.getHours(),2))
				s = s.replace(/mm/,this.nDigits(d.getMinutes()+1,2))
				s = s.replace(/ss/,this.nDigits(d.getSeconds(),2))
			}
			else var s = d.formatYYYYMMDD() + " " + d.toLocaleTimeString()//d.formatHHMMSS(false,true)
			if(bMilli) s = s + this.septime + this.nDigits(d.getMilliseconds(),3)
			return s
		}

		Date.prototype.formatDateString = function(d){
			if(!this.isDate(d)) var d = this
			return this.aDays()[d.getDay()] + ", " + d.getDate() + " " + this.aMonths()[d.getMonth()] + " " + d.getYear() + ", " + d.formatHHMMSS(null,true)
		}		
		
		////

		Date.prototype.getDateDiff = function(d_old,d_new){
			if(!this.isDate(d_new)) var d_new = this
			var d = new Date()
			d.setTime(Math.abs(d_new.getTime()-d_old.getTime()))
			d.setHours(d.getHours()-1)
			return d
		}

		Date.prototype.getLastDay = function(d){
			if(!this.isDate(d)) var d = this
			d.setTime(d.getTime()-this.dayMilliseconds())			
			return d
		}

		Date.prototype.getFirstDayOfWeek = function(d){
			if(!this.isDate(d)) var d = this
			js2.setRegional()
			var dm = this.dayMilliseconds()
			for(var i = this.getDay(); i >= this.firstdayofweek; i--){
				d.setTime(d.getTime()-dm)
			}
			return d
		}
		
		Date.prototype.setFirstDayOfYear = function(d){
			if(!this.isDate(d)) var d = this
			js2.setRegional()
			d.setMonth(0)
			d.setDate(1)
			d.setHours(0)
			d.setMinutes(0)
			d.setSeconds(0)
			d.setMilliseconds(0)
			return d
		}
		
		Date.prototype.getFirstDayOfYear = function(d){
			if(!this.isDate(d)) var d = this
			js2.setRegional()
			return (new Date(d.getYear(),0,1,0,0,0))
		}
		
		Date.prototype.getLastWeekDay = function(d){
			if(!this.isDate(d)) var d = this
			var d2 = new Date()
			d2.setTime(d.getTime())
			var dm = this.dayMilliseconds()
			for(var i = d.getDay(); i < this.weekdays; i++){
				d2.setTime(d2.getTime()+dm)
			}
			return d2
		}
		
		Date.prototype.getDays = function(d){
			if(!this.isDate(d)) var d = this
			return d/this.dayMilliseconds()
		}
		
		Date.prototype.getWeek = function(d){			
			if(!this.isDate(d)) var d = this
			var i = (d.getTime()-(d.getFirstDayOfYear()).getTime())/this.dayMilliseconds()/this.weekdays
			return Math.floor(i+1)
		}
		
		Date.prototype.getDateFirstMonthDay = function(d){
			if(!this.isDate(d)) var d = this
			var dm = this.dayMilliseconds()
			for(var i = this.getDate(); i >= 1; i--){
				d.setTime(d.getTime()-dm)
			}
			return d
		}

		Date.prototype.getDateFirstMonday = function(d){
			if(!this.isDate(d)) var d = this
			var dm = this.dayMilliseconds()
			for(var i = this.getDate(); i >= 1; i--){
				if(i < this.weekdays && d.getDay() == 0) break;
				d.setTime(d.getTime()-dm)
			}
			return d
		}

		////

		Date.prototype.aMonths = function(l){
			if(typeof(l) != "string") var l = "en"
			if(l == "sv") return ['Januari','Februari','Mars','April','Maj','Juni','Juli','Augusti','September','Oktober','November','December'];
			else if(l == "no") return ['Januar','Februar','Mars','April','Mai','Juni','Juli','August','September','Oktober','November','Desember'];
			else return ['January','Febuary','March','April','May','June','July','August','September','October','November','December'];
		}

		Date.prototype.aDays = function(l){
			if(typeof(l) != "string") var l = "en"
			if(l == "sv") return ['Söndag','Måndag','Tisdag','Onsdag','Torsdag','Fredag','Lördag'];
			else if(l == "no") return ['Søndag','Mandag','Tirsdag','Onsdag','Torsdag','Fredag','Lørdag'];
			else return ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
		}
		
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

// Must create array prototype like this. If not, "for(i in o)"will include the inner functions
function _Array(){
	try{ // http://phrogz.net/JS/Classes/ExtendingJavaScriptObjectsAndClasses.html#example4
		this.isArray =  function(o){
			try{
				if(!o || typeof(o) != "object") return false
				// http://blogs.msdn.com/ericlippert/comments/53352.aspx
				return (o.IsPrototypeOf(Array) || o.constructor == _Array || o instanceof Array)
			}
			catch(ee){
				return false
			}
		}

		if(!this.push){ // IE5.0 or lower does not support push
			this.push = function(o){
				this[this.length] = o
				return this
			}
		}

		this.toDict = function(a){
			var d = new ActiveXObject("Scripting.Dictionary")
			if(this.isArray(a)){
				for(var i in a){
					d.Add(new String(a[i]),"d"+i)
				}
			}
			return d
		}

		this.sort2 = function(oo){
			oo = oo ? oo : this.array1
			var o1 = new Array(), o
			for(o in oo){
				o1.push(o)
			}
			o1.sort()
			var o2 = {}
			for(var i = 0; i < o1.length; i++){
				o2[o1[i]] = oo[o1[i]]
			}
			return o2
		};

		if(this.splice && typeof([0].splice(0))=="number") this.splice = null;
		if(!this.splice) this.splice = function(ind,cnt){
			var len = this.length;
			var arglen = arguments.length;
			if(arglen == 0) return ind;
			if(typeof(ind) != "number") ind = 0;
			else if(ind < 0) ind = Math.max(0,len+ind);
			if(ind > len){
				if(arglen > 2) ind = len;
				else return [];
			}
			if(arglen < 2) cnt = len-ind;
			cnt = (typeof(cnt) == "number") ? Math.max(0,cnt) : 0;
			var removeArray = this.slice(ind,ind+cnt);
			var endArray = this.slice(ind+cnt);
			len = this.length = ind;
			for(var i = 2; i < arglen; i++) this[len++] = arguments[i];
			for(var i = 0, len2 = endArray.length; i < len2; i++) this[len++] = endArray[i];
			return removeArray;
		}

		this.unioun = function(a2,bNoSort,bNoCase){
			if(!this.isArray(a2)) return new _Array()
			a2 = this.concat((this.isArray(a2)?a2:new _Array()))
			if(!bNoSort) a2.sort();
			// Removes duplicates
			for(var i = 0, len = a2.length; i < len; i++){
				if(!bNoCase){
					while(a2[i] == a2[i+1]) this.splice(i+1,1), i++;
				}
				else{
					try{
						while((a2[i]).toUpperCase() == (a2[i+1]).toUpperCase()) this.splice(i+1,1), i++;
					}
					catch(ee){}
				}
			}
			return a2
		};

		this.subtract = function(a1,a2){ // a1-a2
			if(!this.isArray(a2)) return a1
			var d = this.toDict(a1)
			for(var i = 0; i < a2.length; i++){
				if(d.Exists(a2[i])){
					d.Remove(a2[i]);
				}
			}
			return (new VBArray(d.Items()).toArray());
		};

		this.shuffle = function(a){ // code based on Fisher-Yates algorithm
			for(var i = a.length; i; ){
				var j = parseInt(Math.random() * i)
				var x = a[--i]
				a[i] = a[j]
				a[j] = x
			}
			return a;
		};

	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}
// 'inherit' from Base by assigning an instance of Base to the prototype
_Array.prototype = new Array();
_Array.prototype.constructor = _Array;

//_Object()
function _Object(){
	try{ // http://devedge-temp.mozilla.org/library/manuals/2000/javascript/1.5/reference/object.html
		// DO NOT USE "Object.prototype" here!! cause the for(o in oo) problem!
		this.show = function(){
			var s = ""
			s = s + "\ntoSource:\t" + this.toSource() //
			s = s + "\ntoString:\t" + this.toString()
			return s
		}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function _List(){
	try{

	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}
_List.prototype = {}
_List.prototype.constructor = _List

function _Class(sClass,sDescription){
	try{
		this.wmi_cimv2 = null
		this.wmi_default = null
		this.wmi_microsofthis = null
		this.wmi_locator = null
		this.computer = null
		this.username = null
		this.password = null
		this.POPUP_WAIT = 15
		this.isStatus = false
		this.isError = true
		this.isDebug = false
		this.isUseLibrary = true
		this.objLogResult = false
		this.objLogError = false
		this.funcScroll = false
		this.isLocalhost = false
		this.class_name = typeof(sClass) == "string" ? sClass.toLowerCase() : ""
		this.class_description = typeof(sDescription) == "string" ? sDescription : ""
		this.name = "_Class"
		this.bInitService = true
		
		// Error
		this.err_number
		this.err_description
		this.err_function

		try{
			WScript.Sleep(1);
			this.isWSScript = true
		}
		catch(ee){
			this.isWSScript = false
		}
		this.subkey = {}
		this.subkey.env = "HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment"
		this.subkey.con = "HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control"
		this.subkey.mem = "HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management"
		
		this.fso = {}
		// oFso.Run()
		this.fso.read = 1
		this.fso.hide = 0
		this.fso.write = 2
		this.fso.TristateUseDefault = -2
		// Window shell directory - https://windowsxp.mvps.org/usersshellfolders.htm
		this.fso.windir = oFso.GetSpecialFolder(0);
		this.fso.system32 = oFso.GetSpecialFolder(1);
		this.fso.temp = oFso.GetSpecialFolder(2);
		this.fso.appdata_all = oWsh.ExpandEnvironmentStrings("%allusersprofile%") + "\\Application Data"
		this.fso.appdata = oWsh.ExpandEnvironmentStrings("%userprofile%") + "\\Application Data"
		this.fso.w32 = oFso.GetSpecialFolder(1) + "\\spool\\drivers\\w32x86";

		this.init = function(oService){
			try{
				this.reset()
				this.echo("## Initializing " + this.class_name + " class prototype on system: " +  this.computer)
				if(this.bInitService){
					if(oService != null && typeof(oService) == "object") this.setService(oService,this.computer,this.username,this.password,true)
					else this.setService(null,this.computer,this.username,this.password,true)
				}
			}
			catch(ee){
				this.error(ee,"init()")
			}
		}

		this.setService = function(oService,sComputer,sUser,sPass,bForce){
			try{
				if(!this.wmi_cimv2 || bForce){
					if(oService) this.wmi_cimv2 = oService
					else {
						this.computer = typeof(sComputer) == "string" ? sComputer : oWno.ComputerName
						this.username = typeof(sUser) == "string" ? sUser : null
						this.password = typeof(sPass) == "string" ? sPass : null
						if(!this.wmi_locator) this.wmi_locator = new ActiveXObject("WbemScripting.SWbemLocator");
						this.wmi_cimv2 = this.wmi_locator.ConnectServer(this.computer,"root\\cimv2",this.username,this.password);
					}
				}
				this.isLocalhost = this.isHostname()
				this.setText()
				return this.wmi_cimv2
			}
			catch(ee){
				this.error(ee,"setService()")
				return null
			}
		}

		this.isHostname = function(){
			if(!this.isString(this.computer)) return false;
			return (this.computer.toUpperCase() == oWno.ComputerName.toUpperCase());
		}

		//// STRINGS

		this.setText = function(){

		}

		this.isString = function(args){
			for(var s, i = 0, len = arguments.length; i < len; i++){
				if(typeof(s = arguments[i]) == "string" && (s.replace(/[ \t]*(.*)[ \t]*$/g,"$1")).length > 0) continue; // "" is not a string
				else return false
			}
			return true
		}
		
		this.isNumber = function(args){
			for(var s, i = 0, len = arguments.length; i < len; i++){
				if(typeof(s = arguments[i]) == "number") continue;
				else return false
			}
			return true
		}
		
		this.trim = function(s){
			if(typeof s != "string") return ""
			var ii
			s = s.replace(/([ \t\n]*)(.*)/g,"$2"); // Removes space at beginning
			if((ii = s.search(/([ \t\n]*$)/m)) > 0) s = s.substring(0,ii);
			return s.replace(/[ \t]+/ig," "); // Replaces tabs, double spaces
		}

		//// ARRAYS

		this.getVBArray = function(vbArray){
			try{
				return (new VBArray(vbArray)).toArray()
			}
			catch(ee){
				return false
			}
		}

		this.setVBArray = function(aJSArray){ // Excellence!!
			var d = new ActiveXObject("Scripting.Dictionary")
			for(var i = j = 0, len = aJSArray.length; i < len; i++){
				if(arguments[i] != null) d.Add("d" + j++,arguments[i])
			}
			return d.Items()
		}
		
		// WINDOW
		
		this.winFullscreen = function(){
			if(!this.winIsFullScreen){
				this.winX = window.screenLeft;
				this.winY = window.screenTop;
				this.winWidth = document.body.offsetWidth + 8
				this.winHeight = document.body.offsetHeight
				window.moveTo(0,0);
				window.resizeTo(window.screen.width,window.screen.height)				
				this.winIsFullScreen = true
			}
			else {
				window.moveTo(this.winX,this.winY);
				window.resizeTo(this.winWidth,this.winHeight);
				this.winIsFullScreen = false
			}
		}
		
		// WMI

		this.wmiDate = function(sDate){
			try{
				if(sDate != null){ // DateTime object
					sDate = sDate.replace(/[\*]/g,0);
					var sDateTime = sDate.substring(0,4) + "-" + sDate.substring(4,6) + "-" + sDate.substring(6,8) + " " + sDate.substring(8,10) + ":" + sDate.substring(10,12) + ":" + sDate.substring(12,14) + "." + sDate.substring(15,18);
					return sDateTime;
				}
				else return false;
			}
			catch(ee){
				return false;
			}
		}

		this.wmiArray = function(aVBarray,sVBName,sBreak){
			try{
				if(aVBarray != null){ // VBArray object
					var aVB = new VBArray(aVBarray).toArray();
					var aArray = new Array()
					sBreak = sBreak ? sBreak : "\n  "
					for(var j = 0, sVB = "", len = aVB.length; j < len; j++){
						 //sVB = sVB + (j == 0 ? "" : sBreak) + sVBName + "-" + (j+1) + ":  " + aVB[j];
						 sVB = sVB + (j == 0 ? "" : sBreak) + aVB[j];
						 aArray.push(aVB[j])
					}
					aVB.stream = (j == 1) ? sVB.replace(/.+-[0-9]{1,2}:  (.+)$/ig,"$1") : sVB;
					//aVB.array = aArray;
					return aVB;
				}
				return false;
			}
			catch(ee){
				return false;
			}
		}

		// ERROR AND LOGS

		this.echo = function(sMessage,sBreak){
			if(typeof(sMessage) != "string") return;
			if(this.isStatus){
				try{					
					if(this.isUseLibrary) js_log_print("log_result",sMessage,sBreak)
					else if(this.objLogResult){
						var d = (new Date()).formatDateTime(), a
						if((a = sMessage.split(/[\n\r]/g)) && (len = a.length) > 1){
							for(var i = 0, s = ""; i < a.length; i++){
								s = s + "\n[" + d + "] " + a[i];
							}
							this.kill(a)
							sMessage = s
						}
						else sMessage = "[" + d + "] " + sMessage
						sBreak = sBreak ? sBreak : "\n"
						this.objLogResult.innerText += sBreak + sMessage
						if(this.funcScroll) this.funcScroll(this.objLogResult)
					}
				}
				catch(ee){
					WScript.Echo(sMessage)
				}
			}
		}

		this.popup = function(sMessage){
			if(!this.isString(sMessage)) return;
			try{
				alert(sMessage)
			}
			catch(ee){
				oWsh.Popup(sMessage,this.POPUP_WAIT,this.class_title,32 + 1)
			}
		}

		this.reset = function(){
			this.err_number = this.err_description = this.err_function = ""
			
			/*
			this.computer = typeof(this.computer) == "string" ? this.computer : oWno.ComputerName
			this.username = typeof(this.username) == "string" ? this.username : null
			this.password = typeof(this.password) == "string" ? this.password : null
			*/
		}

		this.sleep = function(iMilli){
			if(typeof iMilli != "number") return;
			if(this.isWSScript) WScript.Sleep(iMilli);
			else oWsh.Run("%comspec% /c sleep -m " + iMilli + ">nul",this.fso.hide,true)
		}

		this.random = function(hval,lval){
			if(typeof hval != "number") return hval;
			if(typeof lval != "number") lval = 0;
			return Math.floor(Math.random()*hval + lval)
		}
		
		this.matrix2D = function(x,y,c){
			if(typeof x != "number") x = 0;
			if(typeof y != "number") y = 0;
			c = c ? c : "#"
			var a = new Array()
			for(var i = 0; i < x; i++){
				a[i] = new Array()
				for(var j = 0; j < y; j++){
					a[i][j] = c
				}
			}
			return a
		}
		
		this.about = function(){

		}

		this.error = function(oErr,sFunc){
			this.err_number = (oErr.number & 0xFFFF0000)
			this.err_description = this.trim(oErr.description)
			this.err_function = sFunc = this.name + "::" + sFunc
			this.ERR_WMI_ACCESSDENIED = this.err_number == -2147121510 ? true : false
			
			if(this.isError){
				try{
					if(this.isUseLibrary) js_log_error(2,oErr,sFunc)
					else if(this.objLogError){
						this.objLogError.innerText += "\n\n" + oErr.getErrorText(false,sFunc)
						if(this.funcScroll) this.funcScroll(this.objLogError)
					}
				}
				catch(ee){					
					this.echo("## " + this.class_name + " => " + sFunc + ": " + this.err_number + " " + oErr.description)					
					switch(ee.number) {
						case 462 :
						case -2146827826 : // The remote server machine does not exist or is unavailable
							break;
						case 429 :
						case -2146827859 : // ActiveX component can't create object							
							break;
						case -2147217405 : // Missing permissions to connect to this host
							break;
						case -2146828218 : // "Permission denied" - DCOM not enabled error message
							break;
						case -2147418111 :
							break;
						case -2147217394 : // "Unable to connect to namespace: ********"
							break;
						case 4099 :
						case -2147121510 : // Access Denied
						default : break;
					}
				}
			}
		}
		
		this.kill = function(){
			try{
				for(var i = 0, l = arguments.length; i < l; i++){
					if(typeof arguments[i] == "undefined" || arguments[i] == null) continue
					else if(arguments[i] instanceof Function) continue
					else if(arguments[i] instanceof Array){
						for(var j = 0, l2 = arguments[i].length; j < l2; j++){
							if(arguments[i][j] instanceof Object){
								for(var o in arguments[i][j]) delete arguments[i][j][o]
							}
							delete arguments[i][j]
						}
						arguments[i].length = 0
					}
					else if(arguments[i] instanceof Object){				
						if(!(arguments[i] instanceof Enumerator) && !(arguments[i] instanceof Date)){
							for(var o in arguments[i]) this.kill(arguments[i][o])
						}
					}
					else if(typeof arguments[i] == "object"){
						if(typeof arguments[i].RemoveAll == "unknown") arguments[i].RemoveAll()
					}
					delete arguments[i] // delete: http://users.adelphia.net/~daansweris/js/special_operators.html
				}
				return true;
			}
			catch(e){
				return false;
			}
		}
		
		this.close = function(){
			this.kill(this)
		}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

_Error()
function _Error(){
	try{
		Error.prototype.isError =  function(o){
			try{
				if(!o || typeof(o) != "object") return false
				// http://blogs.msdn.com/ericlippert/comments/53352.aspx
				return (o.IsPrototypeOf(Error) || o.constructor == Error || o instanceof Error)
			}
			catch(ee){
				return false
			}
		}
		
		Error.prototype.isErrorWMI = function(){
			return true
		}

		////

		Error.prototype.getCode = function(){
			// An error number is a 32-bit value.
			// The upper 16-bit word is the facility code, while the lower word is the actual error code.
			return (this.number & 0xFFFF)
		}

		Error.prototype.getFacility = function(){
			return (this.number >> 16 & 0x1FFF)
		}

		Error.prototype.getMessage = function(){
			return this.message
		}

		Error.prototype.getSource = function(){
			return (this.source ? this.source : "") // Only in VBScript
		}

		Error.prototype.getCallers = function(f,bArgs){
			var a = new Array()
			if(typeof(f) != "function") return a
			while(true){
				a.push(this.getCaller(f,bArgs))
				if(!f.caller) break;
				f = f.caller
			}
			return a
		}

		Error.prototype.getCaller = function(f,bArgs){
			if(typeof(f) != "function") return null
			var oRe = /[ \t]*function[ \t]+([a-z0-9_]+)[ \t]*([a-z0-9_, \t\(\)]{2,})[ \t]*[{]*.*$/ig
			if((f.toString().split("\n")[0]).search(oRe) > -1) return RegExp.$1 + (bArgs ? (RegExp.$2).replace(/ /g,"") : "()")
			return null
		}

		Error.prototype.caller = null
		
		Error.prototype.getResource = function(){
			try{
				this.resource = js.resource
			}
			catch(e){
				this.resource = oWno.CompterName
			}
			return this.resource
		}
		
		Error.prototype.getErrorText = function(bArgs,sFunc){
			if(!this.caller) this.caller = this.getErrorText.caller
			var s = "\n"
			s = s + "Resource:    " + this.getResource() // can globally be defined... can be computer, user, file etc
			s = s + "\nDate:        " + (new Date()).formatDateTime()
			s = s + "\nType:        " + this.toString()
			s = s + "\nName:        " + this.name
			s = s + "\nCode:        " + this.getCode()
			s = s + "\nFacility:    " + this.getFacility()
			s = s + "\nDescription: " + this.description
			if(sFunc) s = s + "\nFunction:    " + sFunc
			var a = this.getCallers(this.caller,bArgs,sFunc)
			s = s + "\nTrace:       " + a[0]
			for(var i = 1; i < a.length; i++) s = s + "\n             " + a[i]
			return s
		}

		Error.prototype.getErrorObject = function(bArgs,sFunc){
			if(!this.caller) this.caller = this.getErrorObject.caller
			var o = {}
			o.Resource = this.getResource()
			o.Date = (new Date()).formatDateTime()
			o.Type = this.toString()
			o.Name = this.name
			o.Number = o.Code = this.getCode()
			o.Facility = this.getFacility()
			o.Description = this.description
			o.Func = sFunc ? sFunc : ""
			o.Trace = this.getCallers(this.caller,bArgs)
			return o
		}

	}
	catch(e){
		try{
			alert(e.description);
		}
		catch(ee){
			WScript.Echo(e.description);
		}
		return false;
	}
}
