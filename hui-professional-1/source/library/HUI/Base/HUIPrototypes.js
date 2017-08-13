// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**


///////////////////////////////////////////////////////////////////////////////////////////
////////// PROTOTYPES FOR OBJECT
///////////////////////////////////////////////////////////////////////////////////////////


if(!Object.prototype.hasOwnProperty){
	Object.prototype.hasOwnProperty = function(property){
		try{
			var prototype = this.constructor.prototype;
			while(prototype){
				if(prototype[property] == this[property]){
					return false;
				}
				prototype = prototype.prototype;
			}
			return true;
		}
		catch(e){
			return false;
		}
	}
}

Object.extend = function(destination,source,overwrite){
	if(!(source instanceof Object)) throw new Error(8881,"Argument [1] is not an Object. Unable to extend.")
	if(destination == null){
		destination = this
	}
	for(var property in source){
		if((overwrite || !destination[property])){
			destination[property] = source[property]
		}
	}
	return destination
}

Object.extend(Object.prototype,{
	isObject : function(o){
		return o && typeof(o) == "object" && (o.constructor == Object || o instanceof Object)
	},
	
	// extend : function(destination,source,overwrite){
		// if(!this.isObject(source)) throw new Error(8881,"Argument [1] is not an Object. Unable to extend.")
		// if(destination == null){
			// destination = this
		// }
		// for(var property in source){
			// if((overwrite || !destination[property])){
				// destination[property] = source[property]
			// }
		// }
		// return destination
	// },
	
	hasProperty : function(property){
		return (!(property instanceof Function) && !!this[property])
	},
		
	toJSON : function(){
		if(this.valueOf() === null) return 'null';

		var a = [];
		for(var property in this){
			if(!this.hasOwnProperty(property) || !this[property]) continue
			a.push('"' + property + '":' + __H.toJSON(this[property]))
		}

		return '{' + a.join(', ') + '}';
	},

	keys : function(){
		var a = []
		for(var property in this){
			if(this.hasOwnProperty(property) && typeof(this[property]) != "function"){
				a.push(property)
			}
		}
		return a
	},

	values : function(){
		var a = []
		for(var property in this){
			if(this.hasOwnProperty(property) && typeof(this[property]) != "function"){
				a.push(this[property])
			}
		}
		return a
	},

	methods : function(){
		var a = []
		for(var property in this){
			if(this.hasOwnProperty(property) && typeof(this[property]) == "function"){
				a.push(property)
			}
		}
		return a
	},

	empty : function(){
		for(var property in this){
			if(!this.hasOwnProperty(property)) continue
			delete this[property]
		}
	},

	clone : function(){
		return {}.extend(null,this)
	},

	compact : function(bReplace){
		var o = !bReplace ? this.clone() : this
		for(var property in o){
			if(!o.hasOwnProperty(property)) continue
			if(o[property] == null || typeof(o[property]) == "undefined"){
				delete o[property]
			}
		}
		return o
	},

	clear : function(){
		return this.empty()
	},

	sort : function(){ // NOT WORKING!!
		var a = []
		var d = new ActiveXObject("Scripting.Dictionary")
		for(var i in this){
			if(!this.hasOwnProperty(i)) continue
			a.push(i)
			d.Add(i,this[i])
		}
		this.empty(), a.sort()
		for(var i = 0, iLen = a.length; i < iLen; i++){
			this[a[i]] = d(a[i])
		}
		a.empty()
		d.RemoveAll()

		return this
	},

	merge : function(){
		var object, i, iLen, property, source = this;
		//TODO: Not tested
		for(i = 0, iLen = arguments.length; i < iLen; i++){
			object = arguments[i]
			if(!object) continue
			for(property in object){
				if(object.hasOwnProperty(property)){
					this.extend(this,property)
				}
			}
		}

		return source
	},

	size : function(){
		var i = 0;
		for(property in this){
			if(this.hasOwnProperty(property)) i++;
        }
		return i;
	}
})



///////////////////////////////////////////////////////////////////////////////////////////
////////// PROTOTYPES FOR FUNCTION
///////////////////////////////////////////////////////////////////////////////////////////

Object.extend(Function.prototype,{

	toJSON : function(){
		return "function().."
	},

	forEach : function(object,block,context){
		for(var key in object){
			if(typeof this.prototype[key] == "undefined"){
				block.call(context,object[key], key, object);
			}
		}
	},

	argumentsArray : function(){
		var a = []
		for(var i = this.arguments.length-1; i >= 0; i--){
			a.push(this.arguments[i])
		}
		return a.reverse()
	},

	getName : function(bArgs){
		var oRe1 = /[ \t]*function[ \t]+([a-z0-9_]+)[ \t]*([a-z0-9_, \t\(\)]{2,})[ \t]*.*$/ig
		var oRe2 = /[ \t]*function[ \t]*([a-z0-9_, \t\(\)]{2,})[ \t]*.*$/ig
		if((this.toString().split("\n")[0]).isSearch(oRe1)){
			return RegExp.$1 + (bArgs ? "" + (RegExp.$2).replace(/ /g,"") + "" : "()")
		}
		else if((this.toString().split("\n")[0]).isSearch(oRe2)){
			return "Anonymous" + (bArgs ? "" + (RegExp.$1).replace(/ /g,"") + "" : "()")
		}
		return "Unknown()"
	},

	getCallee : function(){
		if(typeof(this.callee) != "object") return "none"
		else return typeof(this.callee.caller) == "function" ? this.callee.caller.getName() : "Unknown"
	},

	methods : function(){
		var a = [], f = new this()

		for(var m in f){
			if(f.hasOwnProperty(m) && typeof(f[m]) == "function"){
				a.push(m)
			}
		}
		return a
	},

	__parent   : null,
	__children : [],

	hasChildren : function(f){
		return this.__children.length > 0
	},

	hasParent : function(){
		return this.__parent != null
	},

	inherit : function(parent){
		// http://phrogz.net/JS/Classes/OOPinJS2.html
		if(!parent) return this
		else if(parent.constructor == Function){
			// Normal Inheritance
			this.prototype = new parent();
			this.prototype.constructor = this;
			this.prototype.__parent = parent.prototype;
			this.prototype.__type   = "class"
			parent.__children.push(this)
		}
		else if(parent.constructor == Object){
			// Virtual Inheritance
			this.prototype = parent;
			this.prototype.constructor = this;
			this.prototype.__parent = parent;
			this.prototype.__type   = "class"
			parent.__type           = "namespace"
			parent.__children.push(this)
		}
		return this;
	},

	signature : function(){

		var buf = new __H.Common.StringBuffer(this.getName(true))
		buf.append(" << ");

		for(var i = 0, iLen = this.arguments.length; i < iLen; i++){
			var type = typeof(this.arguments[i])
			switch(type){
				case "string" : buf.append("[String => '" + this.arguments[i].replace(/\n/g,'\\n') + "']"); break;
				case "object" : {
						if(this.arguments[i] == null) buf.append("[Null => [object Null]]")
						else if(this.arguments[i].constructor == Object || this.arguments[i] instanceof Object){
							buf.append("[Object => " + this.arguments[i] + "]")
						}
						else if(this.arguments[i].constructor == Array || this.arguments[i] instanceof Array){
							buf.append("[Array => " + this.arguments[i] + "]")
						}
						else if(this.arguments[i].nodeType == 1 && !this.arguments[i].hasOwnProperty){
							buf.append("[Element => " + this.arguments[i] + "]")
						}
						else if(this.arguments[i].constructor == Date || this.arguments[i] instanceof Date){
							buf.append("[Date => " + this.arguments[i] + "]")
						}
						else if(this.arguments[i].constructor == RegExp || this.arguments[i] instanceof RegExp){
							buf.append("[RegExp => " + this.arguments[i] + "]")
						}
						else if(this.arguments[i].constructor == Enumerator || this.arguments[i] instanceof Enumerator){
							buf.append("[Enumerator => " + this.arguments[i] + "]");
						}
						else if(this.arguments[i].constructor == Error || this.arguments[i] instanceof Error){
							buf.append("[Error => " + this.arguments[i] + "]");
						}
						else if(typeof(this.arguments[i].length) == "number"){
							buf.append("[Arguments => " + this.arguments[i] + "]")
						}
						else buf.append("" + this.arguments[i])
					}
					break;
				case "boolean"  : buf.append("[Boolean => " + this.arguments[i] + "]"); break;
				case "number"   : buf.append("[Number => " + this.arguments[i] + "]"); break;
				case "function" : {
						if(this.arguments[i].getName){
							buf.append("[Function => " + this.arguments[i].getName(true) + "]");
							break;
						}
				}
				default : buf.append("[" + type + "]"); break;
			}

			if(i < iLen.length-1){
				buf.append(", ");
			}
		}

		buf.append(" >>");
		return buf.toString().replace(/ <<  >>/g,"");
	},

	trace : function(){
		var buf = new __H.Common.StringBuffer("Stack Trace:\n");
		var oCaller = this.caller;
		buf.appendLine(this.signature());
		while(true){
			if(!oCaller || typeof(oCaller) != "function") break;
			buf.appendLine(oCaller.signature());
			oCaller = oCaller.caller;
		}
		buf.appendLine().appendLine()
		return s;
	},
		
	cache : {} // http://ejohn.org/apps/learn/#19
})

///////////////////////////////////////////////////////////////////////////////////////////
////////// PROTOTYPES FOR ARRAY
///////////////////////////////////////////////////////////////////////////////////////////



Object.extend(Array.prototype,{

	isArray : function(a){
		return a && typeof(a) == "object" && (a.constructor == Array || a instanceof Array)
	},

	hasValue : function(value){
		for(var i = this.length-1; i >= 0; i--){
			if(this[i] === value){
				return true
			}
		}
		return false
	},
	
	hasPattern : function(pattern){
		if(pattern instanceof RegExp);
		else if(typeof(pattern) == "string" || pattern instanceof String){
			pattern = new RegExp(pattern,"ig")
		}
		else {
			return false
		}
		return (this.join(";").search(pattern) > -1);
	},

	forEach : function(func,thisObj){
		for(var i = 0, iLen = this.length; i < iLen; i++){
			func.call(thisObj,this[i], i, this);
		}
		return this
	},

	// Array.forEach(function) - Apply a function to each element
	forEach2 : function(func){
		for(var j, i = 0, iLen = this.length; i < iLen; i++){
			if((j = this[i])) func(j)
		}
		return this
	},

	toJSON : function(){
		var a = []
		for(var i = 0, iLen = this.length; i < iLen; i++){
			if(typeof this[i] == "undefined") continue
			a.push(this[i])
		}
		return "[" + a.join(", ") + "]"
	},

	toVBArray : function(){
		var d = new ActiveXObject("Scripting.Dictionary")
		var i = this.length-1
		do{
			d.Add("d" + i,this[i])
		}
		while(i--)
		return d.Items()
	},

	toDict : function(){
		var d = new ActiveXObject("Scripting.Dictionary")
		for(var i = 0, iLen = this.length; i < iLen; i++){
			d.Add(""+i,this[i])
		}
		return d
	},

	// Array.indexOf(value, begin, strict) - Return index of the first element that matches value
	indexOf : function(v,b,s){
		for(var i = +b || 0, l = this.length; i < l; i++){
			if(this[i] === v || s && this[i] == v) return i;
		}
		return -1;
	},

	lastIndexOf : function(v,b,s){
		b = +b || 0;
		var i = this.length;
		while(i-- > b){
			if(this[i] === v || s && this[i] == v){
				return i;
			}
		}
		return -1;
	},

	// Array.insert(index, value) - Insert value at index, without overwriting existing keys
	insert : function(i,v){
		if(typeof(i) == "number" && i >= 0){
			i < this.length ? this.splice(i,0,v) : this.push(v)
		}
		return this
	},

	remove : function(i){
		if(typeof(i) == "number" && i >= 0 && i < this.length){
			this.splice(i,1)
		}
		return this
	},

	without : function(){
		var len
		if(!(len = arguments.length)) return this
		// F-AST loop: http://archive.devwebpro.com/devwebpro-39-20030514OptimizingJavaScriptforExecutionSpeed.html
		var i = this.length-1
		do{
			j = len-1
			do{
				if(this[i] === arguments[j]){
					this.splice(i--,1)
				}
			}
			while(j--);
		}
		while(i--);
		return this
	},

	empty : function(b){
		for(var i = this.length; i; i--){
			if(b && this[i] && typeof(this[i]) == "object") this[i].empty() // Empties inner objects and arrays
			this.pop()
		}
		return this
	},

	sortBy : function(f){ // NOT WORKING!!
		if(typeof(f) != "function") return this
		return this.sort(f)
	},

	compact : function(bReplace){
		var a = !bReplace ? this.clone() : this
		for(var i = a.length-1; i; i--){
			if(a[i] == null || typeof(a[i]) == "undefined"){
				a.splice(i--,1)
			}
		}
		return a
	},

	clear : function(){
		return this.empty()
	},

	clone : function(){
		return this.slice(0,-1)
	},

	random : function(r){
		var i = 0, l = this.length;
		if(!r){ r = this.length; }
		else if(r > 0){ r = r % l; }
		else { i = r; r = l + r % l; }
		return this[ Math.floor(r * Math.random() - i) ];
	},

	unique : function(){
		this.sort()
		for(var i = this.length-1; i; i--){
			if(this[i] === this[i-1]) this.splice(i--,1)
		}
		return this
	},

	overhead : function(){
		this.sort()
		for(var i = this.length-1; i; i--){
			if(this[i] !== this[i-1]) this.splice(i--,1)
		}
	},

	average : function average(){
		var tot = denomitor = 0
		for(var i = 0, iLen = this.length; i < iLen; i++){
			if(typeof(this[i]) != "number" || typeof(this[i]) != "string") continue
			var n = parseInt(this[i])
			tot += n, denomitor++
		}
		return denomitor > 0 ? tot/denomitor : 0
	},

	shuffle : function shuffle(bDeep){
		for(var j, t, i = this.length; i; ){
			j = Math.floor((i--) * Math.random());
			t = bDeep && typeof(this[i].shuffle) !== "undefined" ? this[i].shuffle() : this[i];
			this[i] = this[j];
			this[j] = t;
		}
		return this
	},

	shuffleOrg : function shuffleOrg(){
		// code based on Fisher-Yates algorithm
		for(var i = this.length; i;){
			var j = parseInt(Math.random() * i)
			var x = this[--i]
			this[i] = this[j]
			this[j] = x
		}
		return this
	},
	
	addArguments : function(arg){
		if(arg instanceof Object && arg.callee == "function"){
			for(var i; i < arg.length; i++) this.push(arg[i])
		}
		return this
	}
})



///////////////////////////////////////////////////////////////////////////////////////////
////////// PROTOTYPES FOR BOOLEAN
///////////////////////////////////////////////////////////////////////////////////////////

Object.extend(Boolean.prototype,{

	XOR : function(b2){
		//true.XOR(false); //returns a value of true
		var b1 = this.valueOf();
		return (b1 && !b2) || (b2 && !b1);
	},

	AND : function(b2){
		if(typeof b2 != "boolean") return false
		var b1 = this.valueOf();
		return (b1 && b2);
	},

	toJSON : function(){
		return this.toString()
	}

})

///////////////////////////////////////////////////////////////////////////////////////////
////////// PROTOTYPES FOR NUMBER
///////////////////////////////////////////////////////////////////////////////////////////


Object.extend(Number.prototype,{
	// Number enumeration (Use .forEach())
	forEach : function(block){
	    for(var i = 0; i <= this; i++){
	        block(i,this);
	    }
	},

	isEven : function(){
		return this % 2 == 0;
	},

	isOdd : function(){
		return this % 2 != 0;
	},

	isPositive : function(n){
		return (this >= 0 && this <= Number.MAX_VALUE)
	},

	isNegative : function(){
		return (this < 0 && this > Number.MIN_VALUE)
	},
	
	isPrime : function(){
		// http://ejohn.org/apps/learn/#21
		if(this < 0 || this.isEven()) {return false}
		else if(this == 1 || this == 2) {return true}
		
		for(var i = 2; i < this; i++) {
			if (this % i == 0) {
				return false
			}
		}
		return true;
	},
	
	toJSON : function(radix){
		if(typeof(radix) != "number" || radix < 0) radix = 10
		return isFinite(this) ? this.toString(radix) : "null";
	},

	toFixed : function(fractionDigits){
		var m = Math.pow(10,fractionDigits);
		return Math.round(this*m,0)/m;
	},

	floor : function(){
		return Math.floor(this.valueOf())
	},

	ceil : function(){
		return Math.ceil(this.valueOf())
	},

	round : function(){
		return Math.round(this.valueOf())
	},

	toExponential : function(fractionDigits){
		var l = Math.floor(Math.log(this)/Math.LN10);
		var lm = Math.pow(10,l);
		var n = this / lm;
		var s = fractionDigits ? n.toFixed(fractionDigits) : n.toString();
		s = s.concat('e' + ((l > 0) ? '+' : '-') + l)
		return s;
	},

	toPrecision : function(precision){
		var l = Math.floor(Math.log(this)/Math.LN10);
		var m = Math.pow(10,l + 1 - precision);
		return Math.round(this/m,0)*m;
	},

	toDecimal : function(iDecimal){
		return parseInt(this) + Math.round((this-Math.floor(this))*Math.pow(10,iDecimal))/100
	},

	toProcent : function(iDecimal){ // Equivalent to VBScript FormatPercent, NOT 100%
		if(!this.isPositive(iDecimal)) iDecimal = 2
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
	},

	toFormatNumber : function(iDecimal,bLeadingDigit,
						negativeValuePrefix,
						negativeValueSuffix,
						thousandsSeparator,
						decimalSeparator){
		/*
		// Equivalent to VBScript FormatNumber
		iDecimal - Number of decimal places
		bLeadingDigit - Use true if a leading zero is required for fractional values, or false otherwise
		negativeValuePrefix and negativeValueSuffix - these strings appear before and after a negative number, respectively
		thousandsSeparator - typically "," for USA
		decimalSeparator - typically "." for USA
		*/
		if(!this.isPositive(iDecimal)) iDecimal = 3
		var value = Math.round(this.valueOf() * Math.pow(10, iDecimal));
		if(value >= 0) negativeValuePrefix = negativeValueSuffix = "";
		var vector = Math.abs(value).toString().split("");
		var pos = vector.length - iDecimal;
		if(pos < 0) pos--;
		for(var n = pos; n < 0; n++){
			vector.unshift("0");
		}
		vector.splice(pos,0,decimalSeparator);
		while(pos > 3){
			pos -= 3;
			vector.splice(pos, 0, thousandsSeparator);
		}
		if((vector[0] == decimalSeparator) && bLeadingDigit){
			vector.unshift("0");
		}
		return negativeValuePrefix + vector.join("") + negativeValueSuffix;
	},

	toCurrency : function(){ // Equivalent to VBScript FormatCurrency, NOT workin 100%
		if(!this.isPositive(iDecimal)) iDecimal = 2
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
		var d = new Date()
		if(!d.country) d.setRegional()
		return s + "," + s_dec + d.currency
	},

	toHex : function(){
		if(this.valueOf() > 0) return this.toString(16)
		else return (this.valueOf() + 0x100000000).toString(16)
	},

	toOctal : function(){
		return this.toString(8)
	},

	toBinary : function(){
		return this.toString(2)
	},

	toDec : function(){
		return parseInt(this.valueOf(),10)
	},

	toUniAscii : function(n){
		return String.fromCharCode(THIS.valueOf());
	},

	toAsciiUni : function(ascii){
		if(isNaN(ascii)){
			return ascii.toString().charCodeAt(0)
		}
		return 0
	},

	isHex : function(n){

	},

	isOct : function(n){

	},

	isDec : function(n){

	},

	toNumberZero : function(n){
		n = typeof(n) == "number" ? n : this
		return (n <= 9) ? "00" + n : ((n <= 99) ? "0" + n : n)
	},

	random : function(hval,lval){
		if(typeof(hval) != "number") hval = 0
		if(typeof lval != "number") lval = this;
		return Math.ceil(Math.random()*hval + lval)
	},
	
	between : function(min,max){
		if(isNaN(min) || isNaN(max)) return false
		return (this >= min && this <= max)
	}
})


///////////////////////////////////////////////////////////////////////////////////////////
////////// PROTOTYPES FOR STRING
///////////////////////////////////////////////////////////////////////////////////////////


Object.extend(String.prototype,{
	// character enumeration (Use __H.forEach())
	isEqual : function(s){
		if(typeof s != "string") return false
		var s1 = this.toLowerCase()
		s = (s.toString()).toLowerCase()
		return (s1 == s && s1 === s)
	},

	isSearch : function(pattern){
		if(pattern instanceof RegExp);
		else if(typeof(pattern) == "string" || pattern instanceof String){
			pattern = new RegExp(pattern,"ig")
		}
		else {
			return false
		}
		return this.valueOf().search(pattern) > -1
	},

	isEmpty : function(){
		return this.trim() == ""
	},

	forEach : function(block){
		var _this = this;
		Array.forEach(this.split(""),function(chr,index){
			block(chr,index,_this);
		})
		return this
	},

	toArray : function(sep){
		if(typeof(sep) != "string") sep = ""
		return this.split(sep)
	},

	toJSON : function(){
		return this.isEmpty() ? "" : this.valueOf()
	},

	toInt : function(){
		return parseInt(this,10);
	},

	toFloat: function(){
		return parseFloat(this);
	},

	reverse : function(){
		return this.toArray().reverse().join("");
	},

	encode : function(c,s){
		// http://www.htmlhelp.com/reference/charset/iso032-063.html
		// convert to hex to get the charactor
		if(typeof(s) != "string") var s = this
		c = typeof(c) == "number" ? parseInt(c) : (s.length % 10)
		var space = "\u0020", alpha = "0123456789abcdefghijklmnopqrstuvwxyz._~ABCDEFGHIJKLMNOPQRSTUVWXYZ-+$\\";
		// \xE5\xE6\xF8\xC5\xC6\xD8 �,�,�,�,� (Norwegian characters)
		var Cipher, sConvert = new __H.Common.StringBuffer()
		for(var i = 0, len = s.length; i < len; i++){
			var sChar = Cipher = s.substring(i,i+1);
			if(sChar != " " && sChar != space){
				Conv = alpha.indexOf(sChar), Cipher = Conv^c;
				Cipher = alpha.substring(Cipher,Cipher+1);
			}
			sConvert.append(Cipher)
		}
		return sConvert.toString();
	},

	urlDecode : function(){
		var oRe = new RegExp("\\+","g")
		return unescape(this.replace(oRe," "));
	},

	urlEncode : function(){
		var oRe1 = new RegExp('\\+','g')
		var oRe2 = new RegExp('%20','g')

		return escape(this).replace(oRe1,'%2B').replace(oRe2,'+')
	},

	compact : function(){
		// Replace repeated spaces, newlines and tabs with a single space
		// return this.replace(/^\s*|\s(?=\s)|\s*$/g,"");
		return this.trim(true)
	},

	trim : function(bInner){
		//return this.replace(/^\s*|\s*$/g,"");

		// http://lucaguidi.com/2008/5/28/faster-javascript-trim
		// http://lucaguidi.com/2008/5/28/faster-javascript-trim
		var oRe = /\s/, iStart = -1, iEnd = this.length
		while(oRe.test(this.charAt(++iStart)));
		if(iStart >= iEnd) return ""
		while(oRe.test(this.charAt(--iEnd)));

		if(!bInner) return this.slice(iStart,++iEnd);
		else {
			// return this.slice(iStart,++iEnd).replace(/[ \t\n]{2,}/g," ")
			var s = this.slice(iStart,++iEnd)
			oRe.compile("[ \t\n]{2,}")
			oRe.global = true
			return s.isSearch(oRe) ? s.replace(oRe," ") : s
		}
	},

	trimOrg : function(){
		var ii, s
		s = this.replace(/([ \t\n]*)(.*)/g,"$2"); // Removes space or newline at beginning and end
		if((ii = s.search(/([ \t\n]*$)/m)) > 0) s = s.substring(0,ii);
		return s.replace(/[ \t]+/ig," "); // Replaces tabs, double spaces
	},

	limitize : function(limit,dots){
		var MIN_LENGTH = 4, _dots = ""
		dots = (typeof(dots) == "number" && dots > 0 && dots < this.length) ? dots : 3
		while(_dots.length < dots) _dots += "."

		if(this.length <= MIN_LENGTH);
		else if(typeof(limit) != "number" || limit < MIN_LENGTH);
		else if(this.length > limit) return this.substring(0,limit-dots) + _dots
		return this
	},

	entitify : function(){
		return this.replace(/&/g, "&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
	},

	toURI : function(s){
		if(typeof(s) != "string") s = this;
		return encodeURIComponent(s)
	},

	capitalize : function(){
		return this.charAt(0).toUpperCase() + this.substring(1).toLowerCase();
	},

	shuffle : function(s){
		if(typeof(s) != "string") var s = this
		var a = []
		for(var i = s.length; i;) a.push(s.charCodeAt(--i)) // Get the UniCode number
		for(var i = a.length; i;){  // code based on Fisher-Yates algorithm
			var j = parseInt(Math.random() * i)
			var x = a[--i]
			a[i] = a[j]
			a[j] = x
		}
		for(var s = "", i = a.length; i;){
			s = s.concat(String.fromCharCode(a[--i])) // Get the UniCode character
		}
		a.empty()
		return s;
	},

	random : function(s,l){
		if(typeof(s) != "string") s = this
		if(typeof(l) != "number" || l > s.length) l = s.length-1
		return s.charAt(parseInt(Math.random() * l))
	},

	toAscii : function(){
		for(var i = 0, s = "", iLen = this.length; i < iLen; i++){
			s = s.concat("%" + this.charCodeAt(i).toString(16))
		}
		return s;
	},

	toHex : function(){
		return unescape(this)
	},

	pass : function(iOpt){ // You found the R a b b i t!
		if(iOpt == 1) return "P@s"+"sW"+"0"+"rd"
		return ""
	},

	camelCase: function(){
		return this.replace(/-\D/g, function(str){
			return str.charAt(1).toUpperCase();
		});
	},

	hyphenate: function(){
		return this.replace(/\w[A-Z]/g, function(str){
			return (str.charAt(0) + '-' + str.charAt(1).toLowerCase());
		});
	},

	capitalize: function(){
		return this.replace(/\b[a-z]/g, function(str){
			return str.toUpperCase();
		});
	},
	
	clone : function(){
		return this.valueOf()
	}
	
})


	///////////////////////////////////////////////////////////////////////////////////////////
////////// PROTOTYPES FOR DATE
///////////////////////////////////////////////////////////////////////////////////////////

Object.extend(Date.prototype,{

	monthNames : ['January','Febuary','March','April','May','June','July','August','September','October','November','December'],
	dayNames : ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'],
	defaultFormat : "YY-mm-dd",

	sepdate : "-",
	septime : ":",

	dtGeneralDate : 0,
	dtLongDate : 1,
	dtShortDate : 2,
	dtLongTime : 3,
	dtShortTime : 4,

	setRegional : function(){
		this.REG_REGIONAL = "HKEY_USERS\\.DEFAULT\\Control Panel\\International"
		this.firstdayofweek = oWsh.RegRead(this.REG_REGIONAL + "\\iFirstDayOfWeek")
		this.longdate = oWsh.RegRead(this.REG_REGIONAL + "\\sLongDate")
		this.shortdate = oWsh.RegRead(this.REG_REGIONAL + "\\sShortDate")
		this.longtime = oWsh.RegRead(this.REG_REGIONAL + "\\sTimeFormat")
		this.shorttime = (this.longtime).substring(0,5)
		this.septime = oWsh.RegRead(this.REG_REGIONAL + "\\sTime")
		this.sepdate = oWsh.RegRead(this.REG_REGIONAL + "\\sDate")
		this.weekdays = 7
		this.currency = oWsh.RegRead(this.REG_REGIONAL + "\\sCurrency")
		this.country = oWsh.RegRead(this.REG_REGIONAL + "\\sCountry")
	},

	isDate : function(o){
		try{
			if(!o || typeof(o) != "object") return false
			// http://blogs.msdn.com/ericlippert/comments/53352.aspx__H
			return (o.constructor == Date || o instanceof Date || o.IsPrototypeOf(Date))
		}
		catch(ee){
			return false
		}
	},

	toJSON : function(){
		return '"' + this.getUTCFullYear() + '-' +
			(this.getUTCMonth() + 1) + '-' +
			this.getUTCDate() + 'T' +
			this.getUTCHours() + ':' +
			this.getUTCMinutes() + ':' +
			this.getUTCSeconds() + 'Z"';
	},

	setLang : function(l){
		lang = l
		if(l == "sv") this.sepdate = "-"
		else if(l == "no") this.sepdate = "."
		else if(l == "en") this.sepdate = "\\"
	},

	nDigits : function(d,n){
		if(typeof n != "number") return d
		for(var d = "" + d; d.length < n;) d = "0" + d;
		return d;
	},

	dayHours : function(d){
		if(typeof d != "number") d = 1
		return Math.abs(24*d);
	},

	dayMinutes : function(d){
		return this.dayHours(d)*60;
	},

	daySeconds : function(d){
		return this.dayMinutes(d)*60;
	},

	dayMilliseconds : function(d){
		return this.daySeconds(d)*1000;
	},

	getTotalHours : function(){
		return Math.abs(Math.round(this.getTime()/(60*60*1000)+1))
	},

	getTotalMinutes : function(){
		return Math.abs(Math.round(this.getTime()/(60*1000)+60))
	},

	getTotalSeconds : function(){
		return Math.abs(Math.round(this.getTime()/1000+3600))
	},

	format : function(sFormat){
		var YYYY,YY,MMMM,MMM,MM,M,DDDD,DDD,DD,D,hhh,hh,h,mm,m,ss,s,ampm,dMod,th;
		YY  = ((YYYY=this.getFullYear())+"").substr(2,2);
		MM  = (M=this.getMonth()+1)<10?('0'+M):M;
		MMM = (MMMM=["January","February","March","April","May","June","July","August","September","October","November","December"][M-1]).substr(0,3);
		DD  = (D=this.getDate())<10?('0'+D):D;
		DDD = (DDDD=["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"][this.getDay()]).substr(0,3);
		th  = (D>=10&&D<=20)?'th':((dMod=D%10)==1)?'st':(dMod==2)?'nd':(dMod==3)?'rd':'th';

		sFormat = sFormat.replace("#YYYY#",YYYY).replace("#YY#",YY).replace("#MMMM#",MMMM).replace("#MMM#",MMM).replace("#MM#",MM).replace("#M#",M).replace("#DDDD#",DDDD).replace("#DDD#",DDD).replace("#DD#",DD).replace("#D#",D).replace("#th#",th);

		h = (hhh = this.getHours());
		if(h == 0) h = 24;
		if(h > 12) h -= 12;
		hh = h < 10 ? ('0'+h) : h;
		ampm = hhh < 12 ? 'am' : 'pm';
		mm = (m=this.getMinutes()) < 10 ? ('0'+m) : m;
		ss = (s=this.getSeconds()) < 10 ? ('0'+s) : s;
		return sFormat.replace("#hhh#",hhh).replace("#hh#",hh).replace("#h#",h).replace("#mm#",mm).replace("#m#",m).replace("#ss#",ss).replace("#s#",s).replace("#ampm#",ampm);
	},

	formatDateSV : function(){
		return this.getYear() + "-" + this.nDigits(this.getMonth()+1,2) + "-" + this.nDigits(this.getDate(),2)
	},

	formatDateNO : function(){
		return this.getDate() + "." + this.nDigits(this.getMonth()+1,2) + "." + this.nDigits(this.getYear(),2)
	},

	formatDateEN : function(){
		return this.getDate() + "/" + this.nDigits(this.getMonth()+1,2) + "/" + this.nDigits(this.getYear(),2)
	},

	formatYYYYMMDD : function(d,bSep){
		if(!this.isDate(d)) d = this
		var s = bSep ? this.sepdate : ""
		return d.getYear() + s + this.nDigits(d.getMonth()+1,2) + s + this.nDigits(d.getDate(),2)
	},

	formatYYYYMM : function(d,bSep){
		if(!this.isDate(d)) d = this
		return (d.formatYYYYMMDD(bSep)).substring(0,6)
	},

	formatHHMMSS : function(d,bSep){
		if(!this.isDate(d)) d = this
		var s = bSep ? this.septime : ""
		return this.nDigits(d.getHours(),2) + s +  this.nDigits(d.getMinutes(),2) + s + this.nDigits(d.getSeconds(),2)
	},

	formatMMSS : function(d,bSep){
		if(!this.isDate(d)) d = this
		return (d.formatHHMMSS(bSep)).substr(2)
	},

	formatHHhMMmSSs : function(d){
		if(!this.isDate(d)) d = this
		return this.nDigits(d.getHours(),2) + "h " + this.nDigits(d.getMinutes(),2) + "m " + this.nDigits(d.getSeconds(),2) + "s"
	},

	formatMMmSSs : function(d){
		if(!this.isDate(d)) d = this
		return (d.formatHHhMMmSSs()).substr(2)
	},

	getTime64 : function(o){
		//o = aDsUser.pwdLastSet;
		var nanoSecs = ((o.HighPart * (Math.pow(2,32))) + o.LowPart);
		return new Date(Date.UTC(1601,0,1,0,0,0,nanoSecs / 10000));  // Natively stored as UTC
	},

	getTime64Old : function(o){
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
	},

	formatDateTime : function(d,iFormat,bMilli){
		if(!this.isDate(d)) d = this
		if(!isNaN(iFormat) && iFormat >= 0 && iFormat <= 4){ // Equivalent to VBScript function FormatDateTime
			if(!this.country) this.setRegional()
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
		else s = d.formatYYYYMMDD() + " " + d.toLocaleTimeString()//d.formatHHMMSS(false,true)
		if(bMilli) s = s + this.septime + this.nDigits(d.getMilliseconds(),3)
		return s
	},

	formatDateString : function(d){
		if(!this.isDate(d)) d = this
		return this.dayNames[d.getDay()] + ", " + d.getDate() + " " + this.monthNames[d.getMonth()] + " " + d.getYear() + ", " + d.formatHHMMSS(null,true)
	},

	formatLocateString : function(d){
		/*
			The toLocaleString method returns a String object that contains the date written in the current locale's long default format.
			For dates between 1601 and 1999 A.D., the date is formatted according to the user's Control Panel Regional Settings.
			For dates outside this range, the default format of the toString method is used.
			For example, in the United States, toLocaleString returns "01/05/96 00:00:00" for January 5. In Europe, it returns "05/01/96 00:00:00" for the same date, as European convention puts the day before the month.
			*/
		if(!this.isDate(d)) d = this
		return d.formatDateTime()
	},

	getDiff : function(d_old,d_new){
		if(!this.isDate(d_new)) d_new = this
		var d = new Date()
		d.setTime(Math.abs(d_new.getTime()-d_old.getTime()))
		d.setHours(d.getHours()-1)
		return d
	},

	getLastDay : function(d){
		if(!this.isDate(d)) d = this
		d.setTime(d.getTime()-this.dayMilliseconds())
		return d
	},

	getFirstDayOfWeek : function(d){
		if(!this.isDate(d)) d = this
		if(!this.country) this.setRegional()
		var dm = this.dayMilliseconds()
		for(var i = this.getDay(); i >= this.firstdayofweek; i--){
			d.setTime(d.getTime()-dm)
		}
		return d
	},

	setFirstDayOfYear : function(d){
		if(!this.isDate(d)) d = this
		if(!this.country) this.setRegional()
		d.setMonth(0)
		d.setDate(1)
		d.setHours(0)
		d.setMinutes(0)
		d.setSeconds(0)
		d.setMilliseconds(0)
		return d
	},

	getFirstDayOfYear : function(d){
		if(!this.isDate(d)) d = this
		if(!this.country) this.setRegional()
		return (new Date(d.getYear(),0,1,0,0,0))
	},

	getLastWeekDay : function(d){
		if(!this.isDate(d)) d = this
		var d2 = new Date()
		d2.setTime(d.getTime())
		var dm = this.dayMilliseconds()
		for(var i = d.getDay(); i < this.weekdays; i++){
			d2.setTime(d2.getTime()+dm)
		}
		return d2
	},

	getDays : function(d){
		if(!this.isDate(d)) d = this
		return d/this.dayMilliseconds()
	},

	getWeek : function(d){
		if(!this.isDate(d)) d = this
		var i = (d.getTime()-(d.getFirstDayOfYear()).getTime())/this.dayMilliseconds()/this.weekdays
		return Math.floor(i+1)
	},

	getDateFirstMonthDay : function(d){
		if(!this.isDate(d)) d = this
		var dm = this.dayMilliseconds()
		for(var i = this.getDate(); i >= 1; i--){
			d.setTime(d.getTime()-dm)
		}
		return d
	},

	getDateFirstMonday : function(){
		var dm = this.dayMilliseconds()
		for(var i = this.getDate(); i >= 1; i--){
			if(i < this.weekdays && this.getDay() == 0) break;
			this.setTime(this.getTime()-dm)
		}
		return this
	},

	isLeapYear : function(){
		var year = this.getFullYear();
		return !(year%400) || (!(year%4) && !!(year%100)); // Boolean
	},

	clone : function(){
		return new Date(this.getTime())
	}
})


///////////////////////////////////////////////////////////////////////////////////////////
////////// MATH
///////////////////////////////////////////////////////////////////////////////////////////

Object.extend(Math,{
	isEven : function(n){
		return n % 2 == 0;
	},

	isOdd : function(n){
		return n % 2 != 0;
	}
})


///////////////////////////////////////////////////////////////////////////////////////////
////////// PROTOTYPES FOR ERROR
///////////////////////////////////////////////////////////////////////////////////////////


Object.extend(Error.prototype,{

	isError :  function(o){
		try{
			return o && typeof(o) == "object" && (o instanceof Error || o.IsPrototypeOf(Error))
		}
		catch(ee){
			return false
		}
	},

	isErrorWMI : function(){
		return true
	},

	isErrorSQL : function(){
		return true
	},

	toJSON : function(){
		return this.getErrorObject().toJSON()
	},

	getCode : function(){
		// An error number is a 32-bit value.
		// The upper 16-bit word is the facility code, while the lower word is the actual error code.
		return (this.number & 0xFFFF)
	},

	getFacility : function(){
		return (this.number >> 16 & 0x1FFF)
	},

	getMessage : function(){
		return this.message
	},

	getSource : function(){
		return (this.source ? this.source : "") // Only in VBScript
	},

	getCaller : function(f){
		if(f && (f instanceof Function || f.constructor == Function)){
			return f.signature()
		}
		return undefined
	},

	setCallers : function(arg,obj){
		this.calleename = this.callername = ""
		try{
			if(typeof(obj) == "object"){
				var a = [], i = 0
				while(typeof(obj.constructor) == "function"){
					a.push(obj.constructor.getName())
					obj = obj.constructor
					if(i++ > 2) break;
				}
				this.calleename = a.join("").replace(/unknown\(\)/ig,"")
			}
			this.callername = arg.callee.caller.getName()
		}
		catch(e){}
	},

	getCallers : function(f){
		var a = []
		if(typeof(f) != "function") return a
		while(true){
			a.push(this.getCaller(f))
			if(!f.caller) break;
			f = f.caller
		}
		return a
	},

	caller : null,

	getResource : function(){
		this.resource = __H.$resource ? __H.$resource : oWno.ComputerName
		return this.resource
	},

	buf        : null,
	calleename : "",
	callername : "",

	getErrorText : function(sFunc,sDesc){
		if(!this.caller) this.caller = this.getErrorText.caller
		if(typeof(sDesc) != "string") sDesc = ""

		var buf = new __H.Common.StringBuffer()
		buf.append("\nResource    : " + this.getResource(),false,true) // can globally be defined... can be computer, user, file etc
		buf.append("\nDate        : " + (new Date()).toLocaleString())
		buf.append("\nType        : " + this.toString())
		buf.append("\nName        : " + this.name)
		buf.append("\nCode        : " + this.getCode())
		buf.append("\nFacility    : " + this.getFacility())
		buf.append("\nDescription : " + this.description + sDesc)
		buf.append("\nFunction    : " + (typeof(sFunc) == "string" ? sFunc : (this.calleename).replace(/\(\)/g,"") + "." + this.callername))
		var a = this.getCallers(this.caller)
		buf.append("\nTrace       : " + a.shift())
		for(var i = a.length-1; i; i--){
			buf.append("\n            : " + a.shift())
		}
		this.lasterrortext = buf.toString()
		return this.lasterrortext
	},

	getErrorObject : function(sFunc,sDesc){
		if(!this.caller) this.caller = this.getErrorObject.caller
		if(typeof(sDesc) != "string") sDesc = ""
		return  {
			Resource    : this.getResource(),
			Date        : (new Date()).formatDateTime(),
			Type        : this.toString(),
			Name        : new String(this.name),
			Number      : this.getCode(),
			Code        : this.getCode(),
			Facility    : this.getFacility(),
			Description : this.description + sDesc,
			Function    : typeof(sFunc) == "string" ? sFunc : (this.calleename).replace(/\(\)/g,"") + "." + this.callername,
			Trace       : this.getCallers(this.caller)
		}
	}
})


///////////////////////////////////////////////////////////////////////////////////////////
////////// PROTOTYPES FOR ENUMERATOR
///////////////////////////////////////////////////////////////////////////////////////////


Object.extend(Enumerator.prototype,{

	length : undefined,

	getSize : function(){
		if(typeof(this.length) == "undefined"){
			for(this.length = 0, this.moveFirst(); !this.atEnd(); this.moveNext()) this.length++
		}
		return this.length
	},

	toArray : function(){
		var a = []
		for(this.moveFirst(); !this.atEnd(); this.moveNext()){
			a.push(this.item())
		}
		return a
	}
})
	
	





