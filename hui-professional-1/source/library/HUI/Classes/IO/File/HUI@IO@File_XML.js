// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**


// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/xmlsdk/html/31f7af53-6c3f-4a3e-a503-8aed9c4b270e.asp
// http://webdesign.about.com/od/xml/


__H.include("HUI@IO@File.js")

//(function(){
__H.register(__H.IO.File,"XML","XML Manipulation",function XML(){

	var o_this = this
	var b_initialized = false

	var d_xml_docs
	var oReExt

	/////////////////////////////////////
	//// DEFAULT

	var a_xml_progid = [
		"Msxml2.DOMDocument.6.0",
		"Msxml2.DOMDocument.5.0",
		"Msxml2.DOMDocument.4.0",
		"Msxml2.DOMDocument.3.0",
		"Msxml2.DOMDocument.2.6",
		"Msxml2.DOMDocument",
		"Microsoft.XMLDOM"
	];
	
	var o_cur = {
		o_pinode      : null,
		s_progid      : undefined,
		b_progid_sync : false,
		o_xml         : null,
		s_xml         : undefined,
		o_doc         : null,
		o_xml_tmp     : null,
		o_xsl_tmp     : null,		
		o_node        : null,
		s_node        : undefined		
	}
	
	var o_locales = {
		
	
	}
	
	var o_options = {

	}

	function initialize(bForce){
		if(b_initialized && !bForce) return;

		d_xml_docs    = new ActiveXObject("Scripting.Dictionary")
		oReExt        = /xml|xsl|xsd/ig
		o_cur.b_progid_sync = false

		for(var i = 0, iLen = a_xml_progid.length; i < iLen; i++){
			try{
				o_cur.s_progid      = new ActiveXObject(a_xml_progid[i])
				o_cur.b_progid_sync = !!(o_cur.s_progid.search(/.+\.([4-9])\.[0-9]$/g) > -1)
				break;
			}
			catch(ee){
				o_cur.s_progid = undefined
			}
		}
		
		b_initialized = true
	}
	
	this.setOptions = function setOptions(oOptions){
		if(__H.isObject(oOptions)) return Object.extend(o_options,oOptions,true)
		return false
	}

	this.getOptions = function(){
		return o_options
	}

	this.toString = function(){
		return this//.getSections().toString()
	}

	/////////////////////////////////////
	////

	this.isFile = function isFile(sFile){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(arguments[i]
				&& typeof(arguments[i]) == "string"
				&& oFso.FileExists(arguments[i])
				&& (sExt = oFso.GetExtensionName(arguments[i])).search(oReExt) > -1){
				continue
			}
			__HLog.debug("Either invalid or access unavailable for XML file: " + arguments[i])
			return false
		}
		return !!i;
	}

	this.isXML = function isXML(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(arguments[i]
				&& typeof(arguments[i]) == "object"
				&& arguments[i].hasChildNodes
				&& typeof(arguments[i].createDocumentFragment) == "function"){
				continue
			}
			return false
		}
		return !!i;
	}

	this.loadXML = function loadXML(sXml){
		initialize()
		try{			
			if(!this.isFile(sXml)) return null
			
			/*
			 There are two models of the Document: the free-threaded model and the rental threading model.
			 They both behave exactly the same but differ in that the rental-treaded versions offer
			 better performance because the parser doesn't need to manage concurrent access among threads.
			*/
			o_cur.o_xml_tmp       = new ActiveXObject(o_cur.s_progid);
			o_cur.o_xml_tmp.async = false
			// loads and parses an XML file synchronously
			if(o_cur.b_progid_sync){				
				o_cur.o_xml_tmp.validateOnParse  = false; 
				o_cur.o_xml_tmp.resolveExternals = false
			}			
			o_cur.o_xml_tmp.preserveWhiteSpace = false
			o_cur.o_xml_tmp.validateOnParse    = false
			o_cur.o_xml_tmp.onreadystatechange = verify;
			// Pattern-based selection can be based on two standards: XSLT, and XPath.
			// Usually, the XPath standard will satisfy your requirements.
			o_cur.o_xml_tmp.setProperty("SelectionLanguage", "XPath")
			o_cur.o_xml_tmp_fragment = o_cur.o_xml_tmp.createDocumentFragment()
						
			o_cur.o_xml_tmp.load(sXml)
					
			if(o_cur.o_xml_tmp.parseError.errorCode != 0){
				var s = "Unable to load XML file: " + sXml
				__HExp.XMLLoadingError(s + ". Reason: " + (o_cur.o_xml_tmp.parseError.reason).trim())
				__HLog.errorXML(o_cur.o_xml_tmp,null)
				return null
			}
			
			s_cur_file = sXml			
			return o_cur.o_xml_tmp
		}
		catch(ee){
			__HLog.error(ee,this)
			return null
		}
	}
	
	this.setXML = function setXML(sXml){
		try{
			this.reset()
			
			if(!d_xml_docs.Exists(sXml = sXml.toLowerCase())){
				if(this.loadXML(sXml)){
					d_xml_docs.Add(sXml,o_cur.o_xml_tmp)
				}
				else return null
			}
			
			o_cur.o_xml = d_xml_docs(sXml)
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getXML = function getXML(){
		try{
			if(!o_cur.o_xml){
				__HExp.XMLNotInitialized("XML is NOt Set")
			}
			return o_cur.o_xml
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.addXSL = function addXSL(sXml){
		try{
			if(!o_cur.o_xml){
				var oNode = o_cur.o_xml.createProcessingInstruction("xml-stylesheet","type=\"text/xsl\" href=\"" + sXsl + "\"");
				o_cur.o_xml.appendChild(oNode);
			}
			return o_cur.o_xml
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.reset = function reset(){
		try{

		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	function verify(){
		// 0 Object is not initialized
		// 1 Loading object is loading data
		// 2 Loaded object has loaded data
		// 3 Data from object can be worked with
		// 4 Object completely initialized
		if(o_cur.o_xml.readyState != 4){
		  return false;
		}
		return true
	}

	/////////////////////////////////////
	//// XML DOCUMENT

	this.create = function create(sXml,sRoot,oRootAttributes,sComment){
		try{
			if((this.isFile(sXml))){
				o_cur.s_xml = sXml
				if(this.setPINode()
					&& this.setComment(sComment)
					&& this.setRoot(sRoot,oAttributes)
					&& this.save()){					
					return o_cur.o_xml
				}
			}
			return null
		}
		catch(ee){
			__HLog.error(ee,this)
			return null
		}
	}

	this.setPINode = function setPINode(){
		try{
			// standalone must always be last/after encoding -- http://doc.ddart.net/xmlsdk/htm/xml_concepts_00xa.htm
			o_cur.o_pinode = o_cur.o_xml.createProcessingInstruction("xml","version=\"1.0\" encoding=\"ISO-8859-1\" standalone=\"no\"");
	   		if(o_cur.o_xml.hasChildNodes) o_cur.o_xml.insertBefore(o_cur.o_pinode,o_cur.o_xml.firstChild);
	   		else o_cur.o_xml.appendChild(o_cur.o_pinode);
	   		return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.setComment = function setComment(sComment){
		try{
			var oFrag = o_cur.o_xml.createDocumentFragment();
			oFrag.appendChild(o_cur.o_xml.createComment(sComment));
			o_cur.o_xml.appendChild(oFrag);
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.setRoot = function setRoot(sRoot,oAttributes){
		try{
			o_cur_root = o_cur.o_doc.createElement(sRoot)
			this.setAttributes(o_cur_root,oAttributes)
		    o_cur.o_doc.appendChild(o_cur_root)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.addXml = function addXml(sXml){
		try{
			var oXml
			if(oXml = this.loadXML(sXml)){
				var oNode = o_cur.o_xml.documentElement ? o_cur.o_xml.documentElement.firstChild : o_cur.o_xml
				return this.addNode(oNode,oXml.documentElement);
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.save = function save(){
		try{
			o_cur.o_xml.documentElement.appendChild(o_cur.o_xml.createTextNode("\n"));
			o_cur.o_xml.save(o_cur.s_xml);
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
		
	this.toJSON = function toJSON(){
		try{

		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.toHTML = function toHTML(sXml,sXsl){
		try{
			if(!this.isFile(sXml,sXsl)) return undefined
			var oXml = this.loadXML(sXml);
			var oXsl = this.loadXML(sXsl);
			if(oXml && oXsl){				
				return oXml.transformNode(oXsl);
			}
			
		}
		catch(ee){
			__HLog.error(ee,this)			
		}
		return undefined
	}
	
	this.transform = this.toHTML;	

	this.validate = function validate(sXml,sXsl,sXsd,sNode,bNodeRec){
		try{
			try{
				var oXml, oXsd, oError
				var oNode, oNodes
				
				if(!this.isFile(sXml)) return false				
				else if((oXml = this.loadXML(sXsd))) return false
				
			    if(this.isFile(sXsd)){
					oXsd = this.loadXML(sXsd);
					var oSCache = new ActiveXObject("msxml2.XMLSchemaCache.4.0");
					oSCache.add("urn:" + sNode,oXsd);
					oXml.schemas = oSCache;
				 }
			    
				__HLog.debug("@ Validating XML " + sXml);
			    oError = oXml.validate();
			    if(oError.errorCode != 0){
			    	__HLog.errorXML(oXml)
			    	return false
			    }

			    if(sNode){
					__HLog.debug("@@ Validating nodes of '//" + sNode + "'");
					oNodes = oXml.selectNodes("//" + sNode);
					for(var i = 0, iLen = oNodes.length; i < iLen; i++){
						oNode = oNodes.item(i);
						oError = oXml.validateNode(oNode);
						if(oError.errorCode != 0){
							__HLog.errorXML(oXml)
							__HLog.debug("@@@ <" + oNode.nodeName + "> (" + i + ") is not valid because " + (oError.reason).trim())
							return false
						}
					}
				}

			    if(sNode && bNodeRec){
				    oNodes = oXml.selectNodes("//" + sNode + "/*");
				    __HLog.debug("@@ Validating all children of all " + sNode + " nodes, " + "//" + sNode + "/*");
				    for(var i = 0, iLen = oNodes.length; i < iLen; i++){
						oNode = oNodes.item(i);
						oError = oXml.validateNode(oNode);
						if(oError.errorCode != 0){
							__HLog.errorXML(oXml)
							__HLog.debug("@@@ <" + oNode.nodeName + "> (" + i + ") is not valid beacause " + (oError.reason).trim())
							return false
						}
				    }
				}
				return true
			}
			catch(e){
				__HLog.error(e,this)
				return false;
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.reset = function reset(){
		try{

		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	

	/////////////////////////////////////
	//// XML NODES & ELEMENTS
	

	this.addCDATA = function addCDATA(sNode,sText){
		try{
			if(__H.isStringEmpty(sNode,sText)) return null
			var oNode = o_cur.o_xml.createElement(sNode);
    		oCData = o_cur.o_xml.createCDATASection(sText);
    		oNode.appendChild(oCData);
    		o_cur.o_xml.documentElement.appendChild(oNode);
    		return oNode;
		}
		catch(ee){
			__HLog.error(ee,this)
			return null
		}
	}
	
	this.isNode = function isNode(){
		for(var i = arguments.length-1; i >= 0; i--){
			if(arguments[i]
				&& typeof(arguments[i]) == "object"
				&& arguments[i].xml){
				continue
			}
			__HLog.debug("Invalid XML node in Arguments [" + i + "]")
			return false
		}
		return true
	}
	
	this.setNode = function setNode(oNode){
		try{
			if(this.isNode(oNode)){
				o_cur.o_node = oNode
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.getNode = function(){
		return o_cur.o_node
	}
	
	this.delNode = function(){
		// Returns the removed node
		return o_cur.o_node ? o_cur.o_xml.removeChild(o_cur.o_node) : null
	}
	
	this.nextSibling = function(){
		return o_cur.o_node ? o_cur.o_node.nextSibling : null		
	}
	
	this.nextNode = function(){
		return o_cur.o_node ? o_cur.o_node.nextNode : null
	}
	
	this.nextNodes = function nextNodes(){
		var oNodes = oXml.getElementsByTagName(o_cur.o_node.nodeName)
		var a = []
		for(var i = oNodes.length--; i >= 0; i--){
			a.push(oNodes[i])
		}
		return a.reverse()
	}
	
	this.nodeObject = function(){
		return !o_cur.o_node ? null : {
				name  : o_cur.o_node.nodeName,
				value : o_cur.o_node.nodeValue,
				type  : o_cur.o_node.nodeType,
				xml   : o_cur.o_node.xml					
		}
	}
	
	this.cloneLast = function(){
		if(o_cur.o_node && o_cur.o_node.parentNode){
			o_cur.o_node.parentNode.appendChild(o_cur.o_node.cloneNode(true))
			// Returns the cloned node
			return o_cur.o_node.parentNode.lastChild;
		}
		return null
	}
	
	this.nodeTypeText = function(){		
		return !o_cur.o_node ? undefined : ["0-NODE_UNDEFINED",
			"1-NODE_ELEMENT",
			"2-NODE_ATTRIBUTE",
			"3-NODE_TEXT",
			"4-NODE_CDATA_SECTION",
			"5-NODE_ENTITY_REFERENCE",
			"6-NODE_ENTITY",
			"7-NODE_PROCESSING_INSTRUCTION",
			"8-NODE_COMMENT",
			"9-NODE_DOCUMENT",
			"10-NODE_DOCUMENT_TYPE",
			"11-NODE_DOCUMENT_FRAGMENT",
			"12-NODE_NOTATION"
		][o_cur.o_node.nodeType];
	}
	
	this.byXPath = function byXPath(sXPath){
		try{
			var aXPath = o_cur.o_xml.documentElement.selectNodes(sXPath)
			var a = []
			for(var i = aXPath.length-1; i >= 0; i--){
				a.push(aXPath.item(i))
			}
			return a.reverse()
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getXPath = function getXPath(){
		try{
			var oNode = o_cur.o_node.firstChild ? o_cur.o_node.firstChild : o_cur.o_node
			for(var p = ""; oNode.parentNode; ){
				oNode = oNode.parentNode
				p = "\\".concat(oNode.nodeName).concat(p)
			}
			return "\\".concat(p)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.setAttributes = function setAttributes(sName,sValue){
		try{
			if(__H.isStringEmpty(sName,sValue) && !o_cur.o_node) return false
			
			var oAttr = o_cur.o_xml.createAttribute(sName);
			oAttr.appendChild(o_cur.o_xml.createTextNode(sValue));
			o_cur.o_node.setAttributeNode(oAttr);
			
			return oAttr;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.getAttributes = function getAttributes(){
		try{			
			if(!o_cur.o_node) return [];
			var a = [];
			for(var oEnum = new Enumerator(o_cur.o_node.attributes); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				a.push({
					name  : oItem.name,
					value : oItem.nodeValue,
					type  : oItem.nodeType				
				});
			}
		}
		catch(ee){
			__HLog.error(ee,this)			
		}
		return a
	}
	
	this.getDepth = function getDepth(){
		try{
			var oNode = o_cur.o_node.firstChild ? o_cur.o_node.firstChild : o_cur.o_node
			for(var i = 0; oNode.parentNode; i++){
				oNode = oNode.parentNode
			}
			return i
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.addElement = function addElement(sElement,sText){
		try{
			var bNoLine, sSpace, sSpace2
			var oElem, oFrag
			
			if(__H.isStringEmpty(sElement,sText) && !o_cur.o_node) return null
			else if(o_cur.o_node.parentNode.documentElement) bNoLine = !o_cur.o_node.hasChildNodes ? false : true
			else bNoLine = o_cur.o_node.firstChild == null ? false : true
			
			sSpace = sSpace2 = this.getSpace()
			oFrag = o_cur.o_xml.createDocumentFragment()
			if(!bNoLine) oFrag.appendChild(o_cur.o_xml.createTextNode(sSpace));
			else sSpace2 = sSpace.replace(/\t\t/i,"")
			
			oElem = o_cur.o_xml.createElement(sElem);
			if(sText){
				oText = o_cur.o_xml.createTextNode(sText)
				oElem.appendChild(oText)
			}
			oFrag.appendChild(oElem);
			if(sText) oFrag.appendChild(o_cur.o_xml.createTextNode(sSpace2))
			o_cur.o_node.appendChild(oFrag);
			
			return oElem;
		}
		catch(ee){
			__HLog.error(ee,this)
			return null
		}
	}
	
	this.getSpace = function(){
		var s = "\n"
		for(var i = this.getDepth(); i; i--){
			s = s.concat("\t")
		}
		return s
	}
	
	this.getXMLAJAX = function getXMLAJAX(sUrl,oHandler,sUser,sPassword){
		try{
			if(!__H.isUndefined(sUser,sPassword));
			else sUser = sPassword = null
			if(typeof(oHandler) != "function") throw "Arguments 2 must defines as a the handler return function"
			var oHttp = new ActiveXObject("Microsoft.XMLHTTP")
			
			oHttp.open("GET",sUrl,true,sUser,sPassword)
			oHttp.onreadystatechange = function(){
				if(oHttp.readyState == 4 && oHttp.status == 200){
					oHandler(oHttp.responseXML,oHttp);
					// responseText
				}
			}
			oHttp.send(null)
		}
		catch(e){
			__HLog.error(e,this)
			return false;
		}
	}

})
//}());

var __HXML = new __H.IO.File.XML()

////////////////////////////////////////////////////////////////////////////////
///////// DTD FUNCTIONS


