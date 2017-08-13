// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**


// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/xmlsdk/html/31f7af53-6c3f-4a3e-a503-8aed9c4b270e.asp
// http://webdesign.about.com/od/xml/


var xml = new xml_object();

__H.include("HUI@IO@File.js")

__H.register(__H.IO.File,"XML","XML manipulation",function XML(sFile){

})

function xmlCreateDocument(rootName) {
	var xml = null;
	try {
		xml = new ActiveXObject("Msxml2.DOMDocument.4.0");
	} catch (e) {
		try {
			xml = new ActiveXObject("Msxml2.DOMDocument");		
		} catch(e1) {
			return null;
		}
	}
	
	xml.async           = false;
	xml.validateOnParse = false; 
	
	if(rootName != null) {
		var root = xml.createElement(rootName);
		xml.appendChild(root);
	}
	
	return xml;
}

function xmlAppendChild(xml, parent, name, value) {
	var item  = xml.createElement(name);
	if(value != null) {
		item.text = value;
	}
	parent.appendChild(item);
	
	return item;
}

function xmlXSLT(xml_src, xsl_src) {
	var oXSL = xmlCreateDocument(null);
	var oXML = xmlCreateDocument(null);
	oXSL.load(xsl_src);
	oXML.load(xml_src);
	var html = oXML.transformNode(oXSL);
	
	return html;
}

function xml_object(sProgID,sXml,sXsl,sXsd,sDtd,sComment,sRoot,sAttr,sAttrValue){
	try{		
		this.progid = sProgID ? sProgID : null
		this.xmlfile = sXml ? oFso.GetAbsolutePathName(sXml) : ""
		this.xslfile = sXsl ? sXsl : ""
		this.xsdfile = sXsd ? sXsd : ""
		this.dtdfile = sDtd ? sDtd : ""		
		this.comment = sComment ? sComment : ""
		this.root = sRoot ? sRoot : "xml_root"
		this.attr = sAttr ? sAttr : ""
		this.attrval = sAttrValue ? sAttrValue : ""
		this.dom = null
		this.frag = null
	}
	catch(e){
		__HLog.error(e,this)
		return false
	}
}

function xml_dom(sOpt,sXml,sXsl,oDom,bLog){
	try{		
		var oXml = oXsl = null
		
		if(sOpt == "make"){
			var aProgID = new Array("Msxml2.DOMDocument.5.0","Msxml2.DOMDocument.4.0","Msxml2.DOMDocument.3.0","Msxml2.DOMDocument.2.6","Msxml2.DOMDocument","Microsoft.XMLDOM");
			if(!__H.isUndefined(oDom,oDom.progid)) aProgID = [oDom.progid];
			if(__H.isUndefined(oDom)) oDom = new xml_object();
			for(var i = 0, iLen = aProgID.length; i < iLen; i++){
				try{					
					/*
					 There are two models of the Document: the free-threaded model and the rental threading model.
					 They both behave exactly the same but differ in that the rental-treaded versions offer better performance because the parser doesn't need to manage concurrent access among threads.
					*/
					oXml = new ActiveXObject(aProgID[i]);
					oXml.async = false;
					if(aProgID[i] == "Msxml2.DOMDocument.4.0" || aProgID[i] == "Msxml2.DOMDocument.5.0"){
						oXml.validateOnParse = false; // loads and parses an XML file synchronously
						oXml.resolveExternals = false
					}
					oXml.preserveWhiteSpace = oXml.preserveWhiteSpace = arguments[4] ? true : false;
					oXml.onreadystatechange = xml_dom_verify;
					oXml.setProperty("SelectionLanguage", "XPath"); // Pattern-based selection can be based on two standards: XSLT, and XPath. Usually, the XPath standard will satisfy your requirements.
					oDom.progid = aProgID[i];
					oDom.frag = oXml.createDocumentFragment();			
					break;
				}
				catch(ee){ // Automation server can't create object
					oXml = null
				}
			}
		}
		else if(sOpt == "load"){
			if(!oFso.FileExists(sXml)) return null
			if(bLog) __HLog.log("# Loading XML file: " + sXml);
			if(!__H.isObject(oDom)) oDom = new xml_object();
			oXml = xml_dom("make",null,null,oDom);
			oXml.load(sXml), oDom.xmlfile = sXml
			if(oXml.parseError.errorCode == 0){	
				if(!__H.isStringEmpty(sXsl) && oFso.FileExists(sXsl)){
					oXsl = xml_dom("make",null,null,oDom);
					oXsl.load(sXsl), oDom.xslfile = sXsl;
					if(oXsl.parseError.reason == "") oXml.transform = oXml.transformNode(oXsl);
					else {
						__HLog.log("## XSL file " + sXsl + " did not load because " + (oXsl.parseError.reason).trim());
						xml_log_errorxml(oXsl.parseError);
					}
				}
			}
			else {
				__HLog.log("## XML file " + sXml + " did not load because " + (oXml.parseError.reason).trim());
				xml_log_errorxml(oXml.parseError);
			}
		}
		
		return oXml;
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

function xml_dom_verify(){
	try{
		// 0 Object is not initialized
		// 1 Loading object is loading data
		// 2 Loaded object has loaded data
		// 3 Data from object can be worked with
		// 4 Object completely initialized
		if(oXml.readyState != 4){
		  return false;
		}
		return true
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}



////////////////////////////////////////////////////////////////////////////////
///////// XML DOM FUNCTIONS

function xml_dom_document(sOpt,oXml,sXml,sXsl,sDtd,sProgID,sComment,sRoot,sAttr,sAttrValue,sOpt10){
	try{ // http://doc.ddart.net/xmlsdk/htm/dom_concepts_1u9f.htm
		if(sOpt == "create"){
			if(!xml_doc_isvalid(sXml)) return false
			var oDom = new xml_object(sProgID,sXml,sXsl,sDtd,null,sComment,sRoot,sAttr,sAttrValue), bSave = sOpt10;
			oXml = xml_dom("make",null,null,oDom);
			xml_dom_document("settop",oXml);
	   		xml_xsl_document("create",oXml,oDom.xslfile);
			xml_dtd_document("create",oXml,oDom.root,null,oDom.dtdfile);
			xml_dom_document("setcomment",oXml,null,null,null,null,oDom.comment);
		    xml_dom_document("setroot",oXml,null,null,null,null,null,oDom.root,oDom.attr,oDom.attrval);		   
	    	if(bSave){
	    		oXml.documentElement.appendChild(oXml.createTextNode("\n\t"));
	    		xml_dom_document("save",oXml,oDom.xmlfile);
			}
			oDom.dom = oXml
			return oDom;
		}
		else if(sOpt == "settop"){
			if(__H.isObject(oXml)){ // Note! Don't change this xml_is_doc because root may not have been created
				var oPINode = oXml.createProcessingInstruction("xml","version=\"1.0\" encoding=\"ISO-8859-1\" standalone=\"no\""); // standalone must always be last/after encoding -- http://doc.ddart.net/xmlsdk/htm/xml_concepts_00xa.htm
	   			if(oXml.hasChildNodes) oXml.insertBefore(oPINode,oXml.firstChild);
	   			else oXml.appendChild(oPINode);
	   			return true
	   		}
		}
		else if(sOpt == "setcomment"){
			if(!__H.isUndefined(oXml,sComment)){ // Note! Don't change this xml_is_doc because root may not have been created
				var oFrag = oXml.createDocumentFragment(); 
				oFrag.appendChild(oXml.createComment(sComment));
				oXml.appendChild(oFrag);
				return true
			}
		}
		else if(sOpt == "setroot"){
			if(!__H.isUndefined(oXml,sRoot)){ // Note! Don't change this xml_is_doc because root may not have been created
				var oRoot = oXml.createElement(sRoot);
				if(!__H.isUndefined(sAttr,sAttrValue)){
					xml_dom_attribute("add",oXml,oRoot,sAttr,sAttrValue);
				}
		    	oXml.appendChild(oRoot);
		    	return true
	    	}
		}
		else if(sOpt == "setspace"){
			// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/xmlsdk/html/78fce7e6-7c72-4e12-a3aa-6d6bdc8291d1.asp
		}
		else if(sOpt == "getbase"){
			return oXml.baseName;
		}
		else if(sOpt == "adddocnode"){
			if(!xml_dom_isdocs(oXml) || !xml_dom_isnodes(oXml)) return false
			if(oXml2 = xml_dom("load",sXml)){
				var oParentChildNode = oXml.documentElement ? oXml.documentElement.firstChild : oXml
				return xml_dom_node("addnode",oParentChildNode,oXml2.documentElement);
			}
		}
		else if(sOpt == "save"){
			oXml.documentElement.appendChild(oXml.createTextNode("\n"));
			oXml.save(sXml);
			return true
		}
		return false;
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

function xml_dom_validate(sXml,sXsl,sXsd,sNode,bNodeRec){
	try{		
		if(!xml_doc_isvalid(sXml)) return false
		var bResult = true, oNode, oNodes, oError
		// Load an XML document into a DOM instance.
	    var oXml = xml_dom("load",sXml,sXsl);	
	    if(xml_doc_isvalid(sXsd)){
			var oXsd = xml_dom("load",sXsd);	 
			var oSCache = new ActiveXObject("msxml2.XMLSchemaCache.4.0");	    
			oSCache.add("urn:" + sNode,oXsd);
			oXml.schemas = oSCache;
		 }
	    // Validate the entire DOM.
	    __HLog.log("## Validating XML DOM...");
	    oError = oXml.validate();
	    if(oError.errorCode != 0){
	    	__HLog.log("### " + (oError.reason).trim()); // ### Validate failed because the root element had no associated DTD/schema.
	    	xml_log_errorxml(oError);
	    	return false
	    }
	
	    if(sNode){
			__HLog.log("### Validating all nodes of '//" + sNode + "'");
			oNodes = oXml.selectNodes("//" + sNode);			
			for(var i = 0, iLen = oNodes.length; i < iLen; i++){
				oNode = oNodes.item(i);
				oError = oXml.validateNode(oNode);
				if(oError.errorCode != 0){
					__HLog.log("#### <" + oNode.nodeName + "> (" + i + ") is not valid because " + (oError.reason).trim()), bResult = false;
					xml_log_errorxml(oError);
				}
			}
		}
			
	    if(sNode && bNodeRec){
		    oNodes = oXml.selectNodes("//" + sNode + "/*");
		    __HLog.log("### Validating all children of all " + sNode + " nodes, " + "//" + sNode + "/*");
		    for(var i = 0, iLen = oNodes.length; i < iLen; i++){
				oNode = oNodes.item(i);
				oError = oXml.validateNode(oNode);
				if(oError.errorCode != 0){
					__HLog.log("#### <" + oNode.nodeName + "> (" + i + ") is not valid beacause " + (oError.reason).trim()), bResult = false;
					xml_log_errorxml(oError);
				}
		    }
		}
		return bResult
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

function xml_dom_element(sOpt,oXml,oNode,sElem,sText,sOpt6){
	var oElem, oFrag
	try{
		if(!xml_dom_isdocs(oXml)) return false
		if(sOpt == "getspace"){
			if(!xml_dom_isnodes(oNode)) return false
			len = xml_dom_element("getdepth",oXml,oNode);
			for(var i = 0, sSpace = "\n"; i < len; i++){
				sSpace += "\t";
			}
			return sSpace
		}
		else if(sOpt == "addspace"){
			
		}
		else if(sOpt == "addelem"){ // creates a tidy XML code
			if(!xml_dom_isnodes(oNode)) return false
			if(__H.isUndefined(sElem)) return false
			// A pain in the ass... just to tidy automatically.. stupid!!
			if(oNode.parentNode.documentElement) var bNoLine = !oNode.hasChildNodes ? false : true
			else var bNoLine = oNode.firstChild == null ? false : true
			var sSpace = sSpace2 = xml_dom_element("getspace",oXml,oNode);
			oFrag = oXml.createDocumentFragment();
			if(!bNoLine) oFrag.appendChild(oXml.createTextNode(sSpace));
			else sSpace2 = sSpace.replace(/\t\t/i,"")
			oElem = oXml.createElement(sElem);
			if(sText){
				oText = oXml.createTextNode(sText)
				oElem.appendChild(oText)
			}
			oFrag.appendChild(oElem);
			if(sText) oFrag.appendChild(oXml.createTextNode(sSpace2))
			oNode.appendChild(oFrag);
			return oElem;
		}
		else if(sOpt == "getdepth"){
			if(!xml_dom_isnodes(oNode)) return false
			oNode = oNode.firstChild ? oNode.firstChild : oNode
			for(var i = 0; oNode.parentNode; i++){
				oNode = oNode.parentNode
			}
			return i
		}
		else if(sOpt == "getpath"){
			if(!xml_dom_isnodes(oNode)) return false
			oNode = oNode.firstChild ? oNode.firstChild : oNode
			for(var p = ""; oNode.parentNode; ){
				oNode = oNode.parentNode
				p = "\\" + oNode.nodeName + p
			}
			return "\\" + p
		}
		return false
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

function xml_dom_cdata(sOpt,oXml,oNode,sNode,sText){
	try{
		if(!xml_dom_isdocs(oXml)) return false
		if(sOpt == "add"){
			if(__H.isStringEmpty(sNode)) return false
			oNode = oXml.createElement(sNode);
    		oCData = oXml.createCDATASection(sText);
    		oNode.appendChild(oCData);
    		oXml.documentElement.appendChild(oNode);
    		return oNode;
		}
		return false;
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

function xml_dom_attribute(sOpt,oXml,oNode,sName,sValue){
	try{
		if(!xml_dom_isdocs(oXml)) return false
		if(sOpt == "add"){
			if(__H.isUndefined(sName,sValue)) return false
			var oAttr = oXml.createAttribute(sName);
			oAttr.appendChild(oXml.createTextNode(sValue)); 
			oNode.setAttributeNode(oAttr);
			return oAttr;
		}
		else if(sOpt == "get"){
			if(!xml_dom_isnodes(oNode)) return false
			var aNAttr = oNode.attributes;
			var aAttr = [];
			for(var oEnum = new Enumerator(aNAttr); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				o = {};
				o.name = oItem.name;
				o.value = oItem.nodeValue;
				o.type = xml_dom_types("getnodetype",oItem.nodeType);
				aAttr.push(o);
			}
			return aAttr
		}
		return false;
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

function xml_dom_node(sOpt,oXml,oNode,sNode,oONode,aONodes){
	try{		
		if(!xml_dom_isdocs(oXml)) return false
		else if(sOpt == "addnode"){
			oAddChild = oXml.documentElement ? oXml.documentElement : oXml
			if(!xml_dom_isnodes(oAddChild,oNode)) return false // append oNode as sibling
			if(oParent = oAddChild.parentNode){
				oParent.appendChild(oNode);				
				return oParent.lastChild; // Returns the append node
			}
		}
		else if(sOpt == "clone"){
			if(!xml_dom_isnodes(oNode)) return false
			if(oParent = oNode.parentNode){
				var oCNode = oNode.cloneNode(true);
				oParent.appendChild(oCNode);				
				return oParent.lastChild; // Returns the cloned node
			}
		}
		else if(sOpt == "delnode"){
			if(!xml_dom_isnodes(oNode)) return false
			return oXml.removeChild(oNode); // Returns the removed node
		}
		else if(sOpt == "getnode"){
			if(!xml_dom_isnodes(oNode)) return false
			return {
				path : xml_dom_element("getpath",oXml,oNode),
				name : oNode.nodeName,
				value : oNode.nodeValue,
				text : oNode.text,
				parent : oNode.parentNode,
				type : xml_dom_types("getnodetype",oNode.nodeType),
				attr : xml_dom_attribute("get",oXml,oNode)
			}
		}
		else if(sOpt == "getnodes"){
			var aCNodes = oXml.documentElement.childNodes;
			var aNodes = [];
			for(var i = 0, oONode, oItem, iLen = aCNodes.length; i < iLen; i++){
				oItem = aCNodes.item(i);
				if(oONode = xml_dom_node("getnode",oXml,oItem)){
					aNodes.push(oONode);
				}
			}			
			return aNodes;
		}
		else if(sOpt == "getsibling"){
			if(!xml_dom_isnodes(oNode)) return false
			var oSNode = oNode.nextSibling;
			return xml_dom_node("getnode",oXml,oSNode);
		}
		else if(sOpt == "getnext"){
			if(!xml_dom_isnodes(oNode)) return false
			var oNNode = oNode.nextNode;
			return xml_dom_node("getnode",oXml,oNNode);
		}
		else if(sOpt == "getnextall"){
			if(!xml_dom_isnodes(oNode)) return false			
			var oLNode = oXml.getElementsByTagName(sNode)
			var len = oLNode.Length
			var aNodes = []
			for(var i = 0, oNNode; i < len; i++){
				if(oNNode = xml_dom_node("getnext",oXml,oLNode)){
					aNodes.push(oNNode)	
				}
			}
			return aNodes;
		}
		else if(sOpt == "getxml"){
			if(!xml_dom_isnodes(oNode)) return false
			return oNode.xml;
		}
		else if(sOpt == "show"){
			oONode = typeof(oONode) == "object" ? oONode : xml_dom_isnodes(oNode) ? xml_dom_node("getnode",oXml,oNode) : false
			if(typeof(oONode) != "object") return false
			var s = ""
			for(var o in oONode){
				if(o == "attr"){
					s += "\n" +  o, a = oONode[o];
					for(i = 0, iLen = a.length; i < iLen; i++){
						for(var oo in a[i]){
							if(!a[i].hasOwnProperty(oo)) continue
							s += "\n " +  oo + "\t= " + a[i][oo];
						}
					}
				}
				else s += "\n" +  o + "\t= " + oONode[o];
			}
			return s
		}
		else if(sOpt == "showall"){
			aONodes = typeof(aONodes) == "object" ? aONodes : xml_dom_node("getnodes",oXml)
			if(typeof(aONodes) != "object") return false
			for(var i = 0, ss = "", iLen = aONodes.length; i < iLen; i++){
				if(s = xml_dom_node("show",oXml,null,null,aONodes[i])){
					ss +=  s + "\n"
				}
				
			}
			return ss
		}
		return false
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

function xml_dtd_document(sOpt,oXml,sQualName,sPublicId,sSystemId,sSubset,sType){
	try{
		if(!__H.isObject(oXml)) return false
		if(sOpt == "create"){
			return; // CreateDocumentType is only supported from .NET
			
			if(__H.isUndefined(sQualName)) return false
			sPublicId = sPublicId ? sPublicId : null
			sSystemId = xml_doc_isvalid(sSystemId) ? sSystemId : "SYSTEM" // sSystemId: books.dtd
			var oType = oXml.CreateDocumentType(sQualName,sPublicId,sSystemId,"<!ELEMENT " + sQualName + " " + (sType ? sType : "ANY") + ">");
		    oXml.appendChild(oType);
		}
		else if(sOpt == "getname"){	
			return oXml.doctype.name
		}
		return false;
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

function xml_dtd_entities(sOpt,oXml,sQualName,sPublicId,sSystemId,sSubset,sType){
	try{
		if(!__H.isObject(oXml)) return false
		else if(oXml.doctype == null) return false
		else if(sOpt == "add"){
			return; // CreateDocumentType is only supported from .NET
			if(__H.isStringEmpty(sQualName)) return false
			sPublicId = sPublicId ? sPublicId : null
			sSystemId = sSystemId ? sSystemId : "SYSTEM"
			//sSystemId = xml_doc_isvalid(sSystemId) ? sSystemId : "SYSTEM"
			var oType = oXml.CreateDocumentType(sQualName,sPublicId,sSystemId,"<!ELEMENT " + sQualName + " " + (sType ? sType : "ANY") + ">");
		    oXml.appendChild(oType);
    		return oType;
		}
		else if(sOpt == "getentities"){
			var oEntities = oXml.doctype.entities, aEntities = []
			for(var oEnum = new Enumerator(oEntities); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				var o = {}
				o.name = oItem.nodeName
				o.text = oItem.text
				o.type = oItem.dataType
				aEntities.push(o);
			}
    		return aEntities;
		}
		
		return false;
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

function xml_dtd_notation(sOpt,oXml){
	try{ // http://www.devguru.com/technologies/xmldom/quickref/obj_notation.html__H
		if(!__H.isObject(oXml)) return false
		else if(oXml.doctype == null) return false
		else if(sOpt == "getnotations"){
			var oNotations = oXml.doctype.notations, aNotations = []
			for(var oEnum = new Enumerator(oNotations); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				var o = {}
				o.type = oItem.nodeType
				o.name = oItem.nodeName
				o.public = oItem.publicID
				o.system = oItem.systemID
				aNotations.push(o);
			}
    		return aNotations;
		}
		return false;
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}



function xml_dom_types(sOpt,iType){
	try{		
		if(sOpt == "getnodetype"){ // http://www.devguru.com/technologies/xmldom/quickref/obj_node.html#types
			if(iType == 1) return "1-NODE_ELEMENT"
			else if(iType == 2) return "2-NODE_ATTRIBUTE"
			else if(iType == 3) return "3-NODE_TEXT"
			else if(iType == 4) return "4-NODE_CDATA_SECTION"
			else if(iType == 5) return "5-NODE_ENTITY_REFERENCE"
			else if(iType == 6) return "6-NODE_ENTITY"
			else if(iType == 7) return "7-NODE_PROCESSING_INSTRUCTION"
			else if(iType == 8) return "8-NODE_COMMENT"
			else if(iType == 9) return "9-NODE_DOCUMENT"
			else if(iType == 10) return "10-NODE_DOCUMENT_TYPE"
			else if(iType == 11) return "11-NODE_DOCUMENT_FRAGMENT"
			else if(iType == 12) return "12-NODE_NOTATION"
			else return ""
		}
		return false
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

function xml_dom_xpath(sOpt,oXml,sXPath){
	try{ // http://www.devguru.com/technologies/xmldom/quickref/node_selectNodes.html http://www.w3.org/TR/xpath__H
		if(!xml_dom_isdocs(oXml)) return false
		else if(sOpt == "getnodes"){
			var aOXPath = oXml.documentElement.selectNodes(sXPath);
			var aNodes = [];
			for(var i = 0, oONode, oItem, len = aOXPath.length; i < len; i++){
				oItem = aOXPath.item(i);
				if(oONode = xml_dom_node("getnode",oXml,oItem)){
					aNodes.push(oONode);
				}
			}
			return aNodes;
		}
		return false
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}


////////////////////////////////////////////////////////////////////////////////
///////// HTTP FUNCTIONS


function xml_http_document(sOpt,sUrl,bAsync,sUser,sPassword){
	try{
		bAsync = !bAsync ? bAsync : true
		if(!__H.isUndefined(sUser,sPassword));
		else sUser = sPassword = null
		var oHttp = new ActiveXObject("Microsoft.XMLHTTP");
		if(sOpt == "get"){			
			if(oHttp != null){
				oHttp.open("GET",sUrl,true,sUser,sPassword)
				oHttp.onreadystatechange = function(){
      				if(oHttp.readyState == 4 && oHttp.status == 200){
        				xml_http_handler(oHttp.responseXML);
      				}
      			};
				oHttp.send(null);
			}
		}
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

function xml_http_handler(oXml){
	try{
		if(!xml_dom_isdocs(oXml)) return false
		
		return true
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

////////////////////////////////////////////////////////////////////////////////
///////// IS FUNCTIONS

function xml_dom_isdocs(oStreamXml){
	try{
		for(var oXml, i = 0, iLen = arguments.length; i < iLen; i++){
			oXml = arguments[i];
			if(!__H.isObject(oXml)) return false
			else if(typeof(oXml) != "object" && !oXml.hasChildNodes) return false
		}
		return true
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

function xml_dom_isnodes(oStreamNodes){
	try{
		for(var oNode, i = 0, iLen = arguments.length; i < iLen; i++){
			oNode = arguments[i];
			if(__H.isUndefined(oNode)) return false
			else if(typeof(oNode) != "object" && __H.isUndefined(oNode.nodeName)) return false
		}
		return true
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

function xml_doc_isvalid(sXml){
	try{
		if(__H.isStringEmpty(sXml)) return false
		sExt = (oFso.GetExtensionName(sXml)).toLowerCase()
		return (sExt).isSearch(/xml|xsl|xsd|dtd/ig);
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}


////////////////////////////////////////////////////////////////////////////////
///////// ERROR FUNCTIONS


function xml_log_errorxml(oErr){
	try{		
		if(oErr.errorCode != 0){
			sError = "\nXML Text:      " + (__HKeys.err_source = oErr.srcText);
			sError += "\nXML Code:     " + (__HKeys.err_errorCode = (oErr.errorCode & 0xFFFF));
			sError += "\nXML Url:      " + (__HKeys.err_url = oErr.url);
			sError += "\nXML Line:     " + (__HKeys.err_line = oErr.line);
			sError += "\nXML Char:     " + (__HKeys.err_linechar = oErr.linepos);
			sError += "\nXML Position: " + (__HKeys.err_linepos = oErr.filepos);
			sError += "\nXML Reason:\t" + (__HKeys.err_reason = (oErr.reason).trim());
			__HLog.error(2,new Error(__HLog.errorCode("error"),sError),"")
			return true
		}
		return false
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}





