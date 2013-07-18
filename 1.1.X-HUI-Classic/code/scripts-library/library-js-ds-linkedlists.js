// nOsliw Solutions - HTML Aplication Framework.
//**Start Encode**

// http://factsite.co.uk/en/wikipedia/l/li/linked_list.html


function js_list_doublelinked(oData,sName){
	try{
		this.data = oData ? oData : null
		this.name = sName ? new String(sName) : null
		this.next = null
		this.prev = null
		this.size = 0;
		//this.size++;
		
		// TEST functions
		/////////////////////////////////////////////////////////////
		
		this.testNode = function(){
			var oFruits = new js_list_doublelinked("apple"), oFruits1, oFruits2
			oFruits.addTail(null,"pear")
			oFruits1 = oFruits.addTail()
			oFruits1.setData(null,"banana")
			oCoconut = oFruits.addHead(null,"coconut")
			oFruits.addTail(null,"grapefruit")
			oFruits.addHead(null,"olives")
			oCoconut.addNext(null,"orange")
			oFruits.addTail(null,"cucumber")
			this.print("\nTHE LIST")			
			this.print(oFruits.getDataAll(oFruits,"\nGetting: "))
			oFruits2 = oFruits.reverse()
			this.print("\nTHE REVERSE LIST (not working as it should)")
			this.print(oFruits.getDataAll(oFruits2,"\nGetting: "))
			oFruits = this.shuffle(oFruits)
			this.print("\nTHE SHUFFLED LIST") // not working properly
			this.print(oFruits.getDataAll(oFruits,"\nGetting: "))
			oFruits1 = oFruits1 = null
		};
		
		// JScript array compatibility
		/////////////////////////////////////////////////////////////
		
		this.push = function(oData){
			return this.addTail(null,oData)
		};
		
		this.length = function(){
			return this.getSize()
		};
		
		this.reverse = function(oNode){
			oNode = this.isNode(oNode) ? oNode : this
			oNode = oNode.getTail();
			var oReverse = null
			while(oNode.hasPrev()){
				oReverse = oReverse ? oReverse : new js_list_doublelinked(oNode.data,oNode.name) 
				oNode = oNode.prev
				oReverse.addTail(null,oNode.data,oNode.name)				
			}
			return oReverse
		};
		
		// ADD/SET functions
		/////////////////////////////////////////////////////////////
		
		this.addTail = function(oNode,oData,sName){
			oNode = this.isNode(oNode) ? oNode : this
			var oTail = oNode.getTail()
			return oTail.addNext(null,oData,sName);
		};
		
		this.addHead = function(oNode,oData,sName){
			oNode = this.isNode(oNode) ? oNode : this
			var oHead = oNode.getHead()
			return oHead.addNext(null,oData,sName);
		};
		
		this.addNext = function(oNode,oData,sName){
			oNode = this.isNode(oNode) ? oNode : new js_list_doublelinked(oData,sName)
			oNode.prev = this
			oNode.next = this.next
			if(this.next == null) this.setTail(oNode)
			else this.next.prev = oNode
			this.next = oNode
			oNode.data = oData ? oData : null
			oNode.name = sName ? new String(sName) : null
			return oNode
		};
		
		this.addPrev = function(oNode,oData,sName){
			oNode = this.isNode(oNode) ? oNode : new js_list_doublelinked(oData,sName)
			oNode.prev = this.prev
			oNode.next = this
			if(this.prev == null) this.setHead(oNode)
			else this.prev.next = oNode
			this.prev = oNode
			oNode.data = oData ? oData : null
			oNode.name = sName ? new String(sName) : null
			return oNode
		};
		
		this.addByIndex = function(oNode,iIndex,oData,sName){
			oNode = this.isNode(oNode) ? oNode : this
			var oIndex = oNode.getNodeByIndex(null,iIndex)
			if(oIndex != null){
				oIndex.setData(null,oData,sName)
				return true
			}
			return false
		};
		
		this.setData = function(oNode,oData,sName){
			oNode = this.isNode(oNode) ? oNode : this
			oNode.data = oData ? oData : null
			oNode.name = sName ? new String(sName) : null
			return oNode
		};
		
		this.setDataTail = function(oNode,oData,sName){
			oNode = this.isNode(oNode) ? oNode : this
			oNode = oNode.getTail()
			return oNode.setData(null,oData,sName)
		};
		
		this.setDataHead = function(oNode,oData,sName){
			oNode = this.isNode(oNode) ? oNode : this
			oNode = oNode.getHead()
			return oNode.setData(null,oData,sName)
		};
		
		this.setTail = function(oNode){
			oNode = this.isNode(oNode) ? oNode : this
			LIST_DOUBLE_TAIL = oNode // Defines a global variable
		};
		
		this.setHead = function(oNode){
			oNode = this.isNode(oNode) ? oNode : this
			LIST_DOUBLE_HEAD = oNode // Defines a global variable
		};
		
		// BOOLEAN functions
		/////////////////////////////////////////////////////////////
		
		this.isNode = function(oNode){
			return (typeof(oNode) == "object" && oNode != null); // Must be identical object
		};
		
		this.isEmpty = function(oNode){
			oNode = this.isNode(oNode) ? oNode : this
			return (oNode == null || oNode.data == null || oNode.getSize() == 0);
		};	
		
		this.isData = function(oData1,oData2){
			return (oData1 === oData2 && oData1 == oData2); // the operator must be identical
		};
		
		this.isName = function(sName1,sName2){
			return (sName1 === sName2 && sName1 == sName2); // the operator must be identical
		};
		
		this.hasNext = function(oNode){
			oNode =  this.isNode(oNode) ? oNode : this
			return (oNode.next != null);
		};
		
		this.hasPrev = function(oNode){
			oNode = this.isNode(oNode) ? oNode : this
			return (oNode != null && oNode.prev != null);
		};
		
		this.dataExists = function(oNode,oData){
			oNode = this.isNode(oNode) ? oNode : this
			oNode = oNode.getHead()
			while(oNode.hasNext()){
				if(!oNode.isEmpty()){
					oNode.data == oData
					return true
				}
			}
			return false
		};
		
		// GET functions
		/////////////////////////////////////////////////////////////
		
		this.getHead = function(oNode){			
			try{
				throw "Oops I did it again!"
				//return LIST_DOUBLE_HEAD // globally defined
			}
			catch(ee){
				oNode = this.isNode(oNode) ? oNode : this
				while(oNode.hasPrev()){
					oNode = oNode.prev;
				}
				return oNode;
			}
		};
		
		this.getTail = function(oNode){
			try{
				throw "Oops I did it again!"
				//return LIST_DOUBLE_TAIL // defined
			}
			catch(ee){
				oNode = this.isNode(oNode) ? oNode : this
				while(oNode.hasNext()){
					oNode = oNode.next;
				}
				return oNode;
			}
		};
		
		this.getIndex = function(oNode){
			oNode = oNode ? oNode : this
			for(var iIndex = 0; oNode.hasPrev(); iIndex++);
			return iIndex;
		};
		
		this.getNodeByIndex = function(oNode,iIndex){
			oNode = this.isNode(oNode) ? oNode : this
			if(isNaN(iIndex) || iIndex < 0) return null
			oNode = oNode.getHead();
			for(var i = 0; i < iIndex && oNode.hasNext(); i++){
				oNode = oNode.next;
			}
			return oNode;
		};
		
		this.getSize = function(oNode){
			oNode = this.isNode(oNode) ? oNode : this
			oNode = oNode.getHead()
			for(var iSize = 1; oNode.hasNext(); iSize++){
				oNode = oNode.next;
			}
			return iSize
		};		
		
		this.getData = function(oNode){
			oNode = this.isNode(oNode) ? oNode : this
			return oNode.data
		};
		
		this.getDataAll = function(oNode,sSeperator){
			oNode = this.isNode(oNode) ? oNode : this
			oNode = oNode.getHead()
			sSeperator = sSeperator ? sSeperator : "\n"
			var sData = sSeperator + oNode.data
			while(oNode.hasNext()){
				oNode = oNode.next
				sData += sSeperator + oNode.data
			}
			return sData
		};
		
		this.getDataByIndex = function(oNode,iIndex){
			oNode = this.isNode(oNode) ? oNode : this
			if(oNode = oNode.getNodeByIndex(null,iIndex)) return oNode.data
			return null
		};
		
		// CHANGE functions
		/////////////////////////////////////////////////////////////
		
		this.delNode = function(oNode){
			oNode = this.isNode(oNode) ? oNode : this
			if(oNode.prev == null) this.setHead(oNode.next) // this.head = 
			else oNode.prev.next = oNode.next
			if(oNode.next == null) this.setTail(oNode.prev) // this.tail = 
			else oNode.next.prev = oNode.prev
			oNode.data = oNode.name = null;
			oNode = null;
		};
		
		this.delNodeAll = function(oNode){
			oNode = this.isNode(oNode) ? oNode : this
			oNode = this.getTail(oNode), oTemp
			while(oNode.hasPrev()){
				this.delNode(oNode)
			}
			oNode = null
		};
		
		this.delByIndex = function(oNode,iIndex){
			oNode = this.isNode(oNode) ? oNode : this
			oNode = oNode.getNodeByIndex(null,iIndex)
			this.delNode(oNode)
		};
		
		this.delByData = function(oNode,oData){
			oNode = this.isNode(oNode) ? oNode : this
			var oTail = oNode.getTail(), oTemp
			while(oTail.isData(oTail.data,oData)){
				oTemp = oTail.prev
				oTemp.delNode(oTail)
				oTail = oTemp
				oTemp = null
			}
			oNode = oTail = null
		};
		
		this.shuffle = function(oNode){ // NOT working 100%
			oNode = this.isNode(oNode) ? oNode : this
			oNode = oNode.getHead()
			var oShuffle = new js_list_doublelinked()
			var iLen, iRand, oRand
			if(!oNode.hasNext()) oShuffle.setData(null,oNode.data,oNode.name) // one node
			else {
				iLen1 = oNode.getSize()
				for(var i = 0; oNode.hasNext() && i < iLen1; i++){
					iLen = oNode.getSize()
					iRand = this.random(0,iLen)
					oRand = oNode.getNodeByIndex(null,iRand)
					if(i == 0) oShuffle.setData(null,oRand.data,oRand.name)
					else oShuffle.addNext(null,oRand.data,oRand.name)
					oNode.delNode(oRand)
					oNode = oNode.getHead()
				}
			}
			return oShuffle
		};
		
		// SHOW functions
		/////////////////////////////////////////////////////////////
		
		this.print = function(sShow){
			try{
				WScript.Echo(sShow)
			}
			catch(ee){ // HTA and HTML
				alert(sShow)
			}
		};
		
		this.info = function(oNode){
			oNode = this.isNode(oNode) ? oNode : this
			// TODO:
		}
		
		// CONVERT functions
		/////////////////////////////////////////////////////////////
		
		this.toArray = function(oNode){
			oNode = this.isNode(oNode) ? oNode : this
			var aNode = new Array();	
			oNode = oNode.getHead();
			while(oNode.hasNext()){
				aNode.push(oNode.data);
				oNode = oNode.next;				
			}
			return aNode
		};
		
		this.toNode = function(aNode){
			var oNode = null
			if(typeof(aNode) == "object"){
				for(i in aNode){
					if(!oNode) oNode = new js_list_doublelinked(aNode[i])
					else oNode.addTail(null,aNode[i])
				}
			}
			return oNode
		}
		
		this.toDict = function(oNode){
			oNode = this.isNode(oNode) ? oNode : this
			var oDict = new ActiveXObject("Scripting.Dictionary"), i = 1;
			oNode = oNode.getHead()
			while(oNode.hasNext()){
				sName = oNode.name ? oNode.name : i++
				oDict.Add(sName,oNode.data)	
			}
			return oDict;
		};
		
		// OTHER functions
		/////////////////////////////////////////////////////////////

		this.random = function(lval,hval){
			var oDate = new Date(), tmp;
			var tmp = Math.round(Math.abs(Math.sin(oDate.getTime())*1000000000/Math.PI));
			return tmp % (hval-lval+1) + lval;
		};
	}
	catch(e){
		WScript.Echo(e.description);
		return false;
	}
}

function js_list_arrays(aArray1,aArray2){
	try{
		this.array1 = aArray1 ? aArray1 : new Array()
		this.array2 = aArray2 ? aArray2 : new Array()
		
		this.sort = function(oo){
			oo = oo ? oo : this.array1
			var o1 = new Array(), o
			for(o in oo){
				o1.push(o)
			}
			o1.sort()
			var o2 = new Object()
			for(var i = 0; i < o1.length; i++){
				o2[o1[i]] = oo[o1[i]]
			}
			return o2
		};
		
		this.delDuplicate = function(a1,a2){
			this.array1 = this.getDistinct(a1,a2)
		};
		
		this.getDistinct = function(a1,a2){
			a1 = a1 ? a1 : this.array1
			a2 = a2 ? a2 : this.array2
			if(typeof(a2) == "object") a1 = a1.concat(a2)
			a1 = a1.sort();
			a2 = new Array()
			for(var i = 0; i < a1.length; i++){
				a2.push(a1[i])
				while(a1[i] == a1[i+1]) i++;
			}
			return a2;
		};
		
		this.getSubstraction = function(a1,a2){ // a1-a2
			a1 = a1 ? a1 : this.array1
			a2 = a2 ? a2 : this.array2
			var aDict = new ActiveXObject("Scripting.Dictionary");
			for(var i = 0; i < a1.length; i++){
				if(!aDict.Exists(a1[i])){ // Remember case sensitive!!
					aDict.Add(a1[i],i);
				}
			}
			
			for(var i = 0; i < a2.length; i++){
				if(aDict.Exists(a2[i])){
					aDict.Remove(a2[i]);
				}
			}
			 return new VBArray(aDict.Keys()).toArray();
		};
		
	}
	catch(e){
			
	}
}
