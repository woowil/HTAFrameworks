// nOsliw Solutions - HTML Aplication Framework.
//**Start Encode**

// http://factsite.co.uk/en/wikipedia/l/li/linked_list.html

test()
function test(){
	var oFruits = new js_list_doublelinked(), oFruits1, oFruits2
	WScript.Echo(2)
	oFruits.addNode("pear")
	WScript.Echo(3)
	oFruits1 = oFruits.addTail()
	WScript.Echo(4)
	oFruits.setData(oFruits1,"banana")
	WScript.Echo(oFruits.getSize())
	oCoconut = oFruits.addNode("coconut")
	WScript.Echo(6)
	oFruits.addTail(null,"grapefruit")
	WScript.Echo(7)
	oFruits.addHead(null,"olives")
	WScript.Echo(8)
	oCoconut.addNode(null,"orange")
	oFruits.addTail(null,"cucumber")
	oFruits.print("\nTHE LIST")			
	oFruits.print(oFruits.getDataAll("\nGetting: "))
	oFruits2 = oFruits.reverse()
	oFruits.print("\nTHE REVERSE LIST (not working as it should)")
	oFruits.print(oFruits2.getDataAll("\nGetting: "))
	oFruits = oFruits.shuffle()
	oFruits.print("\nTHE SHUFFLED LIST") // not working properly
	oFruits.print(oFruits.getDataAll("\nGetting: "))
	oFruits1 = oFruits1 = null
}

function js_list_doublelinked(){
	try{
		// Private (hidden) declarations	
		var head = null;
		var tail = null;
		var size = 0;
		
		// CONTRUCTOR
		
		this.Node = function(oData){
			this.data = oData ? oData : null
			this.next = null
			this.prev = null
		};		
		
		// JScript array compatibility
		/////////////////////////////////////////////////////////////
		
		this.push = function(oData){
			return this.addNode(oData)
		};
		
		this.length = function(){
			return this.getSize()
		};
		
		this.reverse = function(){
			var oNode = this.getTail();
			var oReverse = new js_list_doublelinked() 
			while(hasPrev(oNode)){
				oReverse.addNode(oNode.data)
				oNode = oNode.prev
			}
			return oReverse
		};
		
		// ADD/SET functions
		/////////////////////////////////////////////////////////////
		
		this.addNode = function(oData){
			var oNode = new this.Node(oData)
			return this.addTail(oNode)
		};
		
		this.addTail = function(oNode,oData){
			oNode = isNode(oNode) ? oNode : new this.Node()
			if(oData) oNode.data = oData
			var oTail = this.getTail()
			if(oTail == null){ // Last Node
				return this.addHead(oData)
			}
			else return this.addNext(oTail,oNode);
		};
		
		this.addHead = function(oNode,oData){
			oNode = isNode(oNode) ? oNode : new this.Node()
			if(oData) oNode.data = oData
			var oHead = this.getHead()
			if(oHead == null){ // First Node
				setHead(oNode)
				setTail(oNode)
				setSize()
				oNode.prev = null
				oNode.next = null
				return oNode
			}
			else return this.addPrev(oHead,oNode);
		};
		
		this.addNext = function(oPrevNode,oNode){
			oNode = isNode(oNode) ? oNode : new this.Node()
			oNode.prev = oPrevNode
			oNode.next = oPrevNode.next
			if(oPrevNode.next == null) setTail(oNode)
			else oPrevNode.next.prev = oNode
			oPrevNode.next = oNode
			setSize()
			return oNode
		};
		
		this.addPrev = function(oNextNode,oNode){
			oNode = isNode(oNode) ? oNode : new this.Node()
			oNode.prev = oNextNode.prev
			oNode.next = oNextNode
			if(oNextNode.next == null) setHead(oNode)			
			else oNextNode.prev.next = oNode
			oNextNode.next = oNode
			setSize()
			return oNode
		};
		
		this.addByIndex = function(iIndex,oData){
			var oIndex = this.getNodeByIndex(iIndex)
			if(oIndex != null){
				this.setData(oData)
				return true
			}
			return false
		};
		
		this.setData = function(oNode,oData){
			if(isNode(oNode)){
				oNode.data = oData ? oData : null
				return true
			}
			return false
		};
		
		var setTail = function(oNode){ // PRIVATE FUNCTION
			tail = oNode
		};
		
		var setHead = function(oNode){ // PRIVATE FUNCTION
			head = oNode
		};
		
		var setSize = function(){ // PRIVATE FUNCTION
			size++
		};
		
		// BOOLEAN functions
		/////////////////////////////////////////////////////////////
		
		var isNode = function(oNode){ // PRIVATE FUNCTION
			return (typeof(oNode) == "object" && oNode != null); // Must be identical object
		};
		
		this.isEmpty = function(oNode){			
			return (oNode == null || oNode.data == null || oNode.getSize() == 0);
		};	
		
		var isData = function(oData1,oData2){ // PRIVATE FUNCTION
			return (oData1 === oData2 && oData1 == oData2); // the operator must be identical
		};
		
		var hasNext = function(oNode){ // PRIVATE FUNCTION
			if(isNode(oNode)){
				return (oNode != null && oNode.next != null);
			}
			return false
		};
		
		var hasPrev = function(oNode){ // PRIVATE FUNCTION
			if(isNode(oNode)){
				return (oNode != null && oNode.prev != null);
			}
			return false
		};
		
		this.dataExists = function(oData){
			var oNode = this.getHead()
			while(hasNext(oNode)){
				if(!this.isEmpty(oNode)){
					if(oNode.data === oData && oNode.data == oData) return true
				}
			}
			return false
		};
		
		// GET functions
		/////////////////////////////////////////////////////////////
		
		this.getHead = function(){			
			return head
		};
		
		this.getTail = function(){
			return tail
		};
		
		
		this.getSize = function(){
			return size;
		};
		
		this.getIndex = function(oNode){
			if(isNode(oNode)){
				for(var iIndex = 0; hasPrev(oNode); iIndex++);
				return iIndex;
			}
			return 0
		};		
		
		this.getNodeByIndex = function(iIndex){
			if(isNaN(iIndex) || iIndex < 0) return null
			oNode = this.getHead();
			for(var i = 0; i < iIndex && hasNext(oNode); i++){
				oNode = oNode.next;
			}
			return oNode;
		};		
		
		this.getData = function(oNode){
			if(isNode(oNode)){
				return oNode.data
			}
			return null
		};
		
		this.getDataAll = function(sSeperator){
			var oNode = this.getHead()
			sSeperator = sSeperator ? sSeperator : "\n"
			var sData = sSeperator + oNode.data
			while(hasNext(oNode)){
				oNode = oNode.next
				sData += sSeperator + oNode.data
			}
			return sData
		};
		
		this.getDataByIndex = function(iIndex){
			if(oNode = this.getNodeByIndex(iIndex)) return oNode.data
			return null
		};
		
		// CHANGE functions
		/////////////////////////////////////////////////////////////
		
		this.delNode = function(oNode){
			if(!isNode(oNode)) return;
			if(oNode.prev == null) setHead(oNode.next) // this.head = 
			else oNode.prev.next = oNode.next
			if(oNode.next == null) setTail(oNode.prev) // this.tail = 
			else oNode.next.prev = oNode.prev
			oNode.data = oNode.name = null;
			oNode = null;
		};
		
		this.delNodeAll = function(){
			var oNode = this.getTail()
			while(hasPrev(oNode)){
				oNode = oNode.prev
				this.delNode(oNode.next)
			}
			oNode = head = tail = null
		};
		
		this.delByIndex = function(iIndex){
			var oNode = this.getNodeByIndex(iIndex)
			this.delNode(oNode)
		};
		
		this.delByData = function(oData){
			var oTail = oNode.getTail()
			while(isData(oTail.data,oData)){
				this.delNode(oTail)
			}
		};
		
		this.shuffle = function(){ // NOT working 100%
			if(this.getSize() == 0) return null
			var oNode = this.getHead(), oRand
			var oShuffle = new js_list_doublelinked()
			for(var i = this.getSize(); hasNext(oNode) && 0 < i; i--){
				oRand = this.getNodeByIndex(random(0,i))
				oShuffle.addNode(oRand.data)
				this.delNode(oRand)
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
		
		// CONVERT functions
		/////////////////////////////////////////////////////////////
		
		this.toArray = function(){
			var aNode = new Array();	
			var oNode = this.getHead();
			while(hasNext(oNode)){
				aNode.push(oNode.data);
				oNode = oNode.next;				
			}
			return aNode
		};
		
		this.toNode = function(aNode){
			var oNode = new js_list_doublelinked()
			if(typeof(aNode) == "object"){
				for(i in aNode){
					oNode.addNode(aNode[i])
				}
			}
			return oNode
		}
		
		this.toDict = function(oNode){
			var oDict = new ActiveXObject("Scripting.Dictionary"), i = 1, sNum;
			var oNode = this.getHead()
			while(hasNext(oNode)){
				sNum = (i++).toString() // Bug
				oDict.Add(sNum,oNode.data)	
				oNode = oNode.next
			}
			return oDict;
		};
		
		// OTHER functions
		/////////////////////////////////////////////////////////////

		var random = function(lval,hval){
			 return Math.ceil(Math.random()*hval + lval)
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
			for(var i = 0, len = o1.length; i < len; i++){
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
			for(var i = 0, len = a1.length; i < len; i++){
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
			
			for(var i = 0, len = a2.length; i < len; i++){
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
