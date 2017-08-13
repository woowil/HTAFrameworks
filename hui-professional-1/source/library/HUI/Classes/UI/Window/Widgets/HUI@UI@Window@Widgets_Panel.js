// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

__H.include("HUI@UI@Window@Widgets.js","HUI@UI@Window@Elements.js")

__H.register(__H.UI.Window.Widgets,"Panel","Panel",function Panel(){
	var o_this = this
	var b_initialized = false
	
	var d_panels
	var d_tabs
	var o_div_panel
	var o_div_panel_tmp
	var o_div_titles
	var o_div_panes
	
	var o_tab_cur = {
		id       : "",
		title    : null,
		pane     : null,
		isactive : false,
		index    : 0
	}
	var o_tab_active = {
		id       : "",
		title    : null,
		pane     : null,
		isactive : false,
		index    : 0,
		clear    : function(){
			this.id = ""
			this.title = this.pane = null
			this.isactive = false
			this.index = 0
		}
	}
	
	/////////////////////////////////////
	//// DEFAULT
	
	var o_options = {
		s_path_pic     : "/data/pic/",
		i_max_tite     : 20,
		i_max_tite_dot : 3,
		i_titles_wd    : 0,
		
		s_class_frame        : "hui-panel-frame",
		s_class_titles       : "hui-panel-titles",
		s_class_title        : "hui-panel-title",
		s_class_title_active : "hui-panel-title-active",
		s_class_title_over   : "hui-panel-title-over",
		s_class_title_out    : "hui-panel-title-out",
		s_class_panes        : "hui-panel-panes",
		s_class_pane         : "hui-panel-pane"
	}
	
	var initialize = function initialize(){
		if(b_initialized) return;
		
		d_panels    = new ActiveXObject("Scripting.Dictionary")
		d_tabs      = new ActiveXObject("Scripting.Dictionary")
		
		b_initialized = true
	}	
	
	this.setOptions = function setOptions(oOptions){		
		try{
			Object.extend(o_options,oOptions,true)
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getOptions = function(){
		return o_options
	}
	
	/////////////////////////////////////
	//// 
	
	this.isPanel = function isPanel(oPanel){
		try{
			if(typeof(oPanel) == "string") oPanel = __H.byIds(oPanel)
			if(__H.isElement(oPanel,"DIV") && !__H.isStringEmpty(oPanel.className,oPanel.id)){
				var o = oPanel.childNodes;
				if(o.length == 1 && __H.isElement(o(0),"DIV")){
					return !__H.isStringEmpty(o(0).className)
				}
				else if(o.length == 2 && __H.isElement(o(0),"DIV") && __H.isElement(o(1),"DIV")){
					if(!__H.isStringEmpty(o(0).className,o(1).className)){
						return o(0).childNodes.length == o(1).childNodes.length
					}
				}
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return false
	}

	this.loadPanel = function loadPanel(oPanel){
		initialize()
		try{
			if(typeof(oPanel) == "string") oPanel = __H.byIds(oPanel)
			if(!this.isPanel(oPanel)) return false;
			else if(oPanel.loaded){
				__HLog.debug("Panel: " + oPanel.id + " already loaded!")
				o_div_panel = d_panels(oPanel.id)
				o_div_panel.onmouseover()
				return true
			}
			
			__HLog.debug("Loading Panel: " + oPanel.id)
			if(!d_panels.Exists(oPanel.id)) d_panels.Add(oPanel.id,oPanel)
			this.setPanel(oPanel.id)
			for(var i = 0, iLen = o_div_panes.childNodes.length; i < iLen; i++){
				o_tab_cur.id = oPanel.id + "_t" + i
				o_tab_cur.isactive = false
				o_tab_cur.pane = o_div_panes.childNodes(i)
				o_tab_cur.title = __H.byClone("button")
				o_tab_cur.title.index = o_tab_cur.index = i
				if(o_tab_cur.pane.title.length > 0){
					o_tab_cur.title.innerHTML = (o_tab_cur.pane.title).limitize(o_options.i_max_tite,o_options.i_max_tite_dot)
				}
				else {
					o_tab_cur.title.innerHTML = "Title #" + (i < 10 ? "0" + i : i)
				}				
				this.setTab(o_tab_cur)				
				o_div_titles.appendChild(o_tab_cur.title)
				d_tabs.Add(o_tab_cur.id,o_tab_cur)
				if(o_tab_cur.pane.disabled) this.disable(oPanel.id,i)
			}			
			oPanel.loaded = true
			//this.activate(oPanel.id,0)
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.loadPanels = function loadPanels(oContainer){
		try{
			var bReturn = true
			var oContainer = __H.isElement(oContainer) ? oContainer : document
			var oElements = oContainer.getElementsByTagName("DIV")
			for(var i = 0, iLen = oElements.length; i < iLen; i++){
				if(oElements[i].className == o_options.s_class_frame){
					if(!this.loadPanel(oElements[i])) bReturn = false
				}
			}
			return bReturn
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.setPanel = function setPanel(id){
		try{
			if(!(o_div_panel = this.getPanel(id))) return false;
			
			o_div_panel.className = o_options.s_class_frame
			o_div_panel.active_index = o_tab_active.index
			o_div_panel.onmouseover = function(){
				o_div_panel = this
				if(o_div_panel.id != this.id){
					o_tab_active.id = this.id
					o_tab_active.title = null
					o_tab_active.pane = null
					o_tab_active.isactive = false
					o_tab_active.index = 0					
				}
				else{
					o_tab_active.id = this.id
					o_tab_active.title = this.childNodes(0).childNodes(this.active_index)
					o_tab_active.pane = this.childNodes(1).childNodes(this.active_index)
					o_tab_active.isactive = true
					o_tab_active.index = this.active_index
				}
			}
			o_div_panel.onmouseout = function(){
				o_div_panel = null
			}
			
			if(o_div_panel.childNodes.length == 1){
				o_div_titles = __H.byClone("div")
				o_div_titles.className = o_options.s_class_titles
				o_div_titles.unselectable = "on"
				o_div_panel.insertBefore(o_div_titles,o_div_panel.firstChild)				
			}
			else o_div_titles = o_div_panel.firstChild
			o_div_panes = o_div_panel.lastChild
			o_div_panes.className = o_options.s_class_panes

			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getPanel = function getPanel(id){
		try{
			if(__H.isStringEmpty(id)) __HExp.ArgumentIllegal();
			if(d_panels.Exists(id)) return d_panels(id)
			return null
		}
		catch(ee){
			__HLog.error(ee,this)
			return null
		}
	}

	this.addPanel = function addPanel(parent,id){
		try{
			if(!__H.isElement(parent) || __H.isStringEmpty(id)) __HExp.ArgumentIllegal()
			if(d_panels.Exists(id)){
				__HLog.debug("Panel id:" + id + " already exist")
				o_div_panel = d_panels(id)
				o_div_panel.onmouseover()
				return o_div_panel
			}
			o_div_panel_tmp = this.create(null,id);
			parent.appendChild(o_div_panel_tmp)		
			return this.loadPanel(id)	 
		}
		catch(ee){
			__HLog.error(ee,this)
			return null
		}
	}

	this.delPanel = function delPanel(id){
		try{
			if(__H.isStringEmpty(id)) return false;
			if(!d_panels.Exists(id)) return true
			d_panels.Remove(id)
			if(o_div_panel_tmp = __H.byIds(id)) o_div_panel_tmp.removeNode(true)
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.create = function create(oDiv,id){
		try{
			if(__H.isStringEmpty(id)) __HExp.ArgumentIllegal();

			if(o_div_panel_tmp = __H.byIds(id));
			else o_div_panel_tmp = __H.isElement(oDiv,"div") ? oDiv : __H.byClone("DIV")
			o_div_panel_tmp.id = id
			o_div_panel_tmp.className = o_options.s_class_frame
			o_div_panel_tmp.appendChild(__H.byClone("DIV"))
			o_div_panel_tmp.childNodes(0).className = o_options.s_class_panes

			return o_div_panel_tmp
		}
		catch(ee){
			__HLog.error(ee,this)
			return null
		}
	}

	this.reset = function reset(){
		try{

			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.hasTitle = function hasTitle(sTitle){
		try{
			if(!__H.isElement(o_div_panel,"div")) return false
			var oTitles = o_div_panel.childNodes(0).childNodes
			for(var i = 0, iLen = oTitles.length; i < iLen; i++){
				if(oTitles(i).innerHTML == sTitle) return true
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.isTab = function isTab(oTab){
		try{
			return __H.isObject(oTab,oTab.title,oTab.pane)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.setTab = function setTab(oTab){
		try{
			if(!this.isTab(oTab)) return false
			
			oTab.title.className = o_options.s_class_title
			oTab.title.unselectable = "on"
			oTab.title.onmouseover = function(){
				if(this.isactive) return;
				this.className = o_options.s_class_title_over
				o_tab_cur.index = this.index
				o_tab_cur.title = this
				o_tab_cur.pane = this.parentNode.nextSibling.childNodes(this.index)
			}
			oTab.title.onmouseout = function(){
				if(!this.isactive) this.className = o_options.s_class_title_out
				this.blur()
			}
			oTab.title.onclick = function(){
				if(this.index == o_tab_active.index && this.index != 0) return;
				
				__HElem.hide(oTab.pane,o_tab_active.pane)
				__HElem.show(oTab.pane.parentNode.childNodes(this.index))
				
				this.className = o_options.s_class_title_active
				this.isactive = true
				
				if(o_tab_active.title != null){
					o_tab_active.title.className = o_options.s_class_title
					//o_tab_active.pane.className = o_options.s_class_pane
					o_tab_active.isactive = o_tab_active.title.isactive = false
					o_tab_active.index = this.index					
					o_div_panel.active_index = this.index					
				}
				
				this.blur()
			}
			oTab.title.ondblclick = function(){
				if(!this.isactive) return;				
				this.isactive = false
				this.onmouseout()
				__HElem.hide(o_tab_cur.pane)
				o_tab_active.clear()
				this.blur()
			}
			
			oTab.pane.className = o_options.s_class_pane			
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.setTabByIndex = function setTabByIndex(idx){
		try{ // TODO Nor working
			if(isNaN(idx) || d_tabs.Count == 0) return false
			if(!idx.between(0,d_tabs.Count)) return false
			
			var a = new VBArray(d_tabs.Keys()).toArray()
			return this.setTab(a[idx])
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getTab = function getTab(id,index){
		try{			
			if(!(o_div_panel = this.getPanel(id)));			
			else if(typeof(index) == "number" && index >= 0){
				o_div_panel.onmouseover()
				if(d_tabs.Exists(o_div_panel.id + "_t" + index)){
					return d_tabs(o_div_panel.id + "_t" + index)
				}
			}
			return null
		}
		catch(ee){
			__HLog.error(ee,this)
			return null
		}
	}

	this.addTab = function addTab(sTitle,sHtml,bActivate){
		try{
			if(!__H.isElement(o_div_panel,"div") || !o_div_panel.loaded){
				__HLog.debug("Panel is not initialized or not set")
				return null
			}
			if(__H.isStringEmpty(sTitle)) __HExp.ArgumentIllegal();
			sTitle = sTitle.limitize(o_options.i_max_tite,o_options.i_max_tite_dot)
			if(this.hasTitle(sTitle)){
				__HLog.debug("Ignoring add. Tab: " + o_div_panel.id + "::'" + sTitle + "' already exists.")
				return null
			}
			//if(o_tab_cur.id == "Title #00") o_tab_cur.removeNode()
			var i = o_div_titles.childNodes.length
			
			o_tab_cur.id = o_div_panel.id + "_t" + i
			o_tab_cur.isactive = false
			
			o_tab_cur.title = __H.byClone("button")
			o_tab_cur.title.innerHTML = o_tab_cur.title.title = sTitle
			o_tab_cur.title.index = o_tab_cur.index = i
			o_tab_cur.title.isactive = false
			
			o_tab_cur.pane = __H.byClone("div")
			o_tab_cur.pane.innerHTML = typeof(sHtml) == "string" ? sHtml : "Pane #" + (i < 10 ? "0" + i : i)
			
			this.setTab(o_tab_cur)
			
			if(o_options.i_titles_wd < o_div_titles.offsetWidth) o_div_titles.appendChild(o_tab_cur.title)
			else o_div_titles.insertBefore(o_tab_cur.title,o_div_titles.firstChild)
			//o_options.i_titles_wd += o_tab_cur.title.offsetWidth
			o_div_panes.appendChild(o_tab_cur.pane)
			
			d_tabs.Add(o_tab_cur.id,o_tab_cur)
			return o_tab_cur
		}
		catch(ee){
			__HLog.error(ee,this)
			return null
		}
		finally{
			if(bActivate) o_tab_cur.title.onclick()
		}
	}

	this.addTabs = function addTabs(aOTab){
		try{
			if(!__H.isArray(aOTab)) return false
			for(var i = 0, iLen = aOTab.length; i < iLen; i++){
				if(this.isTab(aOTab[i])){
					this.addTab(aOTab[i].title,aOTab[i].pane)
				}
			}
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.activate = function activate(id,index){
		if(o_tab_cur_tmp = this.getTab(id,index)){
			if(!o_tab_cur_tmp.title.disabled){
				o_tab_cur_tmp.title.onmouseover()
				o_tab_cur_tmp.title.onclick()
				return true
			}			
		}
		return false
	}

	this.deactivate = function deactivate(id,index){
		if(o_tab_cur_tmp = this.getTab(id,index)){
			if(!o_tab_cur_tmp.title.disabled){
				o_tab_cur_tmp.title.ondblclick()			
				return true
			}
		}
		return false
	}
	
	this.enable = function enable(id,index){
		if(o_tab_cur_tmp = this.getTab(id,index)){
			o_tab_cur_tmp.title.disabled = false
			o_tab_cur_tmp.pane.disabled = false
			return true
		}
		return false	
	}
	
	this.disable = function disable(id,index){
		if(o_tab_cur_tmp = this.getTab(id,index)){
			o_tab_cur_tmp.title.disabled = true
			o_tab_cur_tmp.pane.disabled = true
			return this.deactivate(id,index)
		}
		return false	
	}
	
	this.setColor = function setColor(sElement,sColor){
		try{
			if(__H.isStringEmpty(sElement,sColor) || !sElement.isSearch(/color$/ig)) __HExp.ArgumentIllegal()
			var bReturn = true
			for(var o in o_options){
				if(!o_options.hasOwnProperty(o)) continue
				if(o.search("class") == -1) continue 
				if(!__HStyle.setStyle(o_options(o),sElement,sColor,true)) bReturn = false
			}
			return bReturn
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getColor = function getColor(){
		try{
			
			return o_options
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.setTheme = function setTheme(){
		try{
			
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getTheme = function getTheme(){
		try{
			
			return o_options
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.setPanel2 = function setPanel2(){
		try{

			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
})

var __HPanel = new __H.UI.Window.Widgets.Panel()
