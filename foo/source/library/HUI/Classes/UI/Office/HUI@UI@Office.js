// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**

__H.include("HUI@UI.js")

__H.register(__H.UI,"Office","Office Automation",function Office(){
	this.objectword = null
	this.file = ""	
	
	var filter = "."
	var directory = "c:\\"
	var WORDSTATEMINIMIZE = 2
	var WORDDIALOGFILEOPEN = 0x50
	var WORDDIALOGFILESAVEAS = 0x54		
	var WORDDIALOGFILEOPENMULTIPLE = 0x100
	
	this.setConfig = function setConfig(sFilter,sDirectory){
		filter = sFilter ? sFilter : filter
		directory = sDirectory ? sDirectory : directory
	}
	
	this.getWordObject = function getWordObject(){
		// "Word.Document"; // MS Office 2002
		this.objectword = new ActiveXObject("Word.Application.8");
		this.objectword.Visible = false; // Hides Word
		this.objectword.WindowState = WORDSTATEMINIMIZE; // Keeps Word from popping up, if/when you open a file
		this.objectword.ChangeFileOpenDirectory(directory);	
	}
	
	this.wordDialogOpen = function wordDialogOpen(){ // Open Dialog
		var oWord = this.objectword.Dialogs(WORDDIALOGFILEOPENMULTIPLE);
		oWord.Name = filter;
		var iDlgBtnClicked = oWord.Show(); // Show dialog open file, return button clicked
		//iDlgBtnClicked = oWord.Display() // show dialog, Without opening the file
		switch(iDlgBtnClicked){
			case '-2' : sResult = "Close"; break;
			case '-1' : sResult = "OK"; break;
			case '0' : sResult = "Cancel"; break;
			default: sResult = "Button number: " + iDlgBtnClicked; break;
		}
		this.file = oWord.Name;
	}
	
	this.wordDialogSaveas = function wordDialogSaveas(){ // SaveAs dialog
		var oWord = this.objectword.Dialogs(WORDDIALOGFILESAVEAS); // Note: savesa won't work unless you have a file open 
		oWord.Name = "foobar";
		var iDlgBtnClicked = oWord.Display(); // show dialog, Without opening the file
		this.file = oWord.Name; // the filename selected		
	}
	
	this.wordClose = function wordClose(){
		if(this.objectword){
			this.objectword.Application.Quit(false); // Quits and save changes
			this.objectword = null;
		}
	}
})