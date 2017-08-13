' Copyright© 2003-2004. nOsliw Solutions. All rights reserved.
'**Start Encode**

Function WMIGetService(strServer,strNameSpace,strDomain,strUser,strPass,iImpersonation)
	Set objLocator = CreateObject("WbemScripting.SWbemLocator")
	Set objService = objLocator.ConnectServer(strServer,strNameSpace,strDomain & "\" & strUser,strPass)
	objService.Security_.ImpersonationLevel = iImpersonation
	WMIGetService = objService
End Function

Function GetQualifier(oQual)
	Dim sRetVal  ' as string (return value)
	Dim bIsAnArray  ' as boolean
	Dim oContext  ' as object
	Dim iElement  ' (array element)
	
		On Error Resume Next
	
		Select Case LCase(oQual.Name)  ' by-pass inappropriate types...
			Case "description"
				Exit Function
			Case "id"
				Exit Function
			Case "cimtype"
				Exit Function
			Case "in"
				Exit Function
		End Select
	
		sRetVal = ""  ' initialize result...
	
		bIsAnArray = (TypeName(oQual.Value)="Variant()")
	
		If Not bIsAnArray then
			sRetVal = sRetVal &  oQual.Name & "=" & oQual.Value & "<br>"
	    ' msgbox("[GetQualifier] retval: " & sRetVal)
	
		Else  ' IS an array...
			sRetVal = sRetVal & oQual.Name & " = Array:" & "<br>"
			Set oContext = CreateObject("WBEMScripting.SWBEMNamedValueSet")
			oContext.Add "n1", oQual.Value
	
			For iElement = LBound(oContext("n1")) to UBound(oContext("n1"))
				If oContext("n1")(iElement) <> "" then  ' if not empty...
					sRetVal = sRetVal & CStr(iElement) & "=" & oContext("n1")(iElement) & "<br>"
				End If  ' not empty
			Next  ' ielement
		End If  ' not an array
	
	  On Error GoTo 0  ' turn off error processing
	
	GetQualifier = sRetVal  ' set return result
End Function

' OBJECT

'Declare the Globals we will need for this sample
	Dim Locator
	Dim Service
	
'*****************************************************************
' This function converts the numerical CIMType code into
' a human-readable string
'*****************************************************************
Function TypeAsString (Property)
	TypeAsString = "uint32"
	
	select case Property.cimType
		Case 16
			TypeAsString = "sint8"
		Case 17
			TypeAsString = "uint8"
		Case 2
			TypeAsString = "sint16"
		Case 18
			TypeAsString = "uint16"
		Case 3
			TypeAsString = "uint32"
		Case 20
			TypeAsString = "sint64"
		Case 21
			TypeAsString = "uint64"
		Case 4
			TypeAsString = "real32"
		Case 5
			TypeAsString = "real64"
		Case 11
			TypeAsString = "boolean"
		Case 8
			TypeAsString = "string"
		Case 101
			TypeAsString = "datetime"
		Case 102
			TypeAsString = "ref"
			Set Qualifier = Property.Qualifiers_("cimtype")
			StrongRefArray = Split(Qualifier.Value,":")

			if (UBound(StrongRefArray) > 0) then
				TypeAsString = TypeAsString & " " & StrongRefArray(1)
			end if 
		Case 103
			TypeAsString = "char16"
		Case 13
			TypeAsString = "object"
			Set Qualifier = Property.Qualifiers_("cimtype")
			StrongObjArray = Split(Qualifier.Value,":")

			if (UBound(StrongObjArray) > 0) then
				TypeAsString = TypeAsString & " " & StrongObjArray(1)
			end if 
	end select

	if Property.isArray = true then
		TypeAsString = TypeAsString & " []"
	end if
	
End Function 

'*****************************************************************
' This handles the WMI ActiveX Class Navigator event that informs
' the container that a Class has been selected.  We display the
' Properties of the Class in an HTML table
'*****************************************************************
Sub ClassNav_EditExistingClass(selObj)
	on error resume next
	Dim classStr
	Dim Property
	Dim CIMClass
	
	ErrorMessage.innerText = ""
	
	' Clear the table (apart from the first title row)
	TableHeader.style.visibility = "hidden"
	while (ClassTable.rows.length > 1)
		ClassTable.deleteRow()
	wend
	
	' Extract the Class using the Scripting connection 
	' to the Namespace
	Set CIMClass = Service.Get (selObj)
	
	' Display the class name
	ClassName.innerText = CIMClass.Path_.Class
	
	' Build the property table by iterating through
	' the property collection of this class
	for each Property in CIMClass.Properties_
		Set row = ClassTable.insertRow
		row.insertCell().innerText = Property.Name
		row.insertCell().innerText = TypeAsString (Property)
		row.insertCell().innerText = Property.Origin
	next
		
	ClassTable.refresh
	TableHeader.style.visibility = "visible"
	
	if err <> 0 then
		ErrorMessage.innerText = "Error: " & Err.description & " - " & Err.source
	end if
End Sub

'*****************************************************************
' This handles the WMI ActiveX Class Navigator event that informs
' the container that a Namespace has been opened.  We make a 
' corresponding scripting connection to the same Namespace that we
' will use in the EditExistingClass event handler (above).
'*****************************************************************
Sub ClassNav_NotifyOpenNameSpace(theNameSpace)
	
	' The Class Navigator control will pass us an absolute Namespace
	' Path of the form \\<server>\<relative_namespace>. We must ensure
	' that we split off the server from the Namespace path before passing
	' to the Scripting API ConnectServer call.

	Dim Server
	Dim Namespace
	
	if Left (theNamespace, 2) = "\\" then
		' Extract Server name
		x = InStr (3, theNamespace, "\")
		
		if x <> 0 then
			Server = Mid(theNamespace, 3, x - 3)
			Namespace = Mid(theNamespace, x + 1)
		else
			Server = Mid(theNamespace, 3)
		end if 
		
	else
		Namespace = theNamespace
	end if 
	
	' Now make a scripting connection to WMI
	Set Service = Locator.ConnectServer (Server,Namespace)
	
End Sub

'*****************************************************************
' This handles the WMI ActiveX Class Navigator event that informs
' the container that a Namespace must been opened.  We delegate
' this request to the WMI ActiveX Login control.
'*****************************************************************
Sub ClassNav_GetIWbemServices(lpctstrNamespace, bUpdatePointer, lpsc, lppServices, lpbUserCancel)
	Login.GetIWbemServices lpctstrNamespace, bUpdatePointer, lpsc, lppServices, lpbUserCancel
End Sub


Function WMIDateStringToDate(dtmDate)
    WMIDateStringToDate = CDate(Mid(dtmDate,5,2) & "/" & Mid(dtmDate,7,2) & "/" & Left(dtmDate,4) & " " & Mid (dtmDate,9,2) & ":" & Mid(dtmDate,11,2) & ":" & Mid(dtmDate,13,2))
End Function


'******************
'* Description: Registry functions Provided by the WMI StdRegProv class
'*
'* Date:           28th August 2001
'* Written by:           Andrew Mayberry
'* Email:          nessross@bigpond.com
'*
'*
'* Warning:          Use this code snippet at your own risk. Although I have
'*               tested this code in my environment, there is no guarantee
'*               implied or otherwise. Test it thoroughly in a NON-PRODUCTION
'*               environment first.
'******************

Const HKCR=&H80000000 'HKEY_CLASSES_ROOT
Const HKCU=&H80000001 'HKEY_CURRENT_USER
Const HKLM=&H80000002 'HKEY_LOCAL_MACHINE
Const HKU=&H80000003 'HKEY_USERS
Const HKCC=&H80000005 'HKEY_CURRENT_CONFIG

Const REG_SZ=1
Const REG_EXPAND_SZ=2
Const REG_BINARY=3
Const REG_DWORD=4
Const REG_MULTI_SZ=7

Const KEY_QUERY_VALUE=&H1
Const KEY_SET_VALUE=&H2
Const KEY_CREATE_SUB_KEY=&H4
Const KEY_ENUMERATE_SUB_KEYS=&H8
Const KEY_NOTIFY=&H10
Const KEY_CREATE_LINK=&H20
Const DELETE=&H10000
Const READ_CONTROL=&H20000
Const WRITE_DAC=&H40000
Const WRITE_OWNER=&H80000


'*******************
'*     Function Name:          EnumKey
'*     Inputs:               Key, Subkey
'*     Example:           wscript.echo EnumKey(HKCU, "Software\Microsoft")
'*     Returns:          List of all sub keys
'*******************

Function EnumKey(Key, SubKey)
     Dim Ret()
     oReg.EnumKey Key,SubKey, sKeys
     ReDim Ret(UBound(sKeys))

     For Count = 0 to UBound(sKeys)
          Ret(Count) = sKeys(Count)
     Next

     EnumKey = Join(Ret,vbCrLf)
End Function


'*******************
'*     Function Name:          EnumValues
'*     Inputs:               Key, Subkey
'*     Example           wscript.echo EnumValues(HKCU, "Software\Microsoft")
'*     Returns:          List of all values in the format;
'*                         NAME,TYPE,VALUE
'*******************

Function EnumValues(Key, SubKey)
     Dim Ret()
     oReg.EnumValues Key,SubKey, sKeys, iKeyType 'fill the array
     ReDim Ret(UBound(sKeys))

     For Count = 0 to UBound(sKeys)
          Select Case iKeyType(Count)
               Case REG_SZ
                    oReg.GetStringValue Key,SubKey, sKeys(Count), sValue
                    Ret(Count) = sKeys(Count) & "," & "REG_SZ" & "," & sValue
               Case REG_EXPAND_SZ
                    oReg.GetExpandedStringValue Key,SubKey, sKeys(Count), sValue
                    Ret(Count) = sKeys(Count) & "," & "REG_EXPAND_SZ" & "," & sValue
               Case REG_BINARY
                    oReg.GetBinaryValue Key,SubKey, sKeys(Count), aValue
                    Ret(Count) = sKeys(Count) & "," & "REG_BINARY" & "," & Join(aValue,"")
               Case REG_DWORD
                    oReg.GetDWORDValue Key,SubKey, sKeys(Count), lValue
                    Ret(Count) = sKeys(Count) & "," & "REG_DWORD" & "," & lValue
               Case REG_MULTI_SZ
                    oReg.GetMultiStringValue Key,SubKey, sKeys(Count), sValue
                    Ret(Count) = sKeys(Count) & "," & "REG_MULTI_SZ" & "," & Join(sValue,"")
          End Select
     Next
     EnumValues = Join(Ret,vbCrLf)
End Function


'*******************
'*     Function Name:          CreateKey
'*     Inputs:               Key, Subkey
'*     Example           CreateKey(HKCU, "Software\KillerApp")
'*     Returns:          The error code where 0 is successful
'*******************

Function CreateKey(Key, SubKey)
     CreateKey = oReg.CreateKey(Key,SubKey)
End Function


'*******************
'*     Function Name:          CheckAccess
'*     Inputs:               Key, Subkey, AccessLevel
'*                         where AccessLevel is one of the following:
'*                              KEY_QUERY_VALUE
'*                              KEY_SET_VALUE
'*                              KEY_CREATE_SUB_KEY
'*                              KEY_ENUMERATE_SUB_KEYS
'*                              KEY_NOTIFY
'*                              KEY_CREATE_LINK
'*                              DELETE
'*                              READ_CONTROL
'*                              WRITE_DAC
'*                              WRITE_OWNER
'*
'*     Example           CheckAccess(HKCU, "Software\KillerApp", KEY_CREATE_SUB_KEY)
'*     Returns:          The error code where 0 is successful
'*******************

Function CheckAccess(Key, Subkey, AccessLevel)
     Dim Ret
     Dim bValue
     CheckAccess = False
     oReg.CheckAccess Key,SubKey,AccessLevel,bValue
     
     If bValue = -1 then
          CheckAccess = 0
     Else
          CheckAccess = -1
     End if
End Function


'*******************
'*     Function Name:          DeleteKey
'*     Inputs:               Key, Subkey
'*     Example           DeleteKey(HKCU, "Software\KillerApp")
'*     Returns:          The error code where 0 is successful
'*******************

Function DeleteKey(Key,SubKey)
     DeleteKey = oReg.DeleteKey(Key,SubKey)
End Function


'*******************
'*     Function Name:          DeleteValue
'*     Inputs:               Key, Subkey, ValueName
'*     Example           DeleteValue(HKCU, "Software\KillerApp","SomeSetting")
'*     Returns:          The error code where 0 is successful
'*******************

Function DeleteValue(Key, SubKey, ValueName)
     DeleteValue = oReg.DeleteValue(Key,SubKey,ValueName)
End Function

'*******************
'*     Function Name:          CreateValue
'*     Inputs:               Key,SubKey,ValueName,Value,KeyType
'*     Example           CreateValue(HKCU, "Software\KillerApp","SomeSetting","Hello World",REG_SZ)
'*     Returns:          The error code where 0 is successful
'*******************

Function CreateValue(Key,SubKey,ValueName,Value,KeyType)
     Select Case KeyType
          Case REG_SZ
               CreateValue = oReg.SetStringValue(Key,SubKey,ValueName,Value)
          Case REG_EXPAND_SZ
               CreateValue = oReg.SetExpandedStringValue(Key,SubKey,ValueName,Value)
          Case REG_BINARY
               CreateValue = oReg.SetBinaryValue(Key,SubKey,ValueName,Value)
          Case REG_DWORD
               CreateValue = oReg.SetDWORDValue(Key,SubKey,ValueName,Value)
          Case REG_MULTI_SZ
               CreateValue = oReg.SetMultiStringValue(Key,SubKey,ValueName,Value)
     End Select
End Function










