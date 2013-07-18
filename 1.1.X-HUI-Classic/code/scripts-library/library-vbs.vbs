' Copyright© 2003-2004. nOsliw Solutions. All rights reserved.
'**Start Encode**

' Returning IP Configuration Data
' WMI script that returns configuration data similar to that returned by IpConfig.

Function vbs_sys_iconfig(strComputer)
	On Error Resume Next
	
	Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& strComputer & "\root\cimv2")
	Set colAdapters = objService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True",,48)
	n = 1
	strHtml = vbNewLine & " NETWORK CONFIGURATION ON: " & UCase(strComputer) & vbNewLine & vbNewLine
	
	For Each objAdapter in colAdapters
		
		strHtml = strHtml & "  ==================================" & vbNewLine
		strHtml = strHtml & "  Network Adapter " & n & vbNewLine
		strHtml = strHtml & "  ==================================" & vbNewLine
		strHtml = strHtml & "  Description:" & vbTab & vbTab & objAdapter.Description & vbNewLine
		strHtml = strHtml & "  Physical (MAC) Address:" & objAdapter.MACAddress & vbNewLine
		strHtml = strHtml & "  Host Name:" & vbTab & vbTab & objAdapter.DNSHostName & vbNewLine
		
		If Not IsNull(objAdapter.IPAddress) Then
		   For i = 0 To UBound(objAdapter.IPAddress)
		      strHtml = strHtml & "  IP address:" & vbTab & vbTab & objAdapter.IPAddress(i) & vbNewLine
		   Next
		End If
		
		If Not IsNull(objAdapter.IPSubnet) Then
		   For i = 0 To UBound(objAdapter.IPSubnet)
		      strHtml = strHtml & "  Subnet Mask:" & vbTab & objAdapter.IPSubnet(i) & vbNewLine
		   Next
		End If
		
		If Not IsNull(objAdapter.DefaultIPGateway) Then
		   For i = 0 To UBound(objAdapter.DefaultIPGateway)
		      strHtml = strHtml & "  Default Gateway:" & vbTab & objAdapter.DefaultIPGateway(i) & vbNewLine
		   Next
		End If
		
		strHtml = strHtml & vbNewLine
		strHtml = strHtml & " [DNS]" & vbNewLine
		strHtml = strHtml & "  DNS Servers in Search Order:" & vbNewLine
		
		If Not IsNull(objAdapter.DNSServerSearchOrder) Then
		   For i = 0 To UBound(objAdapter.DNSServerSearchOrder)
		      strHtml = strHtml & "  DNS Server-" & (i+1) & ":" & vbTab & objAdapter.DNSServerSearchOrder(i) & vbNewLine
		   Next
		End If
		
		strHtml = strHtml & "  DNS Domain:" & vbTab & objAdapter.DNSDomain & vbNewLine
		
		If Not IsNull(objAdapter.DNSDomainSuffixSearchOrder) Then
		   For i = 0 To UBound(objAdapter.DNSDomainSuffixSearchOrder)
		      strHtml = strHtml & "  DNS Suffix Search List:" & vbTab & objAdapter.DNSDomainSuffixSearchOrder(i) & vbNewLine
		   Next
		End If
		
		strHtml = strHtml & vbNewLine
		strHtml = strHtml & " [DHCP]" & vbNewLine
		strHtml = strHtml & "  DHCP Enabled:" & vbTab & objAdapter.DHCPEnabled & vbNewLine
		strHtml = strHtml & "  DHCP Server:" & vbTab & objAdapter.DHCPServer & vbNewLine
		
		If Not IsNull(objAdapter.DHCPLeaseObtained) Then
		   utcLeaseObtained = objAdapter.DHCPLeaseObtained
		   strLeaseObtained = WMIDateStringToDate(utcLeaseObtained)
		Else
		   strLeaseObtained = ""
		End If
		strHtml = strHtml & "  DHCP Lease Obtained:" & vbTab & strLeaseObtained & vbNewLine
		
		If Not IsNull(objAdapter.DHCPLeaseExpires) Then
		   utcLeaseExpires = objAdapter.DHCPLeaseExpires
		   strLeaseExpires = WMIDateStringToDate(utcLeaseExpires)
		Else
		   strLeaseExpires = ""
		End If
		strHtml = strHtml & "  DHCP Lease Expires:" & vbTab & strLeaseExpires & vbNewLine
		
		strHtml = strHtml & vbNewLine
		strHtml = strHtml & " [WINS]" & vbNewLine
		strHtml = strHtml & "  Primary WINS Server:" & vbTab & objAdapter.WINSPrimaryServer & vbNewLine
		strHtml = strHtml & "  Secondary WINS Server:" & vbTab & objAdapter.WINSSecondaryServer & vbNewLine
		strHtml = strHtml & vbNewLine & vbNewLine
		
		n = n & 1	
	Next
	
	vbs_sys_iconfig = vbs_err_check(strHtml)
	
End Function

Function WMIDateStringToDate(utcDate)
   WMIDateStringToDate = CDate(Mid(utcDate, 5, 2)  & "/" & _
                               Mid(utcDate, 7, 2)  & "/" & _
                               Left(utcDate, 4)    & " " & _
                               Mid (utcDate, 9, 2) & ":" & _
                               Mid(utcDate, 11, 2) & ":" & _
                               Mid(utcDate, 13, 2))
End Function


'*************************************************************************************
' Report Hardware Function
' Input computer name, output comma delimited hardware list
'*************************************************************************************
Function vbs_sys_hardware(computerName)

On Error Resume Next

Dim cpuCount, videoCount, partitionCount, soundCount, nicCount, memCount
Dim strCpuName, strMemory, strMemDevices, strVideo, strDisks, strSound, strNic, strNicType, nicMAC
Dim i
	js.status 1,1
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate,(Security)}!\\" & computerName & "\root\cimv2")
	If err.Number = 0 Then
		'processor name(s) from WMI
		Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor",,48)
		cpuCount = 0
		For Each objItem In colItems
			cpuCount = cpuCount + 1
			strCpuName = objItem.Name
		Next
		If cpuCount > 1 Then
			strCpuName = cpuCount & "x " & strCpuName & ","
		End If
		strCpuName = "CPU(s): " & replace(strCpuName, "  ", "") & ","	
		js.status 1,strCpuName
		'system memory size from WMI
		Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
		For Each objComputer In colSettings
			strMemory = "Total System Memory Size: " & mySizeMemory(objComputer.TotalPhysicalMemory) & ","
		Next
		'indiviual memory devices
		Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory",,48)
		memCount = 0
		For Each objItem In colItems
			ReDim Preserve aryMem(memCount)
			aryMem(memCount) = objItem.DeviceLocator & ": " & mySizeMemory(objItem.Capacity)
			memCount = memCount + 1
		Next
		If memCount = 1 Then
			strMemDevices = aryMem(0) & ","
		Else
			For i = 0 To memCount - 1
				strMemDevices = strMemDevices & aryMem(i) & ","
			Next
		End If
		'video card information from WMI
		Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoController",,48)
		gpuCount = 0
		strVideo = ""
		For Each objItem In colItems
			ReDim Preserve aryVideo(gpuCount)
			aryVideo(gpuCount) = objItem.VideoProcessor & " " & mySizeMemory(objItem.AdapterRAM)
			gpuCount = gpuCount + 1
		Next
		If gpuCount = 1 Then
			strVideo = "Graphics Adapter: " & aryVideo(0) & ","
		Else
			For i = 0 To gpuCount - 1
				strVideo = strVideo & "Graphics Adapter #" & i + 1 & ": " & aryVideo(i) & ","
			Next
		End If
		js.status 1,strVideo
		'disk information
		Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskPartition",,48)
		partitionCount = 0
		strDisk = ""
		For Each objItem In colItems
			ReDim Preserve aryDisks(partitionCount) 
			aryDisks(partitionCount) = replace(objItem.Name, ",", "") & " -> " & mySizeDisk(objItem.Size)
			partitionCount = partitionCount + 1
		Next
		If partitionCount = 1 Then
			strDisks = "Disk Volume: " & aryDisks(0) & ","
		Else
			For i = 0 To partitionCount - 1
				strDisks = strDisks & "Disk Volume #" & i + 1 & ": " & aryDisks(i) & ","
			Next
		End If	

		'network devices
		Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter",,48)
		nicCount = 0
		strNic = ""
		js.status 1,strDisks
		For Each objItem In colItems
			err.Clear
			nicMAC = objItem.MACAddress
			If err = 0 Then
				'The wan pptp adapter seems to have a default address of this everywhere
				If nicMAC <> "50:50:54:50:30:30" Then
					If objItem.ProductName <> "Packet Scheduler Miniport" Then
						If objItem.ProductName <> "WAN Miniport (PPPOE)" Then
							ReDim Preserve aryNic(nicCount)
							aryNic(nicCount) = objItem.ProductName
							nicCount = nicCount + 1
						End If
					End If
				End If
			End If
		Next
		
		If nicCount > 0 Then
			If nicCount = 1 Then
				strNic = "NIC: " & aryNic(0) & ","
			Else
				For i = 0 To nicCount - 1
					strNic = strNic & "NIC#" & i + 1 & ": " & aryNic(i) & ","
				Next
			End If
		End If
		'sound devices
		Set colItems = objWMIService.ExecQuery("Select * from Win32_SoundDevice")
		soundCount = 0
		strSound = ""
		For Each objItem In colItems
			ReDim Preserve arySound(soundCount)
			arySound(soundCount) = objItem.Description
			soundCount = soundCount + 1
		Next
		If soundCount = 1 Then
			strSound = "Sound Device: " & arySound(0) & ","
		Else
			For i = 0 To soundCount - 1
				strSound = strSound & "Sound Device #" & i + 1 & ": " & arySound(i) & ","
			Next
		End If
		vbs_sys_hardware = strCpuName & strMemory & strMemDevices & strVideo & strDisks & strNic & strSound
		'this kills off that last comma
		vbs_sys_hardware = Left(vbs_sys_hardware, Len(vbs_sys_hardware) - 1)
		If Err <> 0 Then
			js.status 1,Err.Description, "0x" & Hex(Err.Number)
		End If
	Else
		vbs_sys_hardware = "The WMI service was unavailable on the system " & UCASE(computerName) & _
		". No hardware information will be reported for this host. Possible reasons for this may include; " & _
		"the host is running Windows 98 or Windows NT and does not have WMI installed or the host may be offline " & _
		"or the host does not have the client For Microsoft Networks enabled On the NIC interface you are accessing " & _
		"WMI over."
		err.Clear
	End If
err.Clear	
End Function

'*************************************************************************************
' QuickSort Sub Routine
' the venerable quicksort, sorry no compensating for pre-sorted
'*************************************************************************************
Sub qSort(list, first, last)

	Dim midList, pivotValue, temp, up, down

	up = first
	down = last
	midList = int((first + last) / 2)
	pivotValue = list(midList)
	Do
		Do While (lcase(list(up)) < lcase(pivotValue))
			up = up + 1
		Loop
		Do While (lcase(list(down)) > lcase(pivotValue))
			down = down - 1
		Loop
		If (up <= down) Then
    		'swap up and down
			temp = list(up)
			list(up) = list(down)
			list(down) = temp
			up = up + 1
			down = down - 1
		End If
	Loop While up < down
	If (down > first) Then
		qSort list, first, down
	End If
	If (up < last) Then
		qSort list, up, last
	End If

End Sub

'*************************************************************************
' Ping
' Function to secretly call up the command shell and ping a server
' I wish this was in all versions of WMI but alas on XP and 2K3 Server
' have it built in so far, so for compatibilities sake I'll use this
'*************************************************************************
Function ping(hostName)
	Set wshShell = CreateObject("WScript.Shell")
	ping = Not CBool(wshShell.run("ping -n 1 " & hostName,0,True))	
End Function

'*************************************************************************
' mySizeDisk: this function just formats the output the way I like       
' Input integer in Bytes, outputs easily human readable format
'*************************************************************************
Function mySizeDisk(integerSize)

If integerSize >= 1000000000 Then
	mySizeDisk = round(integerSize / 1000000000, 2)
	mySizeDisk = mySizeDisk & " GB"
ElseIf ((1000000000 > integerSize) And (integerSize >= 1000000)) Then
	mySizeDisk = round(integerSize / 1000000, 2)
	mySizeDisk = mySizeDisk & " MB"
ElseIf ((1000000 > integerSize) And (integerSize >= 1000)) Then
	mySizeDisk = round(integerSize / 1000, 2)
	mySizeDisk = mySizeDisk & " KB"
Else
	mySizeDisk = integerSize & " Bytes"
End If

End Function

'*************************************************************************
' mySizeMemory:this function just formats the output the way I like                    
' Input integer in Bytes, outputs easily human readable format
'*************************************************************************
Function mySizeMemory(integerSize)

If integerSize >= 1000000000 Then
	mySizeMemory = Round(integerSize / 1047000000, 1)
	mySizeMemory = mySizeMemory & " GB"
ElseIf ((1000000000 > integerSize) And (integerSize >= 1000000)) Then
	mySizeMemory = Int(integerSize / 1047000)
	mySizeMemory = mySizeMemory & " MB"
ElseIf ((1000000 > integerSize) And (integerSize >= 1000)) Then
	mySizeMemory = Int(integerSize / 1024)
	mySizeMemory = mySizeMemory & " KB"
Else
	mySizeMemory = integerSize & " Bytes"
End If

End Function
'~~[/script]~~

' ERROR FUNCTIONS
'
'
'

Function vbs_err_check(bolReturn)
	If (Err.Number <> 0) Then
		js_err_show 2,Err,FALSE,FALSE,TRUE
		Err.Clear
		bolReturn = FALSE
	End If
	vbs_err_check = bolReturn
End Function

Sub vbs_err_show()
	if Err = 0 Then
		WScript.Echo "Got a class"
	Else
		WScript.Echo ""
		WScript.Echo "Err Information:"
		WScript.Echo ""
		WScript.Echo "  Source:", Err.Source
		WScript.Echo "  Description:", Err.Description
		WScript.Echo "  Number", "0x" & Hex(Err.Number)
	
		'Create the last error object
		set t_Object = CreateObject("WbemScripting.SWbemLastError")
		WScript.Echo ""
		WScript.Echo "WMI Last Error Information:"
		WScript.Echo ""
		WScript.Echo " Operation:", t_Object.Operation
		WScript.Echo " Provider:", t_Object.ProviderName
	
		strDescr = t_Object.Description
		strPInfo = t_Object.ParameterInfo
		strCode = t_Object.StatusCode
	
		if (strDescr <> nothing) Then
			WScript.Echo " Description:", strDescr		
		end if
	
		if (strPInfo <> nothing) Then
			WScript.Echo " Parameter Info:", strPInfo		
		end if
	
		if (strCode <> nothing) Then
			WScript.Echo " Status:", strCode		
		end if
	
		WScript.Echo ""
		Err.Clear
		set t_Object2 = CreateObject("WbemScripting.SWbemLastError")
		if Err = 0 Then
		   WScript.Echo "Got the error object again - this shouldn't have happened!"	
		Else
		   Err.Clear
		   WScript.Echo "Couldn't get last error again - as expected"
		End if
	End If

End Sub

