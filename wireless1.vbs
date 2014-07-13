'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2009
'
' NAME: 
'
' AUTHOR: Mohammed Alyafae , 
' DATE  : 9/12/2011
'
' COMMENT: this script disable wireless connection if the LAN is Connected
'			and vice versa
'==========================================================================
On Error Resume Next
 Dim strComputer
 Dim objWMIService
 Dim colLAN
 Dim objWifi,objLAN
 Dim state
 Dim wireStatus
 Dim wifiStatus

 state=""
 wireStatus=""
 wifiStatus=""

 Do While True
  
 strComputer = "." 
 Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
 Set colLAN = objWMIService.ExecQuery("Select * From Win32_NetworkAdapter Where NetConnectionID like 'Local Area Connection' and PhysicalAdapter='True'" )
 Set colWiFi=objWMIService.ExecQuery ("Select * From Win32_NetworkAdapter Where NetConnectionID =" & "'" &GetWirlessName & "'" & "and PhysicalAdapter='True' ")
 
 For Each objWifi In colWiFi
	 If objWifi.Netconnectionstatus=2 Then
	 wifiStatus=True
	 Else
	 wifiStatus=False
	 End If
 Next
 
 For Each objLAN in colLAN
 
 If objLAN.Netconnectionstatus=2 Then
	 wireStatus=True
	 state=False ' this is very importnat variable to determine when to enable or disbale wireless connection   
	 Else
	 wireStatus=False
	 End If
 Next
 

If True Then
	If state <>  False Then
		 If wifiStatus = False Then
		 EnableWireless GetWirlessName
		 End If
	Else 
		If wifiStatus = True Then
		DisableWireless GetWirlessName
		End If
	End If
	 
 End If 
 

state=""
wireStatus=""
wifiStatus=""

WScript.Sleep  60000

Loop


' Function to get wireless adapter name from the registery
Function GetWirlessName 

Dim strKeyPath
Dim strComputer
Dim objReg
Dim arrSubKeys
Dim SubKey
Dim strValueName
Dim dwValue
Dim strValue
Const HKLM=&H80000002


strKeyPath="SYSTEM\CurrentControlSet\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}"
strComputer="."

Set objReg=GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
objReg.Enumkey HKLM ,strKeyPath,arrSubKeys

For Each SubKey In arrSubKeys
	strValueName="MediaSubType"
	objReg.GetDWORDValue HKLM,strKeyPath & "\" & subkey & "\" & "Connection" ,strValueName,dwValue
	If dwValue=2 Then
		strValueName = "Name"
		objReg.GetStringValue HKLM,strKeyPath & "\" & subkey & "\" & "Connection" ,strValueName,strValue
		Exit For
	End If
Next

GetWirlessName=strValue

End Function



' Subroutine to disable wireless connection 
Sub DisableWireless (strNetConn)

Dim oConnections
dim objShell
Dim objConnections,objConn
Dim strDisable
Dim objNetwork
Dim objDisable
Dim objVerb

Const NETWORK_CONNECTIONS = &H31&

strDisable = "Disa&ble"

Set objShell = CreateObject("Shell.Application")
Set objConnections = objShell.Namespace(NETWORK_CONNECTIONS)

For Each objConn In objConnections.Items
	If objConn.Name = strNetConn Then
		Set objNetwork = objConn
		Exit For
	End If
Next
Set objDisable = Nothing

For Each objVerb in objNetwork.verbs
	If objVerb.name = strDisable Then 
		Set objDisable = objVerb 
		Exit For
	End If

Next
objDisable.DoIt
WScript.Sleep 1000 
End Sub

'Function to enable wireless connection , you can combone these two subtoutines into one
' but I prefer to seperate them just for simplicity
Sub EnableWireless (strNetConn)
Dim oConnections
dim objShell
Dim objConnections,objConn
Dim strEnable
Dim objNetwork
Dim objEnable
Dim objVerb

Const NETWORK_CONNECTIONS = &H31&

strEnable = "En&able"

Set objShell = CreateObject("Shell.Application")
Set objConnections = objShell.Namespace(NETWORK_CONNECTIONS)


For Each objConn In objConnections.Items
If objConn.Name = strNetConn Then
	Set objNetwork = objConn
	Exit For
End If
Next
Set objEnable = Nothing

' Enable NIC 
For Each objVerb in objNetwork.verbs
	If objVerb.name = strEnable Then 
		Set objEnable = objVerb 
		Exit For
	End If

Next

objEnable.DoIt
WScript.Sleep 1000 
End Sub
