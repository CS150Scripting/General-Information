'==========================================================================
'
'
' NAME: <SetIPaddress.vbs>
'
' AUTHOR: Ed Wilson , MS
' DATE  : 3/18/2006
'
' COMMENT: <Illustrates using WMI to set a static WMI address>
'1.THIS SCRIPT will SET a SPECIFIC STATIC TCP/IP address on your machine.
'2.DO NOT RUN this script, unless this is what you desire to do.
'3.You can, of course, modify the strIPAddress and other values as appropriate
'==========================================================================

Option Explicit 
'On Error Resume Next
dim strComputer
dim wmiNS
dim wmiQuery
dim objWMIService
dim colItems
dim objItem 
Dim strIPaddress
Dim strSubnetMask
Dim strGateway
Dim strGatewayMetric
Dim errEnable, errGateways

strComputer = "."
strIPAddress = Array("192.168.1.1")
strSubnetMask = Array("255.255.255.0")
strGateway = Array("192.168.1.1")
strGatewayMetric = Array(1)

wmiNS = "\root\cimv2"
wmiQuery = "Select * from win32_NetworkAdapterConfiguration where IPEnabled=TRUE"
Set objWMIService = GetObject("winmgmts:\\" & strComputer & wmiNS)
Set colItems = objWMIService.ExecQuery(wmiQuery)

WScript.Echo "starting to change the IP address ... " & Now 
For Each objItem in colItems
	errEnable = objItem.EnableStatic(strIPAddress, strSubnetMask)
    errGateways = objItem.SetGateways(strGateway, strGatewayMetric)
    If errEnable = 0 Then
        WScript.Echo "The IP address has been changed. " & now
    ElseIf errEnable = 1 Then
        WScript.Echo "Reboot required to complete the IP change " & Now 
    Else 
    	WScript.Echo "Error " & errEnable & " occurred. " & Now
    End If
Next

