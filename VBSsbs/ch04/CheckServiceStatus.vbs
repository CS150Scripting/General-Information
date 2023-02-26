'==========================================================================
'
'
' NAME: <CheckServiceStatus.vbs>
'
' AUTHOR: Ed Wilson , MS
' DATE  : 3/26/2006
'
' COMMENT: <Returns the startup mode and status of a single WMI service>
'1.Uses win32_service Class
'2.Uses a variable to hold name of service to examine
'3.Can target other computers by changing value of strComputer
'4.Can check other services by changing value of serviceName
'5.Uses funFIX function to add single quotes to name of service. 
'==========================================================================

Option Explicit 
'On Error Resume Next
dim strComputer 	'name of computer to connect to
Dim serviceName 	'Name of service to query
dim wmiROOT				'path into WMI
dim wmiQuery			'The WQL Query
dim objWMIService	'Connection into WMI
Dim objItem				'single item returned by Get Method


strComputer = "."
ServiceName = "spooler"
wmiRoot = "winmgmts:\\" & strComputer & "\root\cimv2" 
wmiQuery = "win32_service.name=" & funFIX(serviceName)

Set objWMIService = GetObject(wmiRoot)
Set objItem = objWMIService.get(wmiQuery)
		WScript.Echo vbTab & (servicename) & " Is: " _
		& objItem.state & " startup mode is: " & objItem.StartMode
	
' **** function is below *****
	
Function funFIX(strIN)
	funFIX = "'" & strIN & "'"
End Function