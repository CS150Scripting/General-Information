'==========================================================================
'
' NAME: <NamedArgCS.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 5/26/2006
' ver.1.2
'
' COMMENT: <Uses named arguments to check the state of a service on a machine>
'1. This can work locally, or remotely. Either one. Just change the name of the
'2. /computer: argument. 
'3. This version of the script will fail if no argument is supplied. So you 
'4. Will want to add checking for arguments to the script to give a complete
'5. solution. 
'6. Usage: cscript namedArgCS.vbs /computer:localhost /service:spooler
'==========================================================================
Option Explicit 
'On Error Resume Next 'Turn back on once checked for errors
Dim computerName
Dim ServiceName
Dim wmiRoot
Dim wmiQuery
Dim objWMIService
Dim colServices
Dim oservice
Dim colNamedArguments

Set colNamedArguments = WScript.Arguments.Named
computerName = colNamedArguments("computer")
serviceName = colNamedArguments("service")
wmiRoot = "winmgmts:\\" & computerName & "\root\cimv2"
wmiQuery = "Select state from Win32_Service" &_
	" where name = " & "'" & ServiceName & "'"

Set objWMIService = GetObject(wmiRoot)
Set colServices = objWMIService.ExecQuery _
   (wmiQuery)
For Each oservice In colServices
	WScript.Echo servicename & " Is: " &_
	oservice.state & " on: " & computerName
Next 