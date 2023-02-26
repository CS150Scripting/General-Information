'==========================================================================
'
' NAME: <CheckNamedArgCS.vbs>
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
'On Error Resume Next
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
ServiceName = colNamedArguments("service")

subCheckArgs

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

Sub subCheckArgs
If colNamedArguments.Count < 2 Then
	If colNamedArguments.exists("computer") Then
		ServiceName = "spooler"
		WScript.Echo "using default service: spooler"
	Else If colNamedArguments.Exists("Service") Then
		computerName = "localHost"
		WScript.Echo "using default computer: localhost"
	Else
		WScript.Echo "you must supply two arguments" _
			& " to this script." & VbCrLf & "Try this: " _
			& "cscript checkNamedArgCS.vbs /computer:" _
			& "localhost /service:spooler"
		WScript.Quit
	End If
	End If
End If
End Sub