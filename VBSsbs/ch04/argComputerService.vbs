'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <ArgComputerService.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 5/25/2006
'
' COMMENT: <Uses two un-named arguments to check on status of service>
'1. while this script works, it can lead to confusion. it is recommended To
'2. Use named arguments when you want two command line inputs. See the 
'3. NamedArgCS.vbs script for the recommended method of performing this action.
'==========================================================================
Option Explicit 
'on error resume next 'Turn back on once checked for errors
Dim computerName
Dim ServiceName
Dim wmiRoot
Dim wmiQuery
Dim objWMIService
Dim colServices
Dim oservice

computerName = WScript.Arguments(0)
serviceName = WScript.Arguments(1)
wmiRoot = "winmgmts:\\" & computerName & "\root\cimv2"
Set objWMIService = GetObject(wmiRoot)
wmiQuery = "Select state from Win32_Service" &_
	" where name = " & "'" & ServiceName & "'"
Set colServices = objWMIService.ExecQuery _
   (wmiQuery)
For Each oservice In colServices
	WScript.Echo (servicename) & " Is: "&_
	oservice.state & (" on: ") & computerName
Next 