'==========================================================================
'
'
' NAME: <QueryAllProcessors.vbs>
'
' AUTHOR: Ed Wilson , MS
' DATE  : 3/16/2006
'
' COMMENT: <Starter for Using If … Then …Else to fix correct syntax Procedure>
'1. uses a function to translate the architecture into a string value
'==========================================================================

Option Explicit 
'On Error Resume Next
dim strComputer
dim wmiNS
dim wmiQuery
dim objWMIService
dim colItems
dim objItem
Dim intArch

strComputer = "."
wmiNS = "\root\cimv2"
wmiQuery = "Select Architecture from win32_Processor"
Set objWMIService = GetObject("winmgmts:\\" & strComputer & wmiNS)
Set colItems = objWMIService.ExecQuery(wmiQuery)

For Each objItem in colItems
			intArch = funArch(objItem.Architecture)
	WScript.Echo intArch 'debug
Next

' ***** Functions are Below *****

Function funArch(intIn)
	If intIn = 0 Then
	    funArch =  "It is an x86 cpu."
	ElseIf intIn = 1 Then
	    funArch =  "It is a MIPS cpu."
	ElseIf intIn = 2 Then
	    funArch =  "It is an Alpha cpu."
	ElseIf intIn = 3 Then
	    funArch =  "It is a PowerPC cpu."
	ElseIf intIn = 6 Then
	    funArch =  "It is an ia64 cpu."
	Else
	    funArch = "Can-not determine cpu type."
	End If
End Function