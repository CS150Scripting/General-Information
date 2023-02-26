'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <CPUType2AllProcessors.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 3/16/2006
'Version 2.0 Added function to Correct Grammer
'use count property on swbemObject
'fixed minor bugs. 
' COMMENT: <CPUType2AllProcessors>
'1.Uses execQuery to query information on all processors
'2.Counts the number of processors on machine.
'==========================================================================
Option Explicit
'On Error Resume Next
Dim strComputer
Dim wmiQuery
Dim wmiRoot
Dim objWMIService
Dim ObjProcessor, processor

strComputer = "."
wmiQuery = "Select * from win32_Processor"
wmiRoot = "winmgmts:\\" & strComputer & "\root\cimv2"
Set objWMIService = GetObject(wmiRoot)
Set objProcessor = objWMIService.execQuery (wmiQuery)
WScript.Echo "there" & funIS(ObjProcessor.count) & _
	"on this computer"
For Each processor In ObjProcessor
	If processor.Architecture = 0 Then
	    WScript.Echo "It is an x86 cpu."
	ElseIf processor.Architecture = 1 Then
	    WScript.Echo "It is a MIPS cpu."
	ElseIf processor.Architecture = 2 Then
	    WScript.Echo "It is an Alpha cpu."
	ElseIf processor.Architecture = 3 Then
	    WScript.Echo "It is a PowerPC cpu."
	ElseIf processor.Architecture = 6 Then
	    WScript.Echo "It is an ia64 cpu."
	Else
	    WScript.Echo "Cannot determine cpu type."
	    
	End If
Next 

' ***** functions are below *****

Function funIS(intIN)
	If intIN <2 Then
	funIS = " iS " & intIN & " processor "
	Else 
	funIS = " are " & intIN & " processors "
	End If
End Function