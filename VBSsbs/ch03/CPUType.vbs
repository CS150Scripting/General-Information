'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <CPUType.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 5/26/2003
'
' COMMENT: <the following are concepts in this script>
' If Then ElseIf
' win32_Processor name Space
' use of Option Explicit
'
'==========================================================================
Option Explicit
'On Error Resume Next
Dim strComputer
Dim cpu
Dim wmiRoot
Dim objWMIService
Dim ObjProcessor

strComputer = "."
cpu = "win32_Processor.deviceID='CPU0'"
wmiRoot = "winmgmts:\\" & strComputer & "\root\cimv2"

Set objWMIService = GetObject(wmiRoot)
Set objProcessor = objWMIService.Get(cpu)
	
If objProcessor.Architecture = 0 Then
  WScript.Echo "This is an x86 cpu."
ElseIf objProcessor.Architecture = 1 Then
  WScript.Echo "This is a MIPS cpu."
ElseIf objProcessor.Architecture = 2 Then
  WScript.Echo "This is an Alpha cpu."
ElseIf objProcessor.Architecture = 3 Then
  WScript.Echo "This is a PowerPC cpu."
ElseIf objProcessor.Architecture = 6 Then
  WScript.Echo "This is an ia64 cpu."
Else
  WScript.Echo "Can-not determine cpu type."
End If
