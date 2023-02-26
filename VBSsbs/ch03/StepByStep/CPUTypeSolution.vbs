'==========================================================================
'
'
' NAME: <CPUTypeSolution.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 5/26/2003
'
' COMMENT: <comment>
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
cpu = "win32_Processor='CPU0'"
wmiRoot = "winmgmts:\\" & strComputer & "\root\cimv2"
Set objWMIService = GetObject(wmiRoot)
Set objProcessor = objWMIService.Get(cpu)
Select Case objProcessor.Architecture
Case 0 
    WScript.Echo "This is an x86 cpu."
Case 1 
     WScript.Echo "This is a MIPS cpu."
Case 2 
     WScript.Echo "This is an Alpha cpu."
Case 3 
     WScript.Echo "This is a PowerPC cpu."
Case 6 
     WScript.Echo "This is an ia64 cpu."
Case Else
     WScript.Echo "Can-not determine cpu type."
End Select 