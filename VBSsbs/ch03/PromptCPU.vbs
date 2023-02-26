'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  3/15/2006
'
' NAME: PromptCPU.vbs
'
' COMMENT: Key concepts are listed below:
'1.Use of the MsgBox
'2.Use of if ... then ... elseif ... else
'3.combine with the CPUtype script to ask if wish to run the
'4.script.
'==========================================================================
Option Explicit
On Error Resume Next
Dim strPrompt
Dim strTitle
Dim intBTN
Dim intRTN

strPrompt = "Do you want to run the script?"
strTitle = "MsgBox DEMO"
intBTN = 3 '4 is yes/no 3 yes/no/cancel

intRTN = MsgBox(strprompt,intBTN,strTitle)


If intRTN = vbYes Then
	WScript.Echo "yes was pressed"
	subCPU
ElseIf intRTN = vbNo Then
	WScript.Echo "no was pressed"
	WScript.Quit
ElseIf intRTN = vbCancel Then
	WScript.Echo "cancel was pressed"
	WScript.quit
Else
	WScript.Echo intRTN & " was pressed"
	WScript.quit
End If


Sub subCPU 		'subroutine below is the CPUtype.vbs script
Dim strComputer 	'the name of the computer to connect to
Dim cpu				'the specific cpu to connect to
Dim wmiRoot			'the name of the wmi namespace 
Dim objWMIService 	'connection into WMI using moniker
Dim ObjProcessor	'contains swbemObject

strComputer = "."
cpu = "win32_Processor='CPU0'"
wmiRoot = "winmgmts:\\" & strComputer & "\root\cimv2"
Set objWMIService = GetObject(wmiRoot)
Set objProcessor = objWMIService.Get(cpu)
WScript.Echo(ObjProcessor.architecture)
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
    WScript.Echo "Cannot determine cpu type."
End If
End Sub