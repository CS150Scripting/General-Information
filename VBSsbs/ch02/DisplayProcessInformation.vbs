'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  3/11/2006
'
' NAME: DisplayProcessInformation.vbs
'
' COMMENT: Key concepts are listed below:
'1.Uses a constant for one hour setting for sleep
'2.Uses For ... Next to make an 8 pass loop
'3.Uses win32_Process wmi class to provide process information
'4.Uses vbNewLine to make a blank line in script
'5.Uses the space function to "tab" over 9 spaces
'6.Uses the count property of SWbemObjectSet object.
'==========================================================================
Option Explicit
'On Error Resume Next
Dim objWMIService 	'an SWbemObjectSet object
Dim objItem 		'an individual process 
Dim i				'a counter variable

Const MAX_LOOPS = 8, ONE_HOUR = 3600000

For i = 1 To MAX_LOOPS
Set objWMIService = GetObject("winmgmts:").ExecQuery _
    ("SELECT * FROM Win32_Process where processID <> 0")
	wscript.Echo "There are " & objWMIService.count &_
 	" processes running " & Now
 	For Each objItem In objWMIService
        WScript.Echo "Process: " & objItem.Name
        WScript.Echo Space(9) & objItem.commandline
        WScript.Echo "Process ID: " & objItem.ProcessID
        WScript.Echo "Thread Count: " & objItem.ThreadCount
        WScript.Echo "Page File Size: " & objItem.PageFileUsage
        WScript.Echo "Page Faults: " & objItem.PageFaults
        WScript.Echo "Working Set Size: " & objItem.WorkingSetSize
        wscript.Echo vbNewLine
    Next
    WScript.Echo "******PASS COMPLETE**********"
    WScript.Sleep ONE_HOUR
Next
