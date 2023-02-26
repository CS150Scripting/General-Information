'==========================================================================
'
'AUTHOR: Ed WIlson , msft,  10/13/2004
' NAME: <DoLoopMonitorForProcessDeletion.vbs>
'
' COMMENT: <Uses WMI Eventing>
'1. uses win32_Process and eventing to check on deletion of processes
'2. Uses a do loop command to look for process deletion
'3. Defines a separate target instance name as well as class.
'4. Any win32_process property can be reported via targetInstance ....
'==========================================================================
Option Explicit 
'On Error Resume Next
dim strComputer 	'Computer to run the script upon.
dim wmiNS 		'The wmi namespace. Here it is the default namespace
dim wmiQuery 		'The wmi event query
dim objWMIService 	'SWbemServicesEx object
dim colItems 		'SWbemEventSource object
dim objItem 		'Individual item in the collection
Dim objName 		'Monitored item. Any Process. 
Dim objTGT 		'Monitored class. A win32_process. 


strComputer = "."
objName = "'Notepad.exe'" 'The single quotes inside the double quotes required
objTGT = "'win32_Process'"
wmiNS = "\root\cimv2"
wmiQuery = "SELECT * FROM __InstanceDeletionEvent WITHIN 10 WHERE " _
        & "TargetInstance ISA " & objTGT & " AND " _
            & "TargetInstance.Name=" & objName
            
Set objWMIService = GetObject("winmgmts:\\" & strComputer & wmiNS)
Set colItems = objWMIService.ExecNotificationQuery(wmiQuery)
Do 
    Set objItem = colItems.NextEvent
    Wscript.Echo "Name: " & objItem.TargetInstance.Name & " " & now
    Wscript.Echo "ProcessID: " & objItem.TargetInstance.ProcessId 
    WScript.echo "user mode time: " & objItem.TargetInstance.UserModeTime 
Loop
WScript.echo "all done"



