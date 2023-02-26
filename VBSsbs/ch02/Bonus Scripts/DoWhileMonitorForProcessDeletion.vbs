'==========================================================================
'
'AUTHOR: Ed WIlson , msft,  10/13/2004
' NAME: <MonitorProcessDeletion.vbs>
'
' COMMENT: <Uses WMI Eventing>
'1. uses win32_Process and eventing to check on deletion of processes
'
'==========================================================================

Option Explicit 
'On Error Resume Next
dim strComputer
dim wmiNS
dim wmiQuery
dim objWMIService
dim colItems
dim objItem
Dim objName 	'Monitored item
Dim objTGT 	'Monitored class
Dim i

strComputer = "."
objName = "'Notepad.exe'" 
objTGT = "'win32_Process'"
wmiNS = "\root\cimv2"
'i = 0

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



