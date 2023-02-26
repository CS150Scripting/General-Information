'==========================================================================
'
' NAME: <ServicesProcessPT4.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 6/27/2006
'
'
' COMMENT: <This script should run, but it is not complete.>
'					 <this script covers to step 15 in the step -step>
'					 <for chapter four>
'					<NOTE: this script NOW requires a command line argument. If Not>
'					<supplied, then you will get a subscript out of range message>
' 1. working with arguments
' 2. Dictionaries
' 3. WMI win32_service
' 4. for Each
' 5. if then Else
' 6. vbtab
'==========================================================================
Option Explicit

Dim objIdDictionary
Dim strComputer
Dim objWMIService
Dim colServices
Dim objService
Dim colProcessIDs
Dim i 
Dim wmiRoot
Dim wmiQuery

set objIdDictionary = CreateObject("Scripting.Dictionary")
strComputer = WScript.Arguments(0)
wmiRoot = "winmgmts:\\" & strComputer & "\root\cimv2"
Set objWMIService = GetObject(wmiRoot)
wmiQuery = "Select * from Win32_Service Where State <> 'Stopped'"
Set colServices = objWMIService.ExecQuery _
    (wmiQuery)
For Each objService in colServices
    If objIdDictionary.Exists(objService.ProcessID) Then
    Else
        objIdDictionary.Add objService.ProcessID, objService.ProcessID
    End If
Next
colProcessIDs = objIdDictionary.Items
For i = 0 to objIdDictionary.Count - 1
wmiQuery = "Select * from Win32_Service Where ProcessID = '" & _
            colProcessIDs(i) & "'"
    Set colServices = objWMIService.ExecQuery _
        (wmiQuery)
    Wscript.Echo "Process ID: " & colProcessIDs(i)
    For Each objService in colServices
        Wscript.Echo VbTab & objService.DisplayName 
    Next
Next

Wscript.echo "all done"