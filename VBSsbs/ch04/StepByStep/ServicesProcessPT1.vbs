'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <ServicesProcessPT.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 6/27/2006
'
' COMMENT: <This script should run, but is incomplete.>
'					 <covers TO STEP 8 in the step - by - step exercise>
'					 <for chapter four>
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
strComputer = "."
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
    Set colServices = objWMIService.ExecQuery _
        ("Select * from Win32_Service Where ProcessID = '" & _
            colProcessIDs(i) & "'")
    Wscript.Echo "Process ID: " & colProcessIDs(i)
    For Each objService in colServices
        Wscript.Echo VbTab & objService.DisplayName 
    Next
Next

