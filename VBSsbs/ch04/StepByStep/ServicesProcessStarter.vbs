'==========================================================================
' NAME: <ServicesProcessStarter.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 6/26/2006
'
' COMMENT: <starter script for the following:>
'	   <NOT a complete script. Used in step by step for chapter four>
' 1. working with arguments
' 2. Dictionaries
' 3. WMI win32_service
' 4. for Each
' 5. if then Else
' 6. vbtab
'==========================================================================


set objIdDictionary = CreateObject("Scripting.Dictionary")
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colServices = objWMIService.ExecQuery _
    ("Select * from Win32_Service Where State <> 'Stopped'")
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
