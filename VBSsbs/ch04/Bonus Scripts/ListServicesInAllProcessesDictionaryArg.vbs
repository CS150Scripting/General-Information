'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <filename>
'
' AUTHOR: ed wilson , mred
' DATE  : 6/9/2003
'
' COMMENT: <comment>
'
'==========================================================================

strMachines = WScript.Arguments.Item(0)
aMachines = split(strMachines, ";")


set objIdDictionary = CreateObject("Scripting.Dictionary")

For Each machine in aMachines
strComputer = Machine
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
next