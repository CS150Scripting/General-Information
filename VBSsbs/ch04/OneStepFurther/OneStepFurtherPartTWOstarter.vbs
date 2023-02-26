'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <OneStepFurtherPartTWOstarter.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 6/26/2006
'
' COMMENT: <starter script part two of one step further in chapter four>
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
Dim colComputers
Dim computer 

If WScript.Arguments.count = 0 Then
WScript.Echo("You must enter a computer name")
else
set objIdDictionary = CreateObject("Scripting.Dictionary")
strComputer = WScript.Arguments(0)
colComputers = split(strComputer, ",")

For Each computer in colComputers
	wmiRoot = "winmgmts:\\" & Computer & "\root\cimv2"
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
Next 

Wscript.echo "all done"
End If 