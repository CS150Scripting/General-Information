'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <OneStepFurtherPT5.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 6/26/2006
'
' COMMENT: <THIS CODE DOES NOT RUN. It is complete to step 14 in the one step>
'					 <further for chapter four.>
' 1. working with arguments
' 2. Dictionaries
' 3. WMI win32_service
' 4. for Each
' 5. if then Else
' 6. vbtab
'==========================================================================
Option Explicit

Dim objIdDictionary
Dim objWMIService
Dim colServices
Dim objService
Dim colProcessIDs
Dim i 
Dim wmiRoot
Dim wmiQuery
Dim computer 

Dim txtFile
Dim objFSO
Dim objTextFile

TxtFile = "Servers.txt"
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (TxtFile, ForReading)

set objIdDictionary = CreateObject("Scripting.Dictionary")

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
