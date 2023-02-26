'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <OneStepFurtherPT7.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 6/26/2006
'
' COMMENT: <This code runs, but is incomplete. It is complete to>
'					 <step 18 in the one step further for chapter four.>
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
Dim j
Dim wmiRoot
Dim wmiQuery
Dim computer 
Dim txtFile
Dim objFSO
Dim objTextFile
Dim strNextLine
Dim arrServerList
TxtFile = "Servers.txt"
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (TxtFile, ForReading)
Do Until objTextFile.AtEndOfStream
	strNextLine = objTextFile.Readline
	arrServerList = Split(strNextLine , ",")
For Each computer in arrServerList
set objIdDictionary = CreateObject("Scripting.Dictionary")
WScript.Echo" Processes and services on " & (computer)
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
	For j = 0 to objIdDictionary.Count - 1
			wmiQuery = "Select * from Win32_Service Where ProcessID = '" & _
		            colProcessIDs(j) & "'"
		  Set colServices = objWMIService.ExecQuery _
		        (wmiQuery)
		   Wscript.Echo "Process ID: " & colProcessIDs(j)
		    For Each objService in colServices
		        Wscript.Echo VbTab & objService.DisplayName 
		 	Next
		 	objIdDictionary.RemoveAll
		Next
	Next 
Loop
Wscript.echo "all done"