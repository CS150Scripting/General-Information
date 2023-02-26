'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <ParseAPPlog.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 9/17/2003
'
' COMMENT: <This script demonstrates using two InStr commands to populate an array. 
' It then uses the split command to create a multi-dimensional array that is used
' to customize the message obtained from the appLog file.>
' Make sure the script has a path to the csv file.
' Also make sure you run this script under CSCRIPT.
'==========================================================================
Option Explicit
On Error Resume Next
Dim arrTxtArray()
Dim appLog 
Dim SearchString
Dim objTextFile 
Dim strNextLine 
Dim intSize 
Dim objFSO 
Dim i 
Dim ErrorString
Dim newArray
intSize = 0 
appLog = "applog.csv" 'Ensure in path
SearchString = ","
ErrorString = "1004"
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (appLog, ForReading)
Do until objTextFile.AtEndOfStream 
    strNextLine = objTextFile.Readline
    if InStr (strNextLine, SearchString)Then
    	If InStr (strNextLine, ErrorString) then
	    	ReDim Preserve arrTxtArray(intSize)
	    	arrTxtArray(intSize) = strNextLine
	    	intSize = intSize + 1
	    End if
    End if
Loop
	objTextFile.close
For i = LBound(arrTxtArray) To UBound(arrTxtArray)
	If InStr (arrTxtArray(i), ",") Then
	newArray = Split (arrTxtArray(i), ",")
		WScript.Echo "Date: " & newArray(0)
		WScript.Echo "Time: " & newArray(1) 
		WScript.Echo "Source: " & newArray(2)& " "& newArray(3)
		WScript.Echo "Server: " & newArray(7) 
		WScript.Echo "Message1: " & newArray(8)
		WScript.Echo "Message2: " & newArray(9)
		WScript.Echo "Message3: " & newArray(10)
		WScript.Echo " "
	End if
Next
WScript.Echo("all done")