'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  3/26/2006
'
' NAME: <ArrayReadTxtFile.vbs>
'ver.1.2
' COMMENT: Key concepts are listed below:
'1.Uses filesystem object to read a text file
'2.Uses split function to create an array 
'3.Uses echo to print out items stored in Array
'4.Uses ubound function to find upper element of array
'==========================================================================

Option Explicit    	     ' is used to force the scripter to declare variables
'On Error Resume Next ' is used to tell vbscript to go to the next line if it encounters an Error
Dim objFSO
Dim objTextFile
Dim arrServiceList
Dim strNextLine
Dim i
Dim TxtFile

TxtFile = "ServersAndServices.txt"
Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (TxtFile, ForReading)
Do Until objTextFile.AtEndOfStream
    strNextLine = objTextFile.Readline
    arrServiceList = Split(strNextLine , ",")
    Wscript.Echo "Server name: " & arrServiceList(0)
    For i = 1 to Ubound(arrServiceList)
        Wscript.Echo vbTab & "Service: " & arrServiceList(i)
    Next
Loop
WScript.Echo("all done")
