'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  6/2/2003
'
' NAME: <ArrayReadTxtFileUbound.vbs >
'
' COMMENT: Key concepts are listed below:
'1. illustrates use of Ubound as a means of calculating size of array
'2.
'3. 
'==========================================================================

Option Explicit    	     'Is used to force the scripter to declare variables
On Error Resume Next 'Is used to tell vbscript to go to the next line if it encounters an error
Dim objFSO
Dim objTextFile
Dim arrServiceList
Dim strNextLine
Dim i
Dim TxtFile
dim boundary
TxtFile = "ServersAndServices.txt"
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (TxtFile, ForReading)
Do Until objTextFile.AtEndOfStream
boundary = Ubound(arrServiceList)
wscript.echo "upper boundary = " & boundary
    strNextLine = objTextFile.Readline
    arrServiceList = Split(strNextLine , ",")
    Wscript.Echo "Server name: " & arrServiceList(0)
    For i = 1 to Ubound(arrServiceList)
        Wscript.Echo "Service: " & arrServiceList(i)
    Next
Loop
WScript.Echo("all done")
