'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  3/12/2006
'
' NAME: <ReadTextFile.vbs>
'
' COMMENT: Key concepts are listed below:
'1.Uses FileSystemObject to open and read a text file
'2.Uses do Until to loop to the end of the file
'3. Uses instr to look for a pattern match. 
'==========================================================================
Option Explicit
'On Error Resume Next
Dim strError 
Dim objFSO			
Dim objFile			
Dim strLine 
Dim intResult 

CONST ForReading = 1
strError = "error"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\windows\setuplog.txt", ForReading)
strLine = objFile.ReadLine

Do Until objFile.AtEndofStream 
    strLine = objFile.ReadLine
    intResult = InStr(strLine, strError)
    If intResult <>0 Then
   		WScript.Echo(strLine)
   	End if
Loop
WScript.Echo("all done")
objFile.Close
