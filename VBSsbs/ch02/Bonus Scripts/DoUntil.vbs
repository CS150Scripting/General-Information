'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  1/26/2004
'
' NAME: DoUntil.vbs
'
' COMMENT: Key concepts are listed below:
'1.Using Set to hold the FileSystemObject
'2.Using the OpenTextFile command
'3.Using Do ... Until to walk through the textStreamObject
'==========================================================================

Option Explicit
On Error Resume Next
Dim error1String 
Dim objFSO			
Dim objFile			
Dim strLine 
Dim SearchResult 

error1String = "error"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\windows\setuplog.txt", 1)
strLine = objFile.ReadLine

Do Until objFile.AtEndofStream 
    strLine = objFile.ReadLine
    SearchResult = InStr(strLine, error1String)
    If SearchResult <>0 Then
   WScript.Echo(strLine)
   	End if
Loop
WScript.Echo("all done")
objFile.Close
