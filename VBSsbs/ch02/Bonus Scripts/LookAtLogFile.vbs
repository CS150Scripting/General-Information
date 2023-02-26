'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  1/26/2004
' 
' NAME: <LookAtLogFile.vbs>
' 
' COMMENT: Key concepts are listed below:
'1.Using FileSystemObject
'2.Using OpenTextFile command
'3.Using Do Until
'4.Using Instr 
'=========================================================================

Option Explicit
'On Error Resume Next
Dim error1String 	'Holds error string value to search for
Dim objFSO		'Creates an instance of file system object		
Dim objFile 		'Opens the text file 			
Dim strLine 		'Holds the value of one line of text
Dim SearchResult 	'What comes back from searching a line of text for the error
Dim LogFile  		'The file you want to search

error1String = "#E361"
LogFile = "c:\windows\setupapi.log"
WScript.Echo "starting script " & now
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(LogFile, 1)
strLine = objFile.ReadLine

Do Until objFile.AtEndofStream 
    strLine = objFile.ReadLine
    SearchResult = InStr(strLine, error1String)
    If SearchResult <>0 Then
   WScript.Echo(strLine)
   	End if
Loop
WScript.Echo("all done " & now)
objFile.Close
