'==========================================================================
' NAME: <osfLogIfExista.vbs>
'
' AUTHOR: Ed Wilson , MS
' DATE  : 4/6/2006
'ver.1.2
' COMMENT: <In this script, we define two constants. We then use the FSO
' to allow us to verify the existence of a file. We then use an if, Then
' loop to decide whether to write, or append to the file. >
'
'==========================================================================
Option Explicit
dim logfile
dim objFSO
dim objFile

LogFile = "C:\FSO\fso.txt"
Const ForWriting = 2
Const ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(LogFile) Then
	Set objFile = objFSO.OpenTextFile(LogFile, ForAppending)
	objFile.Write "appending " & Now
Else
	Set objFile = objFSO.CreateTextFile(LogFile)
	objFile.Close
	Set objFile = objFSO.OpenTextFile(LogFile, ForWriting)  
	objfile.write "writing to new file " & now
End If
objFile.Close