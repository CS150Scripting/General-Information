'==========================================================================
' NAME: <VerifyFileExists.vbs>
'
' AUTHOR: Ed Wilson , MS
' DATE  : 4/6/2006
'ver. 1.2 'cleaned up code, tightened logic
' COMMENT: <In this script, we define a constant. We then use the FSO
' to allow us to verify the existence of a file. We then use an if, Then
' loop to decide whether to write, or append to the file. >
'==========================================================================
LogFile = "C:\FSO\fso.txt"
Const ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(LogFile) Then
		Set objFile = objFSO.OpenTextFile(LogFile, ForAppending)
		objFile.Write "appending " & Now
	Else
		Set objFile = objFSO.CreateTextFile(LogFile)
			objfile.write "writing to new file " & now
	End If
'objFile.Close

subOpenLog

Sub subOpenLog
Dim wshshell
Set wshshell = CreateObject("WScript.Shell")
wshshell.Run(LogFile)
End Sub