'==========================================================================
' NAME: <DisplayAdminTools_logged.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 4/6/2006
'ver.1.2 'adapted to use new code. minor modifications.
' COMMENT: <Modified DisplayAdminTools script from ch. 1 to include logging of 
' the run. Note this illustrates use of file system object to use create obj
' and write line to enable writing to a log file. >
'
'==========================================================================
LogFile = "C:\fso\fso.txt"
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(LogFile, ForWriting)

Set objshell = CreateObject("Shell.Application")
Set objNS = objshell.namespace(&h2f)
Set colitems = objNS.items

objFile.WriteLine "Process started at " & Now 
	For Each objitem In colitems
		WScript.Echo objitem.name
	Next
objFile.WriteLine "Process completed at " & Now 
objFile.Close





