'==========================================================================
' NAME: <createTextFile.vbs>
'
' AUTHOR: Ed Wilson , MS
' DATE  : 10/20/2003
'
' COMMENT: <uses the createtextfile method of the filesystem object>
'
'==========================================================================

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.CreateTextFile("C:\FSO.txt")
