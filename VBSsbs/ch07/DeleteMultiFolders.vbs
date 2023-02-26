'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: deleteMultiFolders.vbs
'
' AUTHOR: ed wilson , mred
' DATE  : 10/25/2003
'
' COMMENT: <Uses deleteFolder method>
'
'==========================================================================
Option Explicit
Dim numFolders
Dim folderPath
Dim folderPrefix
Dim objFSO
Dim objFolder
Dim i

numFolders = 10
folderPath = "C:\"
folderPrefix = "TempUser"

For i = 1 To numFolders
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.deleteFolder(folderPath & folderPreFix & i)
Next

WScript.Echo(i - 1 & " folders deleted")
