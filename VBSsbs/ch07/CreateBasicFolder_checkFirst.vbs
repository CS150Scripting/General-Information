'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: CreateBasicFolder_checkFirst.vbs
'
' AUTHOR: ed wilson , mred
' DATE  : 10/25/2003
'
' COMMENT: <Uses createFolder method>
'
'==========================================================================

Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists ("C:\fso1") Then
WScript.Echo("folder exists and will be deleted")
objFSO.deleteFolder ("C:\fso1")
WScript.Echo("clean folder created")
Set objFolder = objFSO.CreateFolder("C:\fso1")
Else
WScript.Echo("folder does not exist and will be created")
Set objFolder = objFSO.CreateFolder("C:\fso1")
End if