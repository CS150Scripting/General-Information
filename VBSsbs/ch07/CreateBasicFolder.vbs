'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: CreateBasicFolder.vbs
'
' AUTHOR: ed wilson , mred
' DATE  : 10/25/2003
'
' COMMENT: <Uses createFolder method>
'
'==========================================================================

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.CreateFolder("c:\fso1")
