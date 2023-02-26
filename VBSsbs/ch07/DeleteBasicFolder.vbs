'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: DeleteBasicFolder.vbs
'
' AUTHOR: ed wilson , mred
' DATE  : 10/25/2003
'
' COMMENT: <Uses DeleteFolder method>
'
'==========================================================================

Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.DeleteFolder("c:\fso1")
