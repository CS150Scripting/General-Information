'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: moveFolder.vbs
'
' AUTHOR: ed wilson , mred
' DATE  : 10/26/2003
'
' COMMENT: <Uses the moveFolder method of the fileSystemObject.>
'
'==========================================================================



Set objFSO = CreateObject ("scripting.fileSystemObject")
objFSO.moveFolder "c:\fso","C:\fso2"

