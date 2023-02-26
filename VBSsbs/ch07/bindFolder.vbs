'==========================================================================
' NAME: bindFolder.vbs
'
' AUTHOR: ed wilson , mred
' DATE  : 10/26/2003
'
' COMMENT: <comment>
'
'==========================================================================

Set objFSO = CreateObject("Scripting.filesystemobject")
Set objFolder = objFSO.getfolder("c:\fso")
WScript.Echo("folder is bound")