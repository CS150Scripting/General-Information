'==========================================================================
' NAME: CopyFolder.vbs
'
' AUTHOR: ed wilson , mred
' DATE  : 4/9/2006
' COMMENT: <Uses the CopyFolder method of the fileSystemObject.>
'1. DEMO CODE. Please do not write scripts like this.
'==========================================================================
Set objFSO = CreateObject("scripting.fileSystemObject")
objFSO.CopyFolder "C:\fso", "C:\myFSO"
