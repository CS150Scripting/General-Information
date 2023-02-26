'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: CopyFolderCheckError.vbs
'
' AUTHOR: ed wilson , mred
' DATE  : 4/9/2006
' ver.1.2 cleaned up code, added additional comments
' COMMENT: <Uses the CopyFolder method of the fileSystemObject.>
'
'==========================================================================
Option Explicit
On Error Resume Next
Dim objFSO		'the filesystem object
Dim strSource	'Source folder
Dim strDestination	'Destination location

strSource = "c:\fso"					'source can be local location
strDestination = "q:\fsoED" 'destination can be local, or UNC

Set objFSO = CreateObject("scripting.fileSystemObject")
objFSO.CopyFolder strSource, strDestination


If Err.Number <> 0 Then
	WScript.Echo "An error occurred copying " & strSource &_
		" To " & strDestination & VbCrLf & "The error that " &_
		"occurred Was " & Err.Number & VbCrLf & Err.Source & _
		VbCrLf & Err.Description
End If
