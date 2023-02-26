'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: CopyFolder.vbs
'
' AUTHOR: ed wilson , mred
' DATE  : 10/26/2003
'
' COMMENT: <Uses the CopyFolder method of the fileSystemObject.>
'
'==========================================================================
Option Explicit
Dim objFSO 				'the filesystemobject
Dim strSource			'source files
Dim strDestination'destination files
Dim startTime	'timestamp for timer function
Dim endTime		'timestamp for timer function

Const OverWriteFiles = True 
startTime = Timer
WScript.Echo " beginning copy ..."
strSource = "c:\Documents and Settings"
strDestination = "\\London\fileBU"

Set objFSO = CreateObject ("scripting.fileSystemObject")
objFSO.CopyFolder strSource, strDestination , OverWriteFiles
endTime = Timer 
WScript.Echo "ending copy. It took: " & _
	Round(endtime-startTime) & " seconds to copy"
