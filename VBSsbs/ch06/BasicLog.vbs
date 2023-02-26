'==========================================================================
' NAME: basicLog.vbs
'
' AUTHOR: Ed Wilson , MS
' DATE  : 4/2/2006
'
' COMMENT: <This file illustrates the following:
' 1. ForWriting constant - allows overwriting of file
' 2. OpenTextFile method to write to a pre-existing file
' 3. WriteLine method to add new lines to the text file
' 4. Close method >
'
'==========================================================================
Option Explicit
On Error Resume Next 
Dim LogFile 'The name of the file to create
Dim objFSO  'Contains the File System Object
Dim objFile 'Contains a file object
LogFile = "C:\FSO\fso.txt"
Const ForWriting = 2
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(LogFile, ForWriting)
objFile.WriteLine "beginning process " & Now
objFile.WriteLine "working on process " & Now
objFile.WriteLine "Process completed at " & Now 
objFile.Close
