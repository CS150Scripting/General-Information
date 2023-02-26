'==========================================================================
' NAME: sbsLog.vbs
'
' AUTHOR: Ed Wilson , MS
' DATE  : 4/8/2006
' ver. 1.2 Cleaned up code a little.
' COMMENT: <This file illustrates the following:
' 1. ForWriting constant - allows overwriting of file
' 2. OpenTextFile method to write to a pre-existing file
' 3. WriteLine method to add new lines to the text file
' 4. Close method >
'==========================================================================
Option Explicit
Dim logfile ' holds path to the log file
Dim objFSO ' holds connection to the fileSystemObject
Dim objFile 'used by OpenTextFile command to allow writing to file

LogFile = "C:\FSO\fso.txt"
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(LogFile, ForWriting)

objFile.WriteLine "beginning logging " & Now
objFile.WriteLine "working on process " & Now
objFile.WriteLine "Logging completed at " & Now 
objFile.Close
