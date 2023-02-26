'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <GetCommentsTimed.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 3/16/2006
'
' COMMENT: <This scripts shows the following concepts>
' 1. using constants
' 2. use of filesysteobjectt
' 3. use of do while construct
' 4. use of if Then
' 5. use of vbcrlf
' 6. use of InStr
' 7. Uses the timer function to see how long the script runs
' 8. Uses the formatNumber function to clean up time.
'==========================================================================
Option Explicit
'On Error Resume next
Dim scriptFile
Dim commentFile
Dim objScriptFile
Dim objFSO
Dim objCommentFile
Dim strCurrentLine
Dim intIsComment
Dim startTime, endTime

Const ForReading = 1
Const ForWriting = 2
scriptFile = "displayComputerNames.vbs"
commentFile = "comments.txt"
startTime = Timer
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objScriptFile = objFSO.OpenTextFile _
    (scriptFile, ForReading)
Set objCommentFile = objFSO.OpenTextFile(commentFile, _ 
    ForWriting, TRUE)
Do While objScriptFile.AtEndOfStream <> TRUE
    strCurrentLine = objScriptFile.ReadLine
    intIsComment = Instr(1,strCurrentLine,"'")
    If intIsComment > 0 Then
        objCommentFile.Write strCurrentLine & VbCrLf
    End If
Loop
endTime = Timer
WScript.Echo "script complete. " & round(endTime-startTime, 2)
objScriptFile.Close
objCommentFile.Close