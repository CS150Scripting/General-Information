'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <GetComments.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 5/26/2003
'
' COMMENT: <This scripts shows the following concepts>
' 1. using constants
' 2. use of filesysteobjectt
' 3. use of do while construct
' 4. use of if Then
' 5. use of vbcrlf
' 6. use of InStr
'
'==========================================================================
Option Explicit
On Error Resume Next
Dim scriptFile
Dim commentFile
Dim objScriptFile
Dim objFSO
Dim objCommentFile
Dim strCurrentLine
Dim intIsComment
Const ForReading = 1
Const ForWriting = 2
scriptFile = "displayComputerNames.vbs"
commentFile = "comments.txt"
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
WScript.Echo("script complete")
objScriptFile.Close
objCommentFile.Close