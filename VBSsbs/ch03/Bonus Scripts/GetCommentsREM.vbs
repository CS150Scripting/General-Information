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
'On Error Resume next
Dim scriptFile			Rem: holds the name of the script to search for comments
Dim commentFile			Rem: hold the resulting comments
Dim objScriptFile		Rem: file to open up
Dim objFSO
Dim objCommentFile	Rem: file I write to
Dim strCurrentLine
Dim intIsComment

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const CreateFile = true
scriptFile = "C:\Labs\ch3\GetComments.vbs" 'displayComputerNames.vbs"
commentFile = "C:\Labs\ch3\comments.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objScriptFile = objFSO.OpenTextFile _
	(scriptFile, ForReading)
Set objCommentFile = objFSO.OpenTextFile(commentFile, _ 
    ForWriting, CreateFile)

Do until objScriptFile.AtEndOfStream 
    strCurrentLine = objScriptFile.ReadLine
    intIsComment = Instr(1,strCurrentLine,"'")
	    If intIsComment > 0 Then 'change to = 0 for NO COMMENTS in script!
	        objCommentFile.Write Right(strCurrentLine, Len(strCurrentLine)-intIsComment+1) & VbCRLF
	    Else
	    		intIsComment = InStr(1, UCase(strCurrentLine), "REM")
	    		If intIsComment > 0 Then
	    				objCommentFile.Write Right(strCurrentLine, Len(strCurrentLine)-intIsComment+1) & VbCRLF
	   			End If
	    End If
Loop

WScript.Echo("script complete")

objScriptFile.Close	' not needed
objCommentFile.Close' not needed