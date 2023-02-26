'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  4/15/2006
'
' NAME: <RecursiveListOfFolders.vbs>
'
' COMMENT: Key concepts are listed below:
'1.Uses the file system object,folder and objFolder commands
'2.does a recursive listing of folders by using a subRoutine
'3.Checks if folder exists prior to attempting recursion. If
'4.The folder DOES NOT exist, then it offers to create same using
'5.MsgBox. MsgBox also used in chapter 3. 
'==========================================================================
Option Explicit
'On Error Resume Next
Dim strTarget	'The place to begin recursive folder listing
Dim objFSO		'The file system object

strTarget = "c:\fso\mred"

Set objFSO = CreateObject("Scripting.FileSystemObject")

subCheck	'verifies existence of folder - offers to create same. 

' *** subs below ***
Sub subCheck
Dim strPrompt	'msgbox prompt
Dim strTitle	'title of msgbox
Dim errRTN		'return code from the msgbox Function

strPrompt = strTarget & " Does not exist." &_
		vbNewLine & "Would you like to Create it?"
strTitle = strTarget & " not found!"

	If objFSO.FolderExists(strtarget) Then
		SubRecursiveFolders objFSO.GetFolder(strTarget)
	Else
		errRTN = MsgBox(strPrompt,vbYesNo+vbQuestion,strTitle)
			If errRTN = vbYes Then
				objFSO.CreateFolder(strTarget)
			End If
	End If
End Sub
	
Sub subRecursiveFolders(Folder)
Dim objFolder
    For Each objFolder In Folder.subFolders
        Wscript.Echo objFolder.Path
        subRecursiveFolders objFolder
    Next
End Sub



