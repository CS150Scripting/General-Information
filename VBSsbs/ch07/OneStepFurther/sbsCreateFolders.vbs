'==========================================================================
'
' NAME: sbsCreateFolders.vbs
'
' AUTHOR: ed wilson , mred
' DATE  : 8/2/2006
'v.2.0 moved folder location to special folder
' COMMENT: <Uses createFolder method
' uses folderExists method to check for folder first>
' uses wshShell specialfolders property to obtain path to my documents
'==========================================================================
Option Explicit
Dim numFolders
Dim folderPath
Dim folderPrefix
Dim objFSO
Dim objFolder
Dim i
Dim objShell
Dim strDocPath

Set objShell = CreateObject("WScript.Shell")
strDocPath = objShell.SpecialFolders("mydocuments")

numFolders = 10
folderPath = strDocPath & "\"
folderPrefix = "Student"

For i = 1 To numFolders
	Set objFSO = CreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(folderPath & folderPrefix & i) Then
			WScript.Echo(folderPath & folderPrefix & i & " exists." _
			& " folder not created")
		Else
	Set objFolder = objFSO.CreateFolder(folderPath & folderPreFix & i)
			WScript.Echo(folderPath & folderPrefix & i & " folder created")
		End If 
Next
