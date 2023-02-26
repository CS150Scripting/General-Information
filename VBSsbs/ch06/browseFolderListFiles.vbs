'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  4/5/2006
'
' NAME: BrowseFolderListFiles.vbs
'
' COMMENT: Key concepts are listed below:
'1.Uses the shell.application browseForFolder method to obtain path to folder
'2.Uses checkForWScript subroutine to identify script running in wscript
'3.Uses the files method from filesystemobject to obtain list of files
'4.Uses getFoler method to get folder object. This takes the path obtained
'5.From browseFOrFOlder method which is fully documented in the platform SDK
'==========================================================================
Option Explicit 
On Error Resume Next
Dim FolderPath 	'Path to the folder to be searched for files
Dim objFSO			'The fileSystemObject
Dim objFolder		'The folder object
Dim colFiles		'Collection of files from files method
Dim objFile			'individual file object
Dim strOUT 			'Single output variable

subCheckWscript	'Ensures script is running under wscript
subGetFolder		'Calls the browseForFOlder method

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(FolderPath)
Set colFiles = objFolder.Files

For Each objFile in colFiles
   strOUT = strOUT & objFile.Name & vbTab & objFile.Size _
   & " bytes" & VbCrLf
Next

WScript.Echo strOUT


' ****** subs below ******

Sub subCheckWscript
If UCase(Right(WScript.FullName, 11)) = "CSCRIPT.EXE" Then
    WScript.Echo "This script must be run under WScript."
    WScript.Quit
End If
End Sub

Sub subGetFolder
Dim objShell, objFOlder, objFolderItem
Const windowHandle = 0
Const folderOnly = 0
const folderAndFiles = &H4000&

Set objShell = CreateObject("Shell.Application")      
Set objFolder = objShell.BrowseForFolder(windowHandle, _
		"Select a folder:", folderOnly)       
Set objFolderItem = objFolder.Self   
FolderPath = objFolderItem.Path
End Sub

