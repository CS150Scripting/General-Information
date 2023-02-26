'==========================================================================
'
'
' NAME: CreateMultiFolders.vbs
'
' AUTHOR: ed wilson , mred
' DATE  : 8/1/2006
' ver.2.0 ' changed location to my documents folder.
' COMMENT: <Uses createFolder method>
'1.Uses the specialFolders property from the WshShell object to retrieve
'2.The path to the my documents folder. 
'3.Uses concatneation to create 10 temp user folders.
'4.Uses the createFolder method from the filesystemobect
'==========================================================================
Option Explicit
Dim numFolders 	'The number of folders to create
Dim folderPath	'The path for the folders
Dim folderPrefix'The first part of each folder name
Dim objFSO			'The file system object
Dim objFolder		'The folder object
Dim i						'Counter used to determine how many folders get created
Dim objSHell
Dim myDocs

Set objSHell = CreateObject("wscript.shell")
myDocs = objSHell.SpecialFolders("mydocuments")

numFolders = 10
folderPath = myDocs & "\"
folderPrefix = "TempUser"

Set objFSO = CreateObject("Scripting.FileSystemObject")
For i = 1 To numFolders
	Set objFolder = objFSO.CreateFolder(folderPath & folderPreFix & i)
Next
WScript.Echo(i - 1 & " folders created")
