'==========================================================================
'
'
' NAME: EfficientFolder.vbs
'
' AUTHOR: ed wilson , mred
' DATE  : 4/14/2006
'
' COMMENT: <Uses createFolder method>
'1.Creates multiple folders off the root of the c drive.
'2.Uses timer function to see how long it takes to create the folders
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

Dim	startTime, EndTime, TotalTime	'Used for timer Function

startTime = Timer
Set objSHell = CreateObject("wscript.shell")
myDocs = objSHell.SpecialFolders("mydocuments")

folderPath = myDocs & "\"
numFolders = 100
folderPrefix = "TempUser"

Set objFSO = CreateObject("Scripting.FileSystemObject")
For i = 1 To numFolders
		Set objFolder = objFSO.CreateFolder(folderPath & folderPrefix & i)
Next
EndTime = Timer
TotalTime = EndTime-startTime
WScript.Echo(i - 1 & " folders created")
WScript.Echo "It took " & TotalTime &" seconds"

