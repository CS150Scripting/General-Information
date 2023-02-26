'==========================================================================
' NAME: EfficientFolderLogging.vbs
'
' AUTHOR: ed wilson , mred
' DATE  : 4/14/2006
'
' COMMENT: <Uses createFolder method>
'1.Creates multiple folders off the root of the c drive.
'2.Uses timer function to see how long it takes to create the folders
'3.Uses logging subroutine to log results of running the script.
'4.Uses the specialFolders method form wshShell object to find path to the
'5.Desktop. Uses run method from wshShell object to open up the log file 
'6.automatically. 
'==========================================================================
Option Explicit
Dim numFolders 	'The number of folders to create
Dim folderPath	'The path for the folders
Dim folderPrefix'The first part of each folder name
Dim objFSO			'The file system object
Dim objFolder		'The folder object
Dim i						'Counter used to determine how many folders get created
Dim	startTime, EndTime, TotalTime	'Used for timer Function

startTime = Timer
numFolders = 100
folderPath = "C:\"
folderPrefix = "TempUser"

For i = 1 To numFolders
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFSO.CreateFolder(folderPath & folderPrefix & i)
Next
endTime = Timer
totalTime = endTime-startTime
WScript.Echo(i - 1 & " folders created")
WScript.Echo "It took " & TotalTime &" seconds"

subLogging

'**** subs are below *****
Sub subLogging	'Logs the time the script was run, and how long it took to run.
Dim objShell	'WshShell object
Dim strDir		'Directory for log file.
dim strFile		'Path to the log file
Dim objFile		'The file object from opentextfile method

Set objShell = CreateObject("wscript.shell")
strDir = objShell.SpecialFolders("desktop")
strFile = strDir & "\myLog.txt"
Const forAppending = 8
Const blnCreate = True 'Will create the text file if it does not exist
Const intWindowPos = 4 'use most recent window position
Const blnWait = True 'script will wait until I manually close log file.

Set objFile = objFSO.OpenTextFile (strFile,ForAppending,blnCreate)
objFile.WriteLine("Running script" & VbCrLf & Now & " took " & totalTime)
strFile = """" & strfile & """"
objShell.run strFile,intWindowPos,blnWait
End Sub