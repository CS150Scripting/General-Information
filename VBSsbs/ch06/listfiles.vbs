'==========================================================================
' NAME: <listFiles.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 4/5/2006
'ver.1.2 'cleaned up code, added comments
' COMMENT: <demonstrates use of filesystem object to list files in a folder
' the following commands are germane:
'1. createobject to create filesystemobject
'2. getfolder method of the filesystemobject
'3. the files command to talk to files
'4. the for each loop to walk though the list of files>
'5. Make sure you modify value for folderPath 
'==========================================================================
Option Explicit
On Error Resume Next 
Dim FolderPath 	'Path to the folder to be searched for files
Dim objFSO			'The fileSystemObject
Dim objFolder		'The folder object
Dim colFiles		'Collection of files from files method
Dim objFile			'individual file object

FolderPath = "c:\fso"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(FolderPath)
Set colFiles = objFolder.Files

For Each objFile in colFiles
    WScript.Echo objFile.Name, objFile.Size & " bytes"
Next
