'==========================================================================
' NAME: <dictionary.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 3/31/2006
'
' COMMENT: <Illustrates programatically adding items to the dictionary object>
'1.creates scripting.dictionary object
'2.creates scripting.filesystemobject
'3.uses add method from dictionary object
'==========================================================================
Option Explicit
Dim objDictionary	'The dictionary object
Dim objFSO		'The filesystemobject object
Dim objFolder		'Created by getfolder method
Dim colFiles		'Collection of files from files method
Dim objFile		'Individual file
Dim aryKeys		'Array of keys
Dim strKey		'Individual key from array of keys
Dim strFolder 		'The folder to obtain listing of files

strFolder = "c:\windows" 'Ensure correct path

Set objDictionary = CreateObject("scripting.dictionary")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strFolder)
Set colFiles = objFolder.Files
For Each objFile in colFiles
    objDictionary.add objFile.Name, objFile.Size
Next

aryKeys = objDictionary.Keys

WScript.Echo "Directory listing of " & strFolder
WScript.Echo "***There are " & objDictionary.count & " files"
For Each strKey In aryKeys
	WScript.Echo "The file: " & strKey & " is: " & _
		 objDictionary.Item(strKey) & " bytes"
Next 


