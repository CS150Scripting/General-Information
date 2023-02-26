'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  4/8/2006
'
' NAME: CreateTempFileNameAndOpenInNotepad.vbs
'
' COMMENT: Key concepts are listed below:
'1.A function that creates a temporary file and folder. Pass it a
'2.Filesystem object!
'3.Returns the path to the temporary folder and the temp file name
'4.that was created. It then creates the temporary file using createtextfile
'5.Then it writes to the tempfile, and opens same in notepad. 
'==========================================================================
Option Explicit
Dim objFSO		'filesystem object
Dim objfile		'file object
Dim objshell	'wshshell object
Dim strpath		'path to temp file. From FunTempFile 

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objshell = CreateObject("wscript.shell")
strpath = FunTempFile(objFSO)

Set objFile = objFSO.CreateTextFile(strpath)
objfile.Write("Writing to a temporary file ") & Now
objshell.Run("notepad " & strPath)

' **** Function below *****

Function FunTempFile(objFSO)	'Creates temp folder, and temp file name
Dim objfolder 								'temporary folder object
Dim  strName									'Temporary file name

Const TemporaryFolder = 2			'File system object constant value

Set objfolder = objfso.GetSpecialFolder(TemporaryFolder)
   	strName = objfso.GetTempName
   	strName = objfolder & "\" & strName   
  	FunTempFile = strName  
End Function