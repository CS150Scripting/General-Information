'==========================================================================
' NAME: <ListMultipleSpecialFolders_logged.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 4/6/2006
'	ver.1.2 'adapted to use new code. 
' COMMENT: <Modified DisplayAdminTools script from ch. 1 to include logging of 
' the run. Note this illustrates use of file system object to use create obj
' and write line to enable writing to a log file. >
' There are further modifications that can / should be made to this script:
' 1. move entire logging into perhaps a Function
' 2. build up output variable from tool enumeration this will enable single 
' 3. write to the text file
' 4. implement some real error handling
'==========================================================================
Option Explicit
On Error Resume Next
Dim objFSO 		'The filesystemobject
Dim objFILE 	'file object
Dim logFIle 	'path to log file
Dim objShell 	'shell application object
Dim objNS 		'special folder to connect to
Dim colItems	'collection of items in the folder
Dim objItem 	'single file in the folder
Dim intNS			'individual ns value
Dim strMSG		'The root message written to Log
Dim aryNS			'array of namespace names

strMSG = "Enumerating items: "

aryNS = array(&ha,&h20,&h6)	'special folder values See Appendix E.
LogFile = "C:\fso\fso.txt"
Const ForWriting = 2


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFILE = objFSO.OpenTextFile(LogFile, ForWriting)

For Each intNS In aryNS
	objFile.WriteLine strMSG & " in folder " & intNS & _
		" Started " & Now 
	
	Set objshell = CreateObject("Shell.Application")
	Set objNS = objshell.namespace(intNS)
	Set colitems = objNS.items
		For Each objitem In colItems
			objFILE.writeline objitem.name
		Next
	objFile.WriteLine strMSG & "completed " & Now 
Next

subOpenLog 'opens the log file automatically

' *** subs below ****
Sub subOpenLog 'from SubOpenLogFile.vbs in utilities folder
Dim wshshell
Set wshshell = CreateObject("WScript.Shell")
wshshell.Run(LogFile)
End Sub




