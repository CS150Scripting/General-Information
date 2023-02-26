'==========================================================================
'
' VBScript:  AUTHOR: edwilson , msft,  3/13/2006
'
' NAME: <RunNetStat.vbs>
'Version 2.0 'cleaned up code, changed to readAll(). Added comments
' COMMENT: Key concepts are listed below:
'1. using the Wscript.shell exec method to run programs
'2. using the StdOut method to capture output
'3. KB281336 talks about using netstat. 
'==========================================================================

Option Explicit    	'Is used to force the scripter to declare variables
'On Error Resume Next 	'Is used to tell vbscript to go to the next line if it encounters an Error
Dim objShell		'Holds WshShell object	
Dim objExecObject	'Holds what comes back from executing the command
Dim strText		'Holds the text stream from the exec command. 
Dim command 		'The command to run

command = "cmd /c netstat -ano"
WScript.echo "starting program " & Now 'Used to mark when program begins
Set objShell = CreateObject("WScript.Shell")
Set objExecObject = objShell.Exec(command)

Do Until objExecObject.StdOut.AtEndOfStream
    strText = objExecObject.StdOut.ReadAll()
    	Wscript.Echo strText
Loop
WScript.echo "complete" 'Lets me know program is done running

