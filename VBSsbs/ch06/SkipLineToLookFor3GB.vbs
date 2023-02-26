'==========================================================================
' VBScript:  AUTHOR: Ed Wilson , MS,  4/8/2006
'
' NAME: <SkipLineToLookFor3GB.vbs>
'
' COMMENT: Key concepts are listed below:
'1.Reads the boot.ini file and looks for presence of /3gb (a common problem)
'2.Uses testboot.ini to practice on.
'3.Uses readLine method to read entire file into memory, then uses instr to 
'4.Find the search string. It only reports line number where found
'5.Uses .line property to know where are at in the text file. If the current 
'6.line is less than the line where found match, then we use skipline to go
'7.to next line. However, if it is equal to line where was found, we print
'8.out the line by using readline. If greater, we end the script. We also use
'9.The exit sub command to quit subroutine early, and wscript.quit to end.
'==========================================================================
Option Explicit
On Error Resume Next
Dim objFSO		'The fileSystemObject
Dim objFile		'The file object
Dim strFIle		'Path to the file
Dim strSearch	'String to search for
Dim strText		'text of the textBoot.ini file
Dim intLine 	'used to hold the line representing the start of ini file
Dim i

strFIle = "C:\fso\testBoot.ini" 'the ini file to parse
strSearch = "/3GB"							'the string to search For

Set objFSO = CreateObject("scripting.FileSystemObject")

sublook

WScript.Echo "opening the file a second time ..."

Set objFile = objFSO.OpenTextFile(strFIle)
Do Until objfile.AtEndOfStream
	If objFile.Line < intLine Then
			objFile.SkipLine
		ElseIf objFIle.Line = intLine Then
			WScript.Echo "The beginning of the ini file is: "_
				& vbNewLine & Space(5)& objfile.readLine
		Else
			WScript.Echo "the script is over"
			WScript.quit 
	End If
Loop

' **** functions below ****
Function funLookUP(strText,strSearch)	
Const blnInsensitive = 1
	If InStr (1,strText, strSearch,blnInsensitive) Then
		funLookUP=strSearch & " was found"
	Else
		funLookUP=strSearch & " was not found"
	End If
End Function

Sub sublook
Dim strLine
strSearch = "["
Set objFile = objFSO.OpenTextFile(strFIle)
	Do Until objFile.AtEndOfStream
		strText = objFile.ReadLine							'reads one line at a time
		strLine = funLookUP(strText,strSearch)	'uses function to parse line
		If InStr (strLine, "not") Then
			intLine = (objFile.Line -1) 					'because line method adds an extra one
			WScript.Echo intLine & _
				" not at the beginning of the ini"
		Else
			intLine = (objFile.Line -1) 					'because line method adds an extra one
			WScript.Echo intLine & _
				" is the beginning of the ini"
				objFile.Close 											'have to close file to reset pointer in it.
				Exit Sub 														'leaves the sub routine early.
		End If
	Loop
End Sub