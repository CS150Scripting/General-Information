'==========================================================================
' VBScript:  AUTHOR: Ed Wilson , MS,  4/6/2006
'
' NAME: <LookFor3GB.vbs>
'
' COMMENT: Key concepts are listed below:
'1.Reads the boot.ini file and looks for presence of /3gb (a common problem)
'2.Uses testboot.ini to practice on.
'3.Uses readALL method to read entire file into memory, then uses instr to 
'4.Find the search string. It only reports if it is present or not. 
'==========================================================================
Option Explicit
On Error Resume Next
Dim objFSO
Dim objFile
Dim strFIle
Dim strSearch, strText

strFIle = "C:\fso\testBoot.ini"
strSearch = "/3GB"	


Set objFSO = CreateObject("scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(strFIle)
strText = objFile.ReadAll

WScript.Echo funLookUP(strText,strSearch)

' **** functions below ****
Function funLookUP(strText,strSearch)	
Const blnInsensitive = 1
	If InStr (1,strText, strSearch,blnInsensitive) Then
		funLookUP=strSearch & " was found"
	Else
		funLookUP=strSearch & " was not found"
	End If
End Function