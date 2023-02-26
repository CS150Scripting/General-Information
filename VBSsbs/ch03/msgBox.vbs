'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  3/15/2006
'
' NAME: MsgBox.vbs
'
' COMMENT: Key concepts are listed below:
'1.Use of the MsgBox
'2.Use of if ... then ... elseif ... else
'3. 
'==========================================================================
Option Explicit
Dim strPrompt
Dim strTitle
Dim intBTN
Dim intRTN

strPrompt = "Do you want to run the script?"
strTitle = "MsgBox DEMO"
intBTN = 3 '4 is yes/no 3 yes/no/cancel

intRTN = MsgBox(strprompt,intBTN,strTitle)


If intRTN = vbYes Then
	WScript.Echo "yes was pressed"
ElseIf intRTN = vbNo Then
	WScript.Echo "no was pressed"
ElseIf intRTN = vbCancel Then
	WScript.Echo "cancel was pressed"
Else
	WScript.Echo intRTN & " was pressed"
End If