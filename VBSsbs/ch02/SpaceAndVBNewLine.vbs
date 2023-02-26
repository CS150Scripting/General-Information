'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  3/12/2006
'
' NAME: SpaceAndVBNewLine.vbs
'
' COMMENT: Key concepts are listed below:
'1.Uses the space function and the vbnewline constant
'==========================================================================
Option Explicit

WScript.Echo Space(10) & "this is a 10 space line at the beginning"
wscript.Echo "This line ends with vbnewline" & vbNewLine
WScript.Echo "This is an embedded 5 spaces" & Space(5) & "in the line"
