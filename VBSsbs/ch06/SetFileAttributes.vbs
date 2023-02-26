'==========================================================================
'
' VBScript Source File -- 
'
' NAME: <SetfileAttributes.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 4/6/2006
'Version 2.0 Added new function
' COMMENT: <demonstrates how to report, interpret, and set file attributes
' the following commands are germane:
'1. createobject to create filesystemobject
'2. attributes method 
'3. GetFile command
'4. Uses function to interpret existing file attributes
'5. Assigns specific value for desired file attrib
'==========================================================================
Option Explicit
On Error Resume Next  
Dim objFSO			'The file system object
Dim objFile			'The file object
Dim strTarget		'Path to target file
Dim intAttrib		'desired file attribute combination

strTarget = "C:\fso\test.txt"
intAttrib = 0
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strTarget)

WScript.Echo "The file is: " & strTarget
WScript.Echo "OLD bitmap number is: " & objFile.Attributes & _
	" " & funAttrib(objFile.attributes) & vbNewLine

SubsetAttrib

' *** subs and functions below *****

Function funAttrib(intMask)
Dim strAttrib
If IntMask = 0 Then strAttrib =  "No attributes"
If intMask And 1 Then strAttrib = strAttrib & "Read Only, "
If intMask And 2 Then strAttrib = strAttrib & "Hidden, "
If intMask And 4 Then strAttrib = strAttrib & "System, "
If intMask And 8 Then strAttrib = strAttrib & "Volume, "
If intMask And 16 Then strAttrib = strAttrib & "Directory, "
If intMask And 32 Then strAttrib = strAttrib & "Archive, "
If intMask And 64 Then strAttrib = strAttrib & "Alias, "
If intMask And 2048 Then strAttrib = strAttrib & "Compressed, "
funAttrib = strAttrib
End Function


Sub SubsetAttrib 
objFile.Attributes = intAttrib
WScript.Echo "The new attibutes are: " & funAttrib(objFile.Attributes)
End Sub