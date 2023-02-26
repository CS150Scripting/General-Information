'==========================================================================
'
' VBScript Source File -- 
'
' NAME: <fileAttributes.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 4/6/2006
'Version 2.0 Complete re-wrote function
' COMMENT: <demonstrates use of attributes bitmap to Echo out file attributes
' the following commands are germane:
' 1. createobject to create filesystemobject
' 2. attributes method 
' 3. GetFile command
' 4. the select case construct to parse the bitmap>
'
'==========================================================================
Option Explicit
On Error Resume Next  
Dim objFSO
Dim objFile
Dim Target

Target = "C:\fso\test.txt"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(Target)

WScript.Echo "The file is: " & target
WScript.Echo "bitmap number is: " & objFile.Attributes & _
	" " & funAttrib(objFile.attributes)


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
