'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  5/5/2004
'
' NAME: <filename>
'
' COMMENT: Key concepts are listed below:
'1.
'2.
'3. 
'==========================================================================

Option Explicit
On Error Resume Next  
Dim objFSO
Dim objFile
Dim Target

Target = "C:\boot.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(Target)

WScript.Echo "The file is: " & target
WScript.Echo "bitmap number is: " & (objFile.attributes)
WScript.Echo "This translates to: " & funAttrib(objFile.attributes)


Function funAttrib (inMask)
Dim intMask
If inMask AND 0 Then intMask = intMask & "No Attributes, "
If inMask AND 1 Then intMask = intMask & "Read-Only, "
If inMask AND 2 Then intMask = intMask & "Hidden, "
If inMask AND 4 Then intMask = intMask & "System, "
If inMask AND 32 Then intMask = intMask & "Archive Bit Set, "
If inMask AND 64 Then intMask = intMask & "Link or ShortCut, "
If inMask AND 2048 Then intMask = intMask & "Compressed, "
funAttrib = intMask
End function