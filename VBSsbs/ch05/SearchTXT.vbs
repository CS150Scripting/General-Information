'==========================================================================
'
' NAME: <SearchTXT.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 9/17/2003
'
' COMMENT: <comment>
'1.Modify myFile as required to point to specific path of file.
'==========================================================================

Option Explicit
On Error Resume Next
Dim arrTxtArray()
Dim myFile 
Dim SearchString
Dim objTextFile 
Dim strNextLine 
Dim intSize 
Dim objFSO 
Dim i 
intSize = 0 
myFile = "c:\windows\setuplog.txt" 'Modify as required
SearchString = "Error"
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (myFile, ForReading)
Do until objTextFile.AtEndOfStream 
    strNextLine = objTextFile.Readline
    if InStr (strNextLine, SearchString)then
    	ReDim Preserve arrTxtArray(intSize)
    	arrTxtArray(intSize) = strNextLine
    	intSize = intSize + 1
    End if
Loop
objTextFile.close
For i = LBound(arrTxtArray) To UBound(arrTxtArray)
	WScript.Echo arrTxtArray(i)
Next
WScript.Echo("all done")
