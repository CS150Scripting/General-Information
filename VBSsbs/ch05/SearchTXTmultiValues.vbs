'==========================================================================
' NAME: <searchTXTmultiValues.vbs>
' AUTHOR: ed wilson , mred
' DATE  : 3/29/2006
' version 1.2 added function, cleaned up code, re-did array
' COMMENT: <The following items are illustrated>
' 1. Using Split command to break up a line at comma
' 2. using FSO to read a text file
' 3. use of constants, and do Until
' 4. Creating Dynamic array and using reDim Preserve
' 5. use of Ubound
'==========================================================================
Option Explicit
'On Error Resume Next
Dim arrTxtArray()
Dim myFile 
Dim SearchString, SearchItem, Item
Dim objTextFile 
Dim strNextLine 
Dim intSize 
Dim objFSO 
Dim i 
Dim strPrompt,strTitle,strDefault 'used for input box


intSize = 0 
myFile = "c:\windows\setuplog.txt"
strPrompt = "Enter error words to search for in: " & _
	VbCrLf & myFile
strTitle = "Error locator"
strDefault = "Error,failed,unable to,could not,was NOT"
SearchString = InputBox(strPrompt,strTitle,strDefault)
SearchItem = Split(SearchString, ",")
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")

For Each Item In SearchItem
Set objTextFile = objFSO.OpenTextFile(myFile)
Do until objTextFile.AtEndOfStream 
    strNextLine = objTextFile.Readline
      If InStr (strNextLine, Item)then
       	ReDim Preserve arrTxtArray(intSize)
	    	arrTxtArray(intSize) = strNextLine
	    	intSize = intSize + 1
			End If
Loop
	objTextFile.close
WScript.Echo funLine("There are " & ubound(arrTxtArray)+1 &_
	 " Lines with " & """" & Item & """" & " in them")
	For i = 0 To UBound(arrTxtArray)
		WScript.Echo arrTxtArray(i)
	Next
		intSize = 0
		ReDim arrTxtArray(intSize)
Next

WScript.Echo "all done"

' *** functions below *****
Function funLine(lineOfText)
Dim numEQs, separator, i
numEQs = Len(lineOfText)
For i = 1 To numEQs
	separator= separator & "="
Next
 FunLine = VbCrLf & lineOfText & vbcrlf & separator
End Function



