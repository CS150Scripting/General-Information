'==========================================================================
' NAME: <multiValuesSearch.vbs>
' AUTHOR: ed wilson , mred
' DATE  : 3/29/2006
' version 1.2 added function, cleaned up code, re-did array
' COMMENT: <Starter Script for searchTXTmultiValues.vbs>
' 1. Using Split command to break up a line at comma
' 2. using FSO to read a text file
' 3. use of constants, and do Until
' 4. Creating Dynamic array and using reDim Preserve
' 5. use of Ubound
'==========================================================================
Option Explicit
'On Error Resume Next
Dim arrTxtArray() 'Dynamic array holds values from search
Dim myFile 	  'File to search 
Dim SearchString, SearchItem, Item 'Search items
Dim objTextFile   'The textstream object returned by OpenTextFile method
Dim strNextLine   'One line of text. From readline method
Dim intSize 	  'Counter variable used to re-size the array
Dim objFSO 	  'The filesystem object
Dim i 		  'Counter variable used to Walk through the Array

intSize = 0 
myFile = "c:\windows\setuplog.txt"
SearchString = "Error,failed,unable to,could not,was NOT"
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
objTextFile.close 'We have to close the file BEFORE we can obtain a new text stream
									'for the subsequent pass using instr
	
	For i = 0 To UBound(arrTxtArray)
		WScript.Echo arrTxtArray(i)
	Next
	intSize = 0
	ReDim arrTxtArray(intSize) 'Are not using preserve keyword here. We DO NOT want
															'data from previous pass in the array. so we resize 
															'it to 0.
Next

WScript.Echo "all done"

