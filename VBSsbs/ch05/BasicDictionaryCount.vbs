'==========================================================================
' NAME: <basicDictionaryCount.vbs>
'
' AUTHOR: Ed Wilson , MS
' DATE  : 3/11/2006
'
' COMMENT: <Demo code for using count>
'1.illustrates how to create dictionary
'2.add an item to the dictionary
'3.and echo out a key item
'4.Illustrates count property
'==========================================================================
Option Explicit
Dim objDictionary, i

Set objDictionary = CreateObject("scripting.dictionary")
objDictionary.add 1, "server1"
objDictionary.Add 2, "server2"
objDictionary.Add 3, "server3"
objDictionary.Add 4, "server4"

For i = 1 To objDictionary.count
	WScript.Echo objDictionary.item (i)
Next

objDictionary.Add "5", "Server5"
WScript.Echo "The count after adding key ""5"" with ""server5""" & _
		" to the dictionary is " & objDictionary.Count
		
For i = 1 To objDictionary.count
	WScript.Echo objDictionary.item (i)
Next

WScript.Echo "The count after using the second for ... next loop "&_
	"Is " & objDictionary.Count
	
WScript.Echo "Key ""5"" is a " & typename(objdictionary.item("5"))
WScript.Echo "Key 6 is a " & typename(objdictionary.item(6))
