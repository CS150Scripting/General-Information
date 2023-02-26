'==========================================================================
' NAME: <basicDictionaryKEYS.vbs>
'
' AUTHOR: Ed Wilson , MS
' DATE  : 3/11/2006
'
' COMMENT: <Demo code for using Keys>
'1.Illustrates how to create dictionary
'2.Add an item to the dictionary
'3.Echo out a key item
'4.Illustrates count Property
'5.Uses KEYS to obtain an array of keys
'==========================================================================
Option Explicit
Dim objDictionary, i
Dim aryKeys 		'Holds array of keys from keys method
Dim key 		'An individual key in the array
Set objDictionary = CreateObject("scripting.dictionary")
objDictionary.add 1, "server1"
objDictionary.Add 2, "server2"
objDictionary.Add 3, "server3"
objDictionary.Add 4, "server4"

For i = 1 To objDictionary.count
	WScript.Echo objDictionary.item (i)
Next

objDictionary.Add "5", "Server5"
WScript.Echo "The count after adding key ""5"" with ""server5"""&_
		" to the dictionary is " & objDictionary.Count
		
For i = 1 To objDictionary.count
	WScript.Echo objDictionary.item (i)
Next

WScript.Echo "The count after using the second for ... next loop "&_
	"Is " & objDictionary.Count
	
WScript.Echo "Item ""5"" is a " & typename(objdictionary.item("5"))
WScript.Echo "Item 6 is a " & typename(objdictionary.item(6))

aryKeys = objDictionary.Keys
WScript.Echo "aryKeys is " & vartype(aryKeys)

For Each key In aryKeys
	WScript.Echo "key " & key & " is a " & vartype(key)
Next