'==========================================================================
' NAME: <basicDictionaryAddKeys.vbs>
'
' AUTHOR: Ed Wilson , MS
' DATE  : 3/11/2006
'
' COMMENT: <Demo code for adding keys to a dictionary>
'1.illustrates how to create dictionary
'2.add an item to the dictionary
'3.and echo out a key item
'4.Illustrates count Property
'5.Uses KEYS to obtain an array of keys
'6.Uses remove method to remove a key
'6.Uses exists method to avoid errors
'7.YOU REALLY want to run this script from cscript!
'==========================================================================
Option Explicit
Dim objDictionary, i
Dim aryKeys 	'Holds array of keys from keys method
Dim key 	'An individual key in the array
Dim strItem 	'Holds data stored in key "5"
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

WScript.Echo "Before we remove key 6, the count is: " & objDictionary.Count
WScript.Echo "Removing key 6 ..." & objDictionary.Remove(6)
WScript.Echo "After removal of 6, the count is: " & objDictionary.Count


strItem = objDictionary.Item("5")
objDictionary.Remove("5")
objDictionary.Remove(5)
WScript.Echo "After removing two keys, count is: " & objDictionary.Count

objDictionary.Add objdictionary.Count +1,strItem
WScript.Echo "After adding back, the count is: " & objDictionary.count

For i = 1 To objDictionary.Count
	If objDictionary.exists(i) Then
		WScript.Echo objDictionary.Item(i)
	End If
Next