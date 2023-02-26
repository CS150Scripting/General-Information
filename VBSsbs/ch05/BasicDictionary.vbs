'==========================================================================
' NAME: <basicDictionary.vbs>
'
' AUTHOR: Ed Wilson , MS
' DATE  : 6/21/2003
'
' COMMENT: <Example script>
' illustrates how to create dictionary
' add an item to the dictionary
' and echo out a key item
' important thing here is that you reference the ITEM via the KEY value.
' note that 1,2,3,4 are all keys. server1 ... are the items. This can
' be either cool, or a real source of confusion!
'==========================================================================
Option Explicit
Dim objDictionary, i
Set objDictionary = CreateObject("scripting.dictionary")
objDictionary.add 1, "server1"
objDictionary.Add 2, "server2"
objDictionary.Add 3, "server3"
objDictionary.Add 4, "server4"
For i = 1 To 4
WScript.Echo objDictionary.item (i)
next