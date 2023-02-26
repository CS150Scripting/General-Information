'==========================================================================
'
' NAME: <InstrSample1.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 9/21/2003
'
' COMMENT: <demonstrates the use of the Instr command>
'
'==========================================================================

searchString = "searchstring"
textSearched = "The InStr function is used to find a searchstring inside a text stream"

InstrReturn = InStr (textSearched, SearchString)
WScript.Echo(InstrReturn)
