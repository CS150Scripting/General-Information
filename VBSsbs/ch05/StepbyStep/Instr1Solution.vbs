'==========================================================================
'
'
' NAME: <Instr1Solution.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 3/28/2006
'
' COMMENT: <Demonstrates the use of the Instr command>
'
'==========================================================================

searchString = "5"
textSearched = "123456789"
InstrReturn = InStr (textSearched, SearchString)
WScript.Echo(InstrReturn)
