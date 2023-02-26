'==========================================================================
'
'
' NAME: <Instr2solution.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 3/28/2006
'
' COMMENT: <Demonstrates the advanced use of the Instr function>
'
'==========================================================================

searchString = "5"
textSearched = "123456789"
InstrReturn = InStr (1, textSearched, SearchString, 0)
WScript.Echo(InstrReturn)
