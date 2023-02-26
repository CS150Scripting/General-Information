'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  1/22/2004
'
' NAME: <DecimalToBinaryFunction.vbs>
'
' COMMENT: Key concepts are listed below:
'1. uses a function to convert decimal to Binary
'2. Function is called DecToBin takes an integer as input value
'3. 
'==========================================================================


myDecimal= DecToBin(274)
WScript.Echo(mydecimal)

Function DecToBin(intDec)
  dim strResult
  dim intValue
  dim intExp

  strResult = ""

  intValue = intDEC
  intExp = 65536
  while intExp >= 1
    if intValue >= intExp then
      intValue = intValue - intExp
      strResult = strResult & "1"
    else
      strResult = strResult & "0"
    end if
    intExp = intExp / 2
  wend

  DecToBin = strResult
End Function


