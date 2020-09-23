<div align="center">

## ConvertRoman


</div>

### Description

Takes a roman number and convert into decimal.
 
### More Info
 
Inputs a roman number.

Returns a decimal number.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Francesco Meani](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/francesco-meani.md)
**Level**          |Unknown
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/francesco-meani-convertroman__1-1092/archive/master.zip)





### Source Code

```
Option Explicit
'Valid roman numerals and their values
Private Const M = 1000
Private Const D = 500
Private Const C = 100
Private Const L = 50
Private Const X = 10
Private Const V = 5
Private Const I = 1
Private Function IsRoman(ByVal numr As String) As Boolean
  'This function is given a character and returns true if it is
  'a valid roman numeral, false otherwise.
    'Convert digit to UpperCase
    numr = UCase(numr)
    'Test the digit
    Select Case numr
      Case "M"
        IsRoman = True
      Case "D"
        IsRoman = True
      Case "C"
        IsRoman = True
      Case "L"
        IsRoman = True
      Case "X"
        IsRoman = True
      Case "V"
        IsRoman = True
      Case "I"
        IsRoman = True
      Case Else
       IsRoman = False
    End Select
End Function
Private Function ConvertRoman(ByVal numr As String) As String
  'This function is given a roman numeral and returns its value.
  'NULL is returned if the character is not valid
Dim digit As Integer
    'Convert digit to UpperCase
    numr = UCase(numr)
    'Convert the digit
    Select Case numr
      Case "M"
        digit = M
      Case "D"
        digit = D
      Case "C"
        digit = C
      Case "L"
        digit = L
      Case "X"
        digit = X
      Case "V"
        digit = V
      Case "I"
        digit = I
      Case Else
        digit = vbNull
    End Select
    'And return its value
    ConvertRoman = digit
End Function
Public Function GetRoman(ByVal numr As String) As String
  'This function reads the next number in roman numerals from the input
  'and returns it as an integer
Dim rdigit As String
Dim num As Long
Dim DigValue As Long
Dim LastDigValue As String
Dim j As Long
  j = 1
  num = 0
  LastDigValue = M
    'Get the first digit
    rdigit = Mid(numr, j, 1)
    'While it is a roman digit
    Do While IsRoman(rdigit)
      'Convert roman digit to its value
      DigValue = ConvertRoman(rdigit)
      'If previous digit was a prefix digit
      If DigValue > LastDigValue Then
        'Adjust total
        num = num - 2 * LastDigValue + DigValue
      Else
        'Otherwise accumulate the total
        num = num + DigValue
        'Save this digit as previous
        LastDigValue = DigValue
      End If
        'Get next digit
         j = j + 1
         rdigit = Mid(numr, j, 1)
        'End of the string detected, exit
         If Len(rdigit) = 0 Then
           Exit Do
         End If
    Loop
    'Return the number
     GetRoman = num
End Function
```

