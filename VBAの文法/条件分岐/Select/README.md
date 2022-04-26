```vba
Sub A1HanteiSelect()
  Select Case Range("A1").Value
  Case Is > 5
    Range("A2").Value = "Bad"
  Case Is > 8
    Range("A2").Value = "normal"
  Case Else
    Range("A2").Value = "Good"
  End Select
End Sub
```
```vba
Sub fizzbuzzSelect()
  Dim i As Integer
  For i = 1 To 100
    Select Case True
    Case i Mod 15 = 0
      Cells(i, "A").Value = "FizzBuzz"
    Case i Mod 3 = 0
      Cells(i, "A").Value = "Fizz"
    Case i Mod 5 = 0
      Cells(i, "A").Value = "Buzz"
    Case Else
      Cells(i, "A").Value = i
    End Select
  Next
End Sub
```
