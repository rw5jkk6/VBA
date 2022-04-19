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
