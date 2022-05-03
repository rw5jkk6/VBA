- A1の数字を以下のように判定する
    - 0 < 5　　Bad 
    - 5 < 8 normal 
    - 8 < 100 Good
```vba
Sub A1Hantei()
    If Range("A1").Value < 5 Then
        Range("A2").Value = "Bad"
    ElseIf Range("A1").Value < 8 Then
        Range("A2").Value = "normal"
    Else
        Range("A2").Value = "Good"
    End If
End Sub
```
- if and for
```vba
Sub ForIf()
    Dim i As Integer
    For i = 1 To 9
        Cells(i, "A") = i
        If i Mod 2 = 0 Then
            Cells(i, "B") = "Even"
        Else
            Cells(i, "B") = "odd"
        End If
    Next
End Sub
```
