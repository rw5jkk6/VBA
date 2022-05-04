- for文
```vba
Sub ForCell()
　　Dim i
　　For i = 1 To 10
　　　　Cells(i, 1) = "Hello"
　　Next i
End Sub
```
- 行が9、列が10の九九の掛け算
```vba
Sub KUKU()
    Dim i As Integer, j As Integer
    For i = 1 To 9
        For j = 1 To 10
        Cells(i,j).Value = i * j
        Next
    Next
End Sub
```
