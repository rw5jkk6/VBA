## 1. A列に適当に数字をいくつか入れたものをB１に合計する

```vba
Sub SumColumn()
  Dim i As Integer, cnt As Integer
  i = 1
  Do Until Cells(i, "A").Value = ""
    cnt = Cells(i, "A").Value + cnt
    i = i + 1
  Loop
  Range("B1").Value = cnt
End Sub
```
