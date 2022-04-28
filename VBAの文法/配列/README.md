## 配列の文字を各セルに表示する
```vba
Sub Hairetu()
    Dim name As Variant
    name = Array("hitoshi", "serina", "rui")
    Range("A1:C1").Value = name
End Sub
```
## 一つのセルにある文字列を各セルに分割する
- 次のをA1セルにコピーする
`hitoshi,serina,rui`
```vba
Sub StringSplit()
  Dim i As Integer
  Dim arr As Variant
  arr = Split(Range("A1").Value, ",")
  For i = 0 To UBound(arr)
    Cells(1, i + 2).Value = arr(i)
  Next
End Sub
```
