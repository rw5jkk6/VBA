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
## おみくじ
```vba
Sub Omikuji()
    Dim card(3) As String
    Dim a As Integer
    card(1)="大吉"
    card(2)="中吉"
    card(3)="小吉"
    
    Randomize
    a = Int(Rnd * 3) + 1
    MsgBox "あなたの今日の運勢は" & card(a) & "です"
End Sub
```
