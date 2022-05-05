### コードは上から順に実行される
```vba
Sub StringVariant()
  Dim a As Integer
  MsgBox (a)
  a = 3
End Sub
```
### 変数と文字列
```vba
Sub StringVariant()
  Dim a As Integer
  a = 3
  MsgBox (a)
  MsgBox ("a")
End Sub
```
