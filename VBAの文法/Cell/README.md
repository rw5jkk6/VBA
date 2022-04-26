- A列にFizzBuzzを作ったときにする
- セルの文字を中央と右端に揃える
```vba
Sub horizon()
  Range("A1","A5").HorizontalAlignment=xlRight
  Range("A6","A10").HorizontalAlignment=xlCenter
End Sub
```
