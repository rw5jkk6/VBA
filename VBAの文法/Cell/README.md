- Helloを出力
```vba
Sub Hello()
　　Range("A1") = "Hello"
End Sub
```
- clear
```vba
Sub Clear()
    Cells(2,1).Clear
End Sub
```
- 文字の連結
    - A1 に nakayaと入力
    - B1　に  serinaと入力
```vba
Sub Mjinorenketu()
    Dim name As String, miyoji As String
    miyoji = Range("A1").Value
    name = Range("B1").Value
    Range("C1").Value = miyoji & " " & name
End Sub
```
- A列にFizzBuzzを作ったときにする
- セルの文字を中央と右端に揃える
```vba
Sub horizon()
  Range("A1","A5").HorizontalAlignment=xlRight
  Range("A6","A10").HorizontalAlignment=xlCenter
End Sub
```
