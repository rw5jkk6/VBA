### Excel関数を使う
||A|
|:-:|:-:|
|1|1|
|2|3|
|3|2|
- A列から最大の数字を見つけ出す
```vba
Sub fmax()
  MsgBox (WorksheetFunction.Max(Range("A1:A3")))
End Sub
```
