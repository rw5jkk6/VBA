## codeを書く順番
1. 開発
2. 挿入
3. 標準モジュール
4. codeを書く
5. バーの緑右矢印を押す
6. 閉じる　確認して保存
## code
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
- for文
```vba
Sub ForCell()
　　Dim i
　　For i = 1 To 10
　　　　Cells(i, 1) = "Hello"
　　Next i
End Sub
```
- for文のネスト
```vba
Sub ForNesuto()
    Dim i As Integer, j As Integer
    For i = 1 To 3
        For j = 1 To 2
        Cells(i,j).Value = i * j
        Next
    Next
End Sub
```
- 掛け算
```vba
Sub Kakeru()
    Dim i As Integer
    i = 1
    Do While Cells(i, "B").Value <> ""
        Cells(i, "C").Value = Cells(i, "A").Value * Cells(i,"B").Value
        i = i + 1
    Loop
End Sub
```
