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
- 上の九九のA列の合計を一番下に表示する
```vba
Sub SUMCOLLUMA()
    Dim i As Integer, j As Integer
    i = 1
    Do Until Cells(i, "A").Value = ""
        cnt = Cells(i,"A").Value + cnt
        i = i + 1
    Loop
    Cells(i, "A").Value = cnt
End Sub
```
- J9の数字を以下のように判定する
    - 0< 50　　Bad 
    - 50 < 80 normal 
    - 80 < 100 Good
```vba
Sub J9Hantei()
    If Range("J9").Value < 50 Then
        Range("J10").Value = "Bad"
    ElseIf Range("J9").Value < 80 Then
        Range("J10").Value = "normal"
    Else
        Range("J10").Value = "Good"
    End If
End Sub
```
