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
- A1の数字を以下のように判定する
    - 0 < 5　　Bad 
    - 5 < 8 normal 
    - 8 < 100 Good
```vba
Sub A1Hantei()
    If Range("A1").Value < 5 Then
        Range("A2").Value = "Bad"
    ElseIf Range("A1").Value < 8 Then
        Range("A2").Value = "normal"
    Else
        Range("A2").Value = "Good"
    End If
End Sub
```
- if and for
```vba
Sub ForIf()
    Dim i As Integer
    For i = 1 To 9
        Cells(i, "A") = i
        If i Mod 2 = 0 Then
            Cells(i, "B") = "Even"
        Else
            Cells(i, "B") = "odd"
        End If
    Next
End Sub
```
- offset
```vba
Sub DataOffset()
' A１レンジの後にデータを入力する
    Range("A1") = 1
    With Range("A2")
        .Value = .Offset(-1, 0).Value + 1
        .Offset(0,1).Value = "Hitoshi"
        .Offset(0,2).Value = 3
        .Offset(0,3).Value = 5
    End With
End Sub
```
