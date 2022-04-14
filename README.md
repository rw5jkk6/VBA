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
- Block
```vba
Sub block()
    With Range("B5")
        .Offset(0,0).Interior.Color = RGB(0,0,255)
        .Offset(0,1).Interior.Color = RGB(0,0,255)
        .Offset(1,1).Interior.Color = RGB(0,0,255)
        .Offset(1,2).Interior.Color = RGB(0,0,255)
    End With
End Sub
```
- 行と列を正方形にする
```vba
Sub Gyouretu_line()
    With Range("A1:U20")
        .ColumnWidth = 5
        .RowHeight = 30
    End With
End Sub
```
- 配列
```vba
Sub Hairetu()
    Dim name As Variant
    name = Array("hitoshi", "serina", "rui")
    Range("A1:C1").Value = name
End Sub
```
- 日付と時間
```vba
Sub DateTime()
    Range("A1")=Date
    Range("A2")=Time
End Sub
```
- 最終行のアクティベート
  - 事前にA列に1~1000まで出力する
  - ```vba
    Sub EndActivate()
        Range("A1").End(xlDown).Activate
    End Sub
    ```  
