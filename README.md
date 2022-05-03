## codeを書く順番
1. 開発
2. 挿入
3. 標準モジュール
4. codeを書く
5. バーの緑右矢印を押す
6. 閉じる　確認して保存
## code

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
