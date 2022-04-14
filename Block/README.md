```vba
Sub Otiru()
   Dim i As Integer
   For i = 1 To 10
       With Cells(i, "B")
       .Offset(0,0).Interior.Color = RGB(0,0,255)
       .Offset(0,1).Interior.Color = RGB(0,0,255)
       .Offset(1,1).Interior.Color = RGB(0,0,255)
       .Offset(1,2).Interior.Color = RGB(0,0,255)
       End With
       
       Call Application.Wait(Now + TimeValue("00:00:01"))

       With Cells(i, "B")
       .Offset(0,0).Interior.Color = RGB(255,255,255)
       .Offset(0,1).Interior.Color = RGB(255,255,255)
       .Offset(1,1).Interior.Color = RGB(255,255,255)
       .Offset(1,2).Interior.Color = RGB(255,255,255)
       End With
   Next
End Sub
```
