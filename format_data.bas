Attribute VB_Name = "Module1"
Sub format_data()

'  Ctrl + Shift +D

    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.ActiveSheet
    
    If ws.Range("A1").Value = "" Then
        ws.Columns("A").Delete Shift:=xlToLeft
    End If
    
    ws.Columns("A:E").EntireColumn.AutoFit
    
    With ws.Range("A:E")
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Range("A1:E1")
        .Interior.Pattern = xlSolid
        .Interior.ThemeColor = xlThemeColorAccent2
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
    End With
    
    With ws.Range("A1", ws.Cells(ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row, _
                                  ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = 0
    End With
End Sub


