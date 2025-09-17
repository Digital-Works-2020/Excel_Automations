Sub InsertBlankRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    'Work on the active sheet
    Set ws = ActiveSheet
    
    'Find last used row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    'Loop from bottom to top
    For i = lastRow To 1 Step -1
        ws.Rows(i + 1).Insert Shift:=xlDown
    Next i
    
    MsgBox "Blank row inserted after every row.", vbInformation
End Sub
