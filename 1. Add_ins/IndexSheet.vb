Option Explicit

Sub CreateOrUpdateIndexSheet()
    Dim targetWB As Workbook
    Dim idxWS As Worksheet
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim tbl As ListObject
    Dim vNames As Variant, vDescs As Variant
    Dim lastRow As Long
    Dim i As Long
    
    On Error GoTo ErrHandler

    ' Work on active workbook (not the add-in itself)
    Set targetWB = Application.ActiveWorkbook
    If targetWB Is Nothing Or targetWB Is ThisWorkbook Then
        MsgBox "Please activate the workbook you want to index (not the add-in).", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' If Index exists, capture descriptions to preserve them
    On Error Resume Next
    Set idxWS = targetWB.Worksheets("Index")
    On Error GoTo ErrHandler

    If Not idxWS Is Nothing Then
        lastRow = idxWS.Cells(idxWS.Rows.Count, "A").End(xlUp).Row
        If lastRow >= 2 Then
            vNames = idxWS.Range("A2:A" & lastRow).Value
            vDescs = idxWS.Range("B2:B" & lastRow).Value
        End If
        
        ' remove existing tables and clear rows
        If idxWS.ListObjects.Count > 0 Then
            For i = idxWS.ListObjects.Count To 1 Step -1
                idxWS.ListObjects(i).Delete
            Next i
        End If
        idxWS.Range("A2:C" & idxWS.Rows.Count).Clear
    Else
        Set idxWS = targetWB.Worksheets.Add(Before:=targetWB.Sheets(1))
        idxWS.Name = "Index"
        idxWS.Range("A1").Value = "Sheet Name"
        idxWS.Range("B1").Value = "Description"
        idxWS.Range("C1").Value = "Last Updated"
    End If

    ' Header formatting – light gray background, bold, centered
    With idxWS.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242) ' Light gray
        .Font.Color = RGB(50, 50, 50)        ' Dark gray text
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(200, 200, 200)
    End With

    ' Fill in sheet list
    rowNum = 2
    For Each ws In targetWB.Worksheets
        If ws.Name <> idxWS.Name Then
            idxWS.Hyperlinks.Add Anchor:=idxWS.Cells(rowNum, 1), Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=ws.Name
            
            ' Restore description if it existed
            Dim foundDesc As Variant: foundDesc = ""
            If Not IsEmpty(vNames) Then
                If IsArray(vNames) Then
                    For i = LBound(vNames, 1) To UBound(vNames, 1)
                        If CStr(vNames(i, 1)) = ws.Name Then
                            foundDesc = vDescs(i, 1)
                            Exit For
                        End If
                    Next i
                Else
                    If CStr(vNames) = ws.Name Then foundDesc = vDescs
                End If
            End If
            
            idxWS.Cells(rowNum, 2).Value = foundDesc
            idxWS.Cells(rowNum, 3).Value = Format(Now, "yyyy-mm-dd hh:mm")
            rowNum = rowNum + 1
        End If
    Next ws

    ' Convert to a table with gray theme
    If rowNum > 2 Then
        Set tbl = idxWS.ListObjects.Add(SourceType:=xlSrcRange, _
            Source:=idxWS.Range("A1:C" & rowNum - 1), XlListObjectHasHeaders:=xlYes)
        On Error Resume Next
        tbl.TableStyle = "TableStyleLight1" ' Clean, light gray table
        On Error GoTo ErrHandler
    End If

    idxWS.Columns("A:C").AutoFit

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Index created/updated in workbook: " & targetWB.Name, vbInformation
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
