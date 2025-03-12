Attribute VB_Name = "mdlControls"
Sub GoToFirstSheet()
    ' Activate the first sheet in the workbook
    ThisWorkbook.Sheets(1).Activate
End Sub

Sub GoToNextSheet()
    
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet

    If currentSheet.Index < ThisWorkbook.Sheets.Count Then
        ThisWorkbook.Sheets(currentSheet.Index + 1).Activate
    Else
        MsgBox "This is the last sheet in the workbook.", vbInformation
    End If
End Sub

Sub GoToPreviousSheet()
    
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet

    If currentSheet.Index > 1 Then
        ThisWorkbook.Sheets(currentSheet.Index - 1).Activate
    Else
        MsgBox "This is the first sheet in the workbook.", vbInformation
    End If
End Sub

Sub BulkResizeColumns()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Set column widths
    ws.Columns("A").ColumnWidth = 13
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C:Z").ColumnWidth = 20

    ' Set row height for row 1
    ws.Rows("1").RowHeight = 150

    ' Merge cells A1:H1 and align text to left and top
    With ws.Range("A1:H1")
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With
End Sub

