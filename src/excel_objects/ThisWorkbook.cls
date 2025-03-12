VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If frmSelectColumns.Visible Then
        frmSelectColumns.Hide
        frmSelectColumns.Show
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrorHandler
    Dim LO As Object

    ' Ensure the worksheet contains a table object
    If Sh.ListObjects.Count = 0 Then
        Exit Sub
    End If

    ' Set LO to the table
    Set LO = Sh.ListObjects(1)

    ' Check if DataBodyRange exists and is not Nothing
    If LO.DataBodyRange Is Nothing Then
        ' Add the first row if the table is empty
        LO.ListRows.Add
    Else
        With LO.DataBodyRange
            ' Check if Target intersects with DataBodyRange
            If Not Intersect(Target, .Cells) Is Nothing Then
                Application.EnableEvents = False

                ' Debug: Log the count of blank cells in the last row
                Debug.Print "CountBlank in last row: " & WorksheetFunction.CountBlank(.Rows(.Rows.Count))

                If WorksheetFunction.CountBlank(.Rows(.Rows.Count)) < .Columns.Count Then
                    LO.ListRows.Add
                Else
                    Do While WorksheetFunction.CountBlank(.Rows(.Rows.Count - 1)) = .Columns.Count
                        LO.ListRows(.Rows.Count).Delete
                    Loop
                End If

                Application.EnableEvents = True
            End If
        End With
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Error " & Err.Number & ": " & Err.Description
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Application.EnableEvents = True
End Sub



'Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'    On Error GoTo ErrorHandler
'    Dim LO As ListObject
'    Dim lastRow As Range
'    Dim nonEmptyColumnsCount As Long
'
'    ' Ensure the worksheet contains a table object
'    If Sh.ListObjects.Count = 0 Then
'        Exit Sub
'    End If
'
'    ' Set LO to the table
'    Set LO = Sh.ListObjects("obsSummaryGrowth")
'
'    ' Check if DataBodyRange exists and is not Nothing
'    If LO.DataBodyRange Is Nothing Then
'        ' Add the first row if the table is empty
'        LO.ListRows.Add
'    Else
'        With LO.DataBodyRange
'            ' Check if Target intersects with DataBodyRange
'            If Not Intersect(Target, .Cells) Is Nothing Then
'                Application.EnableEvents = False
'
'                ' Get the last row
'                Set lastRow = .Rows(.Rows.Count)
'
'                ' Count non-empty cells, excluding the combined_key column if it exists
'                On Error Resume Next
'                nonEmptyColumnsCount = Application.WorksheetFunction.CountA(lastRow) - _
'                                       Application.WorksheetFunction.CountBlank(LO.ListColumns("combined_key").DataBodyRange.Cells(lastRow.Row - .Row + 1))
'                On Error GoTo 0
'
'                ' Debug: Log the count of non-empty cells in the last row
'                Debug.Print "Non-empty cells in last row (excluding combined_key): " & nonEmptyColumnsCount
'
'                If nonEmptyColumnsCount > 0 Then
'                    LO.ListRows.Add
'                Else
'                    Do While WorksheetFunction.CountBlank(.Rows(.Rows.Count - 1)) = .Columns.Count
'                        LO.ListRows(.Rows.Count).Delete
'                    Loop
'                End If
'
'                Application.EnableEvents = True
'            End If
'        End With
'    End If
'
'    Exit Sub
'
'ErrorHandler:
'    Debug.Print "Error " & Err.Number & ": " & Err.Description
'    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
'    Application.EnableEvents = True
'End Sub
'
'

'Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'    MsgBox "Change detected in worksheet"
'End Sub


'Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'    On Error GoTo ErrorHandler
'    Dim LO As ListObject
'
'    ' Ensure the worksheet contains a table object
'    If Sh.ListObjects.Count = 0 Then
'        Exit Sub
'    End If
'
'    ' Set LO to the first table
'    Set LO = Sh.ListObjects(1)
'
'    ' Check if DataBodyRange exists and is not Nothing
'    If Not LO.DataBodyRange Is Nothing Then
'        With LO.DataBodyRange
'            ' Check if Target intersects with DataBodyRange
'            If Not Intersect(Target, .Cells) Is Nothing Then
'                Application.EnableEvents = False
'
'                ' Add a new row to the table
'                LO.ListRows.Add
'
'                Application.EnableEvents = True
'            End If
'        End With
'    End If
'
'    Exit Sub
'
'ErrorHandler:
'    Debug.Print "Error " & Err.Number & ": " & Err.Description
'    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
'    Application.EnableEvents = True
'End Sub

