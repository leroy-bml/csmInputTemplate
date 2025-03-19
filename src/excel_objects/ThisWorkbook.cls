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
    On Error Resume Next ' Prevent errors if the form isn't loaded

    ' Check if the form exists and is visible
    If VBA.UserForms.Count > 0 Then
        Dim frm As Object
        For Each frm In VBA.UserForms
            If frm.Name = "frmSelectColumns" Then
                If frm.Visible Then
                    frm.Hide
                    frm.Show
                End If
                Exit For
            End If
        Next frm
    End If

    On Error GoTo 0 ' Turn error handling back on
End Sub


Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrorHandler
'    Dim ws As Worksheet
'    Dim checkRange As Range
'    Dim cell As Range
    Dim LO As Object

'    ' Loop through all sheets whose names start with "ddl"
'    For Each ws In ThisWorkbook.Worksheets
'        If Left(ws.Name, 3) = "ddl" Then ' Check if the sheet name starts with "ddl"
'            Set checkRange = Intersect(Target, ws.Columns("A")) ' Monitor changes in column A of the sheet
'
'            If Not checkRange Is Nothing Then
'                Application.EnableEvents = False ' Prevent infinite loops
'                For Each cell In checkRange
'                    If cell.Value <> "" Then ' If column A has data
'                        ' Clear any previous spilled array results in column B and adjacent cells
'                        cell.Offset(0, 1).Resize(, 10).ClearContents ' Adjust '10' to the maximum expected spill range
'
'                        ' Apply the dynamic array formula
'                        cell.Offset(0, 1).Formula2 = "=TRANSPOSE(SORT(FILTER(tillage[tillage_level]&"" | ""&tillage[tillage_treatment_name],tillage[experiment_ID]=" & cell.Address & ")))"
'                    Else
'                        ' Clear contents if column A is empty
'                        cell.Offset(0, 1).Resize(, 10).ClearContents ' Adjust '10' to the expected spill range
'                    End If
'                Next cell
'                Application.EnableEvents = True ' Re-enable events
'            End If
'        End If
'    Next ws

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

'Private Sub Workbook_SheetCalculate(ByVal Sh As Object)
'    Dim ws As Worksheet
'    Dim spillStart As Range
'    Dim spillRange As Range
'    Dim cell As Range
'
'    ' Loop through all sheets whose names start with "ddl"
'    For Each ws In ThisWorkbook.Worksheets
'        If Left(ws.Name, 3) = "ddl" Then ' Check if the sheet name starts with "ddl"
'
'            ' Define the dynamic array starting point (A2)
'            Set spillStart = ws.Range("A2") ' Adjust this if your array starts elsewhere
'
'            ' Dynamically detect the spill range
'            On Error Resume Next
'            If spillStart.Formula2 <> "" Then
'                Set spillRange = spillStart.Resize(ws.Range(spillStart.Address).End(xlDown).Row - spillStart.Row + 1)
'            Else
'                Set spillRange = Nothing ' No dynamic array detected
'            End If
'            On Error GoTo 0
'
'            ' Process the detected spill range
'            If Not spillRange Is Nothing Then
'                Application.EnableEvents = False ' Prevent infinite loops
'
'                For Each cell In spillRange
'                    If Not IsError(cell.Value) And cell.Value <> "" Then ' Ensure valid data
'                        ' Clear any previous spill results in column B
'                        cell.Offset(0, 1).Resize(, 120).ClearContents ' Adjust '10' for max spill size
'                        cell.Offset(0, 120).Resize(, 120).ClearContents
'
'                        ' Apply the dynamic row-wise formula in column B
'                        cell.Offset(0, 1).Formula2 = "=TRANSPOSE(SORT(FILTER(tillage[tillage_level]&"" | ""&tillage[tillage_treatment_name],tillage[experiment_ID]=" & cell.Address & ")))"
'                        cell.Offset(0, 120).Formula2 = "=TRANSPOSE(SORT(FILTER(tillage[tillage_level],tillage[experiment_ID]=" & cell.Address & ")))"
'                    Else
'                        ' Clear column B if column A is empty or contains an error
'                        cell.Offset(0, 1).Resize(, 10).ClearContents ' Adjust '10' for max spill size
'                        cell.Offset(0, 120).Resize(, 120).ClearContents
'
'                    End If
'                Next cell
'
'                Application.EnableEvents = True ' Re-enable events
'            End If
'        End If
'    Next ws
'End Sub







