VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ExitHandler

    Application.EnableEvents = False ' Disable events to prevent recursive triggering

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("EXP_METADATA")

    Dim tbl As ListObject
    Set tbl = ws.ListObjects(1)

    Dim colExperimentID As ListColumn
    Dim colRevisionFirst As ListColumn
    Dim colRevisionLatest As ListColumn

    ' Dynamically find the columns based on header names
    Set colExperimentID = tbl.ListColumns("experiment_ID")
    Set colRevisionFirst = tbl.ListColumns("revision_date_first")
    Set colRevisionLatest = tbl.ListColumns("revision_date_latest")

    ' Exit if Target includes the table's Insert Row
    If Not tbl.InsertRowRange Is Nothing Then
        If Not Intersect(Target, tbl.InsertRowRange) Is Nothing Then
            GoTo ExitHandler ' Skip further code if interacting with the Insert Row
        End If
    End If

    ' Ignore changes to the revision_date_latest column itself
    If Not Intersect(Target, colRevisionLatest.DataBodyRange) Is Nothing Then
        GoTo ExitHandler ' Do nothing if the change is in the revision_date_latest column
    End If

    ' Autofill revision_date_first for newly added rows
    Dim newRow As ListRow
    For Each newRow In tbl.ListRows
        If Not Intersect(Target, newRow.Range) Is Nothing Then
            If newRow.Range.Cells(1, colExperimentID.Index).Value <> "" And IsEmpty(newRow.Range.Cells(1, colRevisionFirst.Index).Value) Then
                newRow.Range.Cells(1, colRevisionFirst.Index).Value = Date
            End If
        End If
    Next newRow

    ' Detect row deletions by checking if rows in the table are blank
    If Intersect(Target, tbl.DataBodyRange) Is Nothing Then
        Dim rowRange As Range
        For Each rowRange In tbl.DataBodyRange.Rows
            If Application.CountA(rowRange) = 0 Then ' Row is empty (deleted)
                rowRange.Cells(1, colRevisionLatest.Index).ClearContents
            End If
        Next rowRange
        GoTo ExitHandler ' Stop further execution after handling row deletions
    End If

    ' Handle changes to the content of the table
    Dim changedRow As ListRow
    For Each changedRow In tbl.ListRows
        If Not Intersect(Target, changedRow.Range) Is Nothing Then
            ' Update revision_date_latest if content in the row changes
            changedRow.Range.Cells(1, colRevisionLatest.Index).Value = Date
        End If
    Next changedRow

ExitHandler:
    Application.EnableEvents = True ' Re-enable events before exiting
    Exit Sub
End Sub

