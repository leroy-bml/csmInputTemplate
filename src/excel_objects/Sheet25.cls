VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long

    ' Set the worksheet and table
    On Error GoTo ErrorHandler
    Set ws = Me
    Set tbl = ws.ListObjects("obsDataLinks")

    ' Check if the changed range intersects with the table's data body range
    If Not Intersect(Target, tbl.DataBodyRange) Is Nothing Then
        ' Disable events to prevent infinite loop
        Application.EnableEvents = False

        ' Loop through each row in the table, except the last row
        For i = 1 To tbl.ListRows.Count - 1

            ' Check if experiment_ID and plot_ID are empty
            If IsEmpty(tbl.DataBodyRange.Cells(i, tbl.ListColumns("experiment_ID").Index).Value) Or _
               IsEmpty(tbl.DataBodyRange.Cells(i, tbl.ListColumns("plot_ID").Index).Value) Then

                ' Clear observ_treat_id field
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("observ_treat_id").Index).Value = ""

            Else
                ' Autofill the observ_treat_ID column based on the formula
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("observ_treat_id").Index).Value = _
                    tbl.DataBodyRange.Cells(i, tbl.ListColumns("experiment_ID").Index).Value & "_" & _
                    tbl.DataBodyRange.Cells(i, tbl.ListColumns("plot_ID").Index).Value
            End If
        
        Next i

    End If

Cleanup:
    ' Re-enable events
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    Resume Cleanup
End Sub

