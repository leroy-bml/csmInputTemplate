VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet49"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet, lookupWs As Worksheet
    Dim tbl As ListObject, lookupTbl As ListObject
    Dim lookupValue As Variant
    Dim i As Long
    Dim treatmentNumber As Variant, replicateNumber As Variant

    ' Set the worksheet and table
    Set ws = Me
    Set tbl = ws.ListObjects(1)
    Set lookupWs = ThisWorkbook.Sheets("OBS_DATA_LINKS")
    Set lookupTbl = lookupWs.ListObjects("obsDataLinks")

    ' Check if the changed range intersects with the table's data body range
    If Not Intersect(Target, tbl.DataBodyRange) Is Nothing Then
        ' Disable events to prevent infinite loop
        Application.EnableEvents = False

        ' Loop through each row in the table
        For i = 1 To tbl.ListRows.Count
            ' Check if experiment or observation IDs are empty
            If IsEmpty(tbl.DataBodyRange.Cells(i, tbl.ListColumns("experiment_ID").Index).Value) Or _
               IsEmpty(tbl.DataBodyRange.Cells(i, tbl.ListColumns("treatment_lev_ID").Index).Value) Then

                ' Clear autofill fields
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("treatment_number").Index).Value = ""
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("replicate_number").Index).Value = ""

            Else
                ' Define the lookup value
                lookupValue = tbl.DataBodyRange.Cells(i, tbl.ListColumns("treatment_lev_ID").Index).Value

                ' Skip if the lookupValue is empty
                If Not IsEmpty(lookupValue) Then
                    
                    ' Fetch treatment and replicate number from the lookup table
                    treatmentNumber = Application.VLookup(lookupValue, _
                                                          lookupTbl.ListColumns("observ_treat_ID").Range.Resize(, 4), _
                                                          3, _
                                                          False)
                    replicateNumber = Application.VLookup(lookupValue, _
                                                          lookupTbl.ListColumns("observ_treat_ID").Range.Resize(, 4), _
                                                          4, _
                                                          False)
                    If Not IsError(treatmentNumber) Then
                        tbl.DataBodyRange.Cells(i, tbl.ListColumns("treatment_number").Index).Value = treatmentNumber
                        tbl.DataBodyRange.Cells(i, tbl.ListColumns("replicate_number").Index).Value = replicateNumber
                    Else
                        tbl.DataBodyRange.Cells(i, tbl.ListColumns("treatment_number").Index).Value = ""
                        tbl.DataBodyRange.Cells(i, tbl.ListColumns("replicate_number").Index).Value = ""
                    End If
                End If

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


