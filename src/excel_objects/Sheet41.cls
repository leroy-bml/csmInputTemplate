VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lookupWs As Worksheet
    Dim lookupTbl As ListObject
    Dim lookupValue As Variant
    Dim i As Long
    Dim treatmentNumber As Variant
    Dim combinationCounts As Object

    ' Create a dictionary to store counts of unique combinations
    Set combinationCounts = CreateObject("Scripting.Dictionary")

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
            ' Check if experiment_ID and plot_ID are empty
            If IsEmpty(tbl.DataBodyRange.Cells(i, tbl.ListColumns("experiment_ID").Index).Value) Or _
               IsEmpty(tbl.DataBodyRange.Cells(i, tbl.ListColumns("plot_ID").Index).Value) Then

                ' Clear autofill fields
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("treatment_lev_ID").Index).Value = ""
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("OID").Index).Value = ""
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("treatment_number").Index).Value = ""
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("observation_number").Index).Value = ""

            Else
                ' Autofill the treatment_lev_ID column based on the formula
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("treatment_lev_ID").Index).Value = _
                    tbl.DataBodyRange.Cells(i, tbl.ListColumns("experiment_ID").Index).Value & "_" & _
                    tbl.DataBodyRange.Cells(i, tbl.ListColumns("plot_ID").Index).Value

                ' Define the lookup value
                lookupValue = tbl.DataBodyRange.Cells(i, tbl.ListColumns("treatment_lev_ID").Index).Value

                ' Skip if the lookupValue is empty
                If Not IsEmpty(lookupValue) Then
                    ' Fetch treatment_number from the lookup table
                    treatmentNumber = Application.VLookup(lookupValue, _
                                                          lookupTbl.ListColumns("observ_treat_ID").Range.Resize(, 4), _
                                                          3, _
                                                          False)
                    If Not IsError(treatmentNumber) Then
                        tbl.DataBodyRange.Cells(i, tbl.ListColumns("treatment_number").Index).Value = treatmentNumber
                    Else
                        tbl.DataBodyRange.Cells(i, tbl.ListColumns("treatment_number").Index).Value = ""
                    End If

                    ' Increment the count for unique plot-treatment combinations
                    If combinationCounts.exists(lookupValue) Then
                        combinationCounts(lookupValue) = combinationCounts(lookupValue) + 1
                    Else
                        combinationCounts.Add lookupValue, 1
                    End If

                    ' Assign the count to the observation_number column
                    tbl.DataBodyRange.Cells(i, tbl.ListColumns("observation_number").Index).Value = combinationCounts(lookupValue)
                End If

                ' Autofill the observation ID column (OID) based on the formula
                If Not IsEmpty(treatmentNumber) And Not IsNull(treatmentNumber) Then
                    Dim treatmentStr As String
                    treatmentStr = CStr(treatmentNumber) ' Convert to string for string operations
                    If InStr(treatmentStr, " | ") > 0 Then
                        tbl.DataBodyRange.Cells(i, tbl.ListColumns("OID").Index).Value = _
                            tbl.DataBodyRange.Cells(i, tbl.ListColumns("treatment_lev_ID").Index).Value & "_" & _
                            Left(treatmentStr, InStr(treatmentStr, " | ") - 1) & "_" & _
                            combinationCounts(lookupValue)

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

