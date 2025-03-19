VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Private Sub Worksheet_Change(ByVal Target As Range)
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim i As Long
'
'    ' Set the worksheet and table
'    On Error GoTo ErrorHandler
'    Set ws = Me
'    Set tbl = ws.ListObjects("obsDataLinks")
'
'    ' Check if the changed range intersects with the table's data body range
'    If Not Intersect(Target, tbl.DataBodyRange) Is Nothing Then
'        ' Disable events to prevent infinite loop
'        Application.EnableEvents = False
'
'        ' Loop through each row in the table, except the last row
'        For i = 1 To tbl.ListRows.Count - 1
'
'            ' Check if experiment_ID and plot_ID are empty
'            If IsEmpty(tbl.DataBodyRange.Cells(i, tbl.ListColumns("experiment_ID").Index).Value) Or _
'               IsEmpty(tbl.DataBodyRange.Cells(i, tbl.ListColumns("treatment_number").Index).Value) Then
'
'                ' Clear observ_treat_id field
'                tbl.DataBodyRange.Cells(i, tbl.ListColumns("observ_treat_id").Index).Value = ""
'                tbl.DataBodyRange.Cells(i, tbl.ListColumns("replicate_number").Index).Value = ""
'
'            Else
'                ' Autofill the observ_treat_ID column based on the formula
'                tbl.DataBodyRange.Cells(i, tbl.ListColumns("experiment_ID").Index).Value = _
'                    tbl.DataBodyRange.Cells(i, tbl.ListColumns("treatment_number").Index).Value & "_" & _
'                    tbl.DataBodyRange.Cells(i, tbl.ListColumns("replicate_number").Index).Value
'            End If
'
'        Next i
'
'    End If
'
'Cleanup:
'    ' Re-enable events
'    Application.EnableEvents = True
'    Exit Sub
'
'ErrorHandler:
'    MsgBox "An error occurred: " & Err.Description
'    Resume Cleanup
'End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim uniqueCombinations As Object
    Dim comboKey As String

    ' Set the worksheet and table
    On Error GoTo ErrorHandler
    Set ws = Me
    Set tbl = ws.ListObjects("obsDataLinks")

    ' Create a dictionary for tracking unique combinations
    Set uniqueCombinations = CreateObject("Scripting.Dictionary")

    ' Check if the changed range intersects with the table's data body range
    If Not Intersect(Target, tbl.DataBodyRange) Is Nothing Then
        ' Disable events to prevent infinite loop
        Application.EnableEvents = False

        ' Loop through each row in the table, except the last row
        For i = 1 To tbl.ListRows.Count - 1
            Dim experimentID As String
            Dim treatmentRaw As String
            Dim treatmentNumber As String
            Dim replicateNumber As Long
            Dim observUserID As String

            experimentID = tbl.DataBodyRange.Cells(i, tbl.ListColumns("experiment_ID").Index).Value
            treatmentRaw = tbl.DataBodyRange.Cells(i, tbl.ListColumns("treatment_number").Index).Value
            observUserID = tbl.DataBodyRange.Cells(i, tbl.ListColumns("observ_user_id").Index).Value

            ' Extract treatmentNumber using the LEFT and SEARCH logic
            If Not IsEmpty(treatmentRaw) And InStr(treatmentRaw, " |") > 0 Then
                treatmentNumber = Left(treatmentRaw, InStr(treatmentRaw, " |") - 1)
            Else
                treatmentNumber = "" ' Default to empty if " |" is not found
            End If

            ' Check if observ_user_id has a value
            If observUserID <> "" Then
                ' Set observ_treat_id using experiment_ID and observ_user_id
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("observ_treat_id").Index).Value = _
                    experimentID & "_" & observUserID
            ElseIf Not IsEmpty(experimentID) And Not IsEmpty(treatmentNumber) Then
                ' Create a unique key for the combination of experiment_ID and treatment_number
                comboKey = experimentID & "_" & treatmentNumber

                ' Check if the combination already exists in the dictionary
                If Not uniqueCombinations.exists(comboKey) Then
                    uniqueCombinations(comboKey) = 1 ' Start with 1 for a new combination
                Else
                    uniqueCombinations(comboKey) = uniqueCombinations(comboKey) + 1 ' Increment the count
                End If

                ' Set replicate_number and observ_treat_id
                replicateNumber = uniqueCombinations(comboKey)
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("replicate_number").Index).Value = replicateNumber
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("observ_treat_id").Index).Value = _
                    experimentID & "_" & treatmentNumber & "_" & replicateNumber
            Else
                ' Clear replicate_number and observ_treat_id if values are missing
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("replicate_number").Index).Value = ""
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("observ_treat_id").Index).Value = ""
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


