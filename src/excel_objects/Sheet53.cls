VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet53"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim uniqueCombinations As Object
    Dim comboKey As String

    ' Set the worksheet and table
    On Error GoTo ErrorHandler
    Set ws = Me
    Set tbl = ws.ListObjects("plantData")

    ' Create a dictionary for tracking unique combinations
    Set uniqueCombinations = CreateObject("Scripting.Dictionary")

    ' Check if the changed range intersects with the table's data body range
    If Not Intersect(Target, tbl.DataBodyRange) Is Nothing Then
        ' Disable events to prevent infinite loop
        Application.EnableEvents = False

        ' Loop through each row in the table, except the last row
        For i = 1 To tbl.ListRows.Count - 1
            Dim experimentID As String
            Dim gsID As String
            Dim observUnitID As String
            Dim sampleNumber As Long

            experimentID = tbl.DataBodyRange.Cells(i, tbl.ListColumns("experiment_id").Index).Value
            gsID = tbl.DataBodyRange.Cells(i, tbl.ListColumns("crop_season_id").Index).Value
            observUnitID = tbl.DataBodyRange.Cells(i, tbl.ListColumns("obs_unit_id").Index).Value

            If Not IsEmpty(experimentID) And Not IsEmpty(gsID) And Not IsEmpty(observUnitID) Then
                ' Create a unique key for the combination of experiment_ID and treatment_number
                comboKey = experimentID & "_" & gsID & "_" & observUnitID

                ' Check if the combination already exists in the dictionary
                If Not uniqueCombinations.exists(comboKey) Then
                    uniqueCombinations(comboKey) = 1 ' Start with 1 for a new combination
                Else
                    uniqueCombinations(comboKey) = uniqueCombinations(comboKey) + 1 ' Increment the count
                End If

                ' Set sample_number
                sampleNumber = uniqueCombinations(comboKey)
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("sample_number").Index).Value = "So" & sampleNumber
            Else
                ' Clear sample_number if values are missing
                tbl.DataBodyRange.Cells(i, tbl.ListColumns("sample_number").Index).Value = ""
            End If
        Next i
    End If
    
    ' Check if the change was in the table
    If Me.ListObjects.Count > 0 Then
        If Not Intersect(Target, Me.ListObjects(1).DataBodyRange) Is Nothing Then
            ' Schedule the ID update to run AFTER the Workbook_Change event
            Application.OnTime Now, "UpdateSheetIDs"
        End If
    End If

CleanUp:
    ' Re-enable events
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    Resume CleanUp
End Sub


