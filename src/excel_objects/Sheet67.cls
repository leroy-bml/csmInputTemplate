VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet67"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    ' Check if the change was in the table
    If Me.ListObjects.Count > 0 Then
        If Not Intersect(Target, Me.ListObjects(1).DataBodyRange) Is Nothing Then
            ' Schedule the ID update to run AFTER the Workbook_Change event
            Application.OnTime Now, "UpdateSheetIDs"
        End If
    End If


    Dim lookupValue As String
    Dim autofillValue As Variant
    Dim validationRange As Range
    Dim correspondingUnits As Range
    
    ' Define the validation and units ranges by using named ranges
    Set validationRange = ThisWorkbook.Names("ddl_envmodifParameter").RefersToRange
    Set correspondingUnits = ThisWorkbook.Names("ddl_envmodifUnit").RefersToRange
    
    ' Ensure the change is in column 4 (D) of the target sheet
    If Not Intersect(Target, Me.Columns(4)) Is Nothing Then
        Application.EnableEvents = False ' Disable events to prevent infinite loop
        
        ' Loop through each changed cell in column 4
        Dim cell As Range
        For Each cell In Intersect(Target, Me.Columns(4))
            lookupValue = cell.Value ' Get the changed value in column 4
            
            ' Perform a lookup in the validation range (ddl_envmodifParameter)
            On Error Resume Next ' Handle errors if no match is found
            autofillValue = Application.VLookup(lookupValue, _
                Application.Union(validationRange, correspondingUnits), 3, False)
            On Error GoTo 0
            
            ' Populate column 7 with the corresponding unit if found
            If Not IsError(autofillValue) Then
                cell.Offset(0, 3).Value = autofillValue ' Offset(0, 3) points to column 7
            Else
                cell.Offset(0, 3).Value = ""
            End If
        Next cell
        
        Application.EnableEvents = True ' Re-enable events
    End If
End Sub

