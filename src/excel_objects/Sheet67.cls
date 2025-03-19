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
                Application.Union(validationRange, correspondingUnits), 2, False)
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

