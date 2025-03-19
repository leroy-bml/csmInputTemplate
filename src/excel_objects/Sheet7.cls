VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim tableRange As ListObject
    Dim array1 As Range, array2 As Range, array3 As Range
    Dim cell As Range
    
    ' Define your table and dynamic arrays
    Set tableRange = Me.ListObjects(1) ' Replace "YourTable" with the name of your table
    Set array1 = Worksheets("ddlIDs").Range("ddl_experiments") ' Replace "ResourcesSheet" and "Array1" accordingly
    Set array2 = Worksheets("ddlIDs").Range("ddl_soilProfiles")
    Set array3 = Worksheets("ddlIDs").Range("ddl_weatherStations")

    ' Check if the change affects the first column of the table
    If Not Intersect(Target, tableRange.ListColumns(1).DataBodyRange) Is Nothing Then
        Application.EnableEvents = False ' Prevent recursive events
        
        For Each cell In Intersect(Target, tableRange.ListColumns(1).DataBodyRange)
            If Not IsEmpty(cell.Value) Then
                If Not IsError(Application.Match(cell.Value, array1, 0)) Then
                    cell.Offset(0, 1).Value = "Experiment"
                ElseIf Not IsError(Application.Match(cell.Value, array2, 0)) Then
                    cell.Offset(0, 1).Value = "Soil data"
                ElseIf Not IsError(Application.Match(cell.Value, array3, 0)) Then
                    cell.Offset(0, 1).Value = "Weather data"
                Else
                    cell.Offset(0, 1).Value = "" ' Handle unmatched values
                End If
            End If
        Next cell
        
        Application.EnableEvents = True ' Re-enable events
    End If
End Sub

