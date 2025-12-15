VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet18"
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
End Sub
