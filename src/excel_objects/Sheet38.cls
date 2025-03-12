VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet38"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    ' Exit if the selected range contains more than one cell
    If Target.Columns.Count > 1 Then Exit Sub
    If Target.Rows.Count > 1 Then Exit Sub

    ' Check for check boxes
    If Target.Text = "þ" Then
        Application.EnableEvents = False
        Target = "¨"
        Target.Offset(0, -2).Font.Color = RGB(128, 128, 128) ' Change text color to grey
        DeleteHyperlink Target.Offset(0, -2)
        HideSheets Target.Offset(0, -1).Text
        Target.Offset(0, -2).IndentLevel = 1
        Target.Offset(0, -2).VerticalAlignment = xlCenter
        Target.Offset(0, 1).Select
        Application.EnableEvents = True
    ElseIf Target.Text = "¨" Then
        Application.EnableEvents = False
        Target = "þ"
        Target.Offset(0, -2).Font.Color = RGB(0, 0, 0) ' Change text color to black
        CreateHyperlink Target.Offset(0, -2), Target.Offset(0, -1).Text
        UnhideSheets Target.Offset(0, -1).Text
        Target.Offset(0, -2).IndentLevel = 1
        Target.Offset(0, -2).VerticalAlignment = xlCenter
        Target.Offset(0, 1).Select
        Application.EnableEvents = True
    End If
End Sub

Private Sub CreateHyperlink(cell As Range, sheetNames As String)
    Dim firstSheetName As String
    firstSheetName = Split(sheetNames, "; ")(0)
    On Error Resume Next
    cell.Hyperlinks.Delete ' Remove any existing hyperlink
    cell.Hyperlinks.Add Anchor:=cell, Address:="", SubAddress:="'" & firstSheetName & "'!A1", TextToDisplay:=cell.Text
    On Error GoTo 0
End Sub

Private Sub DeleteHyperlink(cell As Range)
    On Error Resume Next
    cell.Hyperlinks.Delete ' Remove any existing hyperlink
    On Error GoTo 0
End Sub

Private Sub HideSheets(sheetNames As String)
    Dim sheetNameArray() As String
    Dim i As Integer
    sheetNameArray = Split(sheetNames, "; ")
    On Error Resume Next
    For i = LBound(sheetNameArray) To UBound(sheetNameArray)
        ThisWorkbook.Sheets(sheetNameArray(i)).Visible = xlSheetHidden
    Next i
    On Error GoTo 0
End Sub

Private Sub UnhideSheets(sheetNames As String)
    Dim sheetNameArray() As String
    Dim i As Integer
    sheetNameArray = Split(sheetNames, "; ")
    On Error Resume Next
    For i = LBound(sheetNameArray) To UBound(sheetNameArray)
        ThisWorkbook.Sheets(sheetNameArray(i)).Visible = xlSheetVisible
    Next i
    On Error GoTo 0
End Sub
