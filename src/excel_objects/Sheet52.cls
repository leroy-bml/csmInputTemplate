VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Private Sub Worksheet_Calculate()
'    Dim spillStart As Range
'    Dim spillRange As Range
'    Dim cell As Range
'
'    ' Define the dynamic array starting point (A2)
'    Set spillStart = Me.Range("A2") ' Use "Me" to refer to the current worksheet
'
'    ' Dynamically detect the spill range
'    On Error Resume Next
'    If spillStart.Formula2 <> "" Then
'        Set spillRange = spillStart.Resize(Me.Range(spillStart.Address).End(xlDown).Row - spillStart.Row + 1)
'    Else
'        Set spillRange = Nothing ' No dynamic array detected
'    End If
'    On Error GoTo 0
'
'    ' Process the detected spill range
'    If Not spillRange Is Nothing Then
'        Application.EnableEvents = False ' Prevent infinite loops
'
'        For Each cell In spillRange
'            If Not IsError(cell.Value) And cell.Value <> "" Then ' Ensure valid data
'                ' Clear any previous spill results in column B and additional columns
'                cell.Offset(0, 1).Resize(, 120).ClearContents ' Clear max spill size in column B
'                cell.Offset(0, 120).Resize(, 120).ClearContents ' Clear results further away
'
'                ' Apply the dynamic row-wise formula in column B
'                cell.Offset(0, 1).Formula2 = "=TRANSPOSE(SORT(FILTER(tillage[tillage_level]&"" | ""&tillage[tillage_treatment_name],tillage[experiment_ID]=" & cell.Address & ")))"
'                cell.Offset(0, 120).Formula2 = "=TRANSPOSE(SORT(FILTER(tillage[tillage_level],tillage[experiment_ID]=" & cell.Address & ")))"
'            Else
'                ' Clear column B and extra columns if column A is empty or contains an error
'                cell.Offset(0, 1).Resize(, 120).ClearContents
'                cell.Offset(0, 120).Resize(, 120).ClearContents
'            End If
'        Next cell
'
'        Application.EnableEvents = True ' Re-enable events
'    End If
'End Sub



