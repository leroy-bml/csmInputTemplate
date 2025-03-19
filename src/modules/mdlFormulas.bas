Attribute VB_Name = "mdlFormulas"
'Sub DynamicFormulaFill()
'    Dim lastCol As Long
'    Dim ws As Worksheet
'    Dim col As Long
'    Dim formulaText As String
'
'    ' Use the active sheet or the one passed by the event
'    Set ws = Application.ActiveSheet
'
'    ' Find the last column in row 2 with data
'    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
'
'    ' Loop through columns and apply the formula
'    For col = 1 To lastCol
'        ' Build the formula as a string
'        formulaText = "=SORT(FILTER(tillage[tillage_level]&"" | ""&tillage[tillage_treatment_name], tillage[experiment_ID]=INDEX($A$2:$Z$2, 1, COLUMN())))"
'
'        ' Assign the formula to the cell
'        On Error Resume Next ' Prevent issues from stopping the loop
'        ws.Cells(3, col).Formula = formulaText
'        If Err.Number <> 0 Then
'            Debug.Print "Error " & Err.Number & ": " & Err.Description
'            MsgBox "Failed to apply formula in column " & col & ". Check the formula or references.", vbCritical
'            Exit Sub
'        End If
'        On Error GoTo 0
'    Next col
'End Sub


