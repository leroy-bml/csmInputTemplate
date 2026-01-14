VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectColumns 
   Caption         =   "Select columns"
   ClientHeight    =   7185
   ClientLeft      =   -435
   ClientTop       =   -1710
   ClientWidth     =   7830
   OleObjectBlob   =   "frmSelectColumns.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    ' Set fixed dimensions for the UserForm
    Me.Width = 400 ' Replace with your desired width in points
    Me.Height = 600 ' Replace with your desired height in points
      
    ' Populate checkboxes dynamically with column options from the dictionary
    PopulateControls
End Sub

Private Sub UserForm_Activate()
    ' Populate controls based on the active sheet each time the form is activated
    PopulateControls
End Sub

Private Sub PopulateControls()
    ' Clear existing controls first to avoid duplicates
    ClearControls

    ' Populate checkboxes dynamically with column options from the dictionary
    Dim dictSheet As Worksheet
    Set dictSheet = ThisWorkbook.Sheets("Dictionary")
    
    ' Find column indices by name
    Dim colVarOrder As Long, colVarName As Long, colVarLabel As Long, colColType As Long
    Dim dictHeader As Range
    Set dictHeader = dictSheet.Rows(1)

    On Error Resume Next ' Temporarily suppress errors
    colVarOrder = Application.Match("var_order", dictHeader, 0)
    colVarName = Application.Match("var_name", dictHeader, 0)
    colVarLabel = Application.Match("var_label_en", dictHeader, 0)
    colColType = Application.Match("column_type", dictHeader, 0)
    On Error GoTo 0 ' Resume normal error handling

    ' ERROR HANDLING: missing columns in dictionary
    If colVarOrder = 0 Or colVarName = 0 Or colVarLabel = 0 Or colColType = 0 Then
        MsgBox "Could not find one or more required columns in the 'Dictionary' sheet: var_order, var_name, var_label_en, column_type.", vbCritical, "Missing Columns"
        Exit Sub
    End If
    
    Dim cell As Range
    Dim topPos As Integer
    Dim chkBox As MSForms.CheckBox
    Dim lblSection As MSForms.label
    Dim Frame As MSForms.Frame
    Dim currentSection As String
    Dim maxWidth As Integer
    Dim tbl As ListObject
    Dim colType As String

    ' Determine the current sheet name
    currentSection = Application.ActiveSheet.Name

    ' Set the label based on the section
    Set lblSection = Me.Controls.Add("Forms.Label.1", "lblSection", True)
    With lblSection
        .Caption = currentSection & " Columns:"
        .Top = 10
        .Left = 10
        .AutoSize = False
        .Width = Me.Width - 80 ' Adjust width to use more form space
        .WordWrap = True
    End With

    ' Add a Frame control to the UserForm
    Set Frame = Me.Controls.Add("Forms.Frame.1", "frame", True)
    With Frame
        .Top = lblSection.Top + lblSection.Height + 5
        .Left = 10
        .Width = Me.Width - 80 ' Adjust width to use more form space
        .Height = Me.Height - .Top - 40 ' Adjust height to fit the form and add space at the bottom
        .ScrollBars = fmScrollBarsVertical ' Enable vertical scroll bar
        .ScrollHeight = topPos + 30
    End With


    topPos = 10 ' Initial top position inside the Frame

    ' Calculate the maximum width for checkboxes
    maxWidth = Frame.Width - 20

    ' Get the table from the current sheet
    Set tbl = ThisWorkbook.Sheets(currentSection).ListObjects(1)

    ' Loop using column indices
    Dim varOrderValue As Variant
    Dim varNameValue As String
    
    ' Create checkboxes for the appropriate section
    ' Loop through Column A (sheet name column)
    For Each cell In dictSheet.Range("A2:A" & dictSheet.Cells(dictSheet.Rows.Count, "A").End(xlUp).Row)
        
        ' Get the value from the "var_order" column for the current row
        varOrderValue = dictSheet.Cells(cell.Row, colVarOrder).Value
        ' Filter by section (Column A) and var_order
        If cell.Value = currentSection And Not varOrderValue = -99 And Not varOrderValue = -1 Then
            ' Get column type from the "column_type" column
            colType = dictSheet.Cells(cell.Row, colColType).Value
            If colType <> "fixed" Then
                ' Get variable name from "var_name" column
                varNameValue = dictSheet.Cells(cell.Row, colVarName).Value
                Set chkBox = Frame.Controls.Add("Forms.CheckBox.1", "chk" & varNameValue, True)
                With chkBox
                    ' Get variable label from "var_label_en" column
                    .Caption = dictSheet.Cells(cell.Row, colVarLabel).Value
                    .Top = topPos
                    .Left = 10
                    .AutoSize = False
                    .Width = maxWidth ' Set width to use more frame space
                    .WordWrap = False ' Prevent text from wrapping unnecessarily
                    ' Set initial state based on whether column exists in the table
                    If ColumnExistsInTable(tbl, varNameValue) Then
                        .Value = True
                    Else
                        .Value = False
                    End If
                End With
                topPos = topPos + chkBox.Height + 2 ' Reduce vertical spacing
            End If
        End If
    Next cell

    ' Calculate the required height based on the number of checkboxes
    Dim requiredHeight As Long
    requiredHeight = topPos + 20

    ' Adjust the Frame ScrollHeight to fit all checkboxes and add space at the bottom
    Frame.ScrollHeight = requiredHeight

    ' Enable scrollbars only if required height exceeds frame height
    If requiredHeight > Frame.Height Then
        Frame.ScrollBars = fmScrollBarsVertical
    Else
        Frame.ScrollBars = fmScrollBarsNone
    End If
End Sub

Private Function ColumnExistsInTable(tbl As ListObject, colName As String) As Boolean
    Dim col As ListColumn
    ColumnExistsInTable = False
    For Each col In tbl.ListColumns
        If col.Name = colName Then
            ColumnExistsInTable = True
            Exit Function
        End If
    Next col
End Function

Private Sub ClearControls()
    ' Clear existing controls except CommandButtons
    Dim i As Integer
    For i = Me.Controls.Count - 1 To 0 Step -1
        If Not TypeName(Me.Controls(i)) = "CommandButton" Then
            Me.Controls.Remove Me.Controls(i).Name
        End If
    Next i
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.Hide
        Cancel = True
    End If
End Sub

Private Sub btnUpdate_Click()
    ' Update the tables based on selected columns
    Dim currentSection As String
    currentSection = Application.ActiveSheet.Name

    Call UpdateTable(currentSection, Me)

    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


