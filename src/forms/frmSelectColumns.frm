VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectColumns 
   Caption         =   "Select columns"
   ClientHeight    =   11820
   ClientLeft      =   40
   ClientTop       =   170
   ClientWidth     =   3400
   OleObjectBlob   =   "frmSelectColumns.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
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

    Dim cell As Range
    Dim topPos As Integer
    Dim chkBox As MSForms.CheckBox
    Dim lblSection As MSForms.label
    Dim frame As MSForms.frame
    Dim currentSection As String
    Dim maxWidth As Integer
    Dim tbl As ListObject
    Dim score As String
    Dim scoreCell As Range

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
    Set frame = Me.Controls.Add("Forms.Frame.1", "frame", True)
    With frame
        .Top = lblSection.Top + lblSection.Height + 5
        .Left = 10
        .Width = Me.Width - 80 ' Adjust width to use more form space
        .Height = Me.Height - .Top - 40 ' Adjust height to fit the form and add space at the bottom
        .ScrollBars = fmScrollBarsVertical ' Enable vertical scroll bar
    End With

    topPos = 10 ' Initial top position inside the Frame

    ' Calculate the maximum width for checkboxes
    maxWidth = frame.Width - 20

    ' Get the table from the current sheet
    Set tbl = ThisWorkbook.Sheets(currentSection).ListObjects(1)

    ' Create checkboxes for the appropriate section
    For Each cell In dictSheet.Range("A2:A" & dictSheet.Cells(dictSheet.Rows.Count, "A").End(xlUp).Row)
        'If cell.Value = currentSection And Not (cell.Offset(0, 2).Value Like "*_ID*" Or cell.Offset(0, 2).Value Like "*_lev*" Or cell.Offset(0, 2).Value Like "*_identifier*" Or cell.Offset(0, 2).Value Like "*_name") Then
        If cell.Value = currentSection And cell.Offset(0, 1) <> -99 Then
            Set scoreCell = cell.Offset(0, dictSheet.Rows(1).Find("score").Column - cell.Column)
            score = scoreCell.Value
            If score <> "S" Then
                Set chkBox = frame.Controls.Add("Forms.CheckBox.1", "chk" & cell.Offset(0, 2).Value, True)
                With chkBox
                    .Caption = cell.Offset(0, 3).Value
                    .Top = topPos
                    .Left = 10
                    .AutoSize = False
                    .Width = maxWidth ' Set width to use more frame space
                    .WordWrap = False ' Prevent text from wrapping unnecessarily
                    ' Set initial state based on whether column exists in the table
                    If ColumnExistsInTable(tbl, cell.Offset(0, 2).Value) Then
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
    requiredHeight = topPos + 20

    ' Adjust the Frame ScrollHeight to fit all checkboxes and add space at the bottom
    frame.ScrollHeight = requiredHeight
    
    ' Enable scrollbars only if required height exceeds frame height
    If requiredHeight > frame.Height Then
        frame.ScrollBars = fmScrollBarsVertical
    Else
        frame.ScrollBars = fmScrollBarsNone
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

    Call UpdateTable(currentSection)

    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


