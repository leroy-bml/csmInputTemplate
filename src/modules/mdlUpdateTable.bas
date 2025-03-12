Attribute VB_Name = "mdlUpdateTable"
Sub ShowColumnSelector()
    frmSelectColumns.Show
End Sub

'Sub UpdateTable(sheetName As String)
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim chkBox As MSForms.CheckBox
'    Dim ctrl As Control
'    Dim colName As String
'    Dim colLabel As String
'    Dim dictSheet As Worksheet
'    Dim headerRange As Range
'    Dim lblRange As Range
'    Dim headers As Collection
'    Dim labels As Collection
'    Dim i As Integer
'    Dim cell As Range
'    Dim score As String
'    Dim dictRange As Range
'    Dim filteredRange As Range
'    Dim lastRow As Long
'
'    ' Disable events to prevent interference
'    Application.EnableEvents = False
'
'    Set ws = ThisWorkbook.Sheets(sheetName)
'    Set tbl = ws.ListObjects(1)
'    Set dictSheet = ThisWorkbook.Sheets("Dictionary")
'    Set headers = New Collection
'    Set labels = New Collection
'
'    ' Filter the dictionary sheet to only include rows for the current sheet
'    lastRow = dictSheet.Cells(dictSheet.Rows.Count, "A").End(xlUp).Row
'    Set dictRange = dictSheet.Range("A1:D" & lastRow)
'    dictRange.AutoFilter Field:=1, Criteria1:=sheetName
'    Set filteredRange = dictSheet.Range("C2:D" & lastRow).SpecialCells(xlCellTypeVisible)
'
'    ' Debugging to verify filter
'    Debug.Print "Filtered Range Address: " & filteredRange.Address
'
'    ' Gather selected headers and their labels
'    Debug.Print "Gathering selected headers and labels"
'    For Each ctrl In frmSelectColumns.Controls
'        If TypeName(ctrl) = "CheckBox" Then
'            Set chkBox = ctrl
'            colName = Replace(chkBox.Name, "chk", "")
'
'            ' Enhanced debugging around VLookup
'            On Error Resume Next
'            colLabel = Application.WorksheetFunction.VLookup(colName, filteredRange, 2, False)
'            If Err.Number <> 0 Then
'                Debug.Print "VLookup Error for column: " & colName & " | Error: " & Err.Description
'                colLabel = "Error"
'            Else
'                Debug.Print "Column Name: " & colName & " | Lookup Result: " & colLabel
'            End If
'            On Error GoTo 0
'
'            If chkBox.Value = True Then
'                headers.Add colName
'                labels.Add colLabel
'                Debug.Print "Selected Column: " & colName & " | Label: " & colLabel
'            End If
'        End If
'    Next ctrl
'
'    ' Clear existing columns except ID columns
'    Debug.Print "Clearing existing columns"
'    For i = tbl.ListColumns.Count To 1 Step -1
'        Set cell = filteredRange.Columns(1).Find(tbl.ListColumns(i).Name)
'        If Not cell Is Nothing Then
'            score = cell.Offset(0, dictSheet.Rows(1).Find("score").Column - cell.Column).Value
'            Debug.Print "Column: " & tbl.ListColumns(i).Name & " | Score: " & score
'            If score <> "S" Then
'                Set lblRange = ws.Cells(tbl.HeaderRowRange.Row - 1, tbl.ListColumns(i).Range.Column)
'                lblRange.ClearContents
'                tbl.ListColumns(i).Delete
'                'Debug.Print "Deleted Column: " & tbl.ListColumns(i).Name
'            End If
'        Else
'            Debug.Print "Column not found in Dictionary: " & tbl.ListColumns(i).Name
'        End If
'    Next i
'
'    ' Add selected columns and their labels
'    Debug.Print "Adding selected columns and their labels"
'    For i = 1 To headers.Count
'        tbl.ListColumns.Add.Name = headers(i)
'        Set headerRange = tbl.HeaderRowRange
'        ws.Cells(headerRange.Row - 1, headerRange.Cells(1, tbl.ListColumns(headers(i)).Index).Column).Value = labels(i)
'        Debug.Print "Added Column: " & headers(i) & " | Label: " & labels(i)
'    Next i
'
'    ' Call SetUpDataValidation to apply data validation
'    SetUpDataValidation
'
'    ' Re-enable events
'    Application.EnableEvents = True
'    Debug.Print "UpdateTable completed"
'
'    ' Remove filter
'    dictSheet.AutoFilterMode = False
'End Sub

Sub UpdateTable(sheetName As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim chkBox As MSForms.CheckBox
    Dim ctrl As Control
    Dim colName As String
    Dim colLabel As String
    Dim dictSheet As Worksheet
    Dim headerRange As Range
    Dim lblRange As Range
    Dim headers As Collection
    Dim labels As Collection
    Dim existingHeaders As Collection
    Dim selectedHeaders As Collection
    Dim i As Integer
    Dim cell As Range
    Dim dictRange As Range
    Dim filteredRange As Range
    Dim lastRow As Long
    Dim score As String
    Dim isMatch As Boolean
    
    ' Disable events to prevent interference
    Application.EnableEvents = False
    
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set tbl = ws.ListObjects(1)
    Set dictSheet = ThisWorkbook.Sheets("Dictionary")
    Set headers = New Collection
    Set labels = New Collection
    Set existingHeaders = New Collection
    Set selectedHeaders = New Collection

    ' Gather existing headers
    For Each cell In tbl.HeaderRowRange
        existingHeaders.Add cell.Value
    Next cell

    ' Filter the dictionary sheet to only include rows for the current sheet
    lastRow = dictSheet.Cells(dictSheet.Rows.Count, "A").End(xlUp).Row
    Set dictRange = dictSheet.Range("A1:D" & lastRow)
    dictRange.AutoFilter Field:=1, Criteria1:=sheetName
    Set filteredRange = dictSheet.Range("C2:D" & lastRow).SpecialCells(xlCellTypeVisible)
    
    ' Gather selected headers and their labels
    For Each ctrl In frmSelectColumns.Controls
        If TypeName(ctrl) = "CheckBox" Then
            Set chkBox = ctrl
            colName = Replace(chkBox.Name, "chk", "")
            colLabel = Application.WorksheetFunction.VLookup(colName, filteredRange, 2, False)
            If chkBox.Value = True Then
                selectedHeaders.Add colName
                headers.Add colName
                labels.Add colLabel
            End If
        End If
    Next ctrl

    ' Delete only columns that are not selected and are not ID columns
    For i = tbl.ListColumns.Count To 1 Step -1
        colName = tbl.ListColumns(i).Name
        Set cell = filteredRange.Columns(1).Find(colName)
        If Not cell Is Nothing Then
            score = cell.Offset(0, dictSheet.Rows(1).Find("score").Column - cell.Column).Value
            If score = "S" Then
                ' Skip deletion for ID columns
                Debug.Print "Preserved ID Column: " & colName
                GoTo SkipDeletion
            End If
        End If
        
        ' Check if the column exists in the selectedHeaders collection manually
        isMatch = False
        For Each col In selectedHeaders
            If col = colName Then
                isMatch = True
                Exit For
            End If
        Next col
        
        If Not isMatch Then
            Set lblRange = ws.Cells(tbl.HeaderRowRange.Row - 1, tbl.ListColumns(i).Range.Column)
            lblRange.ClearContents
            tbl.ListColumns(i).Delete
            Debug.Print "Deleted Column: " & colName
        End If
SkipDeletion:
    Next i

    ' Add only new columns and update labels for existing columns
    For i = 1 To headers.Count
        colName = headers(i)
        colLabel = labels(i)
        isMatch = False
        
        ' Check if the column exists in the existingHeaders collection manually
        For Each col In existingHeaders
            If col = colName Then
                isMatch = True
                Exit For
            End If
        Next col
        
        If Not isMatch Then
            tbl.ListColumns.Add.Name = colName
            Debug.Print "Added New Column: " & colName
        Else
            Debug.Print "Column Already Exists, Preserved: " & colName
        End If
        Set headerRange = tbl.HeaderRowRange
        ws.Cells(headerRange.Row - 1, headerRange.Cells(1, tbl.ListColumns(colName).Index).Column).Value = colLabel
    Next i

    ' Re-enable events
    Application.EnableEvents = True

    ' Remove filter
    dictSheet.AutoFilterMode = False
End Sub

Function GetColumnIndexByHeader(tbl As ListObject, headerName As String) As Integer
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If col.Name = headerName Then
            GetColumnIndexByHeader = col.Index
            Exit Function
        End If
    Next col
    ' Return 0 if header name not found
    GetColumnIndexByHeader = 0
End Function

