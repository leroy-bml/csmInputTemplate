Attribute VB_Name = "mdlDataValidation"
Function GetValidationDictionaryFromSheet(dictSheet As Worksheet) As Object
    Dim dict As Object
    Dim lastRow As Long
    Dim i As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Get the last row with data in the "Dictionary" sheet
    lastRow = dictSheet.Cells(dictSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through the rows to read var_name and validation_list
    For i = 2 To lastRow ' Assuming headers are in the first row
        If Not IsEmpty(dictSheet.Cells(i, 3)) And Not IsEmpty(dictSheet.Cells(i, 10)) Then
            ' Check if the key already exists
            If Not dict.exists(dictSheet.Cells(i, 3).Value) Then
                dict.Add CStr(dictSheet.Cells(i, 3).Value), CStr(dictSheet.Cells(i, 10).Value)
            Else
                Debug.Print "Duplicate key found and skipped: " & dictSheet.Cells(i, 3).Value
            End If
        End If
    Next i
    
    Set GetValidationDictionaryFromSheet = dict
End Function

Sub DisplayValidationDictionary()
    Dim dictSheet As Worksheet
    Dim dict As Object
    Dim key As Variant
    
    ' Set the dictSheet to the "Dictionary" sheet
    Set dictSheet = ThisWorkbook.Sheets("Dictionary")
    
    ' Get the dictionary from the sheet
    Set dict = GetValidationDictionaryFromSheet(dictSheet)
    
    ' Print the dictionary contents to the Immediate window
    For Each key In dict.Keys
        Debug.Print key & ": " & dict(key)
    Next key
End Sub

Sub ApplyDataValidation(tbl As ListObject, headerName As String, validationRangeName As String, dropDownSheet As Worksheet)
    Dim colIndex As Integer
    Dim dataRange As Range
    Dim validationRange As Range
    
    ' Find the column index by header name
    colIndex = GetColumnIndexByHeader(tbl, headerName)
    
    ' Check if the column was found
    If colIndex > 0 Then
        ' Define the range for data validation based on the column index
        Set dataRange = tbl.ListColumns(colIndex).DataBodyRange
        ' Get the validation range using its name
        Set validationRange = dropDownSheet.Range(validationRangeName)
        
        ' Debug: Print validation range address
        Debug.Print "Applying validation for header: " & headerName
        Debug.Print "Validation range address: " & validationRange.Address
        
        ' Apply data validation
        With dataRange.Validation
            .Delete ' Remove existing validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=" & dropDownSheet.Name & "!" & validationRange.Address
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    End If
End Sub

Sub SetUpDataValidation()
    Dim ws As Worksheet
    Dim dropDownSheet As Worksheet
    Dim tbl As ListObject
    Dim dictSheet As Worksheet
    Dim dict As Object
    Dim col As ListColumn
    Dim headerName As String
    Dim cell As Range
    Dim score As String
    
    ' Define the worksheets and table
    Set ws = ActiveSheet
    Set dropDownSheet = ThisWorkbook.Sheets("DropDown")
    
    If ws.ListObjects.Count > 0 Then
        Set tbl = ws.ListObjects(1)
    Else
        MsgBox "No tables found in the active sheet.", vbExclamation
        Exit Sub
    End If
    
    Set dictSheet = ThisWorkbook.Sheets("Dictionary")
    
    ' Get the dictionary of header names and validation ranges from the sheet
    Set dict = GetValidationDictionaryFromSheet(dictSheet)
    
    ' Loop through all columns in the table and apply data validation
    For Each col In tbl.ListColumns
        headerName = col.Name
        If dict.exists(headerName) Then
            ' Column found in the dictionary, apply data validation
            ApplyDataValidation tbl, headerName, dict(headerName), dropDownSheet
        'ElseIf Not (headerName Like "*_ID*" Or headerName Like "*_lev*" Or headerName Like "*_identifier" Or headerName Like "*_name") Then
        
        Else
            Set cell = dictSheet.Range("C1").EntireColumn.Find(headerName)
            If Not cell Is Nothing Then
                score = cell.Offset(0, dictSheet.Rows(1).Find("score").Column - cell.Column).Value
                If score <> "S" Then
        
                    ' Column not found in the dictionary AND not an ID, set data validation to "Any value"
                    With col.DataBodyRange.Validation
                        .Delete ' Remove existing validation
                        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop
                        .IgnoreBlank = True
                        .InCellDropdown = False
                        .ShowInput = True
                        .ShowError = True
                    End With
                End If
            End If
        End If
    Next col
End Sub


