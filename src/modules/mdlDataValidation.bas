Attribute VB_Name = "mdlDataValidation"
Function GetValidationDictionaryFromSheet(dictSheet As Worksheet, targetSheet As Worksheet) As Object
    Dim dict As Object
    Dim lastRow As Long
    Dim headerRow As Range
    Dim i As Long
    Dim key As String, validationList As String, validationType As String, formatType As String, targetSheetName As String
    Dim varOrder As Long
    Dim colVariableName As Long, colValidationList As Long, colValidationType As Long, colFormatType As Long, colVarOrder As Long, colSheetName As Long

    Set dict = CreateObject("Scripting.Dictionary")

    ' Get the last row with data in the "Dictionary" sheet
    lastRow = dictSheet.Cells(dictSheet.Rows.Count, "A").End(xlUp).row

    ' Find the header row (assuming headers are in the first row)
    Set headerRow = dictSheet.Rows(1)
    
    ' Find column numbers for the required fields
    colVariableName = headerRow.Find(What:="var_name", LookIn:=xlValues, LookAt:=xlWhole).Column
    colValidationList = headerRow.Find(What:="validation_list", LookIn:=xlValues, LookAt:=xlWhole).Column
    colValidationType = headerRow.Find(What:="validation_type", LookIn:=xlValues, LookAt:=xlWhole).Column
    colFormatType = headerRow.Find(What:="format", LookIn:=xlValues, LookAt:=xlWhole).Column
    colVarOrder = headerRow.Find(What:="var_order_custom", LookIn:=xlValues, LookAt:=xlWhole).Column
    colSheetName = headerRow.Find(What:="sheet", LookIn:=xlValues, LookAt:=xlWhole).Column

    ' Loop through the rows to read data
    For i = 2 To lastRow ' Assuming data starts in the second row
        targetSheetName = CStr(dictSheet.Cells(i, colSheetName).Value)
        
        ' Apply the target sheet filter
        If targetSheetName = targetSheet.Name Then
            key = dictSheet.Cells(i, colVariableName).Value
            If Len(key) > 0 Then
                ' Read optional fields; use empty strings if cells are empty
                validationList = CStr(dictSheet.Cells(i, colValidationList).Value)
                validationType = CStr(dictSheet.Cells(i, colValidationType).Value)
                formatType = CStr(dictSheet.Cells(i, colFormatType).Value)
                varOrder = dictSheet.Cells(i, colVarOrder).Value
                
                ' Check if the key already exists
                If Not dict.exists(key) Then
                    dict.Add key, Array(validationList, validationType, formatType, varOrder)
                Else
                    Debug.Print "Duplicate key found and skipped: " & key
                End If
            End If
        End If
    Next i

    Set GetValidationDictionaryFromSheet = dict
End Function

Sub DisplayValidationDictionary()
    Dim dictSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim dict As Object
    Dim key As Variant
    
    ' Set the dictSheet to the "Dictionary" sheet
    Set dictSheet = ThisWorkbook.Sheets("Dictionary")
    Set targetSheet = ThisWorkbook.Sheets("FERTILIZERS")
    
    ' Get the dictionary from the sheet
    Set dict = GetValidationDictionaryFromSheet(dictSheet, targetSheet)
    
    ' Print the dictionary contents to the Immediate window
    For Each key In dict.Keys
        Debug.Print key & ": " & Join(dict(key), ", ")
    Next key
End Sub

Sub ApplyDataValidation(tbl As ListObject, headerName As String, validationRangeName As String, validationType As String, dropDownSheet As Worksheet)
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
        
        On Error Resume Next
        Set validationRange = dropDownSheet.Range(validationRangeName)
        If validationRange Is Nothing Then
            MsgBox "Named range '" & validationRangeName & "' does not exist on sheet '" & dropDownSheet.Name & "'.", vbCritical
            Exit Sub
        End If
        On Error GoTo 0
        
        
        'Set validationRange = dropDownSheet.Range(validationRangeName)
        ' Apply data validation
        If validationType = "list_strict" Then
            With dataRange.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                    xlBetween, Formula1:="=" & dropDownSheet.Name & "!" & validationRange.Address  ' Only values from the list
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
        ElseIf validationType = "list_flexible" Then
            With dataRange.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:= _
                    xlBetween, Formula1:="=" & dropDownSheet.Name & "!" & validationRange.Address  ' Allows other values
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = False
            End With
        Else
            dataRange.Validation.Delete
        End If
    End If
End Sub

Sub ApplyColumnFormats(tbl As ListObject, dict As Object)
    Dim col As ListColumn
    Dim headerName As String
    Dim validationInfo As Variant
    Dim formatType As String
    Dim targetRange As Range

    ' Loop through all columns in the table
    For Each col In tbl.ListColumns
        headerName = col.Name
        If dict.exists(headerName) Then
            ' Retrieve format type from the dictionary
            validationInfo = dict(headerName)
            formatType = validationInfo(2) ' Third element (format type)

            ' Print debugging info
            Debug.Print "Header: " & headerName & ", FormatType: " & formatType

            ' Apply formatting to the entire data body range (even if empty)
            Set targetRange = col.DataBodyRange
            If Not targetRange Is Nothing Then
                Select Case formatType
                    Case "date"
                        targetRange.NumberFormat = "dd/mm/yyyy"
                    Case "date_year"
                        targetRange.NumberFormat = "yyyy"
                    Case "numeric"
                        targetRange.NumberFormat = "0.00"
                    Case "integer"
                        targetRange.NumberFormat = "0"
                    Case "text"
                        targetRange.NumberFormat = "@"
                    Case Else
                        targetRange.NumberFormat = "@"
                End Select
            Else
                Debug.Print "Empty range for column: " & headerName
            End If
        Else
            Debug.Print "Header not found in dictionary: " & headerName
        End If
    Next col
End Sub


'Sub SetUpDataValidation(targetSheet As Worksheet)
'    Dim ws As Worksheet
'    Dim dropDownSheet As Worksheet
'    Dim tbl As ListObject
'    Dim dictSheet As Worksheet
'    Dim dict As Object
'    Dim col As ListColumn
'    Dim headerName As String
'    Dim cell As Range
'    Dim score As String
'    Dim validationInfo As Variant
'    Dim validationRangeName As String
'    Dim validationType As String
'    Dim formatType As String
'
'    ' Define the worksheets and table
'    Set ws = ActiveSheet
'    Set dropDownSheet = ThisWorkbook.Sheets("DropDown")
'
'    If ws.ListObjects.Count > 0 Then
'        Set tbl = ws.ListObjects(1)
'    Else
'        MsgBox "No tables found in the active sheet.", vbExclamation
'        Exit Sub
'    End If
'
'    Set dictSheet = ThisWorkbook.Sheets("Dictionary")
'
'    ' Get the dictionary of header names and validation ranges from the sheet
'    Set dict = GetValidationDictionaryFromSheet(dictSheet, targetSheet)
'
'    ' Loop through all columns in the table and apply data validation
'    For Each col In tbl.ListColumns
'        headerName = col.Name
'        If dict.exists(headerName) Then
'            ' Retrieve the validation information
'            validationInfo = dict(headerName)
'            ' Assign the array elements to individual variables
'            validationRangeName = validationInfo(0) '<-- First element (validation range name)
'            validationType = validationInfo(1)     '<-- Second element (validation type)
'            formatType = validationInfo(2)         '<-- Third element (format type)
'
'            ' Apply data validation
'            If validationRangeName = "none" Then
'                ' Explicitly clear validation for columns with "none"
'                col.DataBodyRange.Validation.Delete
'            Else
'                ' Apply data validation
'                ApplyDataValidation tbl, headerName, validationRangeName, validationType, dropDownSheet
'            End If
'
'    ' Apply column formats
'    ApplyColumnFormats tbl, dict
'
'        Else
'            Set cell = dictSheet.Range("C1").EntireColumn.Find(headerName)
'            If Not cell Is Nothing Then
'                score = cell.Offset(0, dictSheet.Rows(1).Find("score").Column - cell.Column).Value
'                If score <> "S" Then
'                    ' Column not found in the dictionary AND not an ID, set data validation to "Any value"
'                    With col.DataBodyRange.Validation
'                        .Delete ' Remove existing validation
'                        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop
'                        .IgnoreBlank = True
'                        .InCellDropdown = False
'                        .ShowInput = True
'                        .ShowError = True
'                    End With
'                End If
'            End If
'        End If
'    Next col
'End Sub


Sub SetUpDataValidation(targetSheet As Worksheet)
    Dim ws As Worksheet
    Dim dropDownSheet As Worksheet
    Dim tbl As ListObject
    Dim dictSheet As Worksheet
    Dim dict As Object
    Dim col As ListColumn
    Dim headerName As String
    Dim cell As Range
    Dim score As String
    Dim validationInfo As Variant
    Dim validationRangeName As String
    Dim validationType As String
    Dim formatType As String
    
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
    Set dict = GetValidationDictionaryFromSheet(dictSheet, targetSheet)
    
    ' Loop through all columns in the table and apply data validation
    For Each col In tbl.ListColumns
        headerName = col.Name
        
        ' Check if the column exists in the dictionary
        If dict.exists(headerName) Then
            ' Retrieve the validation information
            validationInfo = dict(headerName)
            validationRangeName = validationInfo(0) '<-- First element (validation range name)
            validationType = validationInfo(1)     '<-- Second element (validation type)
            formatType = validationInfo(2)         '<-- Third element (format type)
            
            ' Check if column is a structural variable (score = "S")
            Set cell = dictSheet.Range("C1").EntireColumn.Find(headerName)
            If Not cell Is Nothing Then
                score = cell.Offset(0, dictSheet.Rows(1).Find("score").Column - cell.Column).Value
                If score = "S" Then
                    ' Structural variable: Skip modification of validation
                    Debug.Print "Skipping validation changes for structural variable: " & headerName
                    ApplyColumnFormats tbl, dict
                    GoTo NextColumn
                End If
            End If
            
            ' Handle data validation for non-structural variables
            If validationRangeName = "none" Then
                ' Explicitly clear validation for columns with "none"
                col.DataBodyRange.Validation.Delete
            Else
                ' Apply data validation
                ApplyDataValidation tbl, headerName, validationRangeName, validationType, dropDownSheet
            End If
            
            ' Apply column formats
            ApplyColumnFormats tbl, dict
            
        Else
            ' Handle columns not found in the dictionary
            Debug.Print "Header not found in dictionary: " & headerName
            col.DataBodyRange.Validation.Delete ' Default to removing validation
        End If
        
NextColumn:
    Next col
End Sub

