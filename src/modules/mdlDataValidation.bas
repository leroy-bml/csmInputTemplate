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
    lastRow = dictSheet.Cells(dictSheet.Rows.Count, "A").End(xlUp).Row

    ' Find the header row (assuming headers are in the first row)
    Set headerRow = dictSheet.Rows(1)
    
    ' Define lookup columns
    On Error Resume Next
    colVariableName = Application.Match("var_name", headerRow, 0)
    colValidationList = Application.Match("validation_list", headerRow, 0)
    colValidationType = Application.Match("validation_type", headerRow, 0)
    colFormatType = Application.Match("format", headerRow, 0)
    colVarOrder = Application.Match("var_order", headerRow, 0)
    colSheetName = Application.Match("sheet", headerRow, 0)
    colColType = Application.Match("column_type", headerRow, 0)
    On Error GoTo 0

    ' Check if all columns were found
    If colVariableName = 0 Or colValidationList = 0 Or colValidationType = 0 Or colFormatType = 0 Or colVarOrder = 0 Or colSheetName = 0 Or colColType = 0 Then
        MsgBox "Could not find one or more required columns in the 'Dictionary' sheet. Check: var_name, validation_list, validation_type, format, var_order, sheet, column_type", vbCritical, "Missing Columns"
        Set GetValidationDictionaryFromSheet = dict ' Return empty dictionary
        Exit Function
    End If

    ' Loop through the rows to read data
    For i = 2 To lastRow ' Data starts in the second row
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
                colType = CStr(dictSheet.Cells(i, colColType).Value)
                
                ' Check if the key already exists
                If Not dict.exists(key) Then
                    dict.Add key, Array(validationList, validationType, formatType, varOrder, colType)
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

'Sub ApplyDataValidation(tbl As ListObject, headerName As String, validationRangeName As String, validationType As String, dropDownSheet As Worksheet)
'    Dim colIndex As Integer
'    Dim dataRange As Range
'    Dim validationRange As Range
'
'    ' Find the column index by header name
'    colIndex = GetColumnIndexByHeader(tbl, headerName)
'
'    ' Check if the column was found
'    If colIndex > 0 Then
'        ' Define the range for data validation based on the column index
'        Set dataRange = tbl.ListColumns(colIndex).DataBodyRange
'        ' Get the validation range using its name
'
'        On Error Resume Next
'        Set validationRange = dropDownSheet.Range(validationRangeName)
'        If validationRange Is Nothing Then
'            MsgBox "Named range '" & validationRangeName & "' does not exist on sheet '" & dropDownSheet.Name & "'.", vbCritical
'            Exit Sub
'        End If
'        On Error GoTo 0
'
'
'        'Set validationRange = dropDownSheet.Range(validationRangeName)
'        ' Apply data validation
'        If validationType = "list_strict" Then
'            With dataRange.Validation
'                .Delete
'                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'                    xlBetween, Formula1:="=" & dropDownSheet.Name & "!" & validationRange.Address  ' Only values from the list
'                .IgnoreBlank = True
'                .InCellDropdown = True
'                .ShowInput = True
'                .ShowError = True
'            End With
'        ElseIf validationType = "list_flexible" Then
'            With dataRange.Validation
'                .Delete
'                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:= _
'                    xlBetween, Formula1:="=" & dropDownSheet.Name & "!" & validationRange.Address  ' Allows other values
'                .IgnoreBlank = True
'                .InCellDropdown = True
'                .ShowInput = True
'                .ShowError = False
'            End With
'        Else
'            dataRange.Validation.Delete
'        End If
'    End If
'End Sub

Sub ApplyDataValidation(tbl As ListObject, headerName As String, validationRangeName As String, validationType As String, dropDownSheet As Worksheet)
    Dim colIndex As Integer
    Dim dataRange As Range
    ' Dim validationRange As Range ' This is no longer needed
    
    ' Find the column index by header name
    colIndex = GetColumnIndexByHeader(tbl, headerName)
    
    ' Check if the column was found
    If colIndex > 0 Then
        ' Define the range for data validation based on the column index
        Set dataRange = tbl.ListColumns(colIndex).DataBodyRange
        
        ' --- NEW: Check if the Named Range exists anywhere in the workbook ---
        Dim nm As Name
        On Error Resume Next
        Set nm = ThisWorkbook.Names(validationRangeName)
        On Error GoTo 0
        
        If nm Is Nothing Then
            MsgBox "Data validation setup failed for column '" & headerName & "'." & vbCrLf & vbCrLf & _
                   "The Named Range '" & validationRangeName & "' does not exist.", vbCritical, "Missing Named Range"
            Exit Sub
        End If
        ' --- END NEW CHECK ---
        
        ' Apply data validation
        If validationType = "list_strict" Then
            With dataRange.Validation
                .Delete
                ' --- MODIFIED: Use the Named Range string directly ---
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                    xlBetween, Formula1:="=" & validationRangeName
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
        ElseIf validationType = "list_flexible" Then
            With dataRange.Validation
                .Delete
                ' --- MODIFIED: Use the Named Range string directly ---
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:= _
                    xlBetween, Formula1:="=" & validationRangeName
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

Sub SetUpDataValidation(targetSheet As Worksheet)
    
    Dim ws As Worksheet
    Dim dropDownSheet As Worksheet
    Dim tbl As ListObject
    Dim dictSheet As Worksheet
    Dim dict As Object
    Dim col As ListColumn
    Dim headerName As String
    Dim colType As String
    Dim validationInfo As Variant
    Dim validationRangeName As String
    Dim validationType As String
    Dim formatType As String
    
    ' Define the worksheets and table
    Set ws = targetSheet
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
    
    If dict.Count = 0 Then
        Debug.Print "Validation dictionary is empty for sheet: " & targetSheet.Name
        ' Decide if you want to exit or continue and clear formats/validation
    End If
    
    ' Loop through all columns in the table and apply data validation
    For Each col In tbl.ListColumns
        headerName = col.Name
        
        ' Check if the column exists in the dictionary
        If dict.exists(headerName) Then
            ' Retrieve the validation information
            validationInfo = dict(headerName)
            validationRangeName = validationInfo(0) '<-- First element (validation range name)
            validationType = validationInfo(1)      '<-- Second element (validation type)
            formatType = validationInfo(2)          '<-- Third element (format type)
            ' validationInfo(3) is varOrder
            colType = validationInfo(4)             '<-- Fifth element (column type)
            
            ' ####
            If colType = "fixed" Then
                ' Structural variable: Skip modification of validation
                Debug.Print "Skipping validation changes for structural variable: " & headerName
                GoTo NextColumn ' Skip to the next column
            End If
            
            ' Handle data validation for non-structural variables
            If validationRangeName = "none" Or validationRangeName = "" Then
                ' Explicitly clear validation for columns with "none" or empty
                col.DataBodyRange.Validation.Delete
            Else
                ' Apply data validation
                ApplyDataValidation tbl, headerName, validationRangeName, validationType, dropDownSheet
            End If
            
        Else
            ' Handle columns not found in the dictionary
            Debug.Print "Header not found in dictionary: " & headerName
            col.DataBodyRange.Validation.Delete ' Default to removing validation
        End If
        
NextColumn:
    Next col
    
    ' Apply formats to the entire table after all validation is set
    ApplyColumnFormats tbl, dict

End Sub

