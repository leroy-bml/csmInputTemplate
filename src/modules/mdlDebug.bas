Attribute VB_Name = "mdlDebug"
Sub TestVLookup()
    Dim dictSheet As Worksheet
    Dim colName As String
    Dim colLabel As Variant
    
    Set dictSheet = ThisWorkbook.Sheets("Dictionary")
    
    ' Change "harvest_yld_matur_dry_wt" to any column name you are testing
    colName = "durat_summarization_per"
    
    On Error Resume Next
    colLabel = Application.WorksheetFunction.VLookup(colName, dictSheet.Range("C2:D" & dictSheet.Cells(dictSheet.Rows.Count, "C").End(xlUp).Row), 2, False)
    On Error GoTo 0
    
    Debug.Print "Column Name: " & colName & " | Lookup Result: " & colLabel
End Sub

