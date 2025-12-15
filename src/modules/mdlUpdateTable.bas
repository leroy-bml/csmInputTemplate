Attribute VB_Name = "mdlUpdateTable"
Sub ShowColumnSelector()
    frmSelectColumns.Show
End Sub

Public Sub UpdateSheetIDs()
    Dim LO As ListObject
    Dim Sh As Worksheet
    Dim idName As String
    
    On Error GoTo ErrorHandler
    
    ' This macro runs AFTER the change, so ActiveSheet is correct
    Set Sh = ActiveSheet
    
    ' Find the table on this sheet
    If Sh.ListObjects.Count = 0 Then Exit Sub
    Set LO = Sh.ListObjects(1)
    
    ' Find the correct ID column name based on the table's name
    Select Case LO.Name
        Case "institutions": idName = "institute_id"
        Case "persons": idName = "people_level"
        Case "documents": idName = "document_id"
        Case "fields": idName = "field_level"
        Case "plotLayout": idName = "plot_layout_id"
        Case "treatments": idName = "treatment_number"
        Case "plots": idName = "plot_id"
        Case "growingSeason": idName = "crop_season_id"
        Case "plotSetup": idName = "plot_setup_id"
        Case "genotypes": idName = "genotype_level"
        Case "cropResidues": idName = "initial_conditions_level"
        Case "tillage": idName = "tillage_level"
        Case "tillageEvents": idName = "tillage_event_id"
        Case "plantings": idName = "planting_level"
        Case "irrigation": idName = "irrigation_level"
        Case "irrigationEvents": idName = "irrigation_event_id"
        Case "organicMaterials": idName = "org_materials_applic_lev"
        Case "organicMaterialApplications": idName = "organic_material_event_id"
        Case "fertilizers": idName = "fertilizer_level"
        Case "fertilizersEvents": idName = "fertilizer_event_id"
        Case "chemicals": idName = "chemical_applic_level"
        Case "chemicalApplications": idName = "chemical_event_id"
        Case "harvest": idName = "harvest_operations_level"
        Case "harvestEvents": idName = "harvest_event_id"
        Case "mulches": idName = "mulch_level"
'        Case "soilAnalyses": idName = "soil_analysis_level"  'xxx
'        Case "soilAnalysesLayers": idName = "soil_analysis_layers_id"  'xxx
        Case "environModifications": idName = "environmental_modif_lev"
        Case "environModifLevels": idName = "environmental_modif_levels_id"
        
        ' Weather data
        Case "weatherDaily": idName = "weather_daily_record_id"
        Case "weatherMonthly": idName = "weather_monthly_record_id"
        
        ' Soil data
        Case "soilData": idName = "soil_data_id"
        Case "soilLyrCarbon": idName = "soil_layers_carbon_id"
        Case "soilLyrPhys": idName = "soil_layers_phys_id"
        Case "soilLyrChem": idName = "soil_layers_chem_id"
        Case "soilLyrNPK": idName = "soil_layers_npk_id"
        Case "soilLyrNutrients": idName = "soil_layers_nutrients_id"
        Case "soilLyrRoots": idName = "soil_layers_roots_id"
        Case "soilLyrTemp": idName = "soil_layers_temp_id"
        Case "soilLyrWater": idName = "soil_layers_water_id"
        Case "soilLyrNdyn": idName = "soil_layers_ndyn_id"
        Case "soilLyrPdyn": idName = "soil_layers_pdyn_id"
        Case "soilLyrOM": idName = "soil_layers_om_id"
        
        ' Plant data
        Case "plantData": idName = "plant_data_id"
        Case "plantSmDev": idName = "plant_sm_development_id"
        Case "plantSmGrowth": idName = "plant_sm_growth_id"
        Case "plantSmNutrients": idName = "plant_sm_nutrients_id"
        Case "plantTsGrowthDev": idName = "plant_ts_growthdev_id"
        Case "plantTsNutrients": idName = "plant_ts_nutrients_id"
        Case "plantTsPnD": idName = "plant_ts_pnd_id"
        
        ' Plant data
        Case "envData": idName = "env_data_id"
        Case "envSmWaterDynamics": idName = "env_sm_water_id"
        Case "envTsWaterDynamics": idName = "env_ts_water_id"
        Case "envTsPlanSoilAtmos": idName = "env_ts_psa_id"
        
        ' Add all other table names and their ID columns here
        
        Case Else: Exit Sub ' Not a table we want to manage
    End Select
    
    ' --- ID Generation Logic ---
    Dim idColumn As ListColumn
    Dim idDataCells As Range
    Dim i As Long
    
    On Error Resume Next ' In case ID column doesn't exist
    Set idColumn = LO.ListColumns(idName)
    On Error GoTo ErrorHandler ' Reset
    
    ' Exit if the specified ID column wasn't found
    If idColumn Is Nothing Then Exit Sub
    
    Set idDataCells = idColumn.DataBodyRange
    If idDataCells Is Nothing Then Exit Sub ' Exit if table is empty
    
    ' Disable events to prevent this change from re-triggering anything
    Application.EnableEvents = False
    
    ' Check for "experiment_id" column to determine logic
    Dim expColumn As ListColumn
    Dim expDataCells As Range
    
    On Error Resume Next
    Set expColumn = LO.ListColumns("experiment_id")
    On Error GoTo 0
    
    If expColumn Is Nothing Then
        ' --- STANDARD LOGIC ---
        ' "experiment_id" column NOT found. Use simple 1, 2, 3... count.
        For i = 1 To idDataCells.Rows.Count - 1
            idDataCells.Cells(i, 1).Value = i
        Next i
        
    Else
        ' --- GROUPED LOGIC ---
        ' "experiment_id" column WAS found. Use grouped running count.
        Dim dict As Object
        Dim expID As String
        Dim currentCount As Long
        
        Set dict = CreateObject("Scripting.Dictionary")
        Set expDataCells = expColumn.DataBodyRange
        
        For i = 1 To idDataCells.Rows.Count - 1 ' Loop all rows except the last
            expID = CStr(expDataCells.Cells(i, 1).Value)
            
            ' Increment the count for this specific experiment_id
            ' This adds the key with a value of 1 if it's new
            currentCount = dict(expID) + 1
            dict(expID) = currentCount
            
            ' Write the new count to the ID cell
            idDataCells.Cells(i, 1).Value = currentCount
        Next i
    End If
    ' --- END: Logic Switch ---

    ' Ensure the new blank row's ID cell is always clear
    idDataCells.Cells(idDataCells.Rows.Count, 1).ClearContents
    
    Application.EnableEvents = True
    
    Exit Sub

ErrorHandler:
    MsgBox "Error in UpdateSheetIDs: " & Err.Description
    Application.EnableEvents = True
End Sub


'Public Sub UpdateSheetIDs()
'    Dim LO As ListObject
'    Dim Sh As Worksheet
'    Dim idName As String
'
'    On Error GoTo ErrorHandler
'
'    ' This macro runs AFTER the change, so ActiveSheet is correct
'    Set Sh = ActiveSheet
'
'    ' Find the table on this sheet
'    If Sh.ListObjects.Count = 0 Then Exit Sub
'    Set LO = Sh.ListObjects(1)
'
'    ' --- CONFIGURE YOUR IDs HERE ---
'    ' Find the correct ID column name based on the table's name
'    Select Case LO.Name
'        Case "institutions": idName = "institute_id"
'        Case "persons": idName = "people_level"
'        Case "documents": idName = "document_id"
'        Case "fields": idName = "field_level"
'        Case "treatments": idName = "treatment_number"
'        ' Add all other table names and their ID columns here
'
'        Case Else: Exit Sub ' Not a table we want to manage
'    End Select
'
'    ' --- ID Generation Logic ---
'    Dim idColumn As ListColumn
'    Dim dataCells As Range
'    Dim i As Long
'
'    On Error Resume Next ' In case column doesn't exist
'    Set idColumn = LO.ListColumns(idName)
'    On Error GoTo ErrorHandler ' Reset
'
'    If Not idColumn Is Nothing Then
'        ' Disable events to prevent this change from re-triggering anything
'        Application.EnableEvents = False
'
'        Set dataCells = idColumn.DataBodyRange
'        If Not dataCells Is Nothing Then
'            ' Loop through all rows EXCEPT the last (blank) one
'            For i = 1 To dataCells.Rows.Count - 1
'                dataCells.Cells(i, 1).Value = i
'            Next i
'
'            ' Ensure the new blank row's ID cell is clear
'            dataCells.Cells(dataCells.Rows.Count, 1).ClearContents
'        End If
'
'        Application.EnableEvents = True
'    End If
'
'    Exit Sub
'
'ErrorHandler:
'    MsgBox "Error in UpdateSheetIDs: " & Err.Description
'    Application.EnableEvents = True
'End Sub


Sub UpdateTable(sheetName As String, frm As Object)

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
    Dim filteredVarNameCol As Range
    Dim lastRow As Long
    Dim colType As String
    Dim isMatch As Boolean
    
    ' Find column indices by name
    Dim colVarName As Long, colVarLabel As Long, colColType As Long
    Dim dictHeader As Range
    
    Set dictSheet = ThisWorkbook.Sheets("Dictionary")
    Set dictHeader = dictSheet.Rows(1)

    On Error Resume Next ' Temporarily suppress errors
    colVarName = Application.Match("var_name", dictHeader, 0)
    colVarLabel = Application.Match("var_label_en", dictHeader, 0)
    colColType = Application.Match("column_type", dictHeader, 0)
    On Error GoTo 0 ' Resume normal error handling

    ' Check if all columns were found
    If colVarName = 0 Or colVarLabel = 0 Or colColType = 0 Then
        MsgBox "Could not find one or more required columns in the 'Dictionary' sheet: var_name, var_label_en, column_type.", vbCritical, "Missing Columns"
        Exit Sub
    End If
    
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
    lastCol = dictSheet.Cells(1, dictSheet.Columns.Count).End(xlToLeft).Column
    
    ' Use dynamic last column and assume sheet name is in Col A (Field 1)
    Set dictRange = dictSheet.Range("A1", dictSheet.Cells(lastRow, lastCol))
    dictRange.AutoFilter Field:=1, Criteria1:=sheetName
    
    ' Set filtered range to ONLY the visible "var_name" column
    On Error Resume Next ' Handle cases with no visible cells
    Set filteredVarNameCol = dictSheet.Range(dictSheet.Cells(2, colVarName), dictSheet.Cells(lastRow, colVarName)).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If filteredVarNameCol Is Nothing Then
        ' No matching columns found, just clean up and exit
        Debug.Print "No visible columns in dictionary for sheet: " & sheetName
        GoTo CleanUp
    End If
    
    ' Gather selected headers and their labels
    'For Each ctrl In frmSelectColumns.Controls
    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "CheckBox" Then
            Set chkBox = ctrl
            colName = Replace(chkBox.Name, "chk", "")
            
            ' Find the row in the filtered var_name column, then get label from that row
            Set cell = filteredVarNameCol.Find(What:=colName, LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not cell Is Nothing Then
                colLabel = dictSheet.Cells(cell.Row, colVarLabel).Value
            Else
                colLabel = colName ' Fallback in case of mismatch
            End If
            
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
        ' Search for the column name in the filtered "var_name" range
        Set cell = filteredVarNameCol.Find(What:=colName, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not cell Is Nothing Then
            colType = dictSheet.Cells(cell.Row, colColType).Value
            If colType = "fixed" Then
                ' Skip deletion for ID columns
                Debug.Print "Preserved ID Column: " & colName
                GoTo SkipDeletion
            End If
        End If
        
        ' Check if the column exists in the selectedHeaders collection manually
        isMatch = False
        On Error Resume Next ' Prevent error if collection is empty
        For Each col In selectedHeaders
            If col = colName Then
                isMatch = True
                Exit For
            End If
        Next col
        On Error GoTo 0
        
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
    
CleanUp:

    ' Call SetUpDataValidation to apply data validation
    SetUpDataValidation ws

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

