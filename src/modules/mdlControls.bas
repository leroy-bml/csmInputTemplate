Attribute VB_Name = "mdlControls"
Sub GoToFirstSheet()

    ' ThisWorkbook.Sheets(1).Activate
    Sheets("Menu").Visible = xlSheetVisible
    Sheets("Menu").Activate
End Sub

Sub GoToNextSheet()

    Dim ws As Worksheet
    Dim found As Boolean

    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            If ws.Index > ActiveSheet.Index Then
                ws.Activate
                found = True
                Exit For
            End If
        End If
    Next ws

    If Not found Then MsgBox "No next visible sheet found.", vbInformation
End Sub

Sub GoToPreviousSheet()

    Dim currentSheet As Worksheet
    Dim i As Long
    Dim found As Boolean

    Set currentSheet = ActiveSheet

    ' Check sheets in reverse order (from current-1 down to 1)
    For i = currentSheet.Index - 1 To 1 Step -1
        If ThisWorkbook.Sheets(i).Visible = xlSheetVisible Then
            ThisWorkbook.Sheets(i).Activate
            found = True
            Exit Sub
        End If
    Next i

    If Not found Then MsgBox "No previous visible sheet found.", vbInformation
End Sub

Sub BulkResizeColumns()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Set column widths
    ws.Columns("A").ColumnWidth = 13
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C:Z").ColumnWidth = 20

    ' Set row height for row 1
    ws.Rows("1").RowHeight = 150

    ' Merge cells A1:H1 and align text to left and top
    With ws.Range("A1:H1")
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With
End Sub

Sub GoToSoilLyrChem()
    Sheets("SOIL_LYR_CHEMICAL").Visible = xlSheetVisible
    Sheets("SOIL_LYR_CHEMICAL").Activate
End Sub

Sub GoToSoilLyrPhys()
    Sheets("SOIL_LYR_PHYSICAL").Visible = xlSheetVisible
    Sheets("SOIL_LYR_PHYSICAL").Activate
End Sub

Sub GoToSoilLyrTemp()
    Sheets("SOIL_LYR_TEMPERATURE").Visible = xlSheetVisible
    Sheets("SOIL_LYR_TEMPERATURE").Activate
End Sub

Sub GoToSoilLyrOM()
    Sheets("SOIL_LYR_ORGMATTER").Visible = xlSheetVisible
    Sheets("SOIL_LYR_ORGMATTER").Activate
End Sub

Sub GoToSoilLyrCarbon()
    Sheets("SOIL_LYR_CARBON").Visible = xlSheetVisible
    Sheets("SOIL_LYR_CARBON").Activate
End Sub

Sub GoToSoilLyrNPK()
    Sheets("SOIL_LYR_NPK").Visible = xlSheetVisible
    Sheets("SOIL_LYR_NPK").Activate
End Sub

Sub GoToSoilLyrNutrients()
    Sheets("SOIL_LYR_NUTRIENTS").Visible = xlSheetVisible
    Sheets("SOIL_LYR_NUTRIENTS").Activate
End Sub

Sub GoToSoilLyrNdyn()
    Sheets("SOIL_LYR_NITROGEN").Visible = xlSheetVisible
    Sheets("SOIL_LYR_NITROGEN").Activate
End Sub

Sub GoToSoilLyrPdyn()
    Sheets("SOIL_LYR_PHOSPHORUS").Visible = xlSheetVisible
    Sheets("SOIL_LYR_PHOSPHORUS").Activate
End Sub

Sub GoToSoilLyrWater()
    Sheets("SOIL_LYR_WATER").Visible = xlSheetVisible
    Sheets("SOIL_LYR_WATER").Activate
End Sub

Sub GoToSoilLyrRoots()
    Sheets("SOIL_LYR_ROOTS").Visible = xlSheetVisible
    Sheets("SOIL_LYR_ROOTS").Activate
End Sub

Sub GoToWthDaily()
    Sheets("WEATHER_DAILY").Visible = xlSheetVisible
    Sheets("WEATHER_DAILY").Activate
End Sub

Sub GoToWthMonthly()
    Sheets("WEATHER_MONTHLY").Visible = xlSheetVisible
    Sheets("WEATHER_MONTHLY").Activate
End Sub

Sub GoToPlantSmDev()
    Sheets("PLANT_SM_DEV").Visible = xlSheetVisible
    Sheets("PLANT_SM_DEV").Activate
End Sub

Sub GoToPlantSmGrowth()
    Sheets("PLANT_SM_GROWTH").Visible = xlSheetVisible
    Sheets("PLANT_SM_GROWTH").Activate
End Sub

Sub GoToPlantSmNutrients()
    Sheets("PLANT_SM_NUTRIENTS").Visible = xlSheetVisible
    Sheets("PLANT_SM_NUTRIENTS").Activate
End Sub

Sub GoToPlantTsGrowth()
    Sheets("PLANT_TS_GROWTH_DEV").Visible = xlSheetVisible
    Sheets("PLANT_TS_GROWTH_DEV").Activate
End Sub

Sub GoToPlantTsNutrients()
    Sheets("PLANT_TS_NUTRIENTS").Visible = xlSheetVisible
    Sheets("PLANT_TS_NUTRIENTS").Activate
End Sub

Sub GoToPlantTsPnD()
    Sheets("PLANT_TS_PESTS_DISEASES").Visible = xlSheetVisible
    Sheets("PLANT_TS_PESTS_DISEASES").Activate
End Sub

Sub GoToEnvSmWater()
    Sheets("ENV_SM_WATER_DYNAMICS").Visible = xlSheetVisible
    Sheets("ENV_SM_WATER_DYNAMICS").Activate
End Sub

Sub GoToEnvTsWater()
    Sheets("ENV_TS_WATER_DYNAMICS").Visible = xlSheetVisible
    Sheets("ENV_TS_WATER_DYNAMICS").Activate
End Sub

Sub GoToEnvTsPSA()
    Sheets("ENV_TS_PLANT_SOIL_ATMOS").Visible = xlSheetVisible
    Sheets("ENV_TS_PLANT_SOIL_ATMOS").Activate
End Sub

Sub GoToTiEvents()
    Sheets("TILLAGE_EVENTS").Visible = xlSheetVisible
    Sheets("TILLAGE_EVENTS").Activate
End Sub

Sub GoToOmEvents()
    Sheets("ORGANIC_MATERIAL_APPLICS").Visible = xlSheetVisible
    Sheets("ORGANIC_MATERIAL_APPLICS").Activate
End Sub

Sub GoToFeEvents()
    Sheets("FERTILIZER_APPLICS").Visible = xlSheetVisible
    Sheets("FERTILIZER_APPLICS").Activate
End Sub

Sub GoToIrEvents()
    Sheets("IRRIGATION_APPLICATIONS").Visible = xlSheetVisible
    Sheets("IRRIGATION_APPLICATIONS").Activate
End Sub

Sub GoToChEvents()
    Sheets("CHEMICAL_APPLICS").Visible = xlSheetVisible
    Sheets("CHEMICAL_APPLICS").Activate
End Sub

Sub GoToEmLevels()
    Sheets("ENVIRON_MODIF_LEVELS").Visible = xlSheetVisible
    Sheets("ENVIRON_MODIF_LEVELS").Activate
End Sub

Sub GoToHaEvents()
    Sheets("HARVEST_EVENTS").Visible = xlSheetVisible
    Sheets("HARVEST_EVENTS").Activate
End Sub

Sub GoToWeatherParent()
    Sheets("WEATHER_DATA").Visible = xlSheetVisible
    Sheets("WEATHER_DATA").Activate
End Sub

Sub GoToSoilParent()
    Sheets("SOIL_DATA").Visible = xlSheetVisible
    Sheets("SOIL_DATA").Activate
End Sub

Sub GoToPlantParent()
    Sheets("PLANT_DATA").Visible = xlSheetVisible
    Sheets("PLANT_DATA").Activate
End Sub

Sub GoToEnvParent()
    Sheets("ENV_DATA").Visible = xlSheetVisible
    Sheets("ENV_DATA").Activate
End Sub

Sub GoToParentSheet()
    Dim wsCurrent As Worksheet
    Dim wsParent As Worksheet
    Dim prefix As String
    Dim found As Boolean

    ' Get the current active sheet
    Set wsCurrent = ActiveSheet

    ' Get the first 5 characters of the current sheet name
    prefix = Left(wsCurrent.Name, 4)

    ' Search for a sheet that starts with the same prefix
    found = False
    For Each wsParent In ThisWorkbook.Worksheets
        ' Skip the current sheet and check if name starts with prefix
        If wsParent.Name <> wsCurrent.Name And _
           Left(wsParent.Name, 4) = prefix Then
            found = True
            Exit For
        End If
    Next wsParent

    ' If found, activate the parent sheet
    If found Then
        wsParent.Activate
    Else
        MsgBox "No parent sheet found with matching prefix: " & prefix, vbExclamation
    End If
End Sub

