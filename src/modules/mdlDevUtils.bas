Attribute VB_Name = "mdlDevUtils"
Sub ListAllNames()
    Dim ws As Worksheet
    Dim nm As Name
    Dim i As Long

    ' Add a new worksheet to store the names
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Name_List").Delete ' Delete old list if it exists
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "Name_List"
    
    ' Set up the headers
    ws.Range("A1").Value = "Name"
    ws.Range("B1").Value = "Refers To"
    ws.Range("C1").Value = "Scope"
    ws.Range("D1").Value = "Visible"
    ws.Range("A1:D1").Font.Bold = True
    
    i = 2 ' Start on the second row
    
    ' Loop through all names in the workbook
    For Each nm In ThisWorkbook.Names
        ws.Cells(i, 1).Value = nm.Name
        ws.Cells(i, 2).Value = "'" & nm.RefersTo ' Add apostrophe to show formulas as text
        ws.Cells(i, 3).Value = nm.Parent.Name ' Parent is the sheet or workbook
        ws.Cells(i, 4).Value = nm.Visible
        i = i + 1
    Next nm
    
    ' Auto-fit the columns for readability
    ws.Columns("A:D").AutoFit
    
    MsgBox "Finished! All names have been listed in the 'Name_List' sheet."
End Sub
