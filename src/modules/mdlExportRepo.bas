Attribute VB_Name = "mdlExportRepo"
Sub ExportFilesToRepo()

    ' Tool for exporting VBA-based components to an external folder (say for storing in a git repo)
    ' Source: https://minerupset.com/2022/Git-and-Excel-VBA/

    ' Adjust your base path name here
    Dim basePath As String: basePath = "C:\Users\bmlle\Documents\0_DATA\TUM\HEF\FAIRagro\2-UseCases\UC6_IntegratedModeling\Workflows\csmInputTemplate\src\"

    ' Define subfolders for different types of components
    Dim modulesPath As String: modulesPath = basePath & "modules\"
    Dim formsPath As String: formsPath = basePath & "forms\"
    Dim excelObjectsPath As String: excelObjectsPath = basePath & "excel_objects\"
    
    ' Ensure subfolders exist (create if they don't)
    If Dir(modulesPath, vbDirectory) = "" Then MkDir modulesPath
    If Dir(formsPath, vbDirectory) = "" Then MkDir formsPath
    If Dir(excelObjectsPath, vbDirectory) = "" Then MkDir excelObjectsPath
    
    ' The VBComponent Class represents those objects that make up an Excel Workbook
    Dim vbModule As VBComponent
    Dim codeLines As Long

    ' This loops through each of those VBComponents in the Active Workbook
    For Each vbModule In ActiveWorkbook.VBProject.VBComponents

        ' Some Debug.Print statements for easy testing during development
        Debug.Print vbModule.Name

        ' Skip exporting "Module1.bas"
        If vbModule.Name = "Module1" Then
            Debug.Print "Skipping Module1"
            GoTo NextComponent
        End If
        
        ' Count the lines of code in the module
        codeLines = vbModule.CodeModule.CountOfLines

        ' Skip the export if the module is empty
        If codeLines = 0 Then
            Debug.Print "Skipping empty module " & vbModule.Name
            GoTo NextComponent
        End If

        ' Runs a selection based on the type of module
        Select Case vbModule.Type
            Case vbext_ct_StdModule
                vbModule.Export modulesPath & vbModule.Name & ".bas"
                Debug.Print "Exported to modules"
            Case vbext_ct_ClassModule
                vbModule.Export modulesPath & vbModule.Name & ".cls"
                Debug.Print "Exported to modules"
            Case vbext_ct_MSForm
                vbModule.Export formsPath & vbModule.Name & ".frm"
                Debug.Print "Exported to forms"
            Case vbext_ct_Document
                vbModule.Export excelObjectsPath & vbModule.Name & ".cls"
                Debug.Print "Exported to excel_objects"
            Case Else
                Debug.Print "Not exporting " & vbModule.Name
        End Select

NextComponent:
    Next vbModule
End Sub


