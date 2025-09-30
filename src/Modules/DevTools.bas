Attribute VB_Name = "DevTools"
Option Explicit

' VBA Import/Export Development Tools
' Author: VBA Import/Export Starter
' Purpose: Facilitate development workflow between VS Code and Excel VBE
' Requirements: Microsoft Visual Basic for Applications Extensibility 5.3 reference
'               Trust access to the VBA project object model enabled
' Usage: Designed for use in Personal Add-in with Ribbon controls

' ============================================================================
' PUBLIC RIBBON PROCEDURES
' ============================================================================

Public Sub ExportAllRibbon(control As IRibbonControl)
    ' Ribbon callback for ExportAll functionality
    ' Exports all VBA components from the active workbook to the src/ directory
    
    On Error GoTo ErrorHandler
    
    ' Ensure we're working with the active workbook, not the add-in workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found. Please open a workbook first.", _
               vbExclamation, "DevTools - No Active Workbook"
        Exit Sub
    End If
    
    ' Check if active workbook has VS Code setup
    If Not IsVSCodeProject() Then
        MsgBox "The active workbook is not part of a VS Code setup." & vbCrLf & _
               "Please ensure the workbook is in a folder with a 'src' directory structure.", _
               vbExclamation, "DevTools - Not a VS Code Project"
        Exit Sub
    End If
    
    ' Call the main export procedure
    Call ExportAll
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in ExportAllRibbon: " & Err.Number & " - " & Err.Description, _
           vbCritical, "DevTools - Ribbon Error"
End Sub

Public Sub ImportAllRibbon(control As IRibbonControl)
    ' Ribbon callback for ImportAll functionality
    ' Imports all VBA components from the src/ directory into the active workbook
    
    On Error GoTo ErrorHandler
    
    ' Ensure we're working with the active workbook, not the add-in workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found. Please open a workbook first.", _
               vbExclamation, "DevTools - No Active Workbook"
        Exit Sub
    End If
    
    ' Check if active workbook has VS Code setup
    If Not IsVSCodeProject() Then
        MsgBox "The active workbook is not part of a VS Code setup." & vbCrLf & _
               "Please ensure the workbook is in a folder with a 'src' directory structure.", _
               vbExclamation, "DevTools - Not a VS Code Project"
        Exit Sub
    End If
    
    ' Call the main import procedure
    Call ImportAll
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in ImportAllRibbon: " & Err.Number & " - " & Err.Description, _
           vbCritical, "DevTools - Ribbon Error"
End Sub

' ============================================================================
' ALTERNATIVE RIBBON PROCEDURES (if the above don't work)
' ============================================================================

Public Sub RunExportAll(control As IRibbonControl)
    ' Alternative ribbon callback using Application.Run
    On Error GoTo ErrorHandler
    
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found. Please open a workbook first.", _
               vbExclamation, "DevTools - No Active Workbook"
        Exit Sub
    End If
    
    ' Use Application.Run to explicitly call the procedure from this workbook
    Application.Run ThisWorkbook.Name & "!ExportAllRibbon", control
    Exit Sub
    
ErrorHandler:
    MsgBox "Error running export: " & Err.Number & " - " & Err.Description, _
           vbCritical, "DevTools - Export Error"
End Sub

Public Sub RunImportAll(control As IRibbonControl)
    ' Alternative ribbon callback using Application.Run
    On Error GoTo ErrorHandler
    
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found. Please open a workbook first.", _
               vbExclamation, "DevTools - No Active Workbook"
        Exit Sub
    End If
    
    ' Use Application.Run to explicitly call the procedure from this workbook
    Application.Run ThisWorkbook.Name & "!ImportAllRibbon", control
    Exit Sub
    
ErrorHandler:
    MsgBox "Error running import: " & Err.Number & " - " & Err.Description, _
           vbCritical, "DevTools - Import Error"
End Sub

' ============================================================================
' VALIDATION FUNCTIONS
' ============================================================================

Private Function IsVSCodeProject() As Boolean
    ' Checks if the active workbook is part of a VS Code project setup
    ' Returns True if a src directory exists in the project structure
    
    On Error GoTo NotVSCodeProject
    
    ' Ensure we have an active workbook
    If ActiveWorkbook Is Nothing Then
        IsVSCodeProject = False
        Exit Function
    End If
    
    Dim projectRootPath As String
    Dim srcPath As String
    
    ' Get the project root directory
    projectRootPath = ProjectRoot()
    
    ' Check if we got a valid project root
    If projectRootPath = "" Then
        IsVSCodeProject = False
        Exit Function
    End If
    
    srcPath = projectRootPath & "\src"
    
    ' Check if src directory exists
    If (GetAttr(srcPath) And vbDirectory) = vbDirectory Then
        IsVSCodeProject = True
    Else
        IsVSCodeProject = False
    End If
    
    Exit Function
    
NotVSCodeProject:
    IsVSCodeProject = False
End Function

' ============================================================================
' PUBLIC PROCEDURES (Modified for Add-in use)
' ============================================================================

Public Sub ExportAll()
    ' Exports all VBA components from the active workbook to the src/ directory
    ' Components are organized by type: Modules/, Classes/, Forms/
    
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim vbComp As VBIDE.VBComponent
    Dim exportPath As String
    Dim fileName As String
    Dim fileExt As String
    Dim subFolder As String
    Dim exportCount As Long
    
    Set wb = ActiveWorkbook
    
    ' Verify we have a valid workbook
    If wb Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "DevTools - Export Error"
        Exit Sub
    End If
    
    exportCount = 0
    
    ' Ensure src directories exist
    Call MkDirIfMissing(SrcPath("Modules"))
    Call MkDirIfMissing(SrcPath("Classes"))
    Call MkDirIfMissing(SrcPath("Forms"))
    
    ' Export all non-document components
    For Each vbComp In wb.VBProject.VBComponents
        
        ' Skip document modules (ThisWorkbook, Sheets)
        If vbComp.Type <> vbext_ct_Document Then
            
            ' Determine file extension and subfolder based on component type
            Select Case vbComp.Type
                Case vbext_ct_StdModule
                    fileExt = ".bas"
                    subFolder = "Modules"
                Case vbext_ct_ClassModule
                    fileExt = ".cls"
                    subFolder = "Classes"
                Case vbext_ct_MSForm
                    fileExt = ".frm"
                    subFolder = "Forms"
                Case Else
                    ' Skip unknown component types
                    GoTo NextComponent
            End Select
            
            ' Build export path
            fileName = vbComp.Name & fileExt
            exportPath = SrcPath(subFolder) & "\" & fileName
            
            ' Export the component
            vbComp.Export exportPath
            exportCount = exportCount + 1
            
            Debug.Print "Exported: " & subFolder & "\" & fileName
        End If
        
NextComponent:
    Next vbComp
    
    MsgBox "Export complete. " & exportCount & " components exported to src/ directory." & vbCrLf & _
           "Workbook: " & wb.Name, _
           vbInformation, "DevTools - Export Complete"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during export: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "Ensure VBIDE reference is enabled and trust access is granted.", _
           vbCritical, "DevTools - Export Error"
End Sub

Public Sub ImportAll()
    ' Imports all VBA components from the src/ directory into the active workbook
    ' Removes existing non-document components before importing
    
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim importCount As Long
    
    Set wb = ActiveWorkbook
    
    ' Verify we have a valid workbook
    If wb Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "DevTools - Import Error"
        Exit Sub
    End If
    
    importCount = 0
    
    Debug.Print "=== ImportAll Started ==="
    Debug.Print "Current workbook: " & wb.Name
    Debug.Print "Full path: " & wb.FullName
    
    ' Show paths being used
    Call ShowPaths
    
    ' Remove all existing code (except document modules and DevTools if it exists)
    Debug.Print "=== Starting RemoveAllCode ==="
    Call RemoveAllCode(wb)
    Debug.Print "=== RemoveAllCode Complete ==="
    
    ' Import all components from src subdirectories
    Debug.Print "=== Starting Imports ==="
    
    Debug.Print "--- Importing Modules ---"
    importCount = importCount + ImportFromFolder(SrcPath("Modules"), "*.bas")
    
    Debug.Print "--- Importing Classes ---"
    importCount = importCount + ImportFromFolder(SrcPath("Classes"), "*.cls")
    
    Debug.Print "--- Importing Forms ---"
    importCount = importCount + ImportFromFolder(SrcPath("Forms"), "*.frm")
    
    Debug.Print "=== Import Summary ==="
    Debug.Print "Total imported: " & importCount
    
    ' List components after import
    Call ListAllComponents
    
    MsgBox "Import complete. " & importCount & " components imported from src/ directory." & vbCrLf & _
           "Workbook: " & wb.Name, _
           vbInformation, "DevTools - Import Complete"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ImportAll Error ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Error Source: " & Err.Source
    MsgBox "Error during import: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "Check Immediate Window for detailed debug output." & vbCrLf & _
           "Ensure VBIDE reference is enabled and trust access is granted.", _
           vbCritical, "DevTools - Import Error"
End Sub

' ============================================================================
' PRIVATE HELPER PROCEDURES
' ============================================================================

Private Sub MkDirIfMissing(ByVal dirPath As String)
    ' Creates a directory if it doesn't exist
    ' Uses recursive creation similar to mkdir -p
    
    On Error GoTo CreateDir
    
    ' Test if directory exists by attempting to get its attributes
    If (GetAttr(dirPath) And vbDirectory) = vbDirectory Then
        Exit Sub ' Directory already exists
    End If
    
CreateDir:
    On Error GoTo ErrorHandler
    
    ' Create the directory and any parent directories
    MkDir dirPath
    
    Exit Sub
    
ErrorHandler:
    ' If MkDir fails, try creating parent directory first
    If Err.Number = 76 Then ' Path not found
        Dim parentPath As String
        Dim lastSlash As Long
        
        lastSlash = InStrRev(dirPath, "\")
        If lastSlash > 3 Then ' Not a root directory
            parentPath = Left(dirPath, lastSlash - 1)
            Call MkDirIfMissing(parentPath)
            MkDir dirPath
        Else
            Err.Raise Err.Number, "DevTools.MkDirIfMissing", Err.Description
        End If
    Else
        Err.Raise Err.Number, "DevTools.MkDirIfMissing", Err.Description
    End If
End Sub

Private Sub RemoveAllCode(ByVal targetWB As Workbook)
    ' Removes all non-document VBA components from the specified workbook
    ' Preserves ThisWorkbook, Sheet modules, and any DevTools module
    
    Dim vbComp As VBIDE.VBComponent
    Dim compName As String
    Dim removeCount As Long
    
    removeCount = 0
    
    ' Iterate through components in reverse to avoid collection modification issues
    Dim i As Long
    For i = targetWB.VBProject.VBComponents.Count To 1 Step -1
        Set vbComp = targetWB.VBProject.VBComponents(i)
        
        ' Only remove non-document components, but preserve any DevTools module
        If vbComp.Type <> vbext_ct_Document And vbComp.Name <> "DevTools" Then
            compName = vbComp.Name
            targetWB.VBProject.VBComponents.Remove vbComp
            removeCount = removeCount + 1
            Debug.Print "Removed: " & compName
        ElseIf vbComp.Name = "DevTools" Then
            Debug.Print "Preserved: DevTools module"
        End If
    Next i
    
    Debug.Print "RemoveAllCode: " & removeCount & " components removed"
End Sub

Private Function ImportFromFolder(ByVal folderPath As String, ByVal filePattern As String) As Long
    ' Imports all files matching the pattern from the specified folder
    ' Returns the count of successfully imported files
    
    Dim fileName As String
    Dim filePath As String
    Dim importCount As Long
    Dim fileCount As Long
    
    importCount = 0
    fileCount = 0
    
    Debug.Print "ImportFromFolder called with:"
    Debug.Print "  Folder: " & folderPath
    Debug.Print "  Pattern: " & filePattern
    
    ' Check if folder exists
    On Error GoTo FolderError
    Debug.Print "  Checking if folder exists..."
    If (GetAttr(folderPath) And vbDirectory) <> vbDirectory Then
        Debug.Print "  Folder check failed - going to FolderError"
        GoTo FolderError
    End If
    Debug.Print "  Folder exists: YES"
    On Error GoTo ImportError
    
    ' Count and list all files first
    Debug.Print "  Scanning for files..."
    fileName = Dir(folderPath & "\" & filePattern)
    Do While fileName <> ""
        fileCount = fileCount + 1
        Debug.Print "  Found file #" & fileCount & ": " & fileName
        fileName = Dir()
    Loop
    
    Debug.Print "  Total files found: " & fileCount
    
    ' Now import the files
    If fileCount > 0 Then
        Debug.Print "  Starting import process..."
        fileName = Dir(folderPath & "\" & filePattern)
        Do While fileName <> ""
            filePath = folderPath & "\" & fileName
            
            Debug.Print "  Processing: " & fileName
            Debug.Print "  Full path: " & filePath
            
            ' Skip importing DevTools to avoid duplicates
            If Left(fileName, 8) <> "DevTools" Then
                Debug.Print "  Attempting to import..."
                ' Import the component
                ActiveWorkbook.VBProject.VBComponents.Import filePath
                importCount = importCount + 1
                Debug.Print "  SUCCESS - Imported: " & fileName
            Else
                Debug.Print "  SKIPPED - DevTools file: " & fileName
            End If
            
            fileName = Dir() ' Get next file
        Loop
    End If
    
    Debug.Print "  Import summary for folder:"
    Debug.Print "    Files found: " & fileCount
    Debug.Print "    Files imported: " & importCount
    
    ImportFromFolder = importCount
    Exit Function
    
FolderError:
    Debug.Print "  FOLDER ERROR:"
    Debug.Print "    Error Number: " & Err.Number
    Debug.Print "    Error Description: " & Err.Description
    Debug.Print "    Folder path: " & folderPath
    
    ' Try to provide more specific error info
    If Err.Number = 53 Then
        Debug.Print "    Issue: File not found (folder doesn't exist)"
    ElseIf Err.Number = 76 Then
        Debug.Print "    Issue: Path not found"
    End If
    
    ImportFromFolder = 0
    Exit Function
    
ImportError:
    Debug.Print "  IMPORT ERROR for file: " & fileName
    Debug.Print "    Error Number: " & Err.Number
    Debug.Print "    Error Description: " & Err.Description
    Debug.Print "    File path: " & filePath
    Resume Next
End Function

Private Function ProjectRoot() As String
    ' Returns the project root directory based on active workbook location
    ' Handles OneDrive SharePoint URLs by converting to local paths
    
    Dim workbookPath As String
    Dim workbookDir As String
    Dim lastSlash As Long
    
    ' Use active workbook instead of assuming a specific workbook
    If ActiveWorkbook Is Nothing Then
        ProjectRoot = ""
        Exit Function
    End If
    
    workbookPath = ActiveWorkbook.FullName
    
    ' Check if this is a OneDrive SharePoint URL
    If InStr(workbookPath, "sharepoint.com") > 0 Or InStr(workbookPath, "https://") = 1 Then
        workbookPath = ConvertOneDriveUrl(workbookPath)
    End If
    
    ' Get the workbook's directory
    lastSlash = InStrRev(workbookPath, "\")
    If lastSlash > 0 Then
        workbookDir = Left(workbookPath, lastSlash - 1)
    Else
        workbookDir = workbookPath ' Fallback, though this shouldn't happen
    End If
    
    ' For add-in use, assume the workbook is already in the project root
    ' or look for a parent directory with src folder
    If Dir(workbookDir & "\src", vbDirectory) <> "" Then
        ProjectRoot = workbookDir
    Else
        ' Try going up one level
        lastSlash = InStrRev(workbookDir, "\")
        If lastSlash > 0 Then
            Dim parentDir As String
            parentDir = Left(workbookDir, lastSlash - 1)
            If Dir(parentDir & "\src", vbDirectory) <> "" Then
                ProjectRoot = parentDir
            Else
                ProjectRoot = workbookDir ' Fallback to workbook directory
            End If
        Else
            ProjectRoot = workbookDir
        End If
    End If
End Function

Private Function ConvertOneDriveUrl(ByVal sharePointUrl As String) As String
    ' Converts OneDrive SharePoint URL to local file path
    ' Example: https://technip-my.sharepoint.com/personal/user/Documents/folder/file.xlsm
    ' Becomes: C:\Users\user\OneDrive - TEN\folder\file.xlsm
    
    Dim localPath As String
    Dim userName As String
    Dim documentsIndex As Long
    Dim afterDocuments As String
    
    ' Extract username from URL
    userName = Environ("USERNAME")
    
    ' Find the Documents part in the URL
    documentsIndex = InStr(sharePointUrl, "/Documents/")
    
    If documentsIndex > 0 Then
        ' Get everything after "/Documents/"
        afterDocuments = Mid(sharePointUrl, documentsIndex + 11) ' 11 = Len("/Documents/")
        
        ' Replace forward slashes with backslashes
        afterDocuments = Replace(afterDocuments, "/", "\")
        
        ' Build local OneDrive path
        localPath = "C:\Users\" & userName & "\OneDrive - TEN\" & afterDocuments
    Else
        ' Fallback: try to use current directory
        localPath = Application.ActiveWorkbook.Path
        If localPath = "" Then
            localPath = "C:\Users\" & userName & "\OneDrive - TEN"
        End If
    End If
    
    ConvertOneDriveUrl = localPath
End Function

Private Function SrcPath(ByVal subFolder As String) As String
    ' Returns the full path to a subfolder within the src directory
    ' Example: SrcPath("Modules") returns "C:\Path\To\Project\src\Modules"
    
    SrcPath = ProjectRoot() & "\src\" & subFolder
End Function

' ============================================================================
' UTILITY PROCEDURES FOR DEBUGGING
' ============================================================================

Public Sub ListAllComponents()
    ' Debugging helper: Lists all VBA components in the current workbook
    
    Dim vbComp As VBIDE.VBComponent
    Dim compTypeStr As String
    
    Debug.Print "=== VBA Components in " & ActiveWorkbook.Name & " ==="
    
    For Each vbComp In ActiveWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule: compTypeStr = "Standard Module"
            Case vbext_ct_ClassModule: compTypeStr = "Class Module"
            Case vbext_ct_MSForm: compTypeStr = "UserForm"
            Case vbext_ct_Document: compTypeStr = "Document Module"
            Case Else: compTypeStr = "Unknown (" & vbComp.Type & ")"
        End Select
        
        Debug.Print vbComp.Name & " - " & compTypeStr
    Next vbComp
    
    Debug.Print "=== End Component List ==="
End Sub

Public Sub ShowPaths()
    ' Debugging helper: Shows the paths that DevTools will use for active workbook
    
    Debug.Print "=== DevTools Paths ==="
    If ActiveWorkbook Is Nothing Then
        Debug.Print "No active workbook found"
    Else
        Debug.Print "ActiveWorkbook.Name: " & ActiveWorkbook.Name
        Debug.Print "ActiveWorkbook.FullName: " & ActiveWorkbook.FullName
    End If
    Debug.Print "Project Root: " & ProjectRoot()
    Debug.Print "Modules Path: " & SrcPath("Modules")
    Debug.Print "Classes Path: " & SrcPath("Classes")
    Debug.Print "Forms Path: " & SrcPath("Forms")
    Debug.Print "Is VS Code Project: " & IsVSCodeProject()
    Debug.Print "=== End Paths ==="
End Sub

Public Sub TestFolderAccess()
    ' New debugging procedure to test folder access
    
    Debug.Print "=== Testing Folder Access ==="
    
    Dim testPaths As Variant
    Dim i As Long
    
    testPaths = Array( _
        ProjectRoot(), _
        ProjectRoot() & "\src", _
        SrcPath("Modules"), _
        SrcPath("Classes"), _
        SrcPath("Forms") _
    )
    
    For i = 0 To UBound(testPaths)
        Debug.Print "Testing path: " & testPaths(i)
        
        On Error GoTo TestError
        If (GetAttr(testPaths(i)) And vbDirectory) = vbDirectory Then
            Debug.Print "  STATUS: EXISTS (Directory)"
            
            ' List files in the directory
            Dim fileName As String
            fileName = Dir(testPaths(i) & "\*.*")
            If fileName <> "" Then
                Debug.Print "  Files found:"
                Do While fileName <> ""
                    Debug.Print "    " & fileName
                    fileName = Dir()
                Loop
            Else
                Debug.Print "  Directory is empty"
            End If
        Else
            Debug.Print "  STATUS: NOT A DIRECTORY"
        End If
        GoTo NextPath
        
TestError:
        Debug.Print "  STATUS: ERROR - " & Err.Number & " - " & Err.Description
        
NextPath:
        On Error GoTo 0
    Next i
    
    Debug.Print "=== End Folder Access Test ==="
End Sub

