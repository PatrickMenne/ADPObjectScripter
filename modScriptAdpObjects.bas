Attribute VB_Name = "modScriptAdpObjects"
Option Explicit
Option Compare Database

'===============================================================================
' modScriptAdpObjects
'
' Export/import of client-side scriptable objects (Forms, Reports, Macros,
' Modules) for an Access 2003 ADP, using Application.SaveAsText /
' Application.LoadFromText.
'
' (Server-side objects - Views, Stored Procedures, Functions - live in SQL
' Server, not the ADP file, and are intentionally NOT covered here.)
'
' EXPORT USAGE:
'   Call ExportAllScriptableObjects()                                          ' default: <ADP path>\adp_dump\, no log
'   Call ExportAllScriptableObjects("C:\Export\MyProject\")                    ' custom path, no log
'   Call ExportAllScriptableObjects("C:\Export\MyProject\", WithLogFile:=True) ' custom path, with log
'   Call ExportAllScriptableObjects WithLogFile:=True                         ' default path, with log
'   -- or run the interactive wrapper --
'   Call RunExportAllScriptableObjects
'
' Note: Before each run, only the object subfolders (Forms, Reports, Macros,
' Modules) are cleared. Any other content in the target directory
' (e.g. log files, .git, .svn) is left untouched.
'
' IMPORT USAGE:
'   Call ImportScriptableObject(acForm, "C:\Export\Forms\frmCustomer.frm")
'   Call ImportScriptableObject(acForm, "C:\Export\Forms\frmCustomer.frm", "frmCopy")
'===============================================================================

Public Const MODULE_VERSION As String = "1.8.20260907"

Private mLogFile As Integer
Private mLogPath As String
Private mSuccessCount As Long
Private mFailCount As Long
Private mWithLogFile As Boolean

'###############################################################################
' EXPORT
'###############################################################################

'--------------------------------------------------------------
' Interactive entry point: prompts for a folder, then runs export
'--------------------------------------------------------------
Public Sub RunExportAllScriptableObjects()

    Dim sFolder As String
    sFolder = InputBox("Path to the export directory (leave empty for default: " & _
                        "<ADP path>\adp_dump\):", _
                        "Export ADP Scriptable Objects", "")

    ExportAllScriptableObjects sFolder

End Sub

'--------------------------------------------------------------
' Main export routine
'--------------------------------------------------------------
Public Sub ExportAllScriptableObjects(Optional ByVal ExportFolder As String = "", _
                                       Optional ByVal WithLogFile As Boolean = False)

    Dim sBase As String
    Dim sFormsDir As String, sReportsDir As String
    Dim sMacrosDir As String, sModulesDir As String

    On Error GoTo ErrHandler

    mSuccessCount = 0
    mFailCount = 0
    mWithLogFile = WithLogFile

    If Len(Trim$(ExportFolder)) = 0 Then
        sBase = CurrentProject.Path & "\adp_dump\"
    Else
        sBase = ExportFolder
    End If
    If Right$(sBase, 1) <> "\" Then sBase = sBase & "\"

    EnsureFolder sBase

    sFormsDir = sBase & "Forms\"
    sReportsDir = sBase & "Reports\"
    sMacrosDir = sBase & "Macros\"
    sModulesDir = sBase & "Modules\"

    ' Before each run, only clear the respective object subfolders
    ' (remaining content of sBase, e.g. log files, .git, .svn, is left untouched)
    EnsureFolder sFormsDir
    ClearFolderRecursively sFormsDir

    EnsureFolder sReportsDir
    ClearFolderRecursively sReportsDir

    EnsureFolder sMacrosDir
    ClearFolderRecursively sMacrosDir

    EnsureFolder sModulesDir
    ClearFolderRecursively sModulesDir

    ' Only open the log file if requested
    If mWithLogFile Then
        mLogPath = sBase & "ExportLog_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
        mLogFile = FreeFile
        Open mLogPath For Output As #mLogFile
    End If

    LogLine "Export started: " & Now
    LogLine "Module version: " & MODULE_VERSION
    LogLine "Target folder: " & sBase
    LogLine "----------------------------------------"

    ExportAccessObjects acForm, sFormsDir, "Forms"
    ExportAccessObjects acReport, sReportsDir, "Reports"
    ExportAccessObjects acMacro, sMacrosDir, "Macros"
    ExportAccessObjects acModule, sModulesDir, "Modules"

    LogLine "----------------------------------------"
    LogLine "Export finished: " & Now
    LogLine "Succeeded: " & mSuccessCount & "   Failed: " & mFailCount

    If mWithLogFile Then Close #mLogFile

    MsgBox "Export complete." & vbCrLf & _
           "Succeeded: " & mSuccessCount & vbCrLf & _
           "Failed: " & mFailCount & _
           IIf(mWithLogFile, vbCrLf & "Log: " & mLogPath, ""), _
           vbInformation, "Export ADP Scriptable Objects"

    Exit Sub

ErrHandler:
    On Error Resume Next
    LogLine "FATAL ERROR: " & Err.Number & " - " & Err.Description
    If mWithLogFile Then Close #mLogFile
    MsgBox "Export failed: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

'--------------------------------------------------------------
' Export a client-side AccessObject collection (Forms/Reports/Macros/Modules)
' via SaveAsText
'--------------------------------------------------------------
Private Sub ExportAccessObjects(ByVal ObjType As AcObjectType, _
                                 ByVal TargetDir As String, _
                                 ByVal Label As String)

    Dim obj As AccessObject
    Dim objCol As Object
    Dim sName As String, sFile As String
    Dim sExt As String

    LogLine ">>> Exporting " & Label & " ..."

    Select Case ObjType
        Case acForm
            sExt = ".frm"
        Case acReport
            sExt = ".rpt"
        Case acMacro
            sExt = ".mac"
        Case acModule
            sExt = ".bas"
        Case Else
            sExt = ".txt"
    End Select

    Select Case ObjType
        Case acForm
            Set objCol = CurrentProject.AllForms
        Case acReport
            Set objCol = CurrentProject.AllReports
        Case acMacro
            Set objCol = CurrentProject.AllMacros
        Case acModule
            Set objCol = CurrentProject.AllModules
    End Select

    For Each obj In objCol
        sName = obj.Name
        sFile = TargetDir & CleanFileName(sName) & sExt

        On Error Resume Next
        Err.Clear
        Application.SaveAsText ObjType, sName, sFile
        If Err.Number <> 0 Then
            LogLine "  FAILED: " & Label & " [" & sName & "] - " & _
                    Err.Number & ": " & Err.Description
            mFailCount = mFailCount + 1
        Else
            LogLine "  OK: " & sName & " -> " & sFile
            mSuccessCount = mSuccessCount + 1
        End If
        On Error GoTo 0
    Next obj

End Sub

'###############################################################################
' IMPORT
'###############################################################################

'--------------------------------------------------------------
' Loads a single previously-exported object (Form, Report, Macro, or Module)
' back into the current ADP using Application.LoadFromText.
'
' If ObjectName is omitted, the file's own name (without extension) is used.
'
' NOTE: If an object with the resulting name already exists in the
' destination ADP, LoadFromText overwrites it without prompting.
'--------------------------------------------------------------
Public Function ImportScriptableObject(ByVal ObjType As AcObjectType, _
                                        ByVal FilePath As String, _
                                        Optional ByVal ObjectName As String = "") As Boolean

    ImportScriptableObject = False

    If Len(Dir(FilePath)) = 0 Then
        MsgBox "File not found: " & FilePath, vbCritical, "Import Failed"
        Exit Function
    End If

    If Len(Trim$(ObjectName)) = 0 Then
        ObjectName = GetFileNameWithoutExtension(FilePath)
    End If

    On Error GoTo ErrHandler

    Application.LoadFromText ObjType, ObjectName, FilePath

    ImportScriptableObject = True
    Exit Function

ErrHandler:
    MsgBox "Import failed for '" & ObjectName & "'" & vbCrLf & _
           "File: " & FilePath & vbCrLf & _
           Err.Number & ": " & Err.Description, vbCritical, "Import Failed"
    ImportScriptableObject = False

End Function

'###############################################################################
' SHARED HELPERS
'###############################################################################

Private Sub EnsureFolder(ByVal sPath As String)
    Dim sTest As String
    sTest = sPath
    If Right$(sTest, 1) = "\" Then sTest = Left$(sTest, Len(sTest) - 1)
    If Len(Dir(sTest, vbDirectory)) = 0 Then
        MkDir sTest
    End If
End Sub

'--------------------------------------------------------------
' Recursively deletes all files and subfolders in sPath.
' sPath itself is kept.
'--------------------------------------------------------------
Private Sub ClearFolderRecursively(ByVal sPath As String)

    Dim sEntry As String
    Dim colFiles As Collection
    Dim colDirs As Collection
    Dim itm As Variant

    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    If Len(Dir(sPath, vbDirectory)) = 0 Then Exit Sub ' folder does not exist

    Set colFiles = New Collection
    Set colDirs = New Collection

    ' Read everything first (Dir() only has a single active search state,
    ' so recursive calls would otherwise disturb the running loop)
    sEntry = Dir(sPath, vbNormal Or vbHidden Or vbSystem Or vbDirectory)
    Do While Len(sEntry) > 0
        If sEntry <> "." And sEntry <> ".." Then
            If (GetAttr(sPath & sEntry) And vbDirectory) = vbDirectory Then
                ' Never delete .git or .svn folders (case-insensitive)
                If LCase$(sEntry) <> ".git" And LCase$(sEntry) <> ".svn" Then
                    colDirs.Add sEntry
                End If
            Else
                colFiles.Add sEntry
            End If
        End If
        sEntry = Dir()
    Loop

    For Each itm In colFiles
        On Error Resume Next
        Kill sPath & itm
        On Error GoTo 0
    Next itm

    For Each itm In colDirs
        ClearFolderRecursively sPath & itm & "\"
        On Error Resume Next
        RmDir sPath & itm
        On Error GoTo 0
    Next itm

End Sub

Private Function CleanFileName(ByVal sName As String) As String
    Dim sBad As String
    Dim i As Integer
    Dim sResult As String
    sBad = "\/:*?""<>|"
    sResult = sName
    For i = 1 To Len(sBad)
        sResult = Replace(sResult, Mid$(sBad, i, 1), "_")
    Next i
    CleanFileName = sResult
End Function

'--------------------------------------------------------------
' Returns the base filename (no path, no extension) from a full path
'--------------------------------------------------------------
Private Function GetFileNameWithoutExtension(ByVal FilePath As String) As String

    Dim sName As String
    Dim iSlash As Integer
    Dim iDot As Integer

    sName = FilePath
    iSlash = InStrRev(sName, "\")
    If iSlash > 0 Then sName = Mid$(sName, iSlash + 1)

    iDot = InStrRev(sName, ".")
    If iDot > 0 Then sName = Left$(sName, iDot - 1)

    GetFileNameWithoutExtension = sName

End Function

Private Sub LogLine(ByVal sText As String)
    On Error Resume Next
    If mWithLogFile Then Print #mLogFile, sText
    Debug.Print sText
End Sub
