Attribute VB_Name = "modScriptApdObjects"
Option Explicit
Option Compare Database

'===============================================================================
' modScriptApdObjects
'
' Export/import of client-side scriptable objects (Forms, Reports, Macros,
' Modules) for an Access 2003 ADP, using Application.SaveAsText /
' Application.LoadFromText.
'
' (Server-side objects - Views, Stored Procedures, Functions - live in SQL
' Server, not the ADP file, and are intentionally NOT covered here.)
'
' EXPORT USAGE:
'   Call ExportAllScriptableObjects("C:\Export\MyProject\")
'   -- or run the interactive wrapper --
'   Call RunExportAllScriptableObjects
'
' IMPORT USAGE:
'   Call ImportScriptableObject(acForm, "C:\Export\Forms\frmCustomer.txt")
'   Call ImportScriptableObject(acForm, "C:\Export\Forms\frmCustomer.txt", "frmCopy")
'===============================================================================

Private mLogFile As Integer
Private mLogPath As String
Private mSuccessCount As Long
Private mFailCount As Long

'###############################################################################
' EXPORT
'###############################################################################

'--------------------------------------------------------------
' Interactive entry point: prompts for a folder, then runs export
'--------------------------------------------------------------
Public Sub RunExportAllScriptableObjects()

    Dim sFolder As String
    sFolder = InputBox("Enter the full path to the export folder:", _
                        "Export ADP Scriptable Objects", "C:\ADP_Export\")

    If Len(Trim$(sFolder)) = 0 Then
        MsgBox "Export cancelled.", vbInformation
        Exit Sub
    End If

    ExportAllScriptableObjects sFolder

End Sub

'--------------------------------------------------------------
' Main export routine
'--------------------------------------------------------------
Public Sub ExportAllScriptableObjects(ByVal ExportFolder As String)

    Dim sBase As String
    Dim sFormsDir As String, sReportsDir As String
    Dim sMacrosDir As String, sModulesDir As String

    On Error GoTo ErrHandler

    mSuccessCount = 0
    mFailCount = 0

    sBase = ExportFolder
    If Right$(sBase, 1) <> "\" Then sBase = sBase & "\"
    EnsureFolder sBase

    sFormsDir = sBase & "Forms\"
    sReportsDir = sBase & "Reports\"
    sMacrosDir = sBase & "Macros\"
    sModulesDir = sBase & "Modules\"

    EnsureFolder sFormsDir
    EnsureFolder sReportsDir
    EnsureFolder sMacrosDir
    EnsureFolder sModulesDir

    ' Open log file
    mLogPath = sBase & "ExportLog_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    mLogFile = FreeFile
    Open mLogPath For Output As #mLogFile
    LogLine "Export started: " & Now
    LogLine "Target folder: " & sBase
    LogLine "----------------------------------------"

    ExportAccessObjects acForm, sFormsDir, "Forms"
    ExportAccessObjects acReport, sReportsDir, "Reports"
    ExportAccessObjects acMacro, sMacrosDir, "Macros"
    ExportAccessObjects acModule, sModulesDir, "Modules"

    LogLine "----------------------------------------"
    LogLine "Export finished: " & Now
    LogLine "Succeeded: " & mSuccessCount & "   Failed: " & mFailCount
    Close #mLogFile

    MsgBox "Export complete." & vbCrLf & _
           "Succeeded: " & mSuccessCount & vbCrLf & _
           "Failed: " & mFailCount & vbCrLf & _
           "Log: " & mLogPath, vbInformation, "Export ADP Scriptable Objects"

    Exit Sub

ErrHandler:
    On Error Resume Next
    LogLine "FATAL ERROR: " & Err.Number & " - " & Err.Description
    Close #mLogFile
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

    sExt = ".txt"
    If ObjType = acModule Then sExt = ".bas"

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
    Print #mLogFile, sText
    Debug.Print sText
End Sub
