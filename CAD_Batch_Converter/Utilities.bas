Attribute VB_Name = "Utilities"
'==============================================================================
' Utilities.bas
'------------------------------------------------------------------------------
' Host-independent helper routines shared by every other module:
'   * Logging          (OpenLog / WriteLog / LogResult / CloseLog)
'   * Folder handling  (FolderExists / FileExists / EnsureFolder)
'   * Timing           (NowSeconds / ElapsedSince / TimeStamp)
'   * Small helpers    (GetBaseName / LayerEquals / OkText / SafeRelease)
'
' Nothing in here depends on the AutoCAD or SolidWorks object models, so these
' routines are safe to call from any host and are trivially unit-testable.
' No third-party libraries are used - logging is done with native VBA file I/O.
'==============================================================================
Option Explicit

' Module-level state for the open log file. mLogNum = 0 means "no log open".
Private mLogNum As Integer
Private mLogPath As String

'==============================================================================
' LOGGING
'==============================================================================

' Open (or create+append) the batch log inside the supplied folder.
Public Sub OpenLog(ByVal folderPath As String)
    On Error Resume Next
    mLogPath = folderPath & LOG_FILE_NAME
    mLogNum = FreeFile
    Open mLogPath For Append As #mLogNum
    If Err.Number <> 0 Then
        ' Could not open the log; disable file logging but keep Debug output.
        mLogNum = 0
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' Write a single line to the log file (if open) and to the Immediate window.
Public Sub WriteLog(ByVal text As String)
    On Error Resume Next
    If mLogNum <> 0 Then Print #mLogNum, text
    Debug.Print text
    On Error GoTo 0
End Sub

' Serialise a completed TFileResult into a readable multi-line log block.
Public Sub LogResult(ByRef r As TFileResult)
    WriteLog String(67, "-")
    WriteLog "Timestamp         : " & TimeStamp()
    WriteLog "File              : " & r.FileName
    WriteLog "AutoCAD filter    : " & OkText(r.AutoCadOK)
    WriteLog "SolidWorks import : " & OkText(r.ImportOK)
    WriteLog "Extrusion         : " & OkText(r.ExtrudeOK)
    WriteLog "Text marks        : " & r.TextCount & " word(s) - " & OkText(r.TextOK)
    WriteLog "Save SLDPRT       : " & OkText(r.SaveOK)
    If Len(r.LockedLayers) > 0 Then _
        WriteLog "Locked layers     : " & r.LockedLayers
    If r.OpenContour Then _
        WriteLog "Open contour      : YES (small gaps merged on import)"
    If Len(r.Message) > 0 Then _
        WriteLog "Message           : " & r.Message
    WriteLog "Processing time   : " & Format$(r.ElapsedSeconds, "0.00") & " s"
End Sub

' Flush and close the log file.
Public Sub CloseLog()
    On Error Resume Next
    If mLogNum <> 0 Then
        Close #mLogNum
        mLogNum = 0
    End If
    On Error GoTo 0
End Sub

'==============================================================================
' FOLDER / FILE HANDLING
'==============================================================================

' True if the given directory exists.
Public Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    Dim p As String
    p = folderPath
    If Len(p) > 1 Then
        If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    End If
    FolderExists = (Len(Dir$(p, vbDirectory)) > 0)
    On Error GoTo 0
End Function

' True if the given file exists.
Public Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(filePath, vbNormal)) > 0)
    On Error GoTo 0
End Function

' Recursively create a folder (and any missing parents). Returns True on
' success. Existing folders are treated as success.
Public Function EnsureFolder(ByVal folderPath As String) As Boolean
    On Error GoTo failed

    Dim p As String
    p = folderPath
    If Len(p) > 1 Then
        If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    End If

    If FolderExists(p) Then
        EnsureFolder = True
        Exit Function
    End If

    ' Create the parent chain first.
    Dim pos As Long
    pos = InStrRev(p, "\")
    If pos > 0 Then
        Dim parent As String
        parent = Left$(p, pos - 1)
        ' Stop recursing at a drive root ("U:") or UNC share root.
        If Len(parent) > 2 And InStr(parent, "\") > 0 Then
            EnsureFolder parent
        End If
    End If

    If Not FolderExists(p) Then MkDir p
    EnsureFolder = FolderExists(p)
    Exit Function

failed:
    EnsureFolder = False
End Function

' Strip the directory and extension, returning just the base file name.
Public Function GetBaseName(ByVal filePath As String) As String
    Dim s As String
    s = filePath

    Dim slash As Long
    slash = InStrRev(s, "\")
    If slash > 0 Then s = Mid$(s, slash + 1)

    Dim dot As Long
    dot = InStrRev(s, ".")
    If dot > 0 Then s = Left$(s, dot - 1)

    GetBaseName = s
End Function

'==============================================================================
' TIMING
'==============================================================================

' Current high-resolution second-of-day (VBA Timer, ~10 ms resolution).
Public Function NowSeconds() As Double
    NowSeconds = Timer
End Function

' Elapsed seconds since a NowSeconds() reading, guarding the midnight rollover.
Public Function ElapsedSince(ByVal startSeconds As Double) As Double
    Dim e As Double
    e = Timer - startSeconds
    If e < 0 Then e = e + 86400#   ' clock rolled past midnight
    ElapsedSince = e
End Function

' Human-readable timestamp for log entries.
Public Function TimeStamp() As String
    TimeStamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function

'==============================================================================
' SMALL HELPERS
'==============================================================================

' Case-insensitive, whitespace-tolerant layer name comparison.
Public Function LayerEquals(ByVal a As String, ByVal b As String) As Boolean
    LayerEquals = (StrComp(Trim$(a), Trim$(b), vbTextCompare) = 0)
End Function

' "OK" / "FAILED" text for a boolean stage flag.
Public Function OkText(ByVal flag As Boolean) As String
    If flag Then
        OkText = "OK"
    Else
        OkText = "FAILED"
    End If
End Function

' Reset every field of a TFileResult so it can be reused inside a loop
' (VBA has no block scope, so a Dim inside a loop is the same variable).
Public Sub ClearResult(ByRef r As TFileResult)
    r.FileName = vbNullString
    r.AutoCadOK = False
    r.ImportOK = False
    r.ExtrudeOK = False
    r.TextCount = 0
    r.TextOK = False
    r.SaveOK = False
    r.OpenContour = False
    r.LockedLayers = vbNullString
    r.Message = vbNullString
    r.ElapsedSeconds = 0#
End Sub

' Safe COM object release. Swallows any error so cleanup never aborts a batch.
Public Sub SafeRelease(ByRef obj As Object)
    On Error Resume Next
    Set obj = Nothing
    On Error GoTo 0
End Sub
