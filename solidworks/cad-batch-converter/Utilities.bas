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
    WriteLog "Save SLDPRT       : " & OkText(r.SaveOK)
    If Len(r.LockedLayers) > 0 Then _
        WriteLog "Locked layers     : " & r.LockedLayers
    If r.OpenContour Then _
        WriteLog "Open contour      : YES (small gaps merged on import)"
    If Len(r.Message) > 0 Then _
        WriteLog "Message           : " & r.Message
    WriteLog "Processing time   : " & Format$(r.ElapsedSeconds, "0.00") & " s"
End Sub

' Serialise a text-stamp-pass result (a focused subset of TFileResult).
Public Sub LogStampResult(ByRef r As TFileResult)
    WriteLog String(67, "-")
    WriteLog "Timestamp         : " & TimeStamp()
    WriteLog "File              : " & r.FileName
    WriteLog "Words found       : " & r.TextCount
    WriteLog "Text stamped      : " & OkText(r.TextOK)
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
    FileExists = (Len(Dir$(filePath, vbNormal Or vbReadOnly Or _
                                vbHidden Or vbSystem Or vbArchive)) > 0)
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

' Enumerate every file matching fileSpec in a folder into a 0-based array. Full
' paths are stored so the caller does not depend on the Dir() cursor once the
' loop begins. Returns the number of files found.
Public Function CollectFiles(ByVal folderPath As String, _
                             ByVal fileSpec As String, _
                             ByRef arr() As String) As Long
    Const GROW As Long = 256
    Dim count As Long
    ReDim arr(0 To GROW - 1)

    Dim name As String
    name = Dir$(folderPath & fileSpec, vbNormal Or vbReadOnly Or vbArchive)
    Do While Len(name) > 0
        If count > UBound(arr) Then
            ReDim Preserve arr(0 To UBound(arr) + GROW)
        End If
        arr(count) = folderPath & name
        count = count + 1
        name = Dir$
    Loop

    If count > 0 Then
        ReDim Preserve arr(0 To count - 1)
    Else
        Erase arr
    End If

    CollectFiles = count
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

' Canonicalize layer names so legacy files using CUT-OUTSIDE and current
' orchestrator files using CUT - OUTSIDE compare as the same layer.
Private Function NormalizedLayerName(ByVal value As String) As String
    Dim s As String
    s = Trim$(value)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Replace(s, " - ", "-")
    s = Replace(s, "- ", "-")
    s = Replace(s, " -", "-")
    NormalizedLayerName = s
End Function

' Case-insensitive, whitespace- and hyphen-spacing-tolerant comparison.
Public Function LayerEquals(ByVal a As String, ByVal b As String) As Boolean
    LayerEquals = (StrComp(NormalizedLayerName(a), NormalizedLayerName(b), _
                           vbTextCompare) = 0)
End Function

' "OK" / "FAILED" text for a boolean stage flag.
Public Function OkText(ByVal flag As Boolean) As String
    If flag Then
        OkText = "OK"
    Else
        OkText = "FAILED"
    End If
End Function

' True when layerName matches ANY name in csvList (a comma-separated list of
' layer names, e.g. "PIN STAMP TEXT, PART MARKING"). Whitespace around each
' name is ignored; comparison is case-insensitive like LayerEquals. An empty
' list matches nothing.
Public Function LayerInList(ByVal layerName As String, _
                            ByVal csvList As String) As Boolean
    If Len(Trim$(csvList)) = 0 Then Exit Function

    Dim names() As String
    names = Split(csvList, ",")

    Dim i As Long
    For i = LBound(names) To UBound(names)
        If Len(Trim$(names(i))) > 0 Then
            If LayerEquals(layerName, Trim$(names(i))) Then
                LayerInList = True
                Exit Function
            End If
        End If
    Next i
End Function

' Join two log messages with a separator; either side may be empty.
Public Function AppendMsg(ByVal existing As String, ByVal extra As String) As String
    If Len(existing) = 0 Then
        AppendMsg = extra
    ElseIf Len(extra) = 0 Then
        AppendMsg = existing
    Else
        AppendMsg = existing & " | " & extra
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
