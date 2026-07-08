Attribute VB_Name = "TextStamp"
'==============================================================================
' TextStamp.bas
'------------------------------------------------------------------------------
' SEPARATE, re-runnable pass that stamps the PIN STAMP TEXT words onto the top
' face of parts that RunBatch has already produced.
'
'   RunTextStamp()  <-- run this AFTER RunBatch
'
' Keeping this independent from the extrude batch means:
'   * a text problem can never damage the base parts,
'   * it can be re-run at any time against the staging folder,
'   * the two stages can be verified separately.
'
' Flow
' ----
'   For every SLDPRT in the staging folder:
'     1. Find the matching source DWG (same base name) in the source folder.
'     2. Open that DWG read-only in AutoCAD and harvest TEXT/MTEXT on TEXT_LAYER
'        (the source DWG is untouched, so its text layer is still present).
'     3. Open the SLDPRT in SolidWorks, add the words as native sketch text on
'        the top face (TextMarking.ApplyTextMarks), rebuild and save in place.
'     4. Close the part and continue.
'
' Every file has its own error handling; one bad part never stops the pass.
' Requires references to the AutoCAD 2026 and SolidWorks 2025 type libraries.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Entry point for the text-stamp pass.
'------------------------------------------------------------------------------
Public Sub RunTextStamp()
    Dim runStart As Double
    runStart = NowSeconds()

    EnsureFolder OUTPUT_FOLDER
    OpenLog OUTPUT_FOLDER

    WriteLog ""
    WriteLog String(67, "=")
    WriteLog "TEXT STAMP PASS - run started " & TimeStamp()
    WriteLog "Staging folder  : " & OUTPUT_FOLDER
    WriteLog "Source folder   : " & SOURCE_FOLDER
    WriteLog "Text layer      : " & TEXT_LAYER
    WriteLog String(67, "=")

    If Len(Trim$(TEXT_LAYER)) = 0 Then
        WriteLog "TEXT_LAYER is blank in Config.bas - nothing to stamp."
        CloseLog
        Exit Sub
    End If

    ' --- Snapshot the finished parts ----------------------------------------
    Dim parts() As String
    Dim partCount As Long
    partCount = CollectFiles(OUTPUT_FOLDER, SLDPRT_FILESPEC, parts)
    If partCount = 0 Then
        WriteLog "No SLDPRT files found in the staging folder. Nothing to do."
        CloseLog
        Exit Sub
    End If
    WriteLog "Found " & partCount & " part(s) to stamp."

    ' --- Connect to both applications once ----------------------------------
    Dim acadApp As AcadApplication
    Set acadApp = ConnectAutoCAD()
    If acadApp Is Nothing Then
        WriteLog "FATAL: could not start or attach to AutoCAD."
        CloseLog
        Exit Sub
    End If

    Dim swApp As SldWorks.SldWorks
    Set swApp = ConnectSolidWorks()
    If swApp Is Nothing Then
        WriteLog "FATAL: could not start or attach to SolidWorks."
        ShutdownApps acadApp, swApp
        CloseLog
        Exit Sub
    End If

    ' --- Stamp every part ---------------------------------------------------
    Dim okCount As Long
    Dim failCount As Long
    Dim i As Long
    Dim result As TFileResult

    For i = 0 To partCount - 1
        ClearResult result
        StampOnePart acadApp, swApp, parts(i), result
        LogStampResult result

        If result.TextOK Then
            okCount = okCount + 1
        Else
            failCount = failCount + 1
        End If

        WriteLog "Progress: " & (i + 1) & " / " & partCount & _
                 "   (ok=" & okCount & ", fail=" & failCount & ")"
    Next i

    ' --- Summary ------------------------------------------------------------
    WriteLog String(67, "=")
    WriteLog "TEXT STAMP COMPLETE " & TimeStamp()
    WriteLog "Parts: " & partCount & "   OK: " & okCount & _
             "   Failed: " & failCount
    WriteLog "Total time: " & Format$(ElapsedSince(runStart), "0.0") & " s"
    WriteLog String(67, "=")

    ShutdownApps acadApp, swApp
    CloseLog

    If SHOW_SUMMARY_DIALOG Then
        MsgBox "Text stamp pass complete." & vbCrLf & _
               "Parts: " & partCount & vbCrLf & _
               "OK: " & okCount & vbCrLf & _
               "Failed: " & failCount, vbInformation, "CAD Batch Converter"
    End If
End Sub

'------------------------------------------------------------------------------
' Stamp one part. Self-contained error handling keeps the pass running.
'------------------------------------------------------------------------------
Private Sub StampOnePart(ByVal acadApp As AcadApplication, _
                         ByVal swApp As SldWorks.SldWorks, _
                         ByVal partPath As String, _
                         ByRef r As TFileResult)
    Dim t As Double
    t = NowSeconds()
    On Error GoTo errHandler

    Dim baseName As String
    baseName = GetBaseName(partPath)
    r.FileName = baseName & ".SLDPRT"

    ' --- Locate the source DWG that carries the text ------------------------
    Dim dwgPath As String
    dwgPath = SOURCE_FOLDER & baseName & ".dwg"
    If Not FileExists(dwgPath) Then
        r.Message = "No matching source DWG (" & baseName & ".dwg) - skipped."
        r.TextOK = False
        GoTo finish
    End If

    ' --- Harvest the words from the untouched source DWG --------------------
    TextMarking.ClearMarks
    HarvestFromSource acadApp, dwgPath
    r.TextCount = TextMarking.MarkCount()
    If r.TextCount = 0 Then
        r.Message = "No text on '" & TEXT_LAYER & "' - nothing to stamp."
        r.TextOK = True                       ' not a failure, just empty
        GoTo finish
    End If

    ' --- Open the part, stamp the top face, save in place -------------------
    r.TextOK = StampPart(swApp, partPath, r)

finish:
    r.ElapsedSeconds = ElapsedSince(t)
    Exit Sub

errHandler:
    If Len(r.Message) = 0 Then r.Message = "Text stamp error: " & Err.Description
    r.TextOK = False
    r.ElapsedSeconds = ElapsedSince(t)
End Sub

'------------------------------------------------------------------------------
' Open the source DWG read-only, harvest its text into TextMarking, then close
' it without saving. Read-only means the source is never modified.
'------------------------------------------------------------------------------
Private Sub HarvestFromSource(ByVal acadApp As AcadApplication, _
                              ByVal dwgPath As String)
    Dim doc As AcadDocument
    On Error GoTo cleanup

    Set doc = acadApp.Documents.Open(dwgPath, True)   ' True = read-only
    TextMarking.HarvestTextMarks doc
    doc.Close False
    Set doc = Nothing
    Exit Sub

cleanup:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    Set doc = Nothing
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Open the SLDPRT, apply the harvested text to the top face, rebuild and save in
' place (overwriting the same file). Returns True on success.
'------------------------------------------------------------------------------
Private Function StampPart(ByVal swApp As SldWorks.SldWorks, _
                           ByVal partPath As String, _
                           ByRef r As TFileResult) As Boolean
    Dim swModel As SldWorks.ModelDoc2
    Dim errs As Long
    Dim warns As Long
    On Error GoTo errHandler

    Set swModel = swApp.OpenDoc6(partPath, SW_DOC_PART, SW_OPEN_SILENT, _
                                 vbNullString, errs, warns)
    If swModel Is Nothing Then
        r.Message = "Could not open part (error " & errs & ")."
        StampPart = False
        Exit Function
    End If

    Dim applied As Boolean
    applied = TextMarking.ApplyTextMarks(swModel)
    If Not applied Then
        r.Message = "Text sketch could not be created on the top face."
    End If

    swModel.ForceRebuild3 False

    Dim savedOK As Boolean
    savedOK = swModel.Save3(SW_SAVE_SILENT, errs, warns)
    If Not savedOK Then
        r.Message = "Save failed (error " & errs & ", warning " & warns & ")."
    End If

    swApp.CloseDoc swModel.GetTitle
    Set swModel = Nothing

    StampPart = (applied And savedOK)
    Exit Function

errHandler:
    r.Message = "SolidWorks: " & Err.Description
    On Error Resume Next
    If Not swModel Is Nothing Then swApp.CloseDoc swModel.GetTitle
    Set swModel = Nothing
    On Error GoTo 0
    StampPart = False
End Function
