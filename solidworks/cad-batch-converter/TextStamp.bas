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
' Uses late-bound AutoCAD COM and uses late-bound SolidWorks COM.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Entry point for the text-stamp pass.
'------------------------------------------------------------------------------
Public Sub RunTextStamp()
    Dim runStart As Double
    runStart = NowSeconds()
    RefreshRuntimeConfiguration

    If Len(SOURCE_FOLDER) = 0 Or Len(OUTPUT_FOLDER) = 0 Then
        MsgBox "CAD batch source/output folders are not configured. " & _
               "Launch this macro from the Engineering Job Assistant.", _
               vbCritical, "CAD Batch Converter"
        Exit Sub
    End If

    If Not EnsureFolder(OUTPUT_FOLDER) Then
        MsgBox "Could not create or access the output folder:" & vbCrLf & _
               OUTPUT_FOLDER, vbCritical, "CAD Batch Converter"
        Exit Sub
    End If
    OpenLog OUTPUT_FOLDER

    WriteLog ""
    WriteLog String(67, "=")
    WriteLog "TEXT STAMP PASS - run started " & TimeStamp()
    WriteLog "Staging folder  : " & OUTPUT_FOLDER
    WriteLog "Source folder   : " & SOURCE_FOLDER
    WriteLog "Text layer      : " & TEXT_LAYER
    If TEXT_USE_DWG_OUTLINES Then
        WriteLog "Text rendering  : AutoCAD TXTEXP outlines"
    Else
        WriteLog "Text rendering  : direct text + native stroke font"
    End If
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
    Dim acadApp As Object
    Set acadApp = ConnectAutoCAD()
    If acadApp Is Nothing Then
        WriteLog "FATAL: could not start or attach to AutoCAD."
        CloseLog
        Exit Sub
    End If

    Dim swApp As Object
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
Private Sub StampOnePart(ByVal acadApp As Object, _
                         ByVal swApp As Object, _
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
    Dim harvestProblem As String
    If Not HarvestFromSource(acadApp, dwgPath, harvestProblem) Then
        r.Message = "Could not harvest marking data from the source DWG: " & _
                    harvestProblem
        r.TextOK = False
        GoTo finish
    End If
    r.TextCount = TextMarking.MarkCount()
    If r.TextCount = 0 And TextMarking.SegCount() = 0 Then
        r.Message = "No text or marking geometry on '" & TEXT_LAYER & _
                    "' - nothing to stamp."
        r.TextOK = True                       ' not a failure, just empty
        GoTo finish
    End If

    If r.TextCount = 0 And Not TEXT_USE_DWG_OUTLINES Then
        r.Message = AppendMsg(r.Message, _
                    "No supported text entities were harvested; marking " & _
                    "geometry only.")
    End If

    ' Say which text path is in play: with outline conversion on, surviving
    ' TEXT/MTEXT means TXTEXP did not convert it (stroke font renders it).
    If TEXT_USE_DWG_OUTLINES Then
        If r.TextCount > 0 Then
            r.Message = AppendMsg(r.Message, "TXTEXP did not convert the" & _
                        " text - stroke font used.")
        Else
            r.Message = AppendMsg(r.Message, "DWG font outlines in use (TXTEXP).")
        End If
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
' Open the source DWG, harvest its text + marking geometry into TextMarking,
' then close WITHOUT saving - the file on disk is never modified either way.
'
' When TEXT_USE_DWG_OUTLINES is True, Express Tools' TXTEXP is run first: it
' explodes every TEXT/MTEXT into REAL letter-outline polylines (the drawing's
' own font), which the harvest then picks up as marking geometry - so the
' words land in SolidWorks exactly as they look on the drawing. TXTEXP needs a
' writable document, hence the in-memory (non-read-only) open. If TXTEXP is
' not available the text entities simply survive, the harvest captures them as
' words, and the built-in stroke font renders them - automatic fallback.
'------------------------------------------------------------------------------
Private Function HarvestFromSource(ByVal acadApp As Object, _
                                   ByVal dwgPath As String, _
                                   ByRef problem As String) As Boolean
    Dim doc As Object
    On Error GoTo failed

    Set doc = acadApp.Documents.Open(dwgPath, Not TEXT_USE_DWG_OUTLINES)
    If doc Is Nothing Then
        problem = "AutoCAD returned no document."
        Exit Function
    End If

    Dim outlineLayer As String
    If TEXT_USE_DWG_OUTLINES Then
        outlineLayer = PrepareAndExplodeText(doc)
    End If

    TextMarking.HarvestTextMarks doc, outlineLayer
    doc.Close False                     ' discard the exploded copy
    Set doc = Nothing
    HarvestFromSource = True
    Exit Function

failed:
    problem = Err.Description
    If Len(problem) = 0 Then problem = "unknown AutoCAD error"
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    Set doc = Nothing
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' TXTEXP is WMF-based: its exploded output lands on the CURRENT layer, not on
' the text's own layer - which is exactly how the outlines were vanishing
' (they landed on layer "0" where the layer-filtered harvest never looks).
'
' So, on the in-memory copy (never saved):
'   1. unlock every layer,
'   2. DELETE every TEXT/MTEXT that is NOT on a marking layer, so only the
'      marking words are left to explode,
'   3. park the current layer on a scratch layer,
'   4. run TXTEXP - everything on the scratch layer afterwards is, by
'      construction, marking text as real letter outlines.
' Returns the scratch layer name for the harvest, or "" if setup failed
' (the stroke-font fallback then takes over automatically).
'------------------------------------------------------------------------------
Private Function PrepareAndExplodeText(ByVal doc As Object) As String
    On Error Resume Next
    Const SCRATCH_LAYER As String = "ZZ_TXTEXP_OUT"

    ' Unlock everything so stray text can be deleted.
    Dim lay As Object
    For Each lay In doc.Layers
        lay.Lock = False
    Next lay

    ' Remove text that is not marking text (in-memory copy only).
    Dim ms As Object
    Set ms = doc.ModelSpace
    Dim i As Long
    Dim ent As Object
    For i = ms.Count - 1 To 0 Step -1
        Set ent = ms.Item(i)
        If Not ent Is Nothing Then
            If ent.ObjectName = "AcDbText" Or ent.ObjectName = "AcDbMText" Then
                If Not LayerInList(ent.Layer, TEXT_LAYER) Then ent.Delete
            End If
        End If
        Set ent = Nothing
    Next i

    ' Scratch layer becomes current; TXTEXP output collects there.
    Dim scratch As Object
    Set scratch = doc.Layers.Add(SCRATCH_LAYER)
    If scratch Is Nothing Then Exit Function
    Set doc.ActiveLayer = scratch

    ' Synchronous; without Express Tools this is an unknown command, the text
    ' survives on its own layers, and the stroke font renders it instead.
    doc.SendCommand "._TXTEXP" & vbCr & "_ALL" & vbCr & vbCr

    PrepareAndExplodeText = SCRATCH_LAYER
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Open the SLDPRT, apply the harvested text to the top face, rebuild and save in
' place (overwriting the same file). Returns True on success.
'------------------------------------------------------------------------------
Private Function StampPart(ByVal swApp As Object, _
                           ByVal partPath As String, _
                           ByRef r As TFileResult) As Boolean
    Dim swModel As Object
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

    ' Selection and sketch-text APIs act on the ACTIVE document - make sure
    ' the freshly opened part is it (guarded; failing means it already was).
    On Error Resume Next
    swApp.ActivateDoc3 swModel.GetTitle, False, SW_ACTIVATE_NO_REBUILD, errs
    On Error GoTo errHandler

    Dim placed As Long
    Dim detail As String
    Dim applied As Boolean
    applied = TextMarking.ApplyTextMarks(swModel, placed, detail)
    If Not applied Then
        r.Message = AppendMsg(r.Message, "Marking was incomplete; the part was" & _
                    " closed without saving (" & detail & ").")
        swApp.CloseDoc swModel.GetTitle
        Set swModel = Nothing
        StampPart = False
        Exit Function
    End If

    swModel.ForceRebuild3 False

    Dim saveStart As Date
    saveStart = Now
    Dim savedOK As Boolean
    savedOK = swModel.Save3(SW_SAVE_SILENT, errs, warns)
    If Not savedOK Then
        ' Same quirk as SaveAs: trust a freshly rewritten file over the flag.
        If FreshFileOnDisk(partPath, saveStart) Then
            savedOK = True
            r.Message = AppendMsg(r.Message, "Note: Save3 returned False" & _
                        " (error " & errs & ", warning " & warns & ") but" & _
                        " the part was rewritten - treated as saved.")
        Else
            r.Message = AppendMsg(r.Message, "Save failed (error " & errs & _
                                  ", warning " & warns & ").")
        End If
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
