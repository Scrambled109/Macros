Attribute VB_Name = "Main"
'==============================================================================
' Main.bas
'------------------------------------------------------------------------------
' Entry point and orchestration for the CAD Batch Converter.
'
'   RunBatch()  <-- run this
'
' Flow
' ----
'   1. Ensure the FilteredDWGs and staging folders exist.
'   2. Open the log.
'   3. Enumerate every DWG in the source folder (snapshotted up front so the
'      Dir() cursor is not disturbed by the two applications we then launch).
'   4. Connect to AutoCAD and SolidWorks once, and reuse both for the whole run.
'   5. For each file: filter in AutoCAD, then import + extrude + save in
'      SolidWorks. Every file is wrapped in its own error handler so a single
'      bad drawing can never stop the batch.
'   6. Log a per-file result block and a running progress line.
'   7. Log a summary, optionally shut the applications down, close the log.
'
' Requires references to the AutoCAD 2026 and SolidWorks 2025 type libraries.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Main entry point. Fully unattended once started.
'------------------------------------------------------------------------------
Public Sub RunBatch()
    Dim runStart As Double
    runStart = NowSeconds()

    If Len(SOURCE_FOLDER) = 0 Or Len(FILTERED_FOLDER) = 0 Or _
       Len(OUTPUT_FOLDER) = 0 Then
        MsgBox "CAD batch folders are not configured." & vbCrLf & vbCrLf & _
               "Launch this macro from the Engineering Job Assistant, or " & _
               "configure SourceFolder, FilteredFolder, and OutputFolder " & _
               "for the current Windows user.", vbCritical, _
               "CAD Batch Converter"
        Exit Sub
    End If

    ' --- Folders + log ------------------------------------------------------
    EnsureFolder FILTERED_FOLDER
    EnsureFolder OUTPUT_FOLDER
    OpenLog OUTPUT_FOLDER

    WriteLog ""
    WriteLog String(67, "=")
    WriteLog "CAD BATCH CONVERTER - run started " & TimeStamp()
    WriteLog "Source folder   : " & SOURCE_FOLDER
    WriteLog "Filtered folder : " & FILTERED_FOLDER
    WriteLog "Staging folder  : " & OUTPUT_FOLDER
    WriteLog "Target layer    : " & TARGET_LAYER
    WriteLog "Extrude depth   : " & Format$(EXTRUDE_DEPTH_METERS, "0.00000") & " m"
    WriteLog String(67, "=")

    ' --- Verify the source folder exists ------------------------------------
    If Not FolderExists(SOURCE_FOLDER) Then
        WriteLog "FATAL: source folder not found. Nothing to do."
        CloseLog
        Exit Sub
    End If

    ' --- Snapshot the DWG list ----------------------------------------------
    Dim files() As String
    Dim fileCount As Long
    fileCount = CollectFiles(SOURCE_FOLDER, DWG_FILESPEC, files)
    If fileCount = 0 Then
        WriteLog "No DWG files found in the source folder. Nothing to do."
        CloseLog
        Exit Sub
    End If
    WriteLog "Found " & fileCount & " DWG file(s) to process."

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

    ' --- Process every file -------------------------------------------------
    Dim okCount As Long
    Dim failCount As Long
    Dim i As Long
    Dim result As TFileResult

    For i = 0 To fileCount - 1
        ClearResult result
        ProcessOneFile acadApp, swApp, files(i), result
        LogResult result

        If result.SaveOK Then
            okCount = okCount + 1
        Else
            failCount = failCount + 1
        End If

        WriteLog "Progress: " & (i + 1) & " / " & fileCount & _
                 "   (ok=" & okCount & ", fail=" & failCount & ")"
    Next i

    ' --- Summary ------------------------------------------------------------
    WriteLog String(67, "=")
    WriteLog "RUN COMPLETE " & TimeStamp()
    WriteLog "Processed: " & fileCount & "   OK: " & okCount & _
             "   Failed: " & failCount
    WriteLog "Total time: " & Format$(ElapsedSince(runStart), "0.0") & " s"
    WriteLog String(67, "=")

    ' --- Cleanup ------------------------------------------------------------
    ShutdownApps acadApp, swApp
    CloseLog

    If SHOW_SUMMARY_DIALOG Then
        MsgBox "Batch complete." & vbCrLf & _
               "Processed: " & fileCount & vbCrLf & _
               "OK: " & okCount & vbCrLf & _
               "Failed: " & failCount, vbInformation, "CAD Batch Converter"
    End If
End Sub

'------------------------------------------------------------------------------
' Process a single DWG end to end. Self-contained error handling guarantees the
' loop in RunBatch always continues to the next file.
'------------------------------------------------------------------------------
Private Sub ProcessOneFile(ByVal acadApp As AcadApplication, _
                           ByVal swApp As SldWorks.SldWorks, _
                           ByVal srcPath As String, _
                           ByRef r As TFileResult)
    Dim t As Double
    t = NowSeconds()
    On Error GoTo errHandler

    Dim baseName As String
    baseName = GetBaseName(srcPath)
    r.FileName = baseName & ".dwg"

    Dim filteredPath As String
    Dim outPath As String
    filteredPath = FILTERED_FOLDER & baseName & ".dwg"
    outPath = OUTPUT_FOLDER & baseName & ".SLDPRT"

    ' Steps 1-3: AutoCAD filter.
    r.AutoCadOK = FilterDwg(acadApp, srcPath, filteredPath, r)
    If Not r.AutoCadOK Then GoTo finish

    ' Steps 4-9: SolidWorks import, extrude, save.
    ImportAndExtrude swApp, filteredPath, outPath, r

    ' WORKAROUND: if the DWG-import route did not deliver a saved part (the
    ' unattended import has proven flaky - edit-mode/2D-to-3D state, blank or
    ' unselectable "Model" sketch), rebuild the outline as a NATIVE SolidWorks
    ' sketch from coordinates read straight out of AutoCAD and extrude that.
    ' No DWG import is involved at all on this path.
    If Not r.SaveOK Then
        Dim segs() As TSegment
        Dim segCount As Long
        If HarvestOutline(acadApp, filteredPath, segs, segCount, r) Then
            BuildAndExtrudeNative swApp, segs, segCount, outPath, r
        End If
    End If

finish:
    r.ElapsedSeconds = ElapsedSince(t)
    Exit Sub

errHandler:
    ' Anything unexpected is recorded; the batch keeps going.
    If Len(r.Message) = 0 Then r.Message = "Fatal per-file error: " & Err.Description
    r.ElapsedSeconds = ElapsedSince(t)
End Sub

'------------------------------------------------------------------------------
' Optionally quit the applications, then release both references. Guarded so a
' failure here never masks the batch result. Public so the text-stamp pass can
' reuse it.
'------------------------------------------------------------------------------
Public Sub ShutdownApps(ByRef acadApp As AcadApplication, _
                        ByRef swApp As SldWorks.SldWorks)
    On Error Resume Next
    If QUIT_APPS_ON_FINISH Then
        If Not acadApp Is Nothing Then acadApp.Quit
        If Not swApp Is Nothing Then swApp.ExitApp
    End If
    Set acadApp = Nothing
    Set swApp = Nothing
    On Error GoTo 0
End Sub
