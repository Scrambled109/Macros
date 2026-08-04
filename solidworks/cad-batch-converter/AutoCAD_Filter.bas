Attribute VB_Name = "AutoCAD_Filter"
'==============================================================================
' AutoCAD_Filter.bas
'------------------------------------------------------------------------------
' Everything that touches AutoCAD:
'   * ConnectAutoCAD        - attach to / launch AutoCAD 2026
'   * FilterDwg             - the full per-file AutoCAD workflow (steps 1-3):
'                               open -> strip non-target layers -> AUDIT ->
'                               PURGE -> SaveAs filtered DWG -> close
'   * DeleteNonTargetEntities / UnlockAndReportLayers - helpers
'
' Uses late-bound AutoCAD COM; no AutoCAD type-library reference is required.
'
' Design notes
' ------------
' * Entities are deleted by LAYER only: anything not on TARGET_LAYER is removed,
'   which naturally keeps every line / polyline / arc / circle / spline /
'   ellipse that lives on the cut layer.
' * Model space is walked BACKWARDS by index so deleting an entity never
'   invalidates the iteration (a For Each + Delete loop is unreliable).
' * All work is done through the COM object model - no SendKeys, no UI. The one
'   command-line call (AUDIT) uses the synchronous ActiveX SendCommand, which is
'   the only supported way to invoke AUDIT programmatically.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Attach to a running AutoCAD instance, or start a fresh one. Versioned ProgID
' first, generic fallback second. Returns Nothing on total failure.
'------------------------------------------------------------------------------
Public Function ConnectAutoCAD() As Object
    Dim app As Object

    On Error Resume Next
    Set app = GetObject(, ACAD_PROGID_VERSIONED)
    If app Is Nothing Then Set app = GetObject(, ACAD_PROGID_GENERIC)
    If app Is Nothing Then Set app = CreateObject(ACAD_PROGID_VERSIONED)
    If app Is Nothing Then Set app = CreateObject(ACAD_PROGID_GENERIC)
    On Error GoTo 0

    If Not app Is Nothing Then
        On Error Resume Next
        app.Visible = APP_VISIBLE
        On Error GoTo 0
    End If

    Set ConnectAutoCAD = app
End Function

'------------------------------------------------------------------------------
' Steps 1-3 for a single DWG. Returns True only if the filtered DWG was saved.
' Fills r.LockedLayers (report) and, on failure, r.Message.
'------------------------------------------------------------------------------
Public Function FilterDwg(ByVal acadApp As Object, _
                          ByVal srcPath As String, _
                          ByVal destPath As String, _
                          ByRef r As TFileResult) As Boolean

    Dim doc As Object
    On Error GoTo errHandler

    If StrComp(srcPath, destPath, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 1100, "FilterDwg", _
                  "Source and destination paths are identical; original DWG was not overwritten."
    End If

    ' --- Step 1: open the source drawing ------------------------------------
    Set doc = acadApp.Documents.Open(srcPath)

    ' --- Locked-layer handling (report; optionally unlock) ------------------
    r.LockedLayers = UnlockAndReportLayers(doc)

    ' --- Step 2: verify the target before deleting anything ------------------
    ' Report the actual model-space layers when a mapping contract drifts. The
    ' source document is still closed unsaved on failure, but preflight makes
    ' this diagnostic useful and avoids needlessly deleting the in-memory copy.
    Dim targetCount As Long
    Dim unwantedCount As Long
    targetCount = CountEntitiesByTargetState(doc, True)
    If targetCount = 0 Then
        Err.Raise vbObjectError + 1101, "FilterDwg", _
                  "No model-space entities were found on target layer '" & _
                  TARGET_LAYER & "'. Model-space layers present: " & _
                  ModelSpaceLayerSummary(doc)
    End If

    ' --- Delete everything not on the target layer ---------------------------
    DeleteNonTargetEntities doc
    unwantedCount = CountEntitiesByTargetState(doc, False)

    If unwantedCount > 0 Then
        Err.Raise vbObjectError + 1102, "FilterDwg", _
                  unwantedCount & " non-target model-space entit" & _
                  IIf(unwantedCount = 1, "y remains", "ies remain") & _
                  " after filtering; no filtered DWG was saved."
    End If

    ' --- Step 3a: AUDIT (fix errors) ----------------------------------------
    ' ActiveX SendCommand is synchronous, so PURGE below runs only after AUDIT
    ' has fully completed. "_Y" answers the "Fix any errors detected?" prompt.
    doc.SendCommand "._AUDIT " & vbCr & "_Y" & vbCr

    ' --- Step 3b: PURGE (remove unused named objects) -----------------------
    doc.PurgeAll

    ' --- Step 3c: save the filtered copy and close --------------------------
    ' SaveAsType is omitted so AutoCAD writes its current default DWG version
    ' (2018 file format on AutoCAD 2026), which SolidWorks 2025 imports cleanly.
    doc.SaveAs destPath
    doc.Close False           ' already saved; discard the (identical) buffer
    Set doc = Nothing

    FilterDwg = True
    Exit Function

errHandler:
    r.Message = "AutoCAD: " & Err.Description
    ' Cleanup: force the document closed without saving so the next file starts
    ' from a clean slate. Errors during cleanup are deliberately ignored.
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    Set doc = Nothing
    On Error GoTo 0
    FilterDwg = False
End Function

'------------------------------------------------------------------------------
' Delete every model-space entity that is NOT on the target layer.
' Returns the number of entities deleted. Entities that cannot be deleted
' (e.g. still on a locked layer when unlocking is disabled) are skipped
' silently - the offending layer has already been captured by
' UnlockAndReportLayers for the log.
'------------------------------------------------------------------------------
Private Function DeleteNonTargetEntities(ByVal doc As Object) As Long
    Dim ms As Object
    Set ms = doc.ModelSpace

    Dim i As Long
    Dim deleted As Long
    Dim ent As Object

    For i = ms.Count - 1 To 0 Step -1
        Set ent = ms.Item(i)
        If Not LayerEquals(ent.Layer, TARGET_LAYER) Then
            On Error Resume Next
            ent.Delete
            If Err.Number = 0 Then
                deleted = deleted + 1
            Else
                Err.Clear
            End If
            On Error GoTo 0
        End If
        Set ent = Nothing
    Next i

    Set ms = Nothing
    DeleteNonTargetEntities = deleted
End Function

' Count target or non-target model-space entities after the delete pass. This
' turns a failed delete or a missing target layer into a hard, logged failure
' instead of allowing an empty or contaminated cut profile downstream.
Private Function CountEntitiesByTargetState(ByVal doc As Object, _
                                            ByVal countTarget As Boolean) As Long
    Dim ms As Object
    Set ms = doc.ModelSpace

    Dim i As Long
    Dim isTarget As Boolean
    For i = 0 To ms.Count - 1
        isTarget = LayerEquals(ms.Item(i).Layer, TARGET_LAYER)
        If isTarget = countTarget Then
            CountEntitiesByTargetState = CountEntitiesByTargetState + 1
        End If
    Next i

    Set ms = Nothing
End Function

' Return the unique layer names actually used by model-space entities. This is
' intentionally different from listing doc.Layers, which includes empty seed
' layers and would hide a color-to-layer mapping failure.
Private Function ModelSpaceLayerSummary(ByVal doc As Object) As String
    On Error GoTo failed

    Dim ms As Object
    Set ms = doc.ModelSpace
    If ms.Count = 0 Then
        ModelSpaceLayerSummary = "<model space is empty>"
        Exit Function
    End If

    Dim seen As String
    Dim summary As String
    Dim layerName As String
    Dim key As String
    Dim i As Long
    For i = 0 To ms.Count - 1
        layerName = Trim$(CStr(ms.Item(i).Layer))
        key = "|" & UCase$(layerName) & "|"
        If InStr(1, seen, key, vbBinaryCompare) = 0 Then
            seen = seen & key
            If Len(summary) > 0 Then summary = summary & ", "
            summary = summary & layerName
        End If
    Next i

    If Len(summary) = 0 Then summary = "<no readable layer names>"
    ModelSpaceLayerSummary = summary
    Exit Function

failed:
    ModelSpaceLayerSummary = "<layer inspection failed: " & Err.Description & ">"
End Function

'------------------------------------------------------------------------------
' Scan every layer, build a comma-separated report of the locked ones, and
' (when UNLOCK_LOCKED_LAYERS is True) unlock every locked layer except the
' target layer so its entities become deletable.
'------------------------------------------------------------------------------
Private Function UnlockAndReportLayers(ByVal doc As Object) As String
    Dim lay As Object
    Dim report As String

    For Each lay In doc.Layers
        If lay.Lock Then
            If Len(report) > 0 Then report = report & ", "
            report = report & lay.Name

            If UNLOCK_LOCKED_LAYERS Then
                ' Never disturb the target layer; unlock the rest.
                If Not LayerEquals(lay.Name, TARGET_LAYER) Then
                    On Error Resume Next
                    lay.Lock = False
                    On Error GoTo 0
                End If
            End If
        End If
    Next lay

    Set lay = Nothing
    UnlockAndReportLayers = report
End Function
