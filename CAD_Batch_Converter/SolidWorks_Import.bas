Attribute VB_Name = "SolidWorks_Import"
'==============================================================================
' SolidWorks_Import.bas
'------------------------------------------------------------------------------
' Everything that touches SolidWorks (steps 4-9):
'   * ConnectSolidWorks   - attach to / launch SolidWorks 2025
'   * ImportAndExtrude    - full per-file SolidWorks workflow:
'                             import DWG as a new-part 2D sketch -> locate the
'                             imported sketch -> detect open contours ->
'                             best-effort repair -> blind extrude -> rebuild ->
'                             SaveAs SLDPRT -> close
'   * FindImportedSketch / SketchHasOpenContours / ExtrudeSketch - helpers
'
' Requires a reference to the SolidWorks 2025 Type Library.
'
' Design notes
' ------------
' * The DWG is imported with LoadFile4 + an IImportDxfDwgData object obtained
'   from GetImportFileData. Supplying that object suppresses the interactive
'   DWG/DXF import wizard, so the run stays unattended.
' * CRITICAL: the DWG DEFAULT import method is "create new DRAWING". To get a
'   part, ImportMethod - a per-sheet INDEXED property, ImportMethod(sheetName) -
'   must be explicitly set to swImportDxfDwg_ImportToPartSketch (swconst.tlb).
'   Setting it is therefore NOT guarded: if it cannot be set, the file fails
'   loudly instead of silently converting to a drawing. After LoadFile4 the
'   document type is verified to be a part (swDocPART); a drawing is closed
'   and reported as a failure.
' * Optional import flags (merge coincident points, no dimensions) are set
'   defensively and individually guarded so a member that differs between
'   service packs cannot abort the batch.
' * Open contours are detected geometrically by pairing up sketch-segment end
'   points: any end point that is not shared by a second segment is a free end,
'   which means the profile is open. This is logged; small gaps are already
'   merged during import, and if a blind extrude still fails the reason is
'   recorded and the batch moves on.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Attach to a running SolidWorks instance, or start a fresh one. Versioned
' ProgID first, generic fallback second. Returns Nothing on total failure.
'------------------------------------------------------------------------------
Public Function ConnectSolidWorks() As SldWorks.SldWorks
    Dim app As SldWorks.SldWorks

    On Error Resume Next
    Set app = GetObject(, SW_PROGID_VERSIONED)
    If app Is Nothing Then Set app = GetObject(, SW_PROGID_GENERIC)
    If app Is Nothing Then Set app = CreateObject(SW_PROGID_VERSIONED)
    If app Is Nothing Then Set app = CreateObject(SW_PROGID_GENERIC)
    On Error GoTo 0

    If Not app Is Nothing Then
        On Error Resume Next
        app.Visible = APP_VISIBLE
        On Error GoTo 0
    End If

    Set ConnectSolidWorks = app
End Function

'------------------------------------------------------------------------------
' Steps 4-9 for a single filtered DWG. Returns True only when the SLDPRT was
' saved. Fills r.ImportOK / r.ExtrudeOK / r.SaveOK / r.OpenContour / r.Message.
'------------------------------------------------------------------------------
Public Function ImportAndExtrude(ByVal swApp As SldWorks.SldWorks, _
                                 ByVal dwgPath As String, _
                                 ByVal outPath As String, _
                                 ByRef r As TFileResult) As Boolean

    Dim swModel As SldWorks.ModelDoc2
    Dim errs As Long
    On Error GoTo errHandler

    ' --- Step 4: import the DWG as a new-part 2D sketch ---------------------
    Dim importData As SldWorks.ImportDxfDwgData
    Set importData = swApp.GetImportFileData(dwgPath)
    If importData Is Nothing Then
        r.Message = "GetImportFileData returned nothing - cannot import as a part."
        ImportAndExtrude = False
        Exit Function
    End If

    ' The one setting that decides part vs drawing. Never guarded, never
    ' defaulted: without it SolidWorks converts the DWG to a new DRAWING.
    If Not SetPartImportMethod(importData, dwgPath) Then
        r.Message = "Could not set ImportMethod = swImportDxfDwg_ImportToPartSketch" & _
                    " - aborting this file rather than importing it as a drawing."
        ImportAndExtrude = False
        Exit Function
    End If

    ConfigureImportOptions importData, dwgPath

    Set swModel = swApp.LoadFile4(dwgPath, SW_IMPORT_ARGS, importData, errs)
    If swModel Is Nothing Then
        r.Message = "SolidWorks import failed (LoadFile4 error code " & errs & ")."
        ImportAndExtrude = False
        Exit Function
    End If

    ' Verify a PART actually came back; a drawing is useless downstream.
    If swModel.GetType <> SW_DOC_PART Then
        r.Message = "DWG imported as document type " & swModel.GetType & _
                    IIf(swModel.GetType = SW_DOC_DRAWING, " (a DRAWING)", "") & _
                    " instead of a part - ImportMethod was not honoured."
        CloseModel swApp, swModel
        ImportAndExtrude = False
        Exit Function
    End If
    r.ImportOK = True

    ' --- Step 5: locate the imported sketch ---------------------------------
    Dim sketchName As String
    sketchName = FindImportedSketch(swModel)
    If Len(sketchName) = 0 Then
        r.Message = "Import succeeded but no sketch was found in the part."
        CloseModel swApp, swModel
        ImportAndExtrude = False
        Exit Function
    End If

    ' --- Step 6: detect open contours + best-effort repair ------------------
    r.OpenContour = SketchHasOpenContours(swModel, sketchName)
    RepairSketch swModel, sketchName

    ' --- Step 7: blind extrude, merge result, normal direction --------------
    If Not ExtrudeSketch(swModel, sketchName, EXTRUDE_DEPTH_METERS) Then
        r.Message = "Extrusion failed" & _
                    IIf(r.OpenContour, " - open contour detected.", ".")
        CloseModel swApp, swModel
        ImportAndExtrude = False
        Exit Function
    End If
    r.ExtrudeOK = True

    ' --- Step 8: rebuild and save the SLDPRT (overwrite) --------------------
    ' NB: part-marking text is added by a separate pass (RunTextStamp), not here.
    swModel.ForceRebuild3 False

    Dim warns As Long
    Dim savedOK As Boolean
    savedOK = swModel.Extension.SaveAs(outPath, SW_SAVEAS_CURRENT, _
                                       SW_SAVE_SILENT, Nothing, errs, warns)
    r.SaveOK = savedOK
    If Not savedOK Then
        r.Message = "SaveAs failed (error " & errs & ", warning " & warns & ")."
    End If

    ' --- Step 9: close the document -----------------------------------------
    CloseModel swApp, swModel

    ImportAndExtrude = savedOK
    Exit Function

errHandler:
    r.Message = "SolidWorks: " & Err.Description
    On Error Resume Next
    If Not swModel Is Nothing Then CloseModel swApp, swModel
    On Error GoTo 0
    ImportAndExtrude = False
End Function

'------------------------------------------------------------------------------
' Force the import method to "part sketch". ImportMethod is an INDEXED property
' keyed by sheet name - "" (all/default) per the community-verified examples,
' with the file path as the documented fallback index. The plain assignment
' d.ImportMethod = x compiles late-bound but fails at runtime, which is exactly
' how the old code silently fell back to the DWG default of "new drawing".
' Returns True only if one of the two forms was accepted.
'------------------------------------------------------------------------------
Private Function SetPartImportMethod(ByVal d As SldWorks.ImportDxfDwgData, _
                                     ByVal dwgPath As String) As Boolean
    On Error Resume Next

    Err.Clear
    d.ImportMethod("") = swImportDxfDwg_ImportToPartSketch
    SetPartImportMethod = (Err.Number = 0)

    If Not SetPartImportMethod Then
        Err.Clear
        d.ImportMethod(dwgPath) = swImportDxfDwg_ImportToPartSketch
        SetPartImportMethod = (Err.Number = 0)
    End If

    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Apply the optional (nice-to-have) import options. These only exist for the
' part-sketch import path and are individually guarded because member names
' vary slightly across service packs - an unknown member is skipped, never
' fatal. Both sheet-name index forms are attempted, mirroring the method above.
'------------------------------------------------------------------------------
Private Sub ConfigureImportOptions(ByVal d As Object, ByVal dwgPath As String)
    If d Is Nothing Then Exit Sub

    On Error Resume Next
    ' Snap coincident end points within IMPORT_MERGE_METERS (closes tiny gaps).
    d.SetMergePoints "", True, IMPORT_MERGE_METERS
    d.SetMergePoints dwgPath, True, IMPORT_MERGE_METERS
    ' Do NOT import dimensions.
    d.ImportDimensions("") = False
    d.ImportDimensions(dwgPath) = False
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Walk the feature tree and return the name of the (last) 2D sketch, which is
' the geometry produced by the import. If several sketches exist, the imported
' one is the most recently added and therefore the last ProfileFeature found.
'------------------------------------------------------------------------------
Public Function FindImportedSketch(ByVal swModel As SldWorks.ModelDoc2) As String
    Dim feat As SldWorks.Feature
    Dim name As String

    Set feat = swModel.FirstFeature
    Do While Not feat Is Nothing
        If feat.GetTypeName2 = SW_SKETCH_TYPENAME Then name = feat.Name
        Set feat = feat.GetNextFeature
    Loop

    FindImportedSketch = name
End Function

'------------------------------------------------------------------------------
' Geometric open-contour test. Pairs up every segment end point on a tolerance
' grid: a fully closed profile leaves no unpaired end points, so any remainder
' means at least one open contour. Robust across lines/arcs/splines/ellipses
' because it only relies on ISketchSegment.GetStartPoint2 / GetEndPoint2.
'------------------------------------------------------------------------------
Public Function SketchHasOpenContours(ByVal swModel As SldWorks.ModelDoc2, _
                                      ByVal sketchName As String) As Boolean
    On Error GoTo done

    Dim feat As SldWorks.Feature
    Set feat = FindFeatureByName(swModel, sketchName)
    If feat Is Nothing Then Exit Function

    Dim sk As Object                 ' SldWorks.Sketch
    Set sk = feat.GetSpecificFeature2
    If sk Is Nothing Then Exit Function

    Dim vSegs As Variant
    vSegs = sk.GetSketchSegments
    If IsEmpty(vSegs) Then Exit Function
    If Not IsArray(vSegs) Then Exit Function

    ' Toggle collection: an end point is added on first sight and removed on the
    ' second, so whatever remains is unpaired (a free end).
    Dim unpaired As Collection
    Set unpaired = New Collection

    Dim i As Long
    Dim seg As Object                ' SldWorks.SketchSegment
    For i = LBound(vSegs) To UBound(vSegs)
        Set seg = vSegs(i)
        TogglePoint unpaired, seg.GetStartPoint2
        TogglePoint unpaired, seg.GetEndPoint2
        Set seg = Nothing
    Next i

    SketchHasOpenContours = (unpaired.Count > 0)
    Exit Function

done:
    SketchHasOpenContours = False
End Function

'------------------------------------------------------------------------------
' Add a point key on first sight, remove it on the second (pairing coincident
' end points). Keys are quantised onto POINT_TOLERANCE_METERS so near-identical
' coordinates match.
'------------------------------------------------------------------------------
Private Sub TogglePoint(ByRef col As Collection, ByVal pt As Object)
    If pt Is Nothing Then Exit Sub

    On Error Resume Next
    Dim key As String
    key = PointKey(pt.X, pt.Y, pt.Z)
    If Len(key) = 0 Then Exit Sub

    Dim probe As Variant
    Err.Clear                        ' clear stale code (successful calls do not)
    probe = col(key)                 ' raises error if the key is absent
    If Err.Number = 0 Then
        col.Remove key               ' second sighting -> paired, remove it
    Else
        Err.Clear
        col.Add key, key             ' first sighting -> keep it
    End If
    On Error GoTo 0
End Sub

' Build a tolerance-quantised string key for a 3D point.
Private Function PointKey(ByVal x As Double, _
                          ByVal y As Double, _
                          ByVal z As Double) As String
    On Error GoTo fail
    PointKey = CStr(CLng(x / POINT_TOLERANCE_METERS)) & "|" & _
               CStr(CLng(y / POINT_TOLERANCE_METERS)) & "|" & _
               CStr(CLng(z / POINT_TOLERANCE_METERS))
    Exit Function
fail:
    PointKey = vbNullString          ' coordinate out of Long range -> skip it
End Function

'------------------------------------------------------------------------------
' Best-effort sketch repair. Small gaps are already merged during import; this
' simply (re)selects the sketch so it is the active target for the extrude and
' clears any stray selection. Kept deliberately non-destructive.
'------------------------------------------------------------------------------
Private Sub RepairSketch(ByVal swModel As SldWorks.ModelDoc2, _
                         ByVal sketchName As String)
    On Error Resume Next
    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 sketchName, "SKETCH", 0#, 0#, 0#, _
                                  False, 0, Nothing, 0
    swModel.ClearSelection2 True
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Select the sketch and create a single-ended, blind, merged, normal extrusion
' of the requested depth. Returns True only if a feature was created (a Nothing
' return means the profile could not be extruded, e.g. an open contour).
'------------------------------------------------------------------------------
Public Function ExtrudeSketch(ByVal swModel As SldWorks.ModelDoc2, _
                              ByVal sketchName As String, _
                              ByVal depth As Double) As Boolean
    On Error GoTo errHandler

    swModel.ClearSelection2 True
    Dim selected As Boolean
    selected = swModel.Extension.SelectByID2(sketchName, "SKETCH", _
                                             0#, 0#, 0#, False, 0, Nothing, 0)
    If Not selected Then
        ExtrudeSketch = False
        Exit Function
    End If

    ' FeatureExtrusion3 argument map (single-ended blind boss, merge = True):
    '   Sd, Flip, Dir                     -> True, False, False
    '   T1, T2 (end conditions)           -> Blind, Blind
    '   D1, D2 (depths)                   -> depth, 0
    '   Dchk1, Dchk2 (draft on)           -> False, False
    '   Ddir1, Ddir2 (draft outward)      -> False, False
    '   Dang1, Dang2 (draft angle)        -> 0, 0
    '   OffsetReverse1/2                  -> False, False
    '   TranslateSurface1/2               -> False, False
    '   Merge                             -> True
    '   UseFeatScope, UseAutoSelect       -> True, True
    '   T0 (start condition)              -> Sketch plane
    '   StartOffset, FlipStartOffset      -> 0, False
    Dim swFeat As SldWorks.Feature
    Set swFeat = swModel.FeatureManager.FeatureExtrusion3( _
        True, False, False, _
        SW_END_COND_BLIND, SW_END_COND_BLIND, _
        depth, 0#, _
        False, False, _
        False, False, _
        0#, 0#, _
        False, False, _
        False, False, _
        True, _
        True, True, _
        SW_START_SKETCH_PLANE, 0#, False)

    swModel.ClearSelection2 True
    ExtrudeSketch = Not (swFeat Is Nothing)
    Exit Function

errHandler:
    On Error Resume Next
    swModel.ClearSelection2 True
    On Error GoTo 0
    ExtrudeSketch = False
End Function

'------------------------------------------------------------------------------
' Return the feature whose name matches, or Nothing.
'------------------------------------------------------------------------------
Private Function FindFeatureByName(ByVal swModel As SldWorks.ModelDoc2, _
                                   ByVal name As String) As SldWorks.Feature
    Dim feat As SldWorks.Feature
    Set feat = swModel.FirstFeature
    Do While Not feat Is Nothing
        If feat.Name = name Then
            Set FindFeatureByName = feat
            Exit Function
        End If
        Set feat = feat.GetNextFeature
    Loop
End Function

'------------------------------------------------------------------------------
' Close the document without prompting. Uses the current title (which reflects
' the saved file name after SaveAs). Errors are ignored so cleanup is safe.
'------------------------------------------------------------------------------
Private Sub CloseModel(ByVal swApp As SldWorks.SldWorks, _
                       ByRef swModel As SldWorks.ModelDoc2)
    On Error Resume Next
    Dim title As String
    title = swModel.GetTitle
    swApp.CloseDoc title
    Set swModel = Nothing
    On Error GoTo 0
End Sub
