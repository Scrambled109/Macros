Attribute VB_Name = "NativeSketch"
'==============================================================================
' NativeSketch.bas
'------------------------------------------------------------------------------
' The DWG-import WORKAROUND path: skip SolidWorks' DWG import entirely.
'
'   * HarvestOutline         - read the outline geometry (lines, polylines
'                              with arc bulges, arcs, circles) straight out of
'                              the filtered DWG via the AutoCAD COM model, as
'                              plain coordinate data in DWG units.
'   * BuildAndExtrudeNative  - create a brand-new part from the default part
'                              template, redraw that geometry as a NATIVE
'                              SolidWorks sketch on the front plane, extrude,
'                              and save the SLDPRT.
'
' Why this exists
' ---------------
' LoadFile4's DWG-to-part-sketch import proved unreliable unattended (sketch
' left in edit mode / 2D-to-3D state, unselectable "Model" sketch, blank
' geometry), while the same profiles extrude fine by hand. A sketch DRAWN by
' the SolidWorks sketch API has none of those problems: it is native geometry,
' fully owned by the part, and extrudes exactly like one drawn by hand.
' Main.ProcessOneFile calls this path automatically whenever the import route
' fails to deliver a saved part.
'
' Geometry support
' ----------------
' Lines, lightweight + heavy 2D polylines (bulges converted to true arcs),
' arcs, circles. That covers everything a "CUT-OUTSIDE STRAIGHT" style profile
' layer normally holds. Splines/ellipses are reported by name in the log so an
' odd drawing is diagnosable, and the file fails loudly instead of silently
' producing a wrong outline.
'
' Coordinates are harvested in raw DWG units and scaled once, on the
' SolidWorks side, by DWG_UNITS_TO_METERS (the same constant the text pass
' uses - inches by default).
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Read every supported entity in the filtered DWG's model space into segs().
' Opens the DWG read-only and closes it again. Returns True when at least one
' segment was harvested; on False, r.Message says why.
'------------------------------------------------------------------------------
Public Function HarvestOutline(ByVal acadApp As Object, _
                               ByVal dwgPath As String, _
                               ByRef segs() As TSegment, _
                               ByRef segCount As Long, _
                               ByRef r As TFileResult) As Boolean
    Dim doc As Object
    On Error GoTo errHandler

    segCount = 0
    ReDim segs(0 To 63)

    Set doc = acadApp.Documents.Open(dwgPath, True)      ' read-only

    ' Late-bound on purpose, and point properties are ALWAYS copied into a
    ' Variant before indexing: "ent.StartPoint(0)" style access makes VBA call
    ' the property WITH an argument, which AutoCAD rejects with error 451
    ' ("Property let procedure not defined and property get procedure did not
    ' return an object").
    Dim ent As Object
    Dim p1 As Variant
    Dim p2 As Variant
    Dim ctr As Variant
    Dim unsupported As String
    For Each ent In doc.ModelSpace
        Select Case ent.ObjectName
            Case "AcDbLine"
                p1 = ent.StartPoint
                p2 = ent.EndPoint
                AddLineSeg segs, segCount, p1(0), p1(1), p2(0), p2(1)
            Case "AcDbPolyline"                          ' lightweight polyline
                AddPolyline segs, segCount, ent, 2
            Case "AcDb2dPolyline"                        ' heavy 2D polyline
                AddPolyline segs, segCount, ent, 3
            Case "AcDb3dPolyline"                        ' straight segments only
                AddPolyline segs, segCount, ent, 3
            Case "AcDbArc"
                AddArcEntity segs, segCount, ent
            Case "AcDbCircle"
                ctr = ent.Center
                AddCircleSeg segs, segCount, ctr(0), ctr(1), ent.Radius
            Case Else
                ' Report each unsupported type once, by name.
                If InStr(1, unsupported, ent.ObjectName, vbTextCompare) = 0 Then
                    If Len(unsupported) > 0 Then unsupported = unsupported & ", "
                    unsupported = unsupported & ent.ObjectName
                End If
        End Select
        Set ent = Nothing
    Next ent

    doc.Close False
    Set doc = Nothing

    If Len(unsupported) > 0 Then
        r.Message = AppendMsg(r.Message, "Native-sketch workaround: DWG holds" & _
                    " unsupported entity type(s): " & unsupported & _
                    "; refusing to build an incomplete part.")
        HarvestOutline = False
        Exit Function
    End If

    If segCount = 0 Then
        r.Message = AppendMsg(r.Message, "Native-sketch workaround: no" & _
                    " supported outline geometry found in the filtered DWG.")
        HarvestOutline = False
    Else
        HarvestOutline = True
    End If
    Exit Function

errHandler:
    r.Message = AppendMsg(r.Message, "Native-sketch workaround (AutoCAD): " & _
                          Err.Description)
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    Set doc = Nothing
    On Error GoTo 0
    HarvestOutline = False
End Function

'------------------------------------------------------------------------------
' Create a new part from the default part template, redraw segs() as a native
' sketch on the front plane, extrude it EXTRUDE_DEPTH_METERS, save to outPath
' and close. Fills the same TFileResult fields as the import path so the log
' block reads identically.
'------------------------------------------------------------------------------
Public Function BuildAndExtrudeNative(ByVal swApp As SldWorks.SldWorks, _
                                      ByRef segs() As TSegment, _
                                      ByVal segCount As Long, _
                                      ByVal outPath As String, _
                                      ByRef r As TFileResult) As Boolean
    Dim swModel As SldWorks.ModelDoc2
    On Error GoTo errHandler

    ' --- New empty part from the user's default part template ---------------
    Dim tmpl As String
    tmpl = swApp.GetUserPreferenceStringValue(SW_PREF_DEFAULT_PART_TEMPLATE)
    If Len(tmpl) = 0 Then
        r.Message = AppendMsg(r.Message, "Native-sketch workaround: no default" & _
                    " part template is configured in SolidWorks (Tools >" & _
                    " Options > Default Templates).")
        BuildAndExtrudeNative = False
        Exit Function
    End If

    Set swModel = swApp.NewDocument(tmpl, 0, 0#, 0#)
    If swModel Is Nothing Then
        r.Message = AppendMsg(r.Message, "Native-sketch workaround:" & _
                    " NewDocument returned nothing (template: " & tmpl & ").")
        BuildAndExtrudeNative = False
        Exit Function
    End If

    ' --- Open a sketch on the front plane (first RefPlane in the tree, which
    '     is locale-independent, unlike the plane's display name) -------------
    Dim plane As SldWorks.Feature
    Set plane = FindFirstPlane(swModel)
    If plane Is Nothing Then
        r.Message = AppendMsg(r.Message, "Native-sketch workaround: no" & _
                    " reference plane found in the new part.")
        CloseModel swApp, swModel
        BuildAndExtrudeNative = False
        Exit Function
    End If

    plane.Select2 False, 0
    Dim skm As SldWorks.SketchManager
    Set skm = swModel.SketchManager
    skm.InsertSketch True                    ' enter a sketch on the plane

    ' --- Redraw the harvested geometry, scaled DWG units -> meters ----------
    DrawSegments skm, segs, segCount, DWG_UNITS_TO_METERS
    skm.InsertSketch True                    ' exit the sketch
    r.ImportOK = True                        ' geometry is in SolidWorks

    ' --- Extrude (same helper as the import path, incl. contour fallback) ---
    Dim sketchName As String
    sketchName = FindImportedSketch(swModel)
    If Len(sketchName) = 0 Then
        r.Message = AppendMsg(r.Message, "Native-sketch workaround: the drawn" & _
                    " sketch was not found in the feature tree.")
        CloseModel swApp, swModel
        BuildAndExtrudeNative = False
        Exit Function
    End If

    Dim extrudeReason As String
    If Not ExtrudeSketch(swModel, sketchName, EXTRUDE_DEPTH_METERS, extrudeReason) Then
        r.Message = AppendMsg(r.Message, "Native-sketch workaround: extrusion" & _
                    " failed: " & extrudeReason)
        CloseModel swApp, swModel
        BuildAndExtrudeNative = False
        Exit Function
    End If
    r.ExtrudeOK = True

    ' --- Rebuild, save, close ------------------------------------------------
    swModel.ForceRebuild3 False

    Dim savedOK As Boolean
    savedOK = SavePart(swModel, outPath, r)
    r.SaveOK = savedOK
    If savedOK Then
        r.Message = AppendMsg(r.Message, "Recovered via native-sketch" & _
                    " workaround (outline redrawn from AutoCAD geometry).")
    End If

    CloseModel swApp, swModel
    BuildAndExtrudeNative = savedOK
    Exit Function

errHandler:
    r.Message = AppendMsg(r.Message, "Native-sketch workaround (SolidWorks): " & _
                          Err.Description)
    On Error Resume Next
    If Not swModel Is Nothing Then CloseModel swApp, swModel
    On Error GoTo 0
    BuildAndExtrudeNative = False
End Function

'==============================================================================
' Segment helpers - Public: TextMarking reuses them to harvest and draw the
' reference marking geometry (frame lines etc.) on the finished parts.
'==============================================================================

'------------------------------------------------------------------------------
' Draw segs() into the ACTIVE sketch, scaled by unitScale (DWG units ->
' meters). AddToDB skips inference/snapping so coordinates land exactly as
' given. Returns the number of segments actually created (verified).
'------------------------------------------------------------------------------
Public Function DrawSegments(ByVal skm As SldWorks.SketchManager, _
                             ByRef segs() As TSegment, _
                             ByVal segCount As Long, _
                             ByVal unitScale As Double, _
                             Optional ByVal colorRGB As Long = -1) As Long
    On Error GoTo done
    Dim k As Double
    k = unitScale

    Dim prevAddToDB As Boolean
    prevAddToDB = skm.AddToDB
    skm.AddToDB = True
    skm.DisplayWhenAdded = False

    Dim made As Long
    Dim obj As Object
    Dim i As Long
    For i = 0 To segCount - 1
        Set obj = Nothing
        With segs(i)
            Select Case .Kind
                Case SEG_LINE
                    Set obj = skm.CreateLine(.X1 * k, .Y1 * k, 0#, _
                                             .X2 * k, .Y2 * k, 0#)
                Case SEG_ARC
                    Set obj = skm.CreateArc(.CX * k, .CY * k, 0#, _
                                            .X1 * k, .Y1 * k, 0#, _
                                            .X2 * k, .Y2 * k, 0#, .Direction)
                Case SEG_CIRCLE
                    Set obj = skm.CreateCircleByRadius(.CX * k, .CY * k, 0#, _
                                                       .Radius * k)
            End Select
        End With
        If Not obj Is Nothing Then
            made = made + 1
            If colorRGB >= 0 Then
                On Error Resume Next
                obj.Color = colorRGB
                On Error GoTo done
            End If
        End If
    Next i

done:
    On Error Resume Next
    skm.AddToDB = prevAddToDB
    skm.DisplayWhenAdded = True
    On Error GoTo 0
    DrawSegments = made
End Function

' Append a straight segment, skipping zero-length ones.
Public Sub AddLineSeg(ByRef segs() As TSegment, ByRef segCount As Long, _
                      ByVal x1 As Double, ByVal y1 As Double, _
                      ByVal x2 As Double, ByVal y2 As Double)
    If Abs(x2 - x1) < GEOM_EPS_DWG And Abs(y2 - y1) < GEOM_EPS_DWG Then Exit Sub
    EnsureCapacity segs, segCount
    With segs(segCount)
        .Kind = SEG_LINE
        .X1 = x1: .Y1 = y1: .X2 = x2: .Y2 = y2
    End With
    segCount = segCount + 1
End Sub

'------------------------------------------------------------------------------
' Convert one polyline into line/arc segments. stride = 2 for a lightweight
' polyline (flat X,Y coordinate array), 3 for heavy 2D/3D polylines (X,Y,Z).
' A non-zero bulge on a vertex turns that segment into a true arc.
'------------------------------------------------------------------------------
Public Sub AddPolyline(ByRef segs() As TSegment, ByRef segCount As Long, _
                        ByVal pl As Object, ByVal stride As Long)
    Dim coords As Variant
    coords = pl.Coordinates

    Dim nVerts As Long
    nVerts = (UBound(coords) - LBound(coords) + 1) \ stride
    If nVerts < 2 Then Exit Sub

    Dim isClosed As Boolean
    On Error Resume Next
    isClosed = pl.Closed
    On Error GoTo 0

    Dim segTotal As Long
    segTotal = nVerts - 1
    If isClosed Then segTotal = nVerts

    Dim i As Long
    Dim i2 As Long
    Dim b As Double
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    For i = 0 To segTotal - 1
        i2 = (i + 1) Mod nVerts
        x1 = coords(LBound(coords) + i * stride)
        y1 = coords(LBound(coords) + i * stride + 1)
        x2 = coords(LBound(coords) + i2 * stride)
        y2 = coords(LBound(coords) + i2 * stride + 1)

        b = 0#
        On Error Resume Next                 ' 3D polylines have no bulges
        b = pl.GetBulge(i)
        On Error GoTo 0

        If Abs(b) < GEOM_EPS_DWG Then
            AddLineSeg segs, segCount, x1, y1, x2, y2
        Else
            AddBulgeArc segs, segCount, x1, y1, x2, y2, b
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' Convert a polyline bulge segment into a true arc. The bulge is
' tan(includedAngle / 4), positive = counter-clockwise from start to end.
' Standard identities: sagitta s = |b| * d / 2, radius = d * (1 + b^2) / (4|b|),
' and the centre sits on the chord's normal at (radius - sagitta) from the
' midpoint, on the bulge side.
'------------------------------------------------------------------------------
Private Sub AddBulgeArc(ByRef segs() As TSegment, ByRef segCount As Long, _
                        ByVal x1 As Double, ByVal y1 As Double, _
                        ByVal x2 As Double, ByVal y2 As Double, _
                        ByVal b As Double)
    Dim d As Double
    d = Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
    If d < GEOM_EPS_DWG Then Exit Sub

    Dim ab As Double
    ab = Abs(b)
    Dim sagitta As Double
    sagitta = ab * d / 2#
    Dim radius As Double
    radius = d * (1# + ab * ab) / (4# * ab)

    ' Chord normal (unit), pointing to the centre side for a POSITIVE bulge.
    Dim nx As Double, ny As Double
    nx = -(y2 - y1) / d
    ny = (x2 - x1) / d

    Dim t As Double
    t = radius - sagitta
    If b < 0# Then t = -t

    EnsureCapacity segs, segCount
    With segs(segCount)
        .Kind = SEG_ARC
        .X1 = x1: .Y1 = y1: .X2 = x2: .Y2 = y2
        .CX = (x1 + x2) / 2# + nx * t
        .CY = (y1 + y2) / 2# + ny * t
        .Direction = IIf(b > 0#, 1, -1)
    End With
    segCount = segCount + 1
End Sub

' Append an AutoCAD ARC entity (always counter-clockwise, angles in radians).
Public Sub AddArcEntity(ByRef segs() As TSegment, ByRef segCount As Long, _
                         ByVal ent As Object)
    Dim ctr As Variant                   ' copy point into a Variant, then index
    ctr = ent.Center
    Dim cx As Double, cy As Double, rad As Double
    cx = ctr(0): cy = ctr(1): rad = ent.Radius

    EnsureCapacity segs, segCount
    With segs(segCount)
        .Kind = SEG_ARC
        .CX = cx: .CY = cy
        .X1 = cx + rad * Cos(ent.StartAngle)
        .Y1 = cy + rad * Sin(ent.StartAngle)
        .X2 = cx + rad * Cos(ent.EndAngle)
        .Y2 = cy + rad * Sin(ent.EndAngle)
        .Direction = 1
    End With
    segCount = segCount + 1
End Sub

' Append a full circle.
Public Sub AddCircleSeg(ByRef segs() As TSegment, ByRef segCount As Long, _
                        ByVal cx As Double, ByVal cy As Double, _
                        ByVal radius As Double)
    If radius <= 0# Then Exit Sub
    EnsureCapacity segs, segCount
    With segs(segCount)
        .Kind = SEG_CIRCLE
        .CX = cx: .CY = cy
        .Radius = radius
    End With
    segCount = segCount + 1
End Sub

' Grow segs() geometrically so appends stay cheap.
Public Sub EnsureCapacity(ByRef segs() As TSegment, ByVal needed As Long)
    If needed > UBound(segs) Then
        ReDim Preserve segs(0 To UBound(segs) * 2 + 1)
    End If
End Sub

'==============================================================================
' SolidWorks-side helpers
'==============================================================================

' First reference plane in the tree = the front plane, locale-independent.
Private Function FindFirstPlane(ByVal swModel As SldWorks.ModelDoc2) As SldWorks.Feature
    Dim feat As SldWorks.Feature
    Set feat = swModel.FirstFeature
    Do While Not feat Is Nothing
        If feat.GetTypeName2 = SW_REFPLANE_TYPENAME Then
            Set FindFirstPlane = feat
            Exit Function
        End If
        Set feat = feat.GetNextFeature
    Loop
End Function
