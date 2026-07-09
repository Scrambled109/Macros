Attribute VB_Name = "TextMarking"
'==============================================================================
' TextMarking.bas
'------------------------------------------------------------------------------
' Puts the words onto every part WITHOUT trying to import DWG text as extrudable
' geometry (which is the part that was a pain: TEXT/MTEXT does not import as
' clean closed loops, and letters such as A/O/R create nested islands that fight
' the outline extrude).
'
' Instead the text is treated as DATA:
'   * AutoCAD stage  - HarvestTextMarks() reads every TEXT/MTEXT on the
'                      TEXT_LAYER layer(s) BEFORE the filter deletes them,
'                      capturing string + position + height + rotation into an
'                      in-memory list.
'   * SolidWorks stage - ApplyTextMarks() DRAWS each word as single-stroke
'                      line geometry (built-in stick font, see StrokeFor) in a
'                      sketch on the TOP FACE of the extruded part
'                      (auto-detected: the planar face whose normal is +Z, at
'                      the greatest Z), falling back to the Front plane if the
'                      top face cannot be found. Position, height and rotation
'                      all come from the DWG. The sketch is left UN-extruded -
'                      the words are simply "there for modeling", as requested.
'                      (IModelDoc2::InsertSketchText was abandoned: it returned
'                      Nothing for every word on this install, visible or
'                      hidden - stroke geometry via CreateLine is deterministic
'                      and matches what pin-stamp equipment engraves anyway.)
'
' State is held at module level so the existing FilterDwg / ImportAndExtrude
' signatures do not have to change. Call sequence per file:
'       ClearMarks  ->  HarvestTextMarks (AutoCAD)  ->  ApplyTextMarks (SW)
'
' Requires references to the AutoCAD 2026 and SolidWorks 2025 type libraries.
'==============================================================================
Option Explicit

' In-memory list of harvested marks (UDTs cannot live in a Collection, so a
' plain growable array is used).
Private mMarks() As TTextMark
Private mCount As Long

' Stroke-font metrics (see the STROKE-FONT TEXT RENDERING section below).
' NB: module-level Consts must live here in the declarations section - VBA
' does not allow them between procedures.
Private Const FONT_CAP_GRID As Double = 6#     ' grid rows = cap height
Private Const FONT_ADVANCE As Double = 6#      ' pen advance per character
Private Const FONT_FALLBACK_HEIGHT_M As Double = 0.003  ' if the DWG height is 0

' First drawing failure of the current placement attempt, for the log.
Private mLastDrawErr As String

'==============================================================================
' LIST MANAGEMENT
'==============================================================================

' Reset the list for a new file.
Public Sub ClearMarks()
    mCount = 0
    ReDim mMarks(0 To 63)
End Sub

' Number of words currently captured.
Public Function MarkCount() As Long
    MarkCount = mCount
End Function

' Append one captured word to the list, growing the buffer as needed.
Private Sub AddMark(ByVal text As String, ByVal x As Double, ByVal y As Double, _
                    ByVal height As Double, ByVal rotation As Double)
    If mCount > UBound(mMarks) Then
        ReDim Preserve mMarks(0 To UBound(mMarks) + 64)
    End If
    mMarks(mCount).Text = text
    mMarks(mCount).X = x
    mMarks(mCount).Y = y
    mMarks(mCount).Height = height
    mMarks(mCount).Rotation = rotation
    mCount = mCount + 1
End Sub

'==============================================================================
' AUTOCAD STAGE - HARVEST
'==============================================================================

' Read every TEXT / MTEXT entity on the TEXT_LAYER layer(s) from the open
' drawing and store it. TEXT_LAYER may name several layers, comma-separated.
' Reading does not require the layers to be unlocked, and this runs BEFORE
' the filter deletes anything. No-op when TEXT_LAYER is blank.
Public Sub HarvestTextMarks(ByVal doc As AcadDocument)
    If Len(Trim$(TEXT_LAYER)) = 0 Then Exit Sub

    Dim ms As AcadModelSpace
    Set ms = doc.ModelSpace

    Dim i As Long
    Dim ent As Object            ' late-bound: text properties are not on AcadEntity
    For i = 0 To ms.Count - 1
        On Error Resume Next
        Set ent = ms.Item(i)
        If Not ent Is Nothing Then
            If LayerInList(ent.Layer, TEXT_LAYER) Then
                HarvestOne ent
            End If
        End If
        Set ent = Nothing
        On Error GoTo 0
    Next i

    Set ms = Nothing
End Sub

' Capture a single TEXT or MTEXT entity. Anything else on the layer is ignored.
Private Sub HarvestOne(ByVal ent As Object)
    On Error Resume Next

    Dim kind As String
    kind = ent.ObjectName        ' "AcDbText" or "AcDbMText"

    Dim ip As Variant
    Dim s As String

    Select Case kind
        Case "AcDbText"          ' single-line TEXT
            ip = ent.InsertionPoint
            s = ent.TextString
            AddMark s, ip(0), ip(1), ent.Height, ent.Rotation

        Case "AcDbMText"         ' multi-line MTEXT
            ip = ent.InsertionPoint
            s = CleanMText(ent.TextString)
            AddMark s, ip(0), ip(1), ent.Height, ent.RotationAngle
    End Select

    On Error GoTo 0
End Sub

' Strip the common MTEXT inline formatting codes so only the readable string
' remains. Not exhaustive, but covers fonts, colours, height/width/oblique,
' paragraph breaks and grouping braces.
Private Function CleanMText(ByVal raw As String) As String
    Dim s As String
    s = raw

    ' Paragraph / line breaks and non-breaking space -> a single space.
    s = Replace(s, "\P", " ")
    s = Replace(s, "\~", " ")

    ' Remove backslash codes: a backslash + control letter, up to (and
    ' including) its terminating ';' where present.
    s = RemoveBackslashCodes(s)

    ' Drop grouping braces left behind by font/colour blocks.
    s = Replace(s, "{", "")
    s = Replace(s, "}", "")

    ' Un-escape the literal backslash / brace escapes.
    s = Replace(s, "\\", "\")
    s = Replace(s, "\{", "{")
    s = Replace(s, "\}", "}")

    CleanMText = Trim$(s)
End Function

' Helper for CleanMText: walk the string removing "\<letter>....;" formatting
' runs while preserving ordinary characters.
Private Function RemoveBackslashCodes(ByVal s As String) As String
    Dim out As String
    Dim i As Long
    Dim n As Long
    n = Len(s)

    i = 1
    Do While i <= n
        Dim ch As String
        ch = Mid$(s, i, 1)
        If ch = "\" And i < n Then
            Dim nxt As String
            nxt = Mid$(s, i + 1, 1)
            If nxt Like "[A-Za-z]" Then
                ' Formatting code: skip to just past the next ';' (or to the
                ' code letter itself if no ';' follows before the next space).
                Dim semi As Long
                semi = InStr(i, s, ";")
                If semi > 0 Then
                    i = semi + 1
                Else
                    i = i + 2      ' bare code with no terminator
                End If
            Else
                out = out & ch     ' escaped literal (\\ \{ \} \~ handled later)
                i = i + 1
            End If
        Else
            out = out & ch
            i = i + 1
        End If
    Loop

    RemoveBackslashCodes = out
End Function

'==============================================================================
' SOLIDWORKS STAGE - APPLY
'==============================================================================

' Recreate every harvested word as native sketch text on the TOP FACE of the
' extruded part (the planar face whose normal is +Z, at the greatest Z). Falls
' back to the Front plane if the top face cannot be found, so the words are
' never lost. The sketch is left un-consumed. Call AFTER the base extrude and
' BEFORE the save.
'
' placedCount reports how many words actually landed: each InsertSketchText
' call is VERIFIED by the ISketchText object it returns - never trusted
' silently (the old code swallowed a wrong-argument-count error and reported
' OK while placing nothing). Returns True when at least one word was placed
' (and True when there was nothing to place).
Public Function ApplyTextMarks(ByVal swModel As SldWorks.ModelDoc2, _
                               ByRef placedCount As Long, _
                               ByRef detail As String) As Boolean
    placedCount = 0
    detail = vbNullString
    If mCount = 0 Then
        ApplyTextMarks = True
        Exit Function
    End If

    placedCount = PlaceAllWords(swModel, detail)
    ApplyTextMarks = (placedCount > 0)
End Function

'------------------------------------------------------------------------------
' One complete placement attempt: select the top face (Front-plane fallback),
' open a sketch ON it, insert every harvested word, close the sketch. Returns
' the number of words that actually landed.
'
' Every step is verified AND narrated into `detail` so a failure names the
' exact step in the log: which surface was selected, whether the sketch really
' opened (SketchManager.ActiveSketch), how many words drew, and the first
' drawing error if any. If a sketch will not open on the top face, the front
' plane is retried before giving up.
'------------------------------------------------------------------------------
Private Function PlaceAllWords(ByVal swModel As SldWorks.ModelDoc2, _
                               ByRef detail As String) As Long
    Dim placed As Long
    Dim surface As String
    mLastDrawErr = vbNullString
    On Error GoTo errHandler

    ' --- 1) Pick the sketch surface: top face, else the front plane ---------
    If SelectTopFace(swModel) Then
        surface = "top face"
    ElseIf SelectFrontPlane(swModel) Then
        surface = "front plane (no top face found)"
    Else
        detail = "neither the top face nor a reference plane could be selected"
        Exit Function
    End If

    ' --- 2) Open a sketch on it - verified, with a front-plane retry --------
    swModel.SketchManager.InsertSketch True
    If swModel.SketchManager.ActiveSketch Is Nothing Then
        If SelectFrontPlane(swModel) Then
            surface = "front plane (sketch would not open on the top face)"
            swModel.SketchManager.InsertSketch True
        End If
    End If
    If swModel.SketchManager.ActiveSketch Is Nothing Then
        detail = "a sketch would not open (surface: " & surface & ")"
        Exit Function
    End If

    ' --- 3) Draw each word as single-stroke line geometry - the same proven
    '        SketchManager.CreateLine calls the outline workaround uses.
    '        Position, height AND rotation come from the DWG. ----------------
    Dim i As Long
    For i = 0 To mCount - 1
        If Len(mMarks(i).Text) > 0 Then
            If DrawOneWord(swModel, _
                           mMarks(i).X * DWG_UNITS_TO_METERS, _
                           mMarks(i).Y * DWG_UNITS_TO_METERS, _
                           mMarks(i).Height * DWG_UNITS_TO_METERS, _
                           mMarks(i).Rotation, _
                           mMarks(i).Text) Then
                placed = placed + 1
            End If
        End If
    Next i

    ' --- 4) Close the sketch, leaving it in the tree (un-extruded) ----------
    swModel.SketchManager.InsertSketch True
    swModel.ClearSelection2 True

    detail = "surface: " & surface & "; drew " & placed & " of " & mCount & _
             IIf(Len(mLastDrawErr) > 0, "; first draw problem: " & mLastDrawErr, "")
    PlaceAllWords = placed
    Exit Function

errHandler:
    detail = "runtime error after surface '" & surface & "' with " & placed & _
             " word(s) drawn: " & Err.Description
    On Error Resume Next
    swModel.SketchManager.InsertSketch True   ' make sure sketch mode is closed
    swModel.ClearSelection2 True
    On Error GoTo 0
    PlaceAllWords = placed
End Function

'==============================================================================
' STROKE-FONT TEXT RENDERING
'
' IModelDoc2::InsertSketchText proved unusable unattended on this install (it
' returned Nothing for every word, visible or hidden, both signatures). So the
' words are DRAWN instead, as single-stroke line segments through
' SketchManager.CreateLine - the exact call the outline workaround already
' uses successfully. A single-stroke ("stick") font is also what pin-stamp /
' dot-peen marking equipment actually engraves, so the model matches the
' process.
'
' Each character is defined on a 4-wide x 6-tall grid (cap height = 6). A
' definition is one or more polylines: points "x,y" separated by spaces,
' polylines separated by ";". Coordinates scale so grid 6 = the DWG text
' height, and every point is rotated by the DWG rotation about the insertion
' point. Unknown characters draw as a box so missing glyphs are visible, not
' silent. Lower-case letters are drawn as capitals. (The FONT_* constants
' live in the declarations section at the top of this module.)
'==============================================================================

'------------------------------------------------------------------------------
' Draw one word at (xm, ym) meters, cap height hm meters, rotated rot radians
' about the insertion point. Returns True when at least one segment was drawn.
'------------------------------------------------------------------------------
Private Function DrawOneWord(ByVal swModel As SldWorks.ModelDoc2, _
                             ByVal xm As Double, ByVal ym As Double, _
                             ByVal hm As Double, ByVal rot As Double, _
                             ByVal text As String) As Boolean
    On Error GoTo done

    Dim skm As SldWorks.SketchManager
    Set skm = swModel.SketchManager

    If hm <= 0# Then hm = FONT_FALLBACK_HEIGHT_M
    Dim k As Double
    k = hm / FONT_CAP_GRID

    Dim cosR As Double
    Dim sinR As Double
    cosR = Cos(rot)
    sinR = Sin(rot)

    Dim prevAddToDB As Boolean
    prevAddToDB = skm.AddToDB
    skm.AddToDB = True                 ' exact coordinates, no snapping

    Dim penX As Double                 ' cursor, in grid units
    Dim i As Long
    Dim drawn As Boolean

    For i = 1 To Len(text)
        Dim ch As String
        ch = Mid$(UCase$(text), i, 1)

        If ch <> " " Then
            Dim strokes() As String
            strokes = Split(StrokeFor(ch), ";")

            Dim s As Long
            For s = 0 To UBound(strokes)
                Dim pts() As String
                pts = Split(Trim$(strokes(s)), " ")

                Dim p As Long
                Dim havePrev As Boolean
                Dim prevX As Double
                Dim prevY As Double
                havePrev = False

                For p = 0 To UBound(pts)
                    Dim xy() As String
                    xy = Split(pts(p), ",")

                    ' Grid -> meters (relative to insertion point)...
                    Dim gx As Double
                    Dim gy As Double
                    gx = (penX + Val(xy(0))) * k
                    gy = Val(xy(1)) * k

                    ' ...then rotate about the insertion point and translate.
                    Dim wx As Double
                    Dim wy As Double
                    wx = xm + gx * cosR - gy * sinR
                    wy = ym + gx * sinR + gy * cosR

                    If havePrev Then
                        If Not skm.CreateLine(prevX, prevY, 0#, wx, wy, 0#) _
                           Is Nothing Then drawn = True
                    End If
                    prevX = wx
                    prevY = wy
                    havePrev = True
                Next p
            Next s
        End If

        penX = penX + FONT_ADVANCE
    Next i

done:
    ' Record the FIRST failure so the log can explain a zero-word run.
    If Len(mLastDrawErr) = 0 Then
        If Err.Number <> 0 Then
            mLastDrawErr = "'" & text & "': " & Err.Description
        ElseIf Not drawn Then
            mLastDrawErr = "'" & text & "': CreateLine produced no segments"
        End If
    End If
    On Error Resume Next
    skm.AddToDB = prevAddToDB
    On Error GoTo 0
    DrawOneWord = drawn
End Function

'------------------------------------------------------------------------------
' Single-stroke glyph definitions. Grid: x 0..4, y 0..6 (baseline y = 0).
'------------------------------------------------------------------------------
Private Function StrokeFor(ByVal ch As String) As String
    Select Case ch
        Case "A": StrokeFor = "0,0 2,6 4,0;0.7,2 3.3,2"
        Case "B": StrokeFor = "0,0 0,6 3,6 4,5 4,4 3,3 0,3;3,3 4,2 4,1 3,0 0,0"
        Case "C": StrokeFor = "4,1 3,0 1,0 0,1 0,5 1,6 3,6 4,5"
        Case "D": StrokeFor = "0,0 0,6 2,6 4,4 4,2 2,0 0,0"
        Case "E": StrokeFor = "4,0 0,0 0,6 4,6;0,3 3,3"
        Case "F": StrokeFor = "0,0 0,6 4,6;0,3 3,3"
        Case "G": StrokeFor = "4,5 3,6 1,6 0,5 0,1 1,0 3,0 4,1 4,3 2,3"
        Case "H": StrokeFor = "0,0 0,6;4,0 4,6;0,3 4,3"
        Case "I": StrokeFor = "1,0 3,0;2,0 2,6;1,6 3,6"
        Case "J": StrokeFor = "0,1 1,0 2,0 3,1 3,6;2,6 4,6"
        Case "K": StrokeFor = "0,0 0,6;4,6 0,3 4,0"
        Case "L": StrokeFor = "0,6 0,0 4,0"
        Case "M": StrokeFor = "0,0 0,6 2,3 4,6 4,0"
        Case "N": StrokeFor = "0,0 0,6 4,0 4,6"
        Case "O": StrokeFor = "1,0 3,0 4,1 4,5 3,6 1,6 0,5 0,1 1,0"
        Case "P": StrokeFor = "0,0 0,6 3,6 4,5 4,4 3,3 0,3"
        Case "Q": StrokeFor = "1,0 3,0 4,1 4,5 3,6 1,6 0,5 0,1 1,0;2,2 4,0"
        Case "R": StrokeFor = "0,0 0,6 3,6 4,5 4,4 3,3 0,3;2,3 4,0"
        Case "S": StrokeFor = "0,1 1,0 3,0 4,1 4,2 3,3 1,3 0,4 0,5 1,6 3,6 4,5"
        Case "T": StrokeFor = "0,6 4,6;2,6 2,0"
        Case "U": StrokeFor = "0,6 0,1 1,0 3,0 4,1 4,6"
        Case "V": StrokeFor = "0,6 2,0 4,6"
        Case "W": StrokeFor = "0,6 1,0 2,3 3,0 4,6"
        Case "X": StrokeFor = "0,0 4,6;0,6 4,0"
        Case "Y": StrokeFor = "0,6 2,3 4,6;2,3 2,0"
        Case "Z": StrokeFor = "0,6 4,6 0,0 4,0"
        Case "0": StrokeFor = "1,0 3,0 4,1 4,5 3,6 1,6 0,5 0,1 1,0;1,1 3,5"
        Case "1": StrokeFor = "1,4 2,6 2,0;1,0 3,0"
        Case "2": StrokeFor = "0,5 1,6 3,6 4,5 4,4 0,1 0,0 4,0"
        Case "3": StrokeFor = "0,5 1,6 3,6 4,5 4,4 3,3 1,3;3,3 4,2 4,1 3,0 1,0 0,1"
        Case "4": StrokeFor = "3,0 3,6 0,2 4,2"
        Case "5": StrokeFor = "4,6 0,6 0,3 3,3 4,2 4,1 3,0 1,0 0,1"
        Case "6": StrokeFor = "4,5 3,6 1,6 0,5 0,1 1,0 3,0 4,1 4,2 3,3 1,3 0,2"
        Case "7": StrokeFor = "0,6 4,6 1,0"
        Case "8": StrokeFor = "1,3 0,4 0,5 1,6 3,6 4,5 4,4 3,3 1,3;1,3 0,2 0,1 1,0 3,0 4,1 4,2 3,3"
        Case "9": StrokeFor = "4,4 3,3 1,3 0,4 0,5 1,6 3,6 4,5 4,1 3,0 1,0 0,1"
        Case "-": StrokeFor = "1,3 3,3"
        Case ".": StrokeFor = "1.8,0 2.2,0 2.2,0.4 1.8,0.4 1.8,0"
        Case ",": StrokeFor = "2.2,0.5 1.7,-0.7"
        Case "/": StrokeFor = "0,0 4,6"
        Case "\": StrokeFor = "0,6 4,0"
        Case ":": StrokeFor = "1.8,1 2.2,1 2.2,1.4 1.8,1.4 1.8,1;1.8,4 2.2,4 2.2,4.4 1.8,4.4 1.8,4"
        Case "#": StrokeFor = "1,0 1,6;3,0 3,6;0,2 4,2;0,4 4,4"
        Case "+": StrokeFor = "2,1 2,5;0,3 4,3"
        Case "=": StrokeFor = "0,2 4,2;0,4 4,4"
        Case "(": StrokeFor = "3,6 2,4.5 2,1.5 3,0"
        Case ")": StrokeFor = "1,6 2,4.5 2,1.5 1,0"
        Case "'": StrokeFor = "1.9,5 2.1,6"
        Case "&": StrokeFor = "4,0 0,4 0,5 1,6 2,6 3,5 3,4 0,2 0,1 1,0 2,0 4,2"
        Case "_": StrokeFor = "0,-0.8 4,-0.8"
        Case Else: StrokeFor = "0,0 4,0 4,6 0,6 0,0"   ' unknown -> visible box
    End Select
End Function

' Find and select the top face of the extruded part. Returns False (nothing
' selected) if no suitable face is found.
Private Function SelectTopFace(ByVal swModel As SldWorks.ModelDoc2) As Boolean
    On Error GoTo fail

    Dim face As Object
    Set face = TopFace(swModel)
    If face Is Nothing Then Exit Function     ' returns False

    swModel.ClearSelection2 True
    Dim ent As SldWorks.Entity
    Set ent = face                            ' QI IFace2 -> IEntity
    SelectTopFace = ent.Select4(False, Nothing)
    Exit Function

fail:
    SelectTopFace = False
End Function

' Select the Front plane (first reference plane, locale independent). Returns
' False if it cannot be found.
Private Function SelectFrontPlane(ByVal swModel As SldWorks.ModelDoc2) As Boolean
    On Error GoTo fail

    Dim plane As SldWorks.Feature
    Set plane = FirstRefPlane(swModel)
    If plane Is Nothing Then Exit Function     ' returns False

    swModel.ClearSelection2 True
    SelectFrontPlane = plane.Select2(False, 0)
    Exit Function

fail:
    SelectFrontPlane = False
End Function

' Return the planar face whose normal is +Z (0,0,1) and which sits at the
' greatest Z - i.e. the top of the extrusion. Nothing if none qualifies.
Private Function TopFace(ByVal swModel As SldWorks.ModelDoc2) As Object
    On Error GoTo fail

    Dim swPart As SldWorks.PartDoc
    Set swPart = swModel

    Dim vBodies As Variant
    vBodies = swPart.GetBodies2(SW_SOLID_BODY, False)
    If IsEmpty(vBodies) Then Exit Function
    If Not IsArray(vBodies) Then Exit Function

    Dim bestFace As Object
    Dim bestZ As Double
    Dim haveBest As Boolean
    bestZ = -1E+30

    Dim b As Long
    For b = LBound(vBodies) To UBound(vBodies)
        Dim body As Object
        Set body = vBodies(b)
        If Not body Is Nothing Then

            Dim vFaces As Variant
            vFaces = body.GetFaces
            If IsArray(vFaces) Then
                Dim f As Long
                For f = LBound(vFaces) To UBound(vFaces)
                    Dim face As Object
                    Set face = vFaces(f)

                    Dim surf As Object
                    Set surf = face.GetSurface

                    If Not surf Is Nothing Then
                        If surf.IsPlane Then
                            Dim nrm As Variant
                            nrm = face.Normal          ' outward face normal
                            If Abs(nrm(0)) < FACE_NORMAL_TOL _
                               And Abs(nrm(1)) < FACE_NORMAL_TOL _
                               And nrm(2) > 1# - FACE_NORMAL_TOL Then

                                Dim box As Variant
                                box = face.GetBox       ' [xmin,ymin,zmin,xmax,ymax,zmax]
                                Dim z As Double
                                z = box(5)              ' zmax
                                If (Not haveBest) Or (z > bestZ) Then
                                    bestZ = z
                                    Set bestFace = face
                                    haveBest = True
                                End If
                            End If
                        End If
                    End If

                    Set face = Nothing
                Next f
            End If
        End If
        Set body = Nothing
    Next b

    Set TopFace = bestFace
    Exit Function

fail:
    Set TopFace = Nothing
End Function

' Return the first reference plane in the feature tree (the Front plane by
' default), or Nothing.
Private Function FirstRefPlane(ByVal swModel As SldWorks.ModelDoc2) As SldWorks.Feature
    Dim feat As SldWorks.Feature
    Set feat = swModel.FirstFeature
    Do While Not feat Is Nothing
        If feat.GetTypeName2 = SW_REFPLANE_TYPENAME Then
            Set FirstRefPlane = feat
            Exit Function
        End If
        Set feat = feat.GetNextFeature
    Loop
End Function
