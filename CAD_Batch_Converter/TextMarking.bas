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
'   * AutoCAD stage  - HarvestTextMarks() reads every TEXT/MTEXT on TEXT_LAYER
'                      BEFORE the filter deletes it, capturing string + position
'                      + height + rotation into an in-memory list.
'   * SolidWorks stage - ApplyTextMarks() recreates each word as NATIVE
'                      SolidWorks sketch text on the part, on the Front plane
'                      (coincident with the imported outline, so the words line
'                      up in plan). The sketch is left UN-extruded - the words
'                      are simply "there for modeling", exactly as requested.
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

' Read every TEXT / MTEXT entity on TEXT_LAYER from the open drawing and store
' it. Reading does not require the layer to be unlocked, and this runs BEFORE
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
            If LayerEquals(ent.Layer, TEXT_LAYER) Then
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

' Recreate every harvested word as native sketch text on the part's Front plane
' (coincident with the imported outline). The sketch is left un-consumed. Call
' AFTER the base extrude and BEFORE the save. Returns True on success (and True
' when there is nothing to place). Never throws.
Public Function ApplyTextMarks(ByVal swModel As SldWorks.ModelDoc2) As Boolean
    If mCount = 0 Then
        ApplyTextMarks = True
        Exit Function
    End If

    On Error GoTo errHandler

    ' Select the Front plane (first reference plane, locale independent).
    Dim plane As SldWorks.Feature
    Set plane = FirstRefPlane(swModel)
    If plane Is Nothing Then
        ApplyTextMarks = False
        Exit Function
    End If

    swModel.ClearSelection2 True
    plane.Select2 False, 0

    ' Open a new sketch on that plane.
    swModel.SketchManager.InsertSketch True

    ' Place each word. Coordinates are scaled from DWG units to meters so they
    ' land on top of the imported outline.
    Dim i As Long
    For i = 0 To mCount - 1
        If Len(mMarks(i).Text) > 0 Then
            InsertOneText swModel, _
                          mMarks(i).X * DWG_UNITS_TO_METERS, _
                          mMarks(i).Y * DWG_UNITS_TO_METERS, _
                          mMarks(i).Text
        End If
    Next i

    ' Close the sketch, leaving it in the tree (un-extruded).
    swModel.SketchManager.InsertSketch True
    swModel.ClearSelection2 True

    ApplyTextMarks = True
    Exit Function

errHandler:
    On Error Resume Next
    swModel.SketchManager.InsertSketch True   ' make sure sketch mode is closed
    swModel.ClearSelection2 True
    On Error GoTo 0
    ApplyTextMarks = False
End Function

' Insert one piece of sketch text at (x, y) in the active sketch. Late-bound and
' individually guarded because the exact InsertSketchText argument list has
' varied across releases - a failure here skips that one word, never the batch.
Private Sub InsertOneText(ByVal swModel As SldWorks.ModelDoc2, _
                          ByVal x As Double, ByVal y As Double, _
                          ByVal text As String)
    On Error Resume Next
    Dim md As Object
    Set md = swModel          ' late binding for InsertSketchText
    md.InsertSketchText x, y, 0#, text, _
                        SW_TEXT_ALIGN_LEFT, 0, 0, _
                        SW_TEXT_WIDTH_PCT, SW_TEXT_SPACING_PCT, SW_TEXT_HEIGHT_PCT
    On Error GoTo 0
End Sub

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
