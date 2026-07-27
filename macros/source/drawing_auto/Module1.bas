Attribute VB_Name = "Module1"
Option Explicit

'===============================================================================
' CONFIGURABLE SOLIDWORKS DRAWING AUTOMATION
'
' Manual use:
'   Run main().
'   Enter any combination of:
'       N = Add/update view-name labels
'       D = Autodimension drawing views
'       B = Insert/update an individual-part fabrication table
'
'   Examples:
'       N
'       D
'       NB
'       NDB
'       ALL
'
' Scope:
'   - If one or more drawing views are selected before running the macro,
'     Names and Dimensions are applied only to those selected views.
'   - If no drawing views are selected, all model views on the current sheet
'     are processed.
'   - Pictorial views (*Isometric, *Dimetric, *Trimetric) are skipped during
'     an all-views dimension run. Explicitly selecting one overrides this.
'
' Scheduled/non-interactive use:
'   - Make a copy of the macro.
'   - Set INTERACTIVE_RUN = False.
'   - Set SCHEDULE_ACTIONS and SCHEDULE_SKIP_DIMMED_VIEWS below.
'   - This module still expects a drawing to already be active. A Task Scheduler
'     batch wrapper must open each target drawing before this module is called.
'===============================================================================

'--------------------------- USER SETTINGS -------------------------------------

Private Const INTERACTIVE_RUN As Boolean = True

' Used only when INTERACTIVE_RUN = False.
Private Const SCHEDULE_ACTIONS As String = "NDB"
Private Const SCHEDULE_SKIP_DIMMED_VIEWS As Boolean = True

' Optional general-table template for the individual-part fabrication table.
' Leave blank to let SOLIDWORKS create a standard general table.
'
' If you use a custom template, it must contain at least 2 rows and 7 columns:
' ITEM | PART NUMBER | DESCRIPTION | MATERIAL | SHAPE | LENGTH | QTY
'
' Example:
' Private Const CUSTOM_PART_TABLE_TEMPLATE_PATH As String = _
'     "S:\Engineering\SOLIDWORKS Standards\Table Templates\AUTO_PART_TABLE.sldtbt"
Private Const CUSTOM_PART_TABLE_TEMPLATE_PATH As String = ""

' Internal annotation name. This lets the macro update the existing table
' instead of inserting duplicates.
Private Const PART_TABLE_ANNOTATION_NAME As String = "AUTO_PART_TABLE"

' Part-table placement and sizing. All API lengths are meters.
' False is the safer default because it ignores a missing or incorrectly placed
' General Table anchor in the sheet format.
Private Const USE_GENERAL_TABLE_ANCHOR As Boolean = True
Private Const PART_TABLE_LEFT_MARGIN As Double = 0.01
Private Const PART_TABLE_TOP_MARGIN As Double = 0.01

' The table is automatically stretched from its anchor to this distance from
' the right sheet edge. API distances are meters.
Private Const PART_TABLE_RIGHT_MARGIN As Double = 0.01

' Relative column-width proportions. These must add to 1.0.
Private Const PART_TABLE_RATIO_ITEM As Double = 0.06
Private Const PART_TABLE_RATIO_PART_NUMBER As Double = 0.18
Private Const PART_TABLE_RATIO_DESCRIPTION As Double = 0.25
Private Const PART_TABLE_RATIO_MATERIAL As Double = 0.17
Private Const PART_TABLE_RATIO_SHAPE As Double = 0.13
Private Const PART_TABLE_RATIO_LENGTH As Double = 0.13
Private Const PART_TABLE_RATIO_QTY As Double = 0.08

' False allows you to drag the column boundaries manually after generation.
Private Const LOCK_PART_TABLE_COLUMN_WIDTHS As Boolean = False

Private Const PART_TABLE_HEADER_HEIGHT As Double = 0.006
Private Const PART_TABLE_DATA_HEIGHT As Double = 0.009

' View-label appearance and placement. SOLIDWORKS API distances are meters.
Private Const VIEW_LABEL_OFFSET As Double = 0.01     ' 10 mm below view
Private Const VIEW_LABEL_HEIGHT As Double = 0.003    ' 3 mm text height

' Autodimension behavior.
Private Const DIM_HORIZONTAL_PLACEMENT As Long = -1  ' -1 = below, 1 = above
Private Const DIM_VERTICAL_PLACEMENT As Long = -1    ' -1 = left, 1 = right

'--------------------------- SOLIDWORKS CONSTANTS -------------------------------

Private Const swDocPART As Long = 1
Private Const swDocDRAWING As Long = 3

Private Const swSelDRAWINGVIEWS As Long = 12

Private Const swAutodimEntitiesAll As Long = 1
Private Const swAutodimSchemeBaseline As Long = 1
Private Const swAutodimStatusSuccess As Long = 0

Private Const swBOMConfigurationAnchor_TopLeft As Long = 1
Private Const swBOMConfigurationAnchor_TopRight As Long = 2
Private Const swTableRowColChange_TableSizeCanChange As Long = 0

'--------------------------- MODULE VARIABLES ----------------------------------

Private swApp As Object
Private swModel As Object
Private swDraw As Object

'===============================================================================
' ENTRY POINT
'===============================================================================

Public Sub main()

    On Error GoTo FatalError

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "No active document.", vbExclamation, "Drawing Automation"
        Exit Sub
    End If

    If swModel.GetType <> swDocDRAWING Then
        MsgBox "The active document is not a drawing.", _
               vbExclamation, "Drawing Automation"
        Exit Sub
    End If

    Set swDraw = swModel

    Dim actions As String
    Dim skipDimmedViews As Boolean

    If INTERACTIVE_RUN Then

        actions = InputBox( _
            "Choose the actions to run:" & vbCrLf & vbCrLf & _
            "N = Add/update view names" & vbCrLf & _
            "D = Add automatic dimensions" & vbCrLf & _
            "B = Insert/update individual-part table" & vbCrLf & vbCrLf & _
            "Enter any combination, such as N, D, NB, NDB, or ALL." & _
            vbCrLf & vbCrLf & _
            "Selected drawing views are used when any are selected." & _
            vbCrLf & _
            "Otherwise, all model views on the current sheet are used.", _
            "Drawing Automation", _
            "N")

        If Len(Trim$(actions)) = 0 Then Exit Sub

        actions = NormalizeActions(actions)

        If Not ActionsAreValid(actions) Then
            MsgBox "No valid actions were entered. Use N, D, B, or ALL.", _
                   vbExclamation, "Drawing Automation"
            Exit Sub
        End If

        If HasAction(actions, "D") Then

            Dim dimensionChoice As VbMsgBoxResult

            dimensionChoice = MsgBox( _
                "Skip views that already contain at least one dimension?" & _
                vbCrLf & vbCrLf & _
                "Yes: safer for repeated runs." & vbCrLf & _
                "No: run Autodimension even when dimensions already exist." & _
                vbCrLf & _
                "Cancel: stop the macro.", _
                vbYesNoCancel + vbQuestion, _
                "Automatic Dimensions")

            If dimensionChoice = vbCancel Then Exit Sub

            skipDimmedViews = (dimensionChoice = vbYes)

        End If

    Else

        actions = NormalizeActions(SCHEDULE_ACTIONS)
        skipDimmedViews = SCHEDULE_SKIP_DIMMED_VIEWS

        If Not ActionsAreValid(actions) Then
            MsgBox "SCHEDULE_ACTIONS does not contain N, D, or B.", _
                   vbCritical, "Drawing Automation"
            Exit Sub
        End If

    End If

    Dim selectedViews As Collection
    Dim allViews As Collection
    Dim targetViews As Collection
    Dim selectedScope As Boolean

    Set selectedViews = GetSelectedDrawingViews()
    Set allViews = GetAllModelViews()

    If allViews.Count = 0 Then
        MsgBox "No drawing views were found on the current sheet.", _
               vbExclamation, "Drawing Automation"
        Exit Sub
    End If

    selectedScope = (selectedViews.Count > 0)

    If selectedScope Then
        Set targetViews = selectedViews
    Else
        Set targetViews = allViews
    End If

    Dim summary As String
    Dim processed As Long
    Dim skipped As Long
    Dim failed As Long
    Dim detail As String

    summary = "Drawing automation results:" & vbCrLf

    If HasAction(actions, "N") Then

        processed = 0
        skipped = 0
        failed = 0
        detail = ""

        AddOrUpdateViewNames targetViews, processed, skipped, failed, detail

        summary = summary & vbCrLf & _
                  "View names: " & processed & " updated/created"

        If skipped > 0 Then summary = summary & ", " & skipped & " skipped"
        If failed > 0 Then summary = summary & ", " & failed & " failed"

        If Len(detail) > 0 Then
            summary = summary & vbCrLf & detail
        End If

    End If

    If HasAction(actions, "D") Then

        processed = 0
        skipped = 0
        failed = 0
        detail = ""

        AddAutomaticDimensions _
            targetViews, _
            skipDimmedViews, _
            selectedScope, _
            processed, _
            skipped, _
            failed, _
            detail

        summary = summary & vbCrLf & _
                  "Dimensions: " & processed & " views dimensioned"

        If skipped > 0 Then summary = summary & ", " & skipped & " skipped"
        If failed > 0 Then summary = summary & ", " & failed & " failed"

        If Len(detail) > 0 Then
            summary = summary & vbCrLf & detail
        End If

    End If

    If HasAction(actions, "B") Then

        processed = 0
        skipped = 0
        failed = 0
        detail = ""

        InsertOrUpdatePartTable _
            selectedViews, _
            allViews, _
            processed, _
            skipped, _
            failed, _
            detail

        summary = summary & vbCrLf & _
                  "Part table: " & processed & " inserted/updated"

        If skipped > 0 Then summary = summary & ", " & skipped & " skipped"
        If failed > 0 Then summary = summary & ", " & failed & " failed"

        If Len(detail) > 0 Then
            summary = summary & vbCrLf & detail
        End If

    End If

    swModel.ClearSelection2 True
    swModel.EditRebuild3
    swModel.GraphicsRedraw2

    MsgBox summary, vbInformation, "Drawing Automation"
    Exit Sub

FatalError:
    MsgBox "The macro stopped because of an unexpected error:" & vbCrLf & _
           Err.Number & " - " & Err.description, _
           vbCritical, "Drawing Automation"

End Sub

'===============================================================================
' ACTION PARSING
'===============================================================================

Private Function NormalizeActions(ByVal rawActions As String) As String

    Dim result As String

    result = UCase$(Trim$(rawActions))
    result = Replace(result, " ", "")
    result = Replace(result, ",", "")
    result = Replace(result, ";", "")
    result = Replace(result, "+", "")
    result = Replace(result, "-", "")

    If result = "ALL" Then result = "NDB"

    NormalizeActions = result

End Function

Private Function ActionsAreValid(ByVal actions As String) As Boolean

    Dim i As Long
    Dim ch As String
    Dim hasValidAction As Boolean

    For i = 1 To Len(actions)

        ch = Mid$(actions, i, 1)

        Select Case ch
            Case "N", "D", "B"
                hasValidAction = True
            Case Else
                ActionsAreValid = False
                Exit Function
        End Select

    Next i

    ActionsAreValid = hasValidAction

End Function

Private Function HasAction(ByVal actions As String, ByVal actionLetter As String) As Boolean
    HasAction = (InStr(1, actions, actionLetter, vbTextCompare) > 0)
End Function

'===============================================================================
' VIEW COLLECTION
'===============================================================================

Private Function GetSelectedDrawingViews() As Collection

    Dim views As New Collection
    Dim swSelMgr As Object
    Dim swView As Object
    Dim i As Long
    Dim selectionCount As Long

    Set swSelMgr = swModel.SelectionManager
    selectionCount = swSelMgr.GetSelectedObjectCount2(-1)

    For i = 1 To selectionCount

        If swSelMgr.GetSelectedObjectType3(i, -1) = swSelDRAWINGVIEWS Then

            Set swView = swSelMgr.GetSelectedObject6(i, -1)

            If Not swView Is Nothing Then
                AddViewIfMissing views, swView
            End If

        End If

    Next i

    Set GetSelectedDrawingViews = views

End Function

Private Function GetAllModelViews() As Collection

    Dim views As New Collection
    Dim swView As Object

    ' The first returned view is the sheet view. Skip it.
    Set swView = swDraw.GetFirstView

    If Not swView Is Nothing Then
        Set swView = swView.GetNextView
    End If

    Do While Not swView Is Nothing

        AddViewIfMissing views, swView
        Set swView = swView.GetNextView

    Loop

    Set GetAllModelViews = views

End Function

Private Sub AddViewIfMissing(ByRef views As Collection, ByVal swView As Object)

    Dim existingView As Object

    For Each existingView In views

        If StrComp(existingView.GetName2, swView.GetName2, vbTextCompare) = 0 Then
            Exit Sub
        End If

    Next existingView

    views.Add swView

End Sub

'===============================================================================
' VIEW-NAME LABELS
'===============================================================================

Private Sub AddOrUpdateViewNames( _
    ByVal views As Collection, _
    ByRef processed As Long, _
    ByRef skipped As Long, _
    ByRef failed As Long, _
    ByRef detail As String)

    Dim swView As Object
    Dim outline As Variant
    Dim labelText As String
    Dim noteName As String
    Dim sheetX As Double
    Dim sheetY As Double

    For Each swView In views

        On Error GoTo ViewNameFailure

        outline = swView.GetOutline

        If Not IsArray(outline) Then
            skipped = skipped + 1
            GoTo NextViewName
        End If

        labelText = GetViewLabelText(swView)

        If Len(labelText) = 0 Then
            skipped = skipped + 1
            GoTo NextViewName
        End If

        sheetX = (CDbl(outline(0)) + CDbl(outline(2))) / 2#
        sheetY = CDbl(outline(1)) - VIEW_LABEL_OFFSET

        noteName = "AUTO_VIEW_LABEL_" & SanitizeObjectName(swView.GetName2)

        If CreateOrUpdateNamedNote(noteName, labelText, sheetX, sheetY) Then
            processed = processed + 1
        Else
            failed = failed + 1
            AppendDetail detail, _
                "Could not create/update label for " & swView.GetName2 & "."
        End If

NextViewName:
        On Error GoTo 0

    Next swView

    Exit Sub

ViewNameFailure:
    failed = failed + 1
    AppendDetail detail, _
        "View-name error on " & SafeViewName(swView) & _
        ": " & Err.description
    Err.Clear
    Resume NextViewName

End Sub

Private Function GetViewLabelText(ByVal swView As Object) As String

    Dim orientationName As String

    On Error Resume Next
    orientationName = CStr(swView.GetOrientationName)
    On Error GoTo 0

    orientationName = Replace(orientationName, "*", "")
    orientationName = Trim$(orientationName)

    If Len(orientationName) = 0 Then
        orientationName = swView.GetName2
    End If

    GetViewLabelText = orientationName

End Function

Private Function CreateOrUpdateNamedNote( _
    ByVal noteName As String, _
    ByVal noteText As String, _
    ByVal x As Double, _
    ByVal y As Double) As Boolean

    Dim swNote As Object
    Dim swAnnotation As Object
    Dim status As Boolean

    Set swNote = FindNoteByName(noteName)

    If swNote Is Nothing Then

        Set swNote = swDraw.CreateText2( _
            noteText, _
            x, _
            y, _
            0#, _
            VIEW_LABEL_HEIGHT, _
            0#)

        If swNote Is Nothing Then
            CreateOrUpdateNamedNote = False
            Exit Function
        End If

        status = swNote.SetName(noteName)

    Else

        status = swNote.SetText(noteText)

        Set swAnnotation = swNote.GetAnnotation

        If Not swAnnotation Is Nothing Then
            status = swAnnotation.SetPosition2(x, y, 0#)
        End If

    End If

    CreateOrUpdateNamedNote = True

End Function

Private Function FindNoteByName(ByVal requestedName As String) As Object

    Dim swView As Object
    Dim swNote As Object

    Set swView = swDraw.GetFirstView

    Do While Not swView Is Nothing

        Set swNote = Nothing

        On Error Resume Next
        Set swNote = swView.GetFirstNote
        On Error GoTo 0

        Do While Not swNote Is Nothing

            If StrComp(swNote.GetName, requestedName, vbTextCompare) = 0 Then
                Set FindNoteByName = swNote
                Exit Function
            End If

            Set swNote = swNote.GetNext

        Loop

        Set swView = swView.GetNextView

    Loop

    Set FindNoteByName = Nothing

End Function

'===============================================================================
' AUTOMATIC DIMENSIONS
'===============================================================================

Private Sub AddAutomaticDimensions( _
    ByVal views As Collection, _
    ByVal skipDimmedViews As Boolean, _
    ByVal selectedScope As Boolean, _
    ByRef processed As Long, _
    ByRef skipped As Long, _
    ByRef failed As Long, _
    ByRef detail As String)

    Dim swView As Object
    Dim selectionStatus As Boolean
    Dim autoDimStatus As Long

    For Each swView In views

        On Error GoTo DimensionFailure

        ' During an all-view run, do not clutter pictorial views.
        ' Explicitly selecting the view overrides this protection.
        If (Not selectedScope) And IsPictorialView(swView) Then
            skipped = skipped + 1
            AppendDetail detail, _
                "Skipped pictorial view " & swView.GetName2 & "."
            GoTo NextDimensionView
        End If

        If skipDimmedViews And ViewHasDisplayDimensions(swView) Then
            skipped = skipped + 1
            AppendDetail detail, _
                "Skipped already-dimensioned view " & swView.GetName2 & "."
            GoTo NextDimensionView
        End If

        swModel.ClearSelection2 True

        selectionStatus = swModel.Extension.SelectByID2( _
            swView.GetName2, _
            "DRAWINGVIEW", _
            0#, _
            0#, _
            0#, _
            False, _
            0, _
            Nothing, _
            0)

        If Not selectionStatus Then
            failed = failed + 1
            AppendDetail detail, _
                "Could not select view " & swView.GetName2 & "."
            GoTo NextDimensionView
        End If

        autoDimStatus = swDraw.AutoDimension( _
            swAutodimEntitiesAll, _
            swAutodimSchemeBaseline, _
            DIM_HORIZONTAL_PLACEMENT, _
            swAutodimSchemeBaseline, _
            DIM_VERTICAL_PLACEMENT)

        If autoDimStatus = swAutodimStatusSuccess Then
            processed = processed + 1
        Else
            failed = failed + 1
            AppendDetail detail, _
                "Autodimension failed on " & swView.GetName2 & _
                " with status code " & CStr(autoDimStatus) & "."
        End If

NextDimensionView:
        swModel.ClearSelection2 True
        On Error GoTo 0

    Next swView

    Exit Sub

DimensionFailure:
    failed = failed + 1
    AppendDetail detail, _
        "Dimension error on " & SafeViewName(swView) & _
        ": " & Err.description
    Err.Clear
    Resume NextDimensionView

End Sub

Private Function ViewHasDisplayDimensions(ByVal swView As Object) As Boolean

    Dim swDisplayDimension As Object

    On Error Resume Next

    ' GetFirstDisplayDimension5 remains available in current SOLIDWORKS releases
    ' and also works in older releases than GetFirstDisplayDimension6.
    Set swDisplayDimension = swView.GetFirstDisplayDimension5

    If swDisplayDimension Is Nothing Then
        Set swDisplayDimension = swView.GetFirstDisplayDimension
    End If

    On Error GoTo 0

    ViewHasDisplayDimensions = Not (swDisplayDimension Is Nothing)

End Function

Private Function IsPictorialView(ByVal swView As Object) As Boolean

    Dim orientationName As String

    On Error Resume Next
    orientationName = LCase$(CStr(swView.GetOrientationName))
    On Error GoTo 0

    IsPictorialView = _
        (InStr(orientationName, "isometric") > 0) Or _
        (InStr(orientationName, "dimetric") > 0) Or _
        (InStr(orientationName, "trimetric") > 0)

End Function

 '===============================================================================
' INDIVIDUAL-PART FABRICATION TABLE
'===============================================================================

Private Sub InsertOrUpdatePartTable( _
    ByVal preferredViews As Collection, _
    ByVal allViews As Collection, _
    ByRef processed As Long, _
    ByRef skipped As Long, _
    ByRef failed As Long, _
    ByRef detail As String)

    On Error GoTo PartTableFailure

    Dim swPartView As Object
    Dim swPartModel As Object
    Dim swTable As Object
    Dim configurationName As String
    Dim templatePath As String

    Set swPartView = FindPartView(preferredViews)

    If swPartView Is Nothing Then
        Set swPartView = FindPartView(allViews)
    End If

    If swPartView Is Nothing Then
        failed = failed + 1
        AppendDetail detail, _
            "No individual-part drawing view was found on the current sheet."
        Exit Sub
    End If

    Set swPartModel = GetReferencedPartDocument(swPartView)

    If swPartModel Is Nothing Then
        failed = failed + 1
        AppendDetail detail, _
            "The referenced part document is not loaded or could not be read."
        Exit Sub
    End If

    configurationName = GetReferencedConfigurationName(swPartView, swPartModel)
    templatePath = ResolvePartTableTemplatePath()

    If Len(Trim$(CUSTOM_PART_TABLE_TEMPLATE_PATH)) > 0 And _
       Len(templatePath) = 0 Then

        failed = failed + 1
        AppendDetail detail, _
            "CUSTOM_PART_TABLE_TEMPLATE_PATH does not point to an existing file."
        Exit Sub

    End If

    Set swTable = FindNamedTable(PART_TABLE_ANNOTATION_NAME)

    If swTable Is Nothing Then
        Set swTable = CreatePartTable(templatePath)

        If swTable Is Nothing Then
            failed = failed + 1
            AppendDetail detail, _
                "SOLIDWORKS could not create the general part table."
            Exit Sub
        End If

        NameTableAnnotation swTable, PART_TABLE_ANNOTATION_NAME
    End If

    If swTable.RowCount < 2 Or swTable.ColumnCount < 7 Then
        failed = failed + 1
        AppendDetail detail, _
            "The part-table template must contain at least 2 rows and 7 columns."
        Exit Sub
    End If

    FillPartTable swTable, swPartModel, configurationName
    PositionAndSizePartTable swTable

    processed = processed + 1
    Exit Sub

PartTableFailure:
    failed = failed + 1
    AppendDetail detail, "Part-table error: " & Err.description
    Err.Clear

End Sub

Private Function FindPartView(ByVal views As Collection) As Object

    Dim swView As Object

    For Each swView In views

        If IsPartView(swView) Then
            Set FindPartView = swView
            Exit Function
        End If

    Next swView

    Set FindPartView = Nothing

End Function

Private Function IsPartView(ByVal swView As Object) As Boolean

    Dim referencedDocument As Object
    Dim referencedPath As String

    On Error Resume Next

    Set referencedDocument = swView.referencedDocument

    If Not referencedDocument Is Nothing Then

        If referencedDocument.GetType = swDocPART Then
            IsPartView = True
            On Error GoTo 0
            Exit Function
        End If

    End If

    referencedPath = CStr(swView.GetReferencedModelName)

    On Error GoTo 0

    IsPartView = _
        (LCase$(Right$(referencedPath, 7)) = ".sldprt")

End Function

Private Function GetReferencedPartDocument(ByVal swView As Object) As Object

    Dim referencedDocument As Object

    On Error Resume Next
    Set referencedDocument = swView.referencedDocument
    On Error GoTo 0

    If Not referencedDocument Is Nothing Then

        If referencedDocument.GetType = swDocPART Then
            Set GetReferencedPartDocument = referencedDocument
            Exit Function
        End If

    End If

    Set GetReferencedPartDocument = Nothing

End Function

Private Function GetReferencedConfigurationName( _
    ByVal swView As Object, _
    ByVal swPartModel As Object) As String

    Dim configurationName As String
    Dim swConfiguration As Object

    On Error Resume Next
    configurationName = CStr(swView.ReferencedConfiguration)
    On Error GoTo 0

    If Len(Trim$(configurationName)) = 0 Then

        On Error Resume Next
        Set swConfiguration = swPartModel.GetActiveConfiguration

        If Not swConfiguration Is Nothing Then
            configurationName = swConfiguration.Name
        End If

        On Error GoTo 0

    End If

    GetReferencedConfigurationName = configurationName

End Function

Private Function CreatePartTable(ByVal templatePath As String) As Object

    Dim swTable As Object
    Dim swSheet As Object
    Dim paperSize As Long
    Dim sheetWidth As Double
    Dim sheetHeight As Double

    sheetWidth = 0#
    sheetHeight = 0#

    Set swSheet = swDraw.GetCurrentSheet
    paperSize = swSheet.GetSize(sheetWidth, sheetHeight)

    On Error Resume Next

    If USE_GENERAL_TABLE_ANCHOR Then

        ' Uses the General Table anchor stored in the sheet format.
        Set swTable = swDraw.InsertTableAnnotation2( _
            True, _
            0#, _
            0#, _
            swBOMConfigurationAnchor_TopLeft, _
            templatePath, _
            2, _
            7)

    Else

        ' Force the table's upper-left corner inside the drawing border.
        Set swTable = swDraw.InsertTableAnnotation2( _
            False, _
            PART_TABLE_LEFT_MARGIN, _
            sheetHeight - PART_TABLE_TOP_MARGIN, _
            swBOMConfigurationAnchor_TopLeft, _
            templatePath, _
            2, _
            7)

    End If

    On Error GoTo 0

    Set CreatePartTable = swTable

End Function

Private Sub PositionAndSizePartTable(ByVal swTable As Object)

    Dim swSheet As Object
    Dim swAnnotation As Object
    Dim paperSize As Long

    Dim sheetWidth As Double
    Dim sheetHeight As Double
    Dim tableLeftX As Double
    Dim availableWidth As Double

    Dim annotationPosition As Variant
    Dim lockColumns As Boolean

    If swTable Is Nothing Then Exit Sub

    sheetWidth = 0#
    sheetHeight = 0#

    Set swSheet = swDraw.GetCurrentSheet
    paperSize = swSheet.GetSize(sheetWidth, sheetHeight)

    Set swAnnotation = Nothing

    On Error Resume Next
    Set swAnnotation = swTable.GetAnnotation
    On Error GoTo 0

    ' Keep the table attached to the General Table anchor when enabled.
    ' A Top Left stationary corner makes it grow rightward and downward.
    On Error Resume Next

    If USE_GENERAL_TABLE_ANCHOR Then

        swTable.AnchorType = swBOMConfigurationAnchor_TopLeft
        swTable.Anchored = True

    Else

        swTable.Anchored = False
        swTable.AnchorType = swBOMConfigurationAnchor_TopLeft

        If Not swAnnotation Is Nothing Then
            swAnnotation.SetPosition2 _
                PART_TABLE_LEFT_MARGIN, _
                sheetHeight - PART_TABLE_TOP_MARGIN, _
                0#
        End If

    End If

    On Error GoTo 0

    ' Find the table's actual left edge. For an anchored table this should be
    ' the General Table anchor position in sheet coordinates.
    tableLeftX = PART_TABLE_LEFT_MARGIN

    If Not swAnnotation Is Nothing Then

        annotationPosition = Empty

        On Error Resume Next
        annotationPosition = swAnnotation.GetPosition
        On Error GoTo 0

        If IsArray(annotationPosition) Then

            If UBound(annotationPosition) >= 0 Then
                tableLeftX = CDbl(annotationPosition(0))
            End If

        End If

    End If

    ' Stretch across the usable width of the current sheet.
    availableWidth = sheetWidth - tableLeftX - PART_TABLE_RIGHT_MARGIN

    ' Prevent invalid or extremely narrow dimensions if the anchor is wrong.
    If availableWidth < 0.12 Then
        availableWidth = sheetWidth - _
                         PART_TABLE_LEFT_MARGIN - _
                         PART_TABLE_RIGHT_MARGIN
    End If

    ' Unlock first so an existing table can be resized.
    On Error Resume Next

    swTable.SetLockColumnWidth 0, False
    swTable.SetLockColumnWidth 1, False
    swTable.SetLockColumnWidth 2, False
    swTable.SetLockColumnWidth 3, False
    swTable.SetLockColumnWidth 4, False
    swTable.SetLockColumnWidth 5, False
    swTable.SetLockColumnWidth 6, False

    ' Distribute the available sheet width according to the configured ratios.
    swTable.SetColumnWidth 0, _
        availableWidth * PART_TABLE_RATIO_ITEM, _
        swTableRowColChange_TableSizeCanChange

    swTable.SetColumnWidth 1, _
        availableWidth * PART_TABLE_RATIO_PART_NUMBER, _
        swTableRowColChange_TableSizeCanChange

    swTable.SetColumnWidth 2, _
        availableWidth * PART_TABLE_RATIO_DESCRIPTION, _
        swTableRowColChange_TableSizeCanChange

    swTable.SetColumnWidth 3, _
        availableWidth * PART_TABLE_RATIO_MATERIAL, _
        swTableRowColChange_TableSizeCanChange

    swTable.SetColumnWidth 4, _
        availableWidth * PART_TABLE_RATIO_SHAPE, _
        swTableRowColChange_TableSizeCanChange

    swTable.SetColumnWidth 5, _
        availableWidth * PART_TABLE_RATIO_LENGTH, _
        swTableRowColChange_TableSizeCanChange

    swTable.SetColumnWidth 6, _
        availableWidth * PART_TABLE_RATIO_QTY, _
        swTableRowColChange_TableSizeCanChange

    swTable.SetRowHeight 0, _
        PART_TABLE_HEADER_HEIGHT, _
        swTableRowColChange_TableSizeCanChange

    swTable.SetRowHeight 1, _
        PART_TABLE_DATA_HEIGHT, _
        swTableRowColChange_TableSizeCanChange

    lockColumns = LOCK_PART_TABLE_COLUMN_WIDTHS

    swTable.SetLockColumnWidth 0, lockColumns
    swTable.SetLockColumnWidth 1, lockColumns
    swTable.SetLockColumnWidth 2, lockColumns
    swTable.SetLockColumnWidth 3, lockColumns
    swTable.SetLockColumnWidth 4, lockColumns
    swTable.SetLockColumnWidth 5, lockColumns
    swTable.SetLockColumnWidth 6, lockColumns

    On Error GoTo 0

End Sub

Private Sub FillPartTable( _
    ByVal swTable As Object, _
    ByVal swPartModel As Object, _
    ByVal configurationName As String)

    Dim partNumber As String
    Dim description As String
    Dim material As String
    Dim shape As String
    Dim partLength As String
    Dim quantity As String

    partNumber = GetModelProperty( _
        swPartModel, _
        configurationName, _
        Array("Part Number", "PART NUMBER", "PartNumber", "Number"))

    If Len(partNumber) = 0 Then
        partNumber = GetFileStem(swPartModel.GetPathName)
    End If

    description = GetModelProperty( _
        swPartModel, _
        configurationName, _
        Array("Description", "DESCRIPTION"))

    material = GetModelProperty( _
        swPartModel, _
        configurationName, _
        Array("Material", "MATERIAL", "SW-Material"))

    If Len(material) = 0 Then
        material = GetPartMaterial(swPartModel, configurationName)
    End If

    shape = GetModelProperty( _
        swPartModel, _
        configurationName, _
        Array("Shape", "SHAPE"))

    partLength = GetModelProperty( _
        swPartModel, _
        configurationName, _
        Array("Length", "LENGTH", "SW-Bounding Box Length"))

    quantity = GetModelProperty( _
        swPartModel, _
        configurationName, _
        Array("Quantity", "QTY", "Qty", "SW-Quantity"))

    ' A standalone part does not inherently know how many times it is used
    ' in a parent assembly. Use 1 unless a custom quantity property exists.
    If Len(quantity) = 0 Then quantity = "1"

    SetTableCellText swTable, 0, 0, "ITEM"
    SetTableCellText swTable, 0, 1, "PART NUMBER"
    SetTableCellText swTable, 0, 2, "DESCRIPTION"
    SetTableCellText swTable, 0, 3, "MATERIAL"
    SetTableCellText swTable, 0, 4, "SHAPE"
    SetTableCellText swTable, 0, 5, "LENGTH"
    SetTableCellText swTable, 0, 6, "QTY"

    SetTableCellText swTable, 1, 0, "1"
    SetTableCellText swTable, 1, 1, partNumber
    SetTableCellText swTable, 1, 2, description
    SetTableCellText swTable, 1, 3, material
    SetTableCellText swTable, 1, 4, shape
    SetTableCellText swTable, 1, 5, partLength
    SetTableCellText swTable, 1, 6, quantity

End Sub

Private Sub SetTableCellText( _
    ByVal swTable As Object, _
    ByVal rowIndex As Long, _
    ByVal columnIndex As Long, _
    ByVal value As String)

    On Error Resume Next

    swTable.Text2(rowIndex, columnIndex, True) = value

    If Err.Number <> 0 Then
        Err.Clear
        swTable.Text(rowIndex, columnIndex) = value
    End If

    On Error GoTo 0

End Sub

Private Function GetModelProperty( _
    ByVal swPartModel As Object, _
    ByVal configurationName As String, _
    ByVal candidateNames As Variant) As String

    Dim swPropertyManager As Object
    Dim candidate As Variant
    Dim result As String

    ' Configuration-specific properties take precedence.
    If Len(Trim$(configurationName)) > 0 Then

        On Error Resume Next
        Set swPropertyManager = _
            swPartModel.Extension.CustomPropertyManager(configurationName)
        On Error GoTo 0

        If Not swPropertyManager Is Nothing Then

            For Each candidate In candidateNames

                result = ReadPropertyValue( _
                    swPropertyManager, _
                    CStr(candidate))

                If Len(result) > 0 Then
                    GetModelProperty = result
                    Exit Function
                End If

            Next candidate

        End If

    End If

    Set swPropertyManager = Nothing

    ' Fall back to file-level custom properties.
    On Error Resume Next
    Set swPropertyManager = swPartModel.Extension.CustomPropertyManager("")
    On Error GoTo 0

    If Not swPropertyManager Is Nothing Then

        For Each candidate In candidateNames

            result = ReadPropertyValue( _
                swPropertyManager, _
                CStr(candidate))

            If Len(result) > 0 Then
                GetModelProperty = result
                Exit Function
            End If

        Next candidate

    End If

    GetModelProperty = ""

End Function

Private Function ReadPropertyValue( _
    ByVal swPropertyManager As Object, _
    ByVal propertyName As String) As String

    Dim rawValue As String
    Dim resolvedValue As String
    Dim wasResolved As Boolean
    Dim linkedToProperty As Boolean
    Dim status As Long

    On Error Resume Next

    status = swPropertyManager.Get6( _
        propertyName, _
        False, _
        rawValue, _
        resolvedValue, _
        wasResolved, _
        linkedToProperty)

    If Err.Number <> 0 Then

        Err.Clear
        swPropertyManager.Get2 propertyName, rawValue, resolvedValue

    End If

    On Error GoTo 0

    If Len(Trim$(resolvedValue)) > 0 Then
        ReadPropertyValue = Trim$(resolvedValue)
    Else
        ReadPropertyValue = Trim$(rawValue)
    End If

End Function

Private Function GetPartMaterial( _
    ByVal swPartModel As Object, _
    ByVal configurationName As String) As String

    Dim materialDatabase As String
    Dim materialName As String

    On Error Resume Next

    materialName = swPartModel.GetMaterialPropertyName2( _
        configurationName, _
        materialDatabase)

    If Err.Number <> 0 Then
        Err.Clear
        materialName = swPartModel.MaterialUserName
    End If

    On Error GoTo 0

    GetPartMaterial = Trim$(materialName)

End Function

Private Function GetFileStem(ByVal filePath As String) As String

    Dim fileName As String
    Dim dotPosition As Long

    fileName = Mid$(filePath, InStrRev(filePath, "\") + 1)
    dotPosition = InStrRev(fileName, ".")

    If dotPosition > 1 Then
        fileName = Left$(fileName, dotPosition - 1)
    End If

    GetFileStem = fileName

End Function

Private Function FindNamedTable(ByVal annotationName As String) As Object

    Dim swView As Object
    Dim swTable As Object
    Dim swAnnotation As Object

    Set swView = swDraw.GetFirstView

    Do While Not swView Is Nothing

        Set swTable = Nothing

        On Error Resume Next
        Set swTable = swView.GetFirstTableAnnotation2
        On Error GoTo 0

        Do While Not swTable Is Nothing

            Set swAnnotation = Nothing

            On Error Resume Next
            Set swAnnotation = swTable.GetAnnotation
            On Error GoTo 0

            If Not swAnnotation Is Nothing Then

                If StrComp( _
                    swAnnotation.GetName, _
                    annotationName, _
                    vbTextCompare) = 0 Then

                    Set FindNamedTable = swTable
                    Exit Function

                End If

            End If

            Set swTable = swTable.GetNext

        Loop

        Set swView = swView.GetNextView

    Loop

    Set FindNamedTable = Nothing

End Function

Private Sub NameTableAnnotation( _
    ByVal swTable As Object, _
    ByVal annotationName As String)

    Dim swAnnotation As Object

    On Error Resume Next

    Set swAnnotation = swTable.GetAnnotation

    If Not swAnnotation Is Nothing Then
        swAnnotation.SetName annotationName
    End If

    On Error GoTo 0

End Sub

Private Function ResolvePartTableTemplatePath() As String

    If Len(Trim$(CUSTOM_PART_TABLE_TEMPLATE_PATH)) = 0 Then
        ResolvePartTableTemplatePath = ""
        Exit Function
    End If

    If FileExists(CUSTOM_PART_TABLE_TEMPLATE_PATH) Then
        ResolvePartTableTemplatePath = CUSTOM_PART_TABLE_TEMPLATE_PATH
    Else
        ResolvePartTableTemplatePath = ""
    End If

End Function


'===============================================================================
' GENERAL HELPERS
'===============================================================================

Private Function FileExists(ByVal filePath As String) As Boolean

    On Error Resume Next
    FileExists = (Len(Dir$(filePath, vbNormal Or vbHidden Or vbReadOnly)) > 0)
    On Error GoTo 0

End Function

Private Function SanitizeObjectName(ByVal rawName As String) As String

    Dim result As String
    Dim invalidCharacters As Variant
    Dim item As Variant

    result = rawName

    invalidCharacters = Array( _
        "\", "/", ":", "*", "?", Chr$(34), _
        "<", ">", "|", "@", ".", " ")

    For Each item In invalidCharacters
        result = Replace(result, CStr(item), "_")
    Next item

    SanitizeObjectName = result

End Function

Private Function SafeViewName(ByVal swView As Object) As String

    On Error Resume Next

    If swView Is Nothing Then
        SafeViewName = "(unknown view)"
    Else
        SafeViewName = swView.GetName2
    End If

    If Len(SafeViewName) = 0 Then
        SafeViewName = "(unknown view)"
    End If

    On Error GoTo 0

End Function

Private Sub AppendDetail(ByRef detail As String, ByVal message As String)

    If Len(detail) > 0 Then
        detail = detail & vbCrLf
    End If

    detail = detail & "  - " & message

End Sub
