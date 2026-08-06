Attribute VB_Name = "Module_MOD_2_SECONDARY_"
Option Explicit

Private Const EXCEL_SHEET_NAME As String = "Parts_List"

' Change to True if you want each modified part saved automatically.
Private Const SAVE_UPDATED_PARTS As Boolean = False

Sub main()

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swAssembly As SldWorks.AssemblyDoc
    Dim swSelectionMgr As SldWorks.SelectionMgr
    Dim swComp As SldWorks.Component2
    Dim swPartModel As SldWorks.ModelDoc2

    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlSheet As Object

    Dim selectCount As Long
    Dim i As Long
    Dim excelRow As Long

    Dim fileName As String
    Dim referencedConfigName As String

    Dim parsedDescription As String
    Dim parsedShape As String
    Dim parsedMaterial As String

    Dim processedCount As Long
    Dim noComponentCount As Long
    Dim unavailableCount As Long
    Dim nonPartCount As Long
    Dim notFoundCount As Long
    Dim propertyFailureCount As Long
    Dim plateShapeSkippedCount As Long

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "No SOLIDWORKS document is active.", vbCritical
        Exit Sub
    End If

    If swModel.GetType <> swDocumentTypes_e.swDocASSEMBLY Then
        MsgBox "The active SOLIDWORKS document must be an assembly.", vbCritical
        Exit Sub
    End If

    Set swAssembly = swModel

    ' Resolve lightweight components so their part documents are available.
    swAssembly.ResolveAllLightWeightComponents True

    ' Connect to an already-running Excel instance.
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo 0

    If xlApp Is Nothing Then
        MsgBox "Excel is not currently open. Open the spreadsheet first.", vbCritical
        Exit Sub
    End If

    Set xlWorkbook = xlApp.ActiveWorkbook

    If xlWorkbook Is Nothing Then
        MsgBox "Excel does not have an active workbook.", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    Set xlSheet = xlWorkbook.Worksheets(EXCEL_SHEET_NAME)
    On Error GoTo 0

    If xlSheet Is Nothing Then
        MsgBox _
            "Could not find worksheet '" & EXCEL_SHEET_NAME & _
            "' in workbook '" & xlWorkbook.Name & "'.", _
            vbCritical

        Exit Sub
    End If

    Set swSelectionMgr = swModel.SelectionManager
    selectCount = swSelectionMgr.GetSelectedObjectCount2(-1)

    If selectCount = 0 Then
        MsgBox "No objects are selected in the assembly.", vbExclamation
        Exit Sub
    End If

    Debug.Print String(75, "-")
    Debug.Print "Macro started: " & Now
    Debug.Print "Selected object count: " & selectCount
    Debug.Print "Excel workbook: " & xlWorkbook.Name
    Debug.Print "Excel sheet: " & xlSheet.Name

    For i = 1 To selectCount

        Set swComp = Nothing
        Set swPartModel = Nothing

        ' Retrieves the owning component even when a face or edge is selected.
        On Error Resume Next
        Set swComp = swSelectionMgr.GetSelectedObjectsComponent4(i, -1)
        On Error GoTo 0

        If swComp Is Nothing Then

            noComponentCount = noComponentCount + 1

            Debug.Print i & ": No component found for selected object."
            Debug.Print "   Selection type: " & _
                        swSelectionMgr.GetSelectedObjectType3(i, -1)

            GoTo ContinueLoop

        End If

        If swComp.IsSuppressed Then

            unavailableCount = unavailableCount + 1

            Debug.Print i & ": Component is suppressed: " & swComp.Name2

            GoTo ContinueLoop

        End If

        Set swPartModel = swComp.GetModelDoc2

        If swPartModel Is Nothing Then

            unavailableCount = unavailableCount + 1

            Debug.Print i & ": Part document is unavailable: " & swComp.Name2

            GoTo ContinueLoop

        End If

        If swPartModel.GetType <> swDocumentTypes_e.swDocPART Then

            nonPartCount = nonPartCount + 1

            Debug.Print i & ": Selected component is not a part: " & swComp.Name2

            GoTo ContinueLoop

        End If

        fileName = GetDocumentBaseName(swPartModel)

        Debug.Print i & ": Component = " & swComp.Name2
        Debug.Print "   Search filename = [" & fileName & "]"

        excelRow = FindExcelRow(xlSheet, fileName)

        If excelRow = 0 Then

            notFoundCount = notFoundCount + 1

            Debug.Print "   Filename was not found in Excel column B."

            GoTo ContinueLoop

        End If

        ' Excel columns:
        ' C = Description
        ' F = Shape
        ' I = Raw Material
        parsedDescription = SafeCellText(xlSheet.Cells(excelRow, 3).Value2)
        parsedShape = SafeCellText(xlSheet.Cells(excelRow, 6).Value2)
        parsedMaterial = SafeCellText(xlSheet.Cells(excelRow, 9).Value2)

        Debug.Print "   Excel row = " & excelRow
        Debug.Print "   Description = [" & parsedDescription & "]"
        Debug.Print "   Shape = [" & parsedShape & "]"
        Debug.Print "   Raw Material = [" & parsedMaterial & "]"

        If IsPlateDescription(parsedDescription) Then

            plateShapeSkippedCount = plateShapeSkippedCount + 1

            Debug.Print "   Plate description detected."
            Debug.Print "   Existing Shape property will not be changed."

        End If

        ' Write document-level custom properties.
        If Not WriteProperties( _
            swPartModel, _
            vbNullString, _
            parsedDescription, _
            parsedShape, _
            parsedMaterial) Then

            propertyFailureCount = propertyFailureCount + 1

        End If

        ' Write configuration-specific custom properties using the
        ' configuration referenced by the assembly component.
        referencedConfigName = swComp.ReferencedConfiguration

        If Len(Trim$(referencedConfigName)) > 0 Then

            If Not WriteProperties( _
                swPartModel, _
                referencedConfigName, _
                parsedDescription, _
                parsedShape, _
                parsedMaterial) Then

                propertyFailureCount = propertyFailureCount + 1

            End If

        Else

            Debug.Print "   No referenced configuration was returned."

        End If

        swPartModel.ForceRebuild3 False
        swPartModel.SetSaveFlag

        If SAVE_UPDATED_PARTS Then
            SaveModelSilently swPartModel
        End If

        processedCount = processedCount + 1

        Debug.Print "   Part updated successfully."

ContinueLoop:

    Next i

    swModel.ForceRebuild3 True
    swModel.GraphicsRedraw2

    Debug.Print String(75, "-")
    Debug.Print "Successfully updated: " & processedCount
    Debug.Print "Plate shape writes skipped: " & plateShapeSkippedCount
    Debug.Print "Not found in Excel: " & notFoundCount
    Debug.Print "Suppressed or unavailable: " & unavailableCount
    Debug.Print "No owning component: " & noComponentCount
    Debug.Print "Non-part selections: " & nonPartCount
    Debug.Print "Property write failures: " & propertyFailureCount

    MsgBox _
        "Macro finished." & vbCrLf & vbCrLf & _
        "Successfully updated: " & processedCount & vbCrLf & _
        "Plate shapes preserved: " & plateShapeSkippedCount & vbCrLf & _
        "Not found in Excel: " & notFoundCount & vbCrLf & _
        "Suppressed or unavailable: " & unavailableCount & vbCrLf & _
        "No owning component: " & noComponentCount & vbCrLf & _
        "Non-part selections: " & nonPartCount & vbCrLf & _
        "Property write failures: " & propertyFailureCount & vbCrLf & vbCrLf & _
        "Press Ctrl+G in the VBA editor for detailed results.", _
        vbInformation

End Sub

Private Function WriteProperties( _
    ByVal swModel As SldWorks.ModelDoc2, _
    ByVal configurationName As String, _
    ByVal descriptionValue As String, _
    ByVal shapeValue As String, _
    ByVal materialValue As String) As Boolean

    Dim propMgr As SldWorks.CustomPropertyManager

    Dim descriptionResult As Long
    Dim shapeResult As Long
    Dim materialResult As Long

    Dim writeSucceeded As Boolean
    Dim shouldWriteShape As Boolean

    Set propMgr = swModel.Extension.CustomPropertyManager(configurationName)

    If propMgr Is Nothing Then

        Debug.Print _
            "   Could not get CustomPropertyManager for scope [" & _
            configurationName & "]."

        WriteProperties = False
        Exit Function

    End If

    descriptionResult = propMgr.Add3( _
        "Description", _
        swCustomInfoType_e.swCustomInfoText, _
        descriptionValue, _
        swCustomPropertyAddOption_e.swCustomPropertyReplaceValue)

    materialResult = propMgr.Add3( _
        "Raw_Material", _
        swCustomInfoType_e.swCustomInfoText, _
        materialValue, _
        swCustomPropertyAddOption_e.swCustomPropertyReplaceValue)

    ' Determine whether Shape should be imported.
    ' Shape is skipped whenever the Excel description contains "plate".
    shouldWriteShape = Not IsPlateDescription(descriptionValue)

    If shouldWriteShape Then

        shapeResult = propMgr.Add3( _
            "Shape", _
            swCustomInfoType_e.swCustomInfoText, _
            shapeValue, _
            swCustomPropertyAddOption_e.swCustomPropertyReplaceValue)

        Debug.Print _
            "   Shape written in scope [" & configurationName & _
            "]: [" & shapeValue & "]"

    Else

        ' A skipped property is treated as successful.
        shapeResult = 0

        Debug.Print _
            "   Shape preserved in scope [" & configurationName & _
            "] because the description contains Plate."

    End If

    Debug.Print "   Property scope = [" & configurationName & "]"
    Debug.Print "   Description Add3 result = " & descriptionResult
    Debug.Print "   Shape Add3 result = " & shapeResult
    Debug.Print "   Raw_Material Add3 result = " & materialResult

    writeSucceeded = _
        (descriptionResult >= 0) And _
        (shapeResult >= 0) And _
        (materialResult >= 0)

    WriteProperties = writeSucceeded

End Function

Private Function IsPlateDescription( _
    ByVal descriptionValue As String) As Boolean

    Dim normalizedDescription As String

    normalizedDescription = NormalizeKey(descriptionValue)

    ' Examples detected:
    ' PLATE
    ' STEEL PLATE
    ' 1/4 PLATE
    ' PLATE, A36
    ' BASEPLATE
    IsPlateDescription = _
        (InStr(1, normalizedDescription, "PLATE", vbTextCompare) > 0)

End Function

Private Function FindExcelRow( _
    ByVal sheet As Object, _
    ByVal targetItem As String) As Long

    Dim lastRow As Long
    Dim rowNumber As Long

    Dim excelValue As String
    Dim normalizedTarget As String

    FindExcelRow = 0
    normalizedTarget = NormalizeKey(targetItem)

    ' -4162 is Excel's xlUp constant.
    lastRow = sheet.Cells(sheet.Rows.Count, 2).End(-4162).Row

    For rowNumber = 1 To lastRow

        excelValue = SafeCellText(sheet.Cells(rowNumber, 2).Value2)

        If NormalizeKey(excelValue) = normalizedTarget Then

            FindExcelRow = rowNumber
            Exit Function

        End If

    Next rowNumber

End Function

Private Function GetDocumentBaseName( _
    ByVal swModel As SldWorks.ModelDoc2) As String

    Dim fullPath As String
    Dim documentName As String
    Dim extensionPosition As Long

    fullPath = swModel.GetPathName

    If Len(fullPath) > 0 Then

        documentName = Mid$(fullPath, InStrRev(fullPath, "\") + 1)

    Else

        ' Used for unsaved or virtual parts.
        documentName = swModel.GetTitle

    End If

    extensionPosition = InStrRev(documentName, ".")

    If extensionPosition > 0 Then
        documentName = Left$(documentName, extensionPosition - 1)
    End If

    GetDocumentBaseName = Trim$(documentName)

End Function

Private Function NormalizeKey(ByVal value As String) As String

    ' Replace nonbreaking spaces.
    value = Replace(value, Chr$(160), " ")

    ' Remove tabs and line breaks.
    value = Replace(value, vbTab, " ")
    value = Replace(value, vbCr, "")
    value = Replace(value, vbLf, "")

    NormalizeKey = UCase$(Trim$(value))

End Function

Private Function SafeCellText(ByVal cellValue As Variant) As String

    If IsError(cellValue) Then

        SafeCellText = vbNullString

    ElseIf IsNull(cellValue) Then

        SafeCellText = vbNullString

    ElseIf IsEmpty(cellValue) Then

        SafeCellText = vbNullString

    Else

        SafeCellText = Trim$(CStr(cellValue))

    End If

End Function

Private Sub SaveModelSilently( _
    ByVal swModel As SldWorks.ModelDoc2)

    Dim errors As Long
    Dim warnings As Long
    Dim saveSucceeded As Boolean

    saveSucceeded = swModel.Save3( _
        swSaveAsOptions_e.swSaveAsOptions_Silent, _
        errors, _
        warnings)

    Debug.Print _
        "   Save result = " & saveSucceeded & _
        "; errors = " & errors & _
        "; warnings = " & warnings

End Sub

