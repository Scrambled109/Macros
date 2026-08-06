Attribute VB_Name = "ModifiedPartsListProperties"
' SolidWorks VBA: apply properties from a modified Parts List.
' Import this module and ModifiedPartsListMapper.frm into (MOD)2(SECONDARY).swp.
Option Explicit

Private Const DEFAULT_SHEET_NAME As String = "Parts_List"
Private Const MAX_HEADER_SCAN_ROWS As Long = 25
Private Const SW_DOC_ASSEMBLY As Long = 2
Private Const SW_DOC_PART As Long = 1
Private Const SW_SEL_COMPONENTS As Long = 20
Private Const SW_CUSTOM_TEXT As Long = 30
Private Const SW_PROPERTY_REPLACE As Long = 1
Private Const SW_SAVE_SILENT As Long = 1

Private swApp As SldWorks.SldWorks
Private swAssemblyModel As SldWorks.ModelDoc2
Private swAssembly As SldWorks.AssemblyDoc
Private swView As SldWorks.ModelView
Private uiFrozen As Boolean

Public Sub main()
    On Error GoTo fatalError

    Set swApp = Application.SldWorks
    Set swAssemblyModel = swApp.ActiveDoc
    If swAssemblyModel Is Nothing Then
        MsgBox "Open an assembly before running Modified Parts List Properties.", _
               vbExclamation, "Modified Parts List Properties"
        Exit Sub
    End If
    If swAssemblyModel.GetType <> SW_DOC_ASSEMBLY Then
        MsgBox "Run this macro from an assembly.", vbExclamation, _
               "Modified Parts List Properties"
        Exit Sub
    End If
    Set swAssembly = swAssemblyModel

    Dim excelWasCreated As Boolean
    Dim workbookWasOpened As Boolean
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object

    Set xlApp = GetExcelApplication(excelWasCreated)
    If xlApp Is Nothing Then
        MsgBox "Excel could not be started or attached.", vbCritical, _
               "Modified Parts List Properties"
        Exit Sub
    End If

    Set xlBook = GetInputWorkbook(xlApp, workbookWasOpened)
    If xlBook Is Nothing Then GoTo cleanExit

    Set xlSheet = ResolvePartsListSheet(xlBook)
    If xlSheet Is Nothing Then
        MsgBox "The selected workbook contains no worksheets.", vbExclamation, _
               "Modified Parts List Properties"
        GoTo cleanExit
    End If

    Dim headerRow As Long
    headerRow = FindLikelyHeaderRow(xlSheet)
    If headerRow = 0 Then
        MsgBox "No usable header row was found on sheet '" & xlSheet.Name & "'.", _
               vbExclamation, "Modified Parts List Properties"
        GoTo cleanExit
    End If

    Dim mapper As ModifiedPartsListMapper
    Set mapper = New ModifiedPartsListMapper
    mapper.Configure xlSheet, headerRow, CStr(xlBook.Name), CStr(xlSheet.Name)
    mapper.Show vbModal
    If mapper.Cancelled Then GoTo cleanExit

    Dim partColumn As Long
    Dim descriptionColumn As Long
    Dim materialColumn As Long
    partColumn = mapper.PartNumberColumn
    descriptionColumn = mapper.DescriptionColumn
    materialColumn = mapper.RawMaterialColumn
    Unload mapper
    Set mapper = Nothing

    Dim rowsByPart As Object
    Dim duplicateRows As Long
    Set rowsByPart = BuildRowIndex( _
        xlSheet, headerRow, partColumn, duplicateRows)
    If rowsByPart.Count = 0 Then
        MsgBox "The mapped part-number column contains no data rows.", _
               vbExclamation, "Modified Parts List Properties"
        GoTo cleanExit
    End If

    On Error Resume Next
    swAssembly.ResolveAllLightWeightComponents True
    On Error GoTo fatalError

    Dim components As Collection
    Dim selectedOnly As Boolean
    Set components = CollectComponents(selectedOnly)
    If components.Count = 0 Then
        MsgBox "No part components were available to process.", vbExclamation, _
               "Modified Parts List Properties"
        GoTo cleanExit
    End If

    FreezeAssemblyUI

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim updated As Long
    Dim unmatched As Long
    Dim skipped As Long
    Dim saveFailed As Long
    Dim component As SldWorks.Component2
    Dim i As Long

    For i = 1 To components.Count
        Set component = components(i)
        ProcessComponent component, xlSheet, rowsByPart, descriptionColumn, _
                         materialColumn, seen, updated, unmatched, skipped, _
                         saveFailed
    Next i

    RestoreAssemblyUI
    swAssemblyModel.ForceRebuild3 True

    Dim scopeText As String
    If selectedOnly Then
        scopeText = "selected components"
    Else
        scopeText = "the entire assembly"
    End If

    MsgBox "Modified Parts List properties complete." & vbCrLf & vbCrLf & _
           "Scope: " & scopeText & vbCrLf & _
           "Updated: " & updated & vbCrLf & _
           "Not found in Parts List: " & unmatched & vbCrLf & _
           "Skipped: " & skipped & vbCrLf & _
           "Save failures: " & saveFailed & vbCrLf & _
           "Duplicate spreadsheet keys: " & duplicateRows & vbCrLf & vbCrLf & _
           "Unmatched and failed items are listed in the VBA Immediate window (Ctrl+G).", _
           IIf(saveFailed > 0, vbExclamation, vbInformation), _
           "Modified Parts List Properties"

cleanExit:
    RestoreAssemblyUI
    CloseOwnedExcel xlApp, xlBook, workbookWasOpened, excelWasCreated
    Exit Sub

fatalError:
    Dim errorMessage As String
    errorMessage = "Unexpected error " & Err.Number & ": " & Err.Description
    RestoreAssemblyUI
    CloseOwnedExcel xlApp, xlBook, workbookWasOpened, excelWasCreated
    MsgBox errorMessage, vbCritical, "Modified Parts List Properties"
End Sub

Private Function GetExcelApplication(ByRef created As Boolean) As Object
    Dim xlApp As Object
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then
        Set xlApp = CreateObject("Excel.Application")
        created = Not xlApp Is Nothing
    End If
    On Error GoTo 0
    Set GetExcelApplication = xlApp
End Function

Private Function GetInputWorkbook(ByVal xlApp As Object, _
                                  ByRef openedByMacro As Boolean) As Object
    Dim book As Object
    On Error Resume Next
    Set book = xlApp.ActiveWorkbook
    On Error GoTo 0
    If Not book Is Nothing Then
        Set GetInputWorkbook = book
        Exit Function
    End If

    Dim chosen As Variant
    chosen = xlApp.GetOpenFilename( _
        "Excel Workbooks (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        1, "Select the modified Parts List")
    If VarType(chosen) = vbBoolean Then Exit Function

    On Error Resume Next
    Set book = xlApp.Workbooks.Open(CStr(chosen), False, True)
    On Error GoTo 0
    If Not book Is Nothing Then openedByMacro = True
    Set GetInputWorkbook = book
End Function

Private Function ResolvePartsListSheet(ByVal book As Object) As Object
    Dim sheet As Object
    On Error Resume Next
    Set sheet = book.Worksheets(DEFAULT_SHEET_NAME)
    If sheet Is Nothing Then Set sheet = book.ActiveSheet
    If sheet Is Nothing And book.Worksheets.Count > 0 Then
        Set sheet = book.Worksheets(1)
    End If
    On Error GoTo 0
    Set ResolvePartsListSheet = sheet
End Function

Private Function FindLikelyHeaderRow(ByVal sheet As Object) As Long
    On Error GoTo failed
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstColumn As Long
    Dim lastColumn As Long
    firstRow = sheet.UsedRange.Row
    firstColumn = sheet.UsedRange.Column
    lastRow = firstRow + sheet.UsedRange.Rows.Count - 1
    lastColumn = firstColumn + sheet.UsedRange.Columns.Count - 1
    If lastRow > firstRow + MAX_HEADER_SCAN_ROWS - 1 Then
        lastRow = firstRow + MAX_HEADER_SCAN_ROWS - 1
    End If

    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim populated As Long
    Dim bestCount As Long
    For rowIndex = firstRow To lastRow
        populated = 0
        For columnIndex = firstColumn To lastColumn
            If Len(Trim$(CStr(sheet.Cells(rowIndex, columnIndex).Value2))) > 0 Then
                populated = populated + 1
            End If
        Next columnIndex
        If populated > bestCount Then
            bestCount = populated
            FindLikelyHeaderRow = rowIndex
        End If
    Next rowIndex
    Exit Function
failed:
    FindLikelyHeaderRow = 0
End Function

Private Function BuildRowIndex(ByVal sheet As Object, ByVal headerRow As Long, _
                               ByVal partColumn As Long, _
                               ByRef duplicates As Long) As Object
    Dim index As Object
    Set index = CreateObject("Scripting.Dictionary")
    index.CompareMode = vbTextCompare

    Dim lastRow As Long
    lastRow = sheet.Cells(sheet.Rows.Count, partColumn).End(-4162).Row ' xlUp
    Dim rowIndex As Long
    Dim key As String
    For rowIndex = headerRow + 1 To lastRow
        key = NormalizePartNumber(sheet.Cells(rowIndex, partColumn).Value2)
        If Len(key) > 0 Then
            If index.Exists(key) Then
                duplicates = duplicates + 1
                Debug.Print "DUPLICATE PARTS LIST KEY (first row retained): " & _
                            key & " at row " & rowIndex
            Else
                index.Add key, rowIndex
            End If
        End If
    Next rowIndex
    Set BuildRowIndex = index
End Function

Private Function CollectComponents(ByRef selectedOnly As Boolean) As Collection
    Dim result As New Collection
    Dim selection As SldWorks.SelectionMgr
    Set selection = swAssemblyModel.SelectionManager

    Dim i As Long
    For i = 1 To selection.GetSelectedObjectCount2(-1)
        If selection.GetSelectedObjectType3(i, -1) = SW_SEL_COMPONENTS Then
            Dim selectedComponent As SldWorks.Component2
            Set selectedComponent = selection.GetSelectedObjectsComponent4(i, -1)
            If Not selectedComponent Is Nothing Then result.Add selectedComponent
        End If
    Next i

    If result.Count > 0 Then
        selectedOnly = True
        Set CollectComponents = result
        Exit Function
    End If

    selectedOnly = False
    Dim allComponents As Variant
    allComponents = swAssembly.GetComponents(False)
    If IsEmpty(allComponents) Then
        Set CollectComponents = result
        Exit Function
    End If

    For i = LBound(allComponents) To UBound(allComponents)
        If Not IsEmpty(allComponents(i)) Then
            Dim assemblyComponent As SldWorks.Component2
            Set assemblyComponent = allComponents(i)
            If Not assemblyComponent Is Nothing Then result.Add assemblyComponent
        End If
    Next i
    Set CollectComponents = result
End Function

Private Sub ProcessComponent(ByVal component As SldWorks.Component2, _
                             ByVal sheet As Object, ByVal rowsByPart As Object, _
                             ByVal descriptionColumn As Long, _
                             ByVal materialColumn As Long, ByVal seen As Object, _
                             ByRef updated As Long, ByRef unmatched As Long, _
                             ByRef skipped As Long, ByRef saveFailed As Long)
    On Error GoTo componentError
    If component Is Nothing Then
        skipped = skipped + 1
        Exit Sub
    End If

    Dim model As SldWorks.ModelDoc2
    Set model = component.GetModelDoc2
    If model Is Nothing Or model.GetType <> SW_DOC_PART Then
        skipped = skipped + 1
        Exit Sub
    End If

    Dim partNumber As String
    partNumber = PartNumberForModel(model)
    Dim configurationName As String
    configurationName = CStr(component.ReferencedConfiguration)
    Dim uniqueKey As String
    uniqueKey = LCase$(partNumber & "|" & configurationName)
    If seen.Exists(uniqueKey) Then Exit Sub
    seen.Add uniqueKey, True

    If Not rowsByPart.Exists(NormalizePartNumber(partNumber)) Then
        unmatched = unmatched + 1
        Debug.Print "NOT FOUND IN PARTS LIST: " & partNumber & _
                    "  configuration=" & configurationName
        Exit Sub
    End If

    Dim rowIndex As Long
    rowIndex = CLng(rowsByPart(NormalizePartNumber(partNumber)))
    Dim descriptionValue As String
    Dim materialValue As String
    If descriptionColumn > 0 Then
        descriptionValue = CellText(sheet.Cells(rowIndex, descriptionColumn).Value2)
    End If
    If materialColumn > 0 Then
        materialValue = CellText(sheet.Cells(rowIndex, materialColumn).Value2)
    End If

    WriteMappedProperties model, "", descriptionColumn, descriptionValue, _
                          materialColumn, materialValue
    If Len(configurationName) > 0 Then
        WriteMappedProperties model, configurationName, descriptionColumn, _
                              descriptionValue, materialColumn, materialValue
    End If

    Dim saveErrors As Long
    Dim saveWarnings As Long
    If model.Save3(SW_SAVE_SILENT, saveErrors, saveWarnings) Then
        updated = updated + 1
        Debug.Print "UPDATED: " & partNumber & "  configuration=" & configurationName
    Else
        saveFailed = saveFailed + 1
        Debug.Print "SAVE FAILED: " & partNumber & "  errors=" & saveErrors & _
                    " warnings=" & saveWarnings
    End If
    Exit Sub

componentError:
    skipped = skipped + 1
    Debug.Print "SKIPPED COMPONENT: " & component.Name2 & " - " & Err.Description
    Err.Clear
End Sub

Private Sub WriteMappedProperties(ByVal model As SldWorks.ModelDoc2, _
                                  ByVal configurationName As String, _
                                  ByVal descriptionColumn As Long, _
                                  ByVal descriptionValue As String, _
                                  ByVal materialColumn As Long, _
                                  ByVal materialValue As String)
    Dim manager As SldWorks.CustomPropertyManager
    Set manager = model.Extension.CustomPropertyManager(configurationName)
    If manager Is Nothing Then Exit Sub
    If descriptionColumn > 0 Then
        manager.Add3 "Description", SW_CUSTOM_TEXT, descriptionValue, _
                     SW_PROPERTY_REPLACE
    End If
    If materialColumn > 0 Then
        manager.Add3 "Raw_Material", SW_CUSTOM_TEXT, materialValue, _
                     SW_PROPERTY_REPLACE
    End If
End Sub

Private Function PartNumberForModel(ByVal model As SldWorks.ModelDoc2) As String
    Dim path As String
    path = model.GetPathName
    If Len(path) > 0 Then
        PartNumberForModel = CreateObject("Scripting.FileSystemObject").GetBaseName(path)
    Else
        PartNumberForModel = model.GetTitle
        If LCase$(Right$(PartNumberForModel, 7)) = ".sldprt" Then
            PartNumberForModel = Left$(PartNumberForModel, Len(PartNumberForModel) - 7)
        End If
    End If
End Function

Private Function NormalizePartNumber(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    Dim result As String
    result = Trim$(CStr(value))
    If LCase$(Right$(result, 7)) = ".sldprt" Then
        result = Left$(result, Len(result) - 7)
    End If
    NormalizePartNumber = LCase$(Trim$(result))
End Function

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = ""
    Else
        CellText = Trim$(CStr(value))
    End If
End Function

Private Sub FreezeAssemblyUI()
    On Error Resume Next
    Set swView = swAssemblyModel.ActiveView
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = False
    swAssemblyModel.FeatureManager.EnableFeatureTree = False
    uiFrozen = True
    On Error GoTo 0
End Sub

Private Sub RestoreAssemblyUI()
    If Not uiFrozen Then Exit Sub
    On Error Resume Next
    If Not swView Is Nothing Then swView.EnableGraphicsUpdate = True
    If Not swAssemblyModel Is Nothing Then
        swAssemblyModel.FeatureManager.EnableFeatureTree = True
    End If
    uiFrozen = False
    On Error GoTo 0
End Sub

Private Sub CloseOwnedExcel(ByVal xlApp As Object, ByVal book As Object, _
                            ByVal workbookWasOpened As Boolean, _
                            ByVal excelWasCreated As Boolean)
    On Error Resume Next
    If workbookWasOpened And Not book Is Nothing Then book.Close False
    If excelWasCreated And Not xlApp Is Nothing Then xlApp.Quit
    Set book = Nothing
    Set xlApp = Nothing
    On Error GoTo 0
End Sub
