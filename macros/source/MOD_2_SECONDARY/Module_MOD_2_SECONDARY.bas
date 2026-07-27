Attribute VB_Name = "Module_MOD_2_SECONDARY_"
Option Explicit

Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swSelectionMgr As SldWorks.SelectionMgr
    Dim swComp As SldWorks.Component2
    Dim swPart As SldWorks.PartDoc
    Dim swCustPropMgrGlobal As SldWorks.CustomPropertyManager
    Dim swCustPropMgrConfig As SldWorks.CustomPropertyManager
    Dim swConfig As SldWorks.Configuration

    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlSheet As Object

    Dim selectCount As Long
    Dim i As Long
    Dim fileName As String
    Dim activeConfigName As String

    ' --- USER CONFIGURATION ---
    Const EXCEL_SHEET_NAME As String = "Parts_List"
    ' --------------------------

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Please open an assembly and select components first.", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then
        MsgBox "Could not find an active Excel instance. Please open your spreadsheet first.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    Set xlWorkbook = xlApp.ActiveWorkbook

    On Error Resume Next
    Set xlSheet = xlWorkbook.Sheets(EXCEL_SHEET_NAME)
    If xlSheet Is Nothing Then
        MsgBox "Could not find sheet '" & EXCEL_SHEET_NAME & "' in the active workbook!", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    Set swSelectionMgr = swModel.SelectionManager
    selectCount = swSelectionMgr.GetSelectedObjectCount2(-1)

    If selectCount = 0 Then
        MsgBox "No components selected.", vbExclamation
        Exit Sub
    End If

    Dim processCount As Long
    processCount = 0

    For i = 1 To selectCount
        If swSelectionMgr.GetSelectedObjectType3(i, -1) = swSelectType_e.swSelCOMPONENTS Then
            Set swComp = swSelectionMgr.GetSelectedObjectsComponent4(i, -1)

            If Not swComp Is Nothing Then
                Set swPart = swComp.GetModelDoc2

                If Not swPart Is Nothing And swPart.GetType = swDocumentTypes_e.swDocPART Then
                    ' Get the exact file name
                    fileName = swPart.GetTitle

                    ' Strip off the .sldprt extension if it is visible
                    If InStr(LCase(fileName), ".sldprt") > 0 Then
                        fileName = Left(fileName, InStr(LCase(fileName), ".sldprt") - 1)
                    End If

                    ' Search Column B for the exact part name
                    Dim excelRow As Long
                    excelRow = FindExcelRow(xlSheet, fileName)

                    If excelRow > 0 Then
                        Dim parsedDescription As String
                        Dim parsedMaterial As String

                        ' Column 3 = C (Description), Column 9 = I (Material)
                        parsedDescription = xlSheet.Cells(excelRow, 3).Value
                        parsedMaterial = xlSheet.Cells(excelRow, 9).Value

                        ' Handle blank cells gracefully
                        If IsEmpty(parsedDescription) Then parsedDescription = ""
                        If IsEmpty(parsedMaterial) Then parsedMaterial = ""

                        ' 1. WRITE TO GLOBAL CUSTOM TAB
                        Set swCustPropMgrGlobal = swPart.Extension.CustomPropertyManager(vbNullString)
                        swCustPropMgrGlobal.Add3 "Description", swCustomInfoType_e.swCustomInfoText, parsedDescription, swCustomPropertyAddOption_e.swCustomPropertyReplaceValue
                        swCustPropMgrGlobal.Add3 "Raw_Material", swCustomInfoType_e.swCustomInfoText, parsedMaterial, swCustomPropertyAddOption_e.swCustomPropertyReplaceValue

                        ' 2. WRITE TO CONFIGURATION-SPECIFIC TAB
                        Set swConfig = swPart.GetActiveConfiguration
                        If Not swConfig Is Nothing Then
                            activeConfigName = swConfig.Name
                            Set swCustPropMgrConfig = swPart.Extension.CustomPropertyManager(activeConfigName)

                            swCustPropMgrConfig.Add3 "Description", swCustomInfoType_e.swCustomInfoText, parsedDescription, swCustomPropertyAddOption_e.swCustomPropertyReplaceValue
                            swCustPropMgrConfig.Add3 "Raw_Material", swCustomInfoType_e.swCustomInfoText, parsedMaterial, swCustomPropertyAddOption_e.swCustomPropertyReplaceValue
                        End If

                        processCount = processCount + 1
                    End If
                End If
            End If
        End If
    Next i

    swModel.ForceRebuild3 True
    MsgBox "Done! Successfully found and updated " & processCount & " parts.", vbInformation
End Sub

Function FindExcelRow(sheet As Object, targetItem As String) As Long
    Dim rRange As Object
    FindExcelRow = 0

    ' Searches Column B for the exact string (e.g., "PA071053-18")
    Set rRange = sheet.Columns("B").Find(What:=targetItem, _
                                         LookIn:=-4163, _
                                         LookAt:=1, _
                                         SearchOrder:=1, _
                                         MatchByte:=False)

    If Not rRange Is Nothing Then
        FindExcelRow = rRange.Row
    End If
End Function
