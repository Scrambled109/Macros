Attribute VB_Name = "MaterialSpec_MOD_1"
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
    Dim itemNumberString As String
    Dim itemNumber As Long
    Dim lastDashPos As Long
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
                    fileName = swPart.GetTitle

                    If InStr(LCase(fileName), ".sldprt") > 0 Then
                        fileName = Left(fileName, InStr(LCase(fileName), ".sldprt") - 1)
                    End If

                    lastDashPos = InStrRev(fileName, "-")

                    If lastDashPos > 0 Then
                        itemNumberString = Mid(fileName, lastDashPos + 1)
                        itemNumber = ExtractLeadingNumber(itemNumberString)

                        If itemNumber > 0 Then
                            Dim excelRow As Long
                            excelRow = FindExcelRow(xlSheet, itemNumber)

                            If excelRow > 0 Then
                                Dim parsedDescription As String
                                Dim parsedMaterial As String

                                ' Column 3 = C (Description), Column 9 = I (Material)
                                parsedDescription = xlSheet.Cells(excelRow, 3).Value
                                parsedMaterial = xlSheet.Cells(excelRow, 9).Value

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
            End If
        End If
    Next i

    swModel.ForceRebuild3 True
    MsgBox "Done! Updated " & processCount & " parts on BOTH Global and Configuration layers.", vbInformation
End Sub

Function ExtractLeadingNumber(inputStr As String) As Long
    Dim resultStr As String
    Dim ch As String
    Dim k As Long
    resultStr = ""
    inputStr = Trim(inputStr)
    For k = 1 To Len(inputStr)
        ch = Mid(inputStr, k, 1)
        If IsNumeric(ch) Then
            resultStr = resultStr & ch
        Else
            Exit For
        End If
    Next k
    If resultStr <> "" Then ExtractLeadingNumber = CLng(resultStr) Else ExtractLeadingNumber = 0
End Function

Function FindExcelRow(sheet As Object, targetItemNo As Long) As Long
    Dim rRange As Object
    FindExcelRow = 0

    ' Searches Column B for the part number now
    Set rRange = sheet.Columns("B").Find(What:=targetItemNo, _
                                         LookIn:=-4163, _
                                         LookAt:=1, _
                                         SearchOrder:=1, _
                                         MatchByte:=False)

    If Not rRange Is Nothing Then
        FindExcelRow = rRange.Row
    End If
End Function
