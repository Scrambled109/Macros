Attribute VB_Name = "AutoBOMProperties1"
' SolidWorks VBA: Bounding Box -> LENGTH / SHAPE (v14, Forced Overwrite)
' Run from an open ASSEMBLY.
' References: SolidWorks Type Library + SolidWorks Constant Type Library
Option Explicit

Dim swApp As SldWorks.SldWorks
Dim processedCount As Long
Dim skippedCount As Long

' ---- Output unit state ----
Dim gUnitMode As String      ' "in", "mm", or "cm"
Dim gUnitFactor As Double    ' meters -> chosen unit

Sub main()
    On Error GoTo ErrHandler
    Set swApp = Application.SldWorks

    Dim swAssyModel As SldWorks.ModelDoc2
    Set swAssyModel = swApp.ActiveDoc

    If swAssyModel Is Nothing Then
        MsgBox "Open an assembly first.", vbExclamation
        Exit Sub
    End If
    If swAssyModel.GetType <> swDocumentTypes_e.swDocASSEMBLY Then
        MsgBox "Run this from an Assembly.", vbExclamation
        Exit Sub
    End If

    ' ---- Ask for output unit ONCE ----
    Dim ans As String
    ans = InputBox("Output unit?" & vbCrLf & vbCrLf & _
                   "in = inches (fractions, 1/16)" & vbCrLf & _
                   "mm = millimeters (decimal)" & vbCrLf & _
                   "cm = centimeters (decimal)", _
                   "Bounding Box Units", "in")
    If Len(ans) = 0 Then Exit Sub   ' user hit Cancel

    Select Case LCase(Trim(ans))
        Case "mm"
            gUnitMode = "mm": gUnitFactor = 1000#
        Case "cm"
            gUnitMode = "cm": gUnitFactor = 100#
        Case "in", "inch", "inches", ""
            gUnitMode = "in": gUnitFactor = 39.3700787401575
        Case Else
            MsgBox "Unrecognized unit '" & ans & "'. Use in, mm, or cm.", vbExclamation
            Exit Sub
    End Select

    Dim swAssy As SldWorks.AssemblyDoc
    Set swAssy = swAssyModel

    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swAssyModel.SelectionManager
    Dim selCount As Long
    selCount = swSelMgr.GetSelectedObjectCount2(-1)

    If selCount = 0 Then
        On Error Resume Next
        swAssy.ResolveAllLightWeightComponents True
        On Error GoTo ErrHandler
    End If

    ' ---- SPEED HACK: Freeze the User Interface ----
    Dim swModView As SldWorks.ModelView
    Set swModView = swAssyModel.ActiveView
    If Not swModView Is Nothing Then swModView.EnableGraphicsUpdate = False
    swAssyModel.FeatureManager.EnableFeatureTree = False

    Dim vComponents As Variant
    If selCount > 0 Then
        Dim selComps() As Variant
        ReDim selComps(selCount - 1)
        Dim s As Long
        For s = 1 To selCount
            Set selComps(s - 1) = swSelMgr.GetSelectedObjectsComponent4(s, -1)
        Next s
        vComponents = selComps
    Else
        vComponents = swAssy.GetComponents(False)
    End If

    ' GetComponents can return Empty when the assembly contains no components.
    ' Guard before LBound/UBound so an empty assembly exits cleanly.
    If Not IsArray(vComponents) Then
        If Not swModView Is Nothing Then swModView.EnableGraphicsUpdate = True
        swAssyModel.FeatureManager.EnableFeatureTree = True
        MsgBox "No processable components were found.", vbInformation
        Exit Sub
    End If

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    processedCount = 0
    skippedCount = 0

    Dim i As Long
    Dim swComp As SldWorks.Component2
    Dim swMD As SldWorks.ModelDoc2
    Dim partPath As String

    For i = LBound(vComponents) To UBound(vComponents)
        If Not IsEmpty(vComponents(i)) Then
            Set swComp = vComponents(i)
            If Not swComp Is Nothing Then
                Set swMD = swComp.GetModelDoc2
                If Not swMD Is Nothing Then
                    If swMD.GetType = swDocumentTypes_e.swDocPART Then
                        partPath = swMD.GetPathName
                        If Len(partPath) = 0 Then partPath = swMD.GetTitle

                        If Not dict.Exists(partPath) Then
                            dict.Add partPath, True
                            On Error Resume Next
                            ProcessPart swMD
                            If Err.Number <> 0 Then
                                Debug.Print "ERROR  " & partPath & " - " & Err.Description
                                skippedCount = skippedCount + 1
                                Err.Clear
                            End If
                            On Error GoTo ErrHandler
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' ---- SPEED HACK: Unfreeze the User Interface ----
    If Not swModView Is Nothing Then swModView.EnableGraphicsUpdate = True
    swAssyModel.FeatureManager.EnableFeatureTree = True

    swAssyModel.SetUserPreferenceToggle swUserPreferenceToggle_e.swViewDispGlobalBBox, False
    swAssyModel.ForceRebuild3 True

    MsgBox "Done (" & gUnitMode & ")." & vbCrLf & _
           "Forced Updates: " & processedCount & _
           "   Skipped (No Dims): " & skippedCount & _
           vbCrLf & "Details: Immediate window (Ctrl+G).", vbInformation
    Exit Sub

ErrHandler:
    If Not swModView Is Nothing Then swModView.EnableGraphicsUpdate = True
    If Not swAssyModel Is Nothing Then swAssyModel.FeatureManager.EnableFeatureTree = True
    MsgBox "Unexpected error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Sub ProcessPart(swModel As SldWorks.ModelDoc2)
    If swModel Is Nothing Then Exit Sub
    If swModel.GetType <> swDocumentTypes_e.swDocPART Then Exit Sub

    Dim partName As String
    partName = swModel.GetPathName
    If Len(partName) = 0 Then partName = swModel.GetTitle

    ' ---- CRITICAL: Force the part to rebuild so stale geometry updates ----
    swModel.ForceRebuild3 False

    Dim dims(2) As Double
    Dim gotDims As Boolean
    gotDims = False

    ' ---- Find existing bounding box feature ----
    Dim swBBoxFeat As SldWorks.Feature
    Dim swFeat As SldWorks.Feature
    Set swFeat = swModel.FirstFeature
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "BoundingBox" Then
            Set swBBoxFeat = swFeat
            Exit Do
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop

    ' ---- Insert Bounding Box if missing ----
    Dim status As Long
    Dim wasInserted As Boolean
    wasInserted = False

    If swBBoxFeat Is Nothing Then
        On Error Resume Next
        Set swBBoxFeat = swModel.FeatureManager.InsertGlobalBoundingBox( _
            swGlobalBoundingBoxFitOptions_e.swBoundingBoxType_BestFit, _
            False, False, status)
        On Error GoTo 0
        wasInserted = Not (swBBoxFeat Is Nothing)
    End If

    ' ---- Extract dimensions (raw meters) ----
    If Not swBBoxFeat Is Nothing Then

        If wasInserted Then swModel.ForceRebuild3 False

        Dim dThick As Object, dWidth As Object, dLength As Object
        Set dThick = swModel.Parameter("Thickness@" & swBBoxFeat.Name)
        Set dWidth = swModel.Parameter("Width@" & swBBoxFeat.Name)
        Set dLength = swModel.Parameter("Length@" & swBBoxFeat.Name)

        If Not dThick Is Nothing And Not dWidth Is Nothing And Not dLength Is Nothing Then
            dims(0) = dThick.SystemValue
            dims(1) = dWidth.SystemValue
            dims(2) = dLength.SystemValue
            gotDims = True
        End If

        ' Hide Bounding Box visibility inside the Part
        On Error Resume Next
        swBBoxFeat.Select2 False, 0
        swModel.BlankRefGeom
        swModel.ClearSelection2 True
        swModel.SetUserPreferenceToggle swUserPreferenceToggle_e.swViewDispGlobalBBox, False
        On Error GoTo 0
    End If

    ' ---- FALLBACK: Solid Body Box (raw meters) ----
    If Not gotDims Then
        Dim swPartDoc As SldWorks.PartDoc
        Set swPartDoc = swModel
        Dim vBodies As Variant
        vBodies = swPartDoc.GetBodies2(0, True) ' 0 = swSolidBody

        If Not IsEmpty(vBodies) Then
            Dim minX As Double, minY As Double, minZ As Double
            Dim maxX As Double, maxY As Double, maxZ As Double
            minX = 100000000000000#: minY = 100000000000000#: minZ = 100000000000000#
            maxX = -100000000000000#: maxY = -100000000000000#: maxZ = -100000000000000#

            Dim bb As Long
            For bb = LBound(vBodies) To UBound(vBodies)
                Dim swBody As SldWorks.Body2
                Set swBody = vBodies(bb)
                Dim bBox As Variant
                bBox = swBody.GetBodyBox
                If IsArray(bBox) Then
                    If bBox(0) < minX Then minX = bBox(0)
                    If bBox(1) < minY Then minY = bBox(1)
                    If bBox(2) < minZ Then minZ = bBox(2)
                    If bBox(3) > maxX Then maxX = bBox(3)
                    If bBox(4) > maxY Then maxY = bBox(4)
                    If bBox(5) > maxZ Then maxZ = bBox(5)
                End If
            Next bb

            If minX < 10000000000000# Then
                dims(0) = Abs(maxX - minX)
                dims(1) = Abs(maxY - minY)
                dims(2) = Abs(maxZ - minZ)
                gotDims = True
            End If
        End If
    End If

    If Not gotDims Then
        Debug.Print "SKIP   " & partName & " (no dimensions obtainable)"
        skippedCount = skippedCount + 1
        Exit Sub
    End If

    ' Sort ascending (Thickness x Width x Length)
    Dim ii As Long, jj As Long, tmp As Double
    For ii = 0 To 1
        For jj = ii + 1 To 2
            If dims(ii) > dims(jj) Then
                tmp = dims(ii): dims(ii) = dims(jj): dims(jj) = tmp
            End If
        Next jj
    Next ii

    ' Format using the chosen unit
    Dim lengthVal As String, shapeVal As String
    lengthVal = FormatDim(dims(2))
    shapeVal = "PL " & FormatDim(dims(0)) & " x " & FormatDim(dims(1))

    ' Write to Custom tab + every configuration (FORCED OVERWRITE)
    WriteProps swModel, "", lengthVal, shapeVal
    Dim vConfigs As Variant, c As Long
    vConfigs = swModel.GetConfigurationNames
    If IsArray(vConfigs) Then
        For c = LBound(vConfigs) To UBound(vConfigs)
            WriteProps swModel, CStr(vConfigs(c)), lengthVal, shapeVal
        Next c
    End If

    ' Save
    Dim saveErrs As Long, saveWarns As Long
    If swModel.Save3(swSaveAsOptions_e.swSaveAsOptions_Silent, saveErrs, saveWarns) Then
        Debug.Print "OVERWROTE OK  " & partName & "  L=" & lengthVal & "  Shape=" & shapeVal
        processedCount = processedCount + 1
    Else
        Debug.Print "WROTE (save failed)  " & partName & _
                    " errors=" & saveErrs & " warnings=" & saveWarns
        skippedCount = skippedCount + 1
    End If
End Sub

Sub WriteProps(swModel As SldWorks.ModelDoc2, cfgName As String, _
               lengthVal As String, shapeVal As String)
    Dim mgr As SldWorks.CustomPropertyManager
    Set mgr = swModel.Extension.CustomPropertyManager(cfgName)
    If mgr Is Nothing Then Exit Sub
    mgr.Add3 "LENGTH", swCustomInfoType_e.swCustomInfoText, _
             lengthVal, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
    mgr.Add3 "SHAPE", swCustomInfoType_e.swCustomInfoText, _
             shapeVal, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
End Sub

Function FormatDim(ByVal valMeters As Double) As String
    Dim x As Double
    x = valMeters * gUnitFactor
    If gUnitMode = "in" Then
        FormatDim = DecToFraction16(x)
    Else
        FormatDim = Format(x, "0.0")
    End If
End Function

Function DecToFraction16(ByVal num As Double) As String
    Dim whole As Long
    Dim frac As Double
    Dim numerator As Long
    Dim denominator As Long

    num = Round(num * 16) / 16
    whole = Int(num)
    frac = num - whole

    numerator = Round(frac * 16)
    denominator = 16

    If numerator = 0 Then
        DecToFraction16 = CStr(whole)
        Exit Function
    End If

    If numerator = 16 Then
        DecToFraction16 = CStr(whole + 1)
        Exit Function
    End If

    Do While numerator Mod 2 = 0 And denominator Mod 2 = 0
        numerator = numerator / 2
        denominator = denominator / 2
    Loop

    If whole = 0 Then
        DecToFraction16 = CStr(numerator) & "/" & CStr(denominator)
    Else
        DecToFraction16 = CStr(whole) & " " & CStr(numerator) & "/" & CStr(denominator)
    End If
End Function
