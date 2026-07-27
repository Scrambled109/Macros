Attribute VB_Name = "Test_test1"
Dim swApp As SldWorks.SldWorks

Sub main()
    Set swApp = Application.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc

    ' Check to make sure we are in a part file
    If swModel Is Nothing Then Exit Sub
    If swModel.GetType <> swDocumentTypes_e.swDocPART Then
        MsgBox "Please run this inside the individual Part file for the test!"
        Exit Sub
    End If

    Dim swFeatMgr As SldWorks.FeatureManager
    Set swFeatMgr = swModel.FeatureManager

    ' 1. Insert Global Bounding Box
    Dim status As Long
    swFeatMgr.InsertGlobalBoundingBox swGlobalBoundingBoxFitOptions_e.swBoundingBoxType_BestFit, False, False, status

    swModel.ForceRebuild3 False

    Dim swFeat As SldWorks.Feature
    Set swFeat = swModel.FirstFeature
    Dim swBBoxFeat As SldWorks.Feature

    ' Find the feature
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2() = "GlobalBBoxFeature" Then
            Set swBBoxFeat = swFeat
            Exit Do
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop

    If Not swBBoxFeat Is Nothing Then
        Dim swBBoxData As SldWorks.BoundingBoxFeatureData
        Set swBBoxData = swBBoxFeat.GetDefinition

        If Not swBBoxData Is Nothing Then
            ' Convert to inches
            Dim d(2) As Double
            d(0) = swBBoxData.Thickness * 39.3701
            d(1) = swBBoxData.Width * 39.3701
            d(2) = swBBoxData.Length * 39.3701

            ' Sort ascending
            Dim i As Integer, j As Integer, temp As Double
            For i = 0 To 1
                For j = i + 1 To 2
                    If d(i) > d(j) Then
                        temp = d(i)
                        d(i) = d(j)
                        d(j) = temp
                    End If
                Next j
            Next i

            Dim lengthVal As String
            Dim shapeVal As String
            lengthVal = Round(d(2), 3)
            shapeVal = Round(d(0), 3) & " x " & Round(d(1), 3)

            ' 3. WRITE TO THE CUSTOM TAB WITH THE AGGRESSIVE ADD COMMAND
            Dim swCustPrpMgr As SldWorks.CustomPropertyManager
            Set swCustPrpMgr = swModel.Extension.CustomPropertyManager("")

            ' Using swCustomPropertyDeleteAndAdd (Value = 1) to force row creation
            swCustPrpMgr.Add3 "Box_Length", swCustomInfoType_e.swCustomInfoText, lengthVal, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
            swCustPrpMgr.Add3 "Box_Shape", swCustomInfoType_e.swCustomInfoText, shapeVal, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
        End If

        ' 4. Clean up
        swBBoxFeat.Select2 False, 0
        swModel.Extension.DeleteSelection2 swDeleteSelectionOptions_e.swDelete_Absorbed

        MsgBox "Test Complete! Please check the Custom Properties tab."
    Else
        MsgBox "Error: SolidWorks failed to generate a bounding box for this part."
    End If
End Sub
