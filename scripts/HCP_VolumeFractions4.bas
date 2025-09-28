Attribute VB_Name = "HCP_VolumeFractions4"
Function loopDrawFrame()
    BounceY = 0
    modCount = noMicronsForY - 1
    For n = 1 To noMicronsForY
        f_drawFrame (BounceY)
        edgeLength = micronDiameter * 0.001
        BounceY = BounceY + edgeLength
    Next
End Function
Public Function f_drawFrame(BounceY)
    Dim PositionAX, PositionAY As Double
    Dim PositionBX, PositionBY As Double
    Dim PositionCX, PositionCY As Double
    Dim edgeLength, CentreHeight As Double
    edgeLength = micronDiameter * 0.000000001
    CentreHeight = edgeLength / 2 * Sqr(3)
    PositionAX = edgeLength / 2
    PositionAY = edgeLength / 2
    PoistionAMidY = PositionAY + (edgeLength / 2)
    PositionBX = PositionAX
    PositionBY = PositionAY + edgeLength
    PositionCX = PositionAX + CentreHeight
    PositionCY = edgeLength
    Dim NoOnY As Integer
    NoOnY = hcpVolumeFractionForm.TextBox_ModInY.Text
    'Debug.Print "NumY:( " & NumY & ", " & PositionAY & " )"
    'Debug.Print "PositionA:( " & PositionAX & ", " & PositionAY & " )"
    'Debug.Print "PositionB:( " & PositionBX & ", " & PositionBY & " )"
    'Debug.Print "PositionC:( " & PositionCX & ", " & PositionCY & " )"
    Set part = swApp.ActiveDoc
    part.SketchManager.InsertSketch True
    boolstatus = part.Extension.SelectByID2("Top Plane", "PLANE", -8.06249393634786E-02, 2.83428931870904E-03, 4.06191073188588E-02, False, 0, Nothing, 0)
    part.ClearSelection2 True
    Dim skSegment As Object
    Set skSegment = part.SketchManager.CreateLine(PositionAX, PositionAY, 0#, PositionBX, PositionBY, 0#)
    boolstatus = part.SketchManager.CreateLinearSketchStepAndRepeat(1, 4, 0.01, edgeLength, 0, 1.5707963267949, "", False, False, False, False, True)
    Set skSegment = part.SketchManager.CreateLine(PositionBX, PositionBY, 0#, PositionCX, PositionCY, 0#)
    boolstatus = part.SketchManager.CreateLinearSketchStepAndRepeat(1, 4, 0.01, edgeLength, 0, 1.5707963267949, "", False, False, False, False, True)
    Set skSegment = part.SketchManager.CreateLine(PositionCX, PositionCY, 0#, PositionAX, PositionAY, 0#)
    boolstatus = part.SketchManager.CreateLinearSketchStepAndRepeat(1, NoOnY, 0.01, edgeLength, 0, 1.5707963267949, "", False, False, False, False, True)
    part.ClearSelection2 True
    part.SketchManager.InsertSketch True
End Function
Public Function f_drawMicrons()
    Dim CircleDiameter As Double
    Dim CircleRadius As Double
    Dim myRefPlane As Object
    Dim skSegment As Object
    Dim PickPositionX As Double
    Dim PickPositionY As Double
    Dim myFeature As Object
    Dim CircleOffsetX As Double
    Dim CircleOffsetY As Double
    CircleDiameter = micronDiameter * 0.000001
    CircleRadius = CircleDiameter / 2
    boolstatus = part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
    Set myRefPlane = part.FeatureManager.InsertRefPlane(8, 0.222, 0, 0, 0, 0)
    part.ClearSelection2 True
    boolstatus = part.Extension.SelectByID2("Plane1", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    part.SketchManager.InsertSketch True
    Set skSegment = part.SketchManager.CreateCircle(CircleRadius, CircleRadius, 0, 0, CircleRadius, 0#)
    part.ClearSelection2 True
    Set skSegment = part.SketchManager.CreateLine(CircleRadius, 0, 0#, CircleRadius, CircleDiameter, 0#)
    part.SketchManager.InsertSketch True
    part.ClearSelection2 True
    PickPositionX = CircleRadius * -0.01
    PickPositionY = CircleRadius
    boolstatus = part.Extension.SelectByID2("Arc1@Sketch1", "EXTSKETCHSEGMENT", 1.55920217144306E-03, 2.59336645695081E-02, 0, False, 0, Nothing, 0)
    boolstatus = part.Extension.SelectByID2("Sketch1", "SKETCHCONTOUR", 2.33019493494598E-02, 5.88834973432331E-03, -1.62965789908901E-02, True, 0, Nothing, 0)
    boolstatus = part.Extension.SelectByID2("Sketch1", "SKETCHREGION", 0.030085158882541, 3.36477127675618E-03, -1.81615042526304E-02, True, 0, Nothing, 0)
    boolstatus = part.DeSelectByID("", "SKETCHREGION", 0.030085158882541, 3.36477127675618E-03, -1.81615042526304E-02)
    boolstatus = part.DeSelectByID("", "SKETCHCONTOUR", 2.33019493494598E-02, 5.88834973432331E-03, -1.62965789908901E-02)
    boolstatus = part.Extension.SelectByID2("Line1@Sketch1", "EXTSKETCHSEGMENT", 0.0185, 2.17489742080375E-02, 0, True, 0, Nothing, 0)
    boolstatus = part.Extension.SelectByID2("Sketch1", "SKETCHREGION", 2.98069389697478E-02, 1.26178922878357E-03, -1.65859564303913E-02, True, 0, Nothing, 0)
    part.ClearSelection2 True
    boolstatus = part.Extension.SelectByID2("Line1@Sketch1", "EXTSKETCHSEGMENT", 0.0185, 2.17489742080375E-02, 0, False, 16, Nothing, 0)
    boolstatus = part.Extension.SelectByID2("Sketch1", "SKETCH", 2.98069389697478E-02, 1.26178922878357E-03, -1.65859564303913E-02, True, 2, Nothing, 0)
    part.SelectionManager.EnableContourSelection = True
    boolstatus = part.Extension.SelectByID2("Sketch1", "SKETCHREGION", 2.98069389697478E-02, 1.26178922878357E-03, -1.65859564303913E-02, True, 2, Nothing, 0)
    Set myFeature = part.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, 6.2831853071796, 0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True)
    part.SelectionManager.EnableContourSelection = False
    part.ClearSelection2 True
    boolstatus = part.Extension.SelectByID2("Line1@Sketch1", "EXTSKETCHSEGMENT", 0.0185, 2.29581443538227E-02, 0, False, 1, Nothing, 0)
    boolstatus = part.Extension.SelectByID2("Revolve1", "SOLIDBODY", 8.61078149075004E-03, 0.237602060885024, -1.95144227007494E-02, True, 256, Nothing, 0)
    Set myFeature = part.FeatureManager.FeatureLinearPattern3(noMicronsForY, CircleDiameter, 1, 0.01, False, False, "NULL", "NULL", False, False)
    part.ShowNamedView2 "*Top", 5
    part.ClearSelection2 True
    boolstatus = part.Extension.SelectByID2("Plane1", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    part.SketchManager.InsertSketch True
    CircleOffsetX = CircleRadius + Val(hcpVolumeFractionForm.TextBox_TetraH.Text)
    CircleOffsetY = CircleDiameter
    Set skSegment = part.SketchManager.CreateCircle(CircleOffsetX, CircleOffsetY, 0, CircleRadius, 0, 0#)
    'part.ClearSelection2 True
    'Set skSegment = part.SketchManager.CreateLine(CircleOffsetX, CircleRadius, 0#, CircleOffsetX, CircleOffsetY, 0#)
    part.SketchManager.InsertSketch True
    part.ClearSelection2 True
End Function
Public Function f_drawVolume()
    Dim nanoToMeterX As Double
    Dim nanoToMeterY As Double
    Dim nanoToMeterZ As Double
    nanoToMeterX = boxX * 0.000001
    nanoToMeterY = boxY * 0.000001
    nanoToMeterZ = boxZ * 0.000001
    'Debug.Print "nanoToMeterX " & nanoToMeterX
    'Debug.Print "nanoToMeterY " & nanoToMeterY
    'Debug.Print "nanoToMeterZ " & nanoToMeterX
    boolstatus = part.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    part.SketchManager.InsertSketch True
    'part.ClearSelection2 True
    Dim vSkLines As Variant
    vSkLines = part.SketchManager.CreateCornerRectangle(0, 0, 0, nanoToMeterX, nanoToMeterY, 0)
    part.SketchManager.InsertSketch True
    'part.ClearSelection2 True
    boolstatus = part.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, False, 4, Nothing, 0)
    Dim myFeature As Object
    Set myFeature = part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, nanoToMeterZ, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, True, True, True, 0, 0, False)
    part.SelectionManager.EnableContourSelection = False
    part.ShowNamedView2 "*Dimetric", 9
End Function
