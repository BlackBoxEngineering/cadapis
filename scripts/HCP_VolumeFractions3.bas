Attribute VB_Name = "HCP_VolumeFractions3"
Function f_runStoreRangeDiameter(diameter As Double)
    Dim rangeHigh As Double
    Dim rangeLow As Double
    Dim loopRange As Double
    Dim stepSize As Double
    stepSize = Val(hcpVolumeFractionForm.TextBox_stepSize.Text)
    loopRange = rangeLow
    Pi = 4 * Atn(1)
    f_openExcelObject (excelobjectLocation)
    rangeHigh = Val(hcpVolumeFractionForm.TextBox_RangeHighD.Text)
    rangeLow = Val(hcpVolumeFractionForm.TextBox_RangeLowD.Text)
    area = Val(hcpVolumeFractionForm.TextBox_area) * 1000000000000#
    thickness = Val(hcpVolumeFractionForm.TextBox_thickness)
    hcpVolumeFractionForm.TextBox_diameter.Text = ""
    hcpVolumeFractionForm.Repaint
    Row = 2
    For loopRange = rangeLow To rangeHigh Step stepSize
        hcpVolumeFractionForm.TextBox_diameter.Text = loopRange
        micronDiameter = loopRange
        f_runForMicrons diameter
        excelobjectActive.Cells(Row, 1) = loopRange
        excelobjectActive.Cells(Row, 2) = f_hcpFractionRatioThree
        Row = Row + 1
    Next
    f_closeExcelObject
    resetForm
End Function

Function f_runStoreRangeArea(diameter As Double)
    Dim rangeHigh As Double
    Dim rangeLow As Double
    Dim storeArea As Double
    Dim stepSize As Double
    storeArea = surfaceArea
    f_openExcelObject (excelobjectLocation)
    rangeHigh = Val(hcpVolumeFractionForm.TextBox_RangeHighA.Text)
    rangeLow = Val(hcpVolumeFractionForm.TextBox_RangeLowA.Text)
    micronDiameter = Val(hcpVolumeFractionForm.TextBox_diameter.Text)
    thickness = Val(hcpVolumeFractionForm.TextBox_thickness.Text)
    stepSize = Val(hcpVolumeFractionForm.TextBox_stepSize.Text)
    hcpVolumeFractionForm.TextBox_area.Text = ""
    hcpVolumeFractionForm.Repaint
    Row = 2
    If rangeHigh > rangeLow Then
        For loopRange = rangeLow To rangeHigh Step stepSize
            hcpVolumeFractionForm.TextBox_area.Text = loopRange
            surfaceArea = loopRange * 1000000000000#
            f_runForMicrons diameter
            excelobjectActive.Cells(Row, 4) = loopRange
            excelobjectActive.Cells(Row, 5) = f_hcpFractionRatioThree
            Row = Row + 1
        Next
    Else
        stepSize = stepSize * -1
        For loopRange = rangeLow To rangeHigh Step stepSize
            hcpVolumeFractionForm.TextBox_area.Text = loopRange
            surfaceArea = loopRange * 1000000000000#
            f_runForMicrons diameter
            excelobjectActive.Cells(Row, 4) = loopRange
            excelobjectActive.Cells(Row, 5) = f_hcpFractionRatioThree
            Row = Row + 1
        Next
        stepSize = stepSize * -1
    End If
    surfaceArea = storeArea
    f_closeExcelObject
    resetForm
End Function

