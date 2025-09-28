Attribute VB_Name = "HCP_VolumeFractions5"
Option Explicit

Public Function f_loopForVolumes(diameter)
    f_runForNanos nanoDiameter, nanoVolThree
    f_displayNanoInfo (diameter)
    f_runForNanos nanoDiameter, nanoVolTwo
    f_fixVolumeTwo
    f_mixingRatios
End Function

Public Function f_runForMicrons(diameter As Double)
    f_hcpMakeVolumeOne surfaceArea, surfaceThickness, micronDiameter
    f_assembleSpheres (diameter)
    f_hcpMakeVolumeTwo (diameter)
    f_hcpMakeVolumeThree (diameter)
    f_hcpMicronFractionRatio (diameter)
    f_displayMicronInfo (diameter)
End Function

Public Function f_runForNanos(diameter As Double, volume) As Double
    Dim edgeLength As Double
    Dim nanoArea As Double
    edgeLength = volume ^ (1 / 3)
    nanoArea = edgeLength ^ (2)
    f_hcpMakeVolumeOne nanoArea, edgeLength, nanoDiameter
    f_assembleSpheres (diameter)
    f_hcpMakeVolumeTwo (diameter)
    f_hcpMakeVolumeThree (diameter)
    f_hcpNanoFractionRatio (diameter)
End Function

Public Function f_iniParameters()
    Pi = 4 * Atn(1)
    micronDiameter = Val(hcpVolumeFractionForm.TextBox_diameter.Text)
    nanoDiameter = Val(hcpVolumeFractionForm.TextBox_NanoD.Text)
    surfaceArea = Val(hcpVolumeFractionForm.TextBox_area.Text) * 1000000000000# '1e+12
    surfaceThickness = Val(hcpVolumeFractionForm.TextBox_thickness)
    voidVolumeReduction = 1 - 0.00089560302
    nanoVolOne = Val(hcpVolumeFractionForm.TextBox_volOrig.Text)
    nanoVolTwo = Val(hcpVolumeFractionForm.TextBox_volMod.Text)
    nanoVolThree = Val(hcpVolumeFractionForm.TextBox_volNew.Text)
    If excelobjectLocation = "" Then
        excelobjectLocation = "F:\Medical Computation\MicronsAndNanos\ProgramOutput.xlsx"
        hcpVolumeFractionForm.TextBox_excellPath.Text = excelobjectLocation
    End If
End Function

Sub main()
    f_SetActiveDocument
End Sub

