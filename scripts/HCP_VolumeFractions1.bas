Attribute VB_Name = "HCP_VolumeFractions1"
'////////////////////////////////////////////////////////////////////////////////////////
'// HCP VOLUME FRACTIONS CALCULATOR
'// Hexagonal Close Packed Structure Analysis
'//
'// Author: Matt McPhillips
'// Email: mattmcp@blackboxengineering.co.uk
'//
'// Advanced materials science calculator for analyzing volume fractions
'// in hexagonal close packed (HCP) crystal structures with nano/micro particles
'//
'////////////////////////////////////////////////////////////////////////////////////////

Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////
'// Global Variables - Program State
'//
Public swApp As Object
Public part As Object
Public excelobject As Object
Public excelobjectActive As Object
Public excelobjectLocation As String
Public excelobjectOpen As Boolean
Public Pi As Double
Public X, Y, Z, V As Double
Public voidVolumeReduction As Double
Public noMicronsIntoX, noMicronsIntoY, noMicronsIntoZ As Double
Public noMicronsForX, noMicronsForY, noMicronsForZ As Double
Public modX, modY, modZ, modV, modA As Double
Public remX, remY, remZ, remV, remH As Double
Public boxX, boxY, boxZ, boxV As Double
Public f_hcpFractionRatioOne As Double
Public f_hcpFractionRatioTwo As Double
Public f_hcpFractionRatioThree As Double
Public revOne, revTwo, revThree As Double
Public nanVolOne, nanVolTwo, nanVolThree As Double
Public nanoVolOne, nanoVolTwo, nanoVolThree As Double
' Input
Public micronDiameter As Double
Public nanoDiameter As Double
Public surfaceThickness As Double
Public surfaceArea As Double

Public Function f_displayNanoInfo(diameter As Double)
    hcpVolumeFractionForm.TextBox_nanoVolume.Text = f_SphereVolume(nanoDiameter)
    hcpVolumeFractionForm.TextBox_totalNanos = f_totalNoSpheres
    hcpVolumeFractionForm.TextBox_NanosOnA.Text = f_spheresOnLayerA
    hcpVolumeFractionForm.TextBox_NanosOnB.Text = f_spheresOnLayerB
    hcpVolumeFractionForm.TextBox_NanosOnAX.Text = f_layerCountA
    hcpVolumeFractionForm.TextBox_NanosOnBX.Text = f_layerCountB
    hcpVolumeFractionForm.TextBox_totalNanoVolume.Text = f_totalSphereVolume(f_totalNoSpheres, nanoDiameter)
    hcpVolumeFractionForm.TextBox_XONano.Text = f_eqHeight(nanoDiameter)
    hcpVolumeFractionForm.TextBox_YONano.Text = nanoDiameter
    hcpVolumeFractionForm.TextBox_ZONano.Text = f_tetrahedronHeight(nanoDiameter)
    hcpVolumeFractionForm.TextBox_TetraHNano.Text = f_tetrahedronHeight(nanoDiameter)
    hcpVolumeFractionForm.TextBox_EqTriHNano.Text = f_eqHeight(nanoDiameter)
    hcpVolumeFractionForm.TextBox_TetraVNano.Text = f_tetrahedronVolume(nanoDiameter)
    hcpVolumeFractionForm.TextBox_VFO2.Text = f_hcpFractionRatioOne
    hcpVolumeFractionForm.TextBox_VFN2.Text = f_hcpFractionRatioThree
    hcpVolumeFractionForm.TextBox_VFO3.Text = revOne
    hcpVolumeFractionForm.TextBox_VFN3.Text = revThree
    hcpVolumeFractionForm.TextBox_VFO4.Text = Val(hcpVolumeFractionForm.TextBox_VFO1.Text) + Val(hcpVolumeFractionForm.TextBox_VFO2.Text) * Val(hcpVolumeFractionForm.TextBox_VFO3.Text)
    hcpVolumeFractionForm.TextBox_VFN4.Text = Val(hcpVolumeFractionForm.TextBox_VFN1.Text) + Val(hcpVolumeFractionForm.TextBox_VFN2.Text) * Val(hcpVolumeFractionForm.TextBox_VFN3.Text)
    hcpVolumeFractionForm.TextBox_NanoVFOne.Text = f_hcpFractionRatioOne
    hcpVolumeFractionForm.TextBox_NanoVFThree.Text = f_hcpFractionRatioThree
End Function

Public Function f_displayMicronInfo(diameter As Double)
    hcpVolumeFractionForm.TextBox_NewX.Text = boxX
    hcpVolumeFractionForm.TextBox_NewY.Text = boxY
    hcpVolumeFractionForm.TextBox_NewZ.Text = boxZ
    hcpVolumeFractionForm.TextBox_NewV.Text = boxV
    hcpVolumeFractionForm.TextBox_ModTM.Text = surfaceThickness + (modZ - Z)
    hcpVolumeFractionForm.TextBox_ModTN.Text = surfaceThickness + (boxZ - Z)
    hcpVolumeFractionForm.TextBox_NewInX.Text = noMicronsForX
    hcpVolumeFractionForm.TextBox_NewInY.Text = noMicronsForY
    hcpVolumeFractionForm.TextBox_NewInZ.Text = (boxZ / (f_tetrahedronHeight(micronDiameter)))
    hcpVolumeFractionForm.TextBox_ModX.Text = modX
    hcpVolumeFractionForm.TextBox_ModY.Text = modY
    hcpVolumeFractionForm.TextBox_ModZ.Text = modZ
    hcpVolumeFractionForm.TextBox_ModV.Text = modV
    hcpVolumeFractionForm.TextBox_totalMicrons = f_totalNoSpheres
    hcpVolumeFractionForm.TextBox_MicronsOnA.Text = f_spheresOnLayerA
    hcpVolumeFractionForm.TextBox_MicronsOnB.Text = f_spheresOnLayerB
    hcpVolumeFractionForm.TextBox_MicronsOnAX.Text = f_layerCountA
    hcpVolumeFractionForm.TextBox_MicronsOnBX.Text = f_layerCountB
    hcpVolumeFractionForm.TextBox_ModInX.Text = noMicronsForX
    hcpVolumeFractionForm.TextBox_ModInY.Text = noMicronsForY
    hcpVolumeFractionForm.TextBox_ModInZ.Text = noMicronsForZ
    hcpVolumeFractionForm.TextBox_OrigInX.Text = noMicronsIntoX
    hcpVolumeFractionForm.TextBox_OrigInY.Text = noMicronsIntoY
    hcpVolumeFractionForm.TextBox_OrigInZ.Text = noMicronsIntoZ
    hcpVolumeFractionForm.TextBox_OrigX.Text = X
    hcpVolumeFractionForm.TextBox_OrigY.Text = Y
    hcpVolumeFractionForm.TextBox_OrigZ.Text = Z
    hcpVolumeFractionForm.TextBox_OrigV.Text = V
    hcpVolumeFractionForm.TextBox_XO.Text = f_eqHeight(micronDiameter)
    hcpVolumeFractionForm.TextBox_YO.Text = micronDiameter
    hcpVolumeFractionForm.TextBox_ZO.Text = f_tetrahedronHeight(micronDiameter)
    hcpVolumeFractionForm.TextBox_OrigVF.Text = f_hcpFractionRatioOne
    hcpVolumeFractionForm.TextBox_ModVF.Text = f_hcpFractionRatioTwo
    hcpVolumeFractionForm.TextBox_NewVF.Text = f_hcpFractionRatioThree
    hcpVolumeFractionForm.TextBox_VFO1.Text = f_hcpFractionRatioOne
    hcpVolumeFractionForm.TextBox_VFM1.Text = f_hcpFractionRatioTwo
    hcpVolumeFractionForm.TextBox_VFN1.Text = f_hcpFractionRatioThree
    hcpVolumeFractionForm.TextBox_RemFNanoOrig.Text = revOne
    hcpVolumeFractionForm.TextBox_RemFNanoMod.Text = revTwo
    hcpVolumeFractionForm.TextBox_RemFNanoNew.Text = revThree
    hcpVolumeFractionForm.TextBox_micronVolume.Text = f_SphereVolume(micronDiameter)
    hcpVolumeFractionForm.TextBox_totalMicronVolume.Text = f_totalSphereVolume(f_totalNoSpheres, micronDiameter)
    hcpVolumeFractionForm.TextBox_TetraH.Text = f_tetrahedronHeight(micronDiameter)
    hcpVolumeFractionForm.TextBox_EqTriH.Text = f_eqHeight(micronDiameter)
    hcpVolumeFractionForm.TextBox_TetraV.Text = f_tetrahedronVolume(micronDiameter)
    hcpVolumeFractionForm.TextBox_volOrig.Text = nanVolOne
    hcpVolumeFractionForm.TextBox_volMod.Text = nanVolTwo
    hcpVolumeFractionForm.TextBox_volNew.Text = nanVolThree
End Function

Public Function f_fixVolumeTwo()
    hcpVolumeFractionForm.TextBox_NanoVFTwo.Text = f_hcpFractionRatioTwo
    hcpVolumeFractionForm.TextBox_VFM2.Text = f_hcpFractionRatioTwo
    hcpVolumeFractionForm.TextBox_VFM3.Text = revTwo
    hcpVolumeFractionForm.TextBox_VFM4.Text = Val(hcpVolumeFractionForm.TextBox_VFM1.Text) + Val(hcpVolumeFractionForm.TextBox_VFM2.Text) * Val(hcpVolumeFractionForm.TextBox_VFM3.Text)
End Function
