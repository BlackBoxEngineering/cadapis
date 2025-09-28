Attribute VB_Name = "HCP_VolumeFractions2"
'////////////////////////////////////////////////////////////////////////////////////////
'// HCP VOLUME FRACTIONS CALCULATOR - MATHEMATICAL FUNCTIONS
'// Hexagonal Close Packed Structure Analysis - Core Calculations
'//
'// Author: Matt McPhillips
'// Email: mattmcp@blackboxengineering.co.uk
'//
'// Advanced pharmaceutical engineering calculator for determining optimal
'// nanoparticle packing in non-Euclidean void spaces between HCP structures.
'// Developed for pharmaceutical applications to maximize drug delivery efficiency.
'//
'////////////////////////////////////////////////////////////////////////////////////////

Option Explicit

Public Function f_tetrahedronVolume(diametre As Double) As Double
    Dim tetVolume As Double
    tetVolume = (diametre * diametre * diametre) / (6 * Sqr(2))
    f_tetrahedronVolume = tetVolume
    'Debug.Print "tetVolume " & tetVolume
End Function

Public Function f_tetrahedronHeight(diametre As Double) As Double
    Dim tetHeight As Double
    tetHeight = Sqr(2 / 3) * diametre
    f_tetrahedronHeight = tetHeight
    'Debug.Print "tetHeight " & tetHeight
End Function

Public Function f_eqHeight(diametre As Double) As Double
    Dim eqHeight As Double
    eqHeight = diametre / 2 * Sqr(3)
    f_eqHeight = eqHeight
    'Debug.Print "f_eqHeight " & f_eqHeight
End Function

Public Function f_spheresIntoX(modMicronX As Double) As Double
    noMicronsIntoX = X / modMicronX
    noMicronsForX = Round(noMicronsIntoX - 0.5, 0)
    f_spheresIntoX = noMicronsForX
End Function

Public Function f_spheresIntoY(modMicronY As Double) As Double
    noMicronsIntoY = Y / modMicronY
    If noMicronsIntoY = 1 Then
        noMicronsForY = 1
    Else
        noMicronsForY = Round(noMicronsIntoY - 0.5, 0)
    End If
    f_spheresIntoY = noMicronsForY
End Function

Public Function f_spheresIntoZ(modMicronZ As Double) As Double
    noMicronsIntoZ = Z / modMicronZ
    noMicronsForZ = Round(noMicronsIntoZ - 0.5, 0)
    f_spheresIntoZ = noMicronsForZ
End Function

Public Function f_spheresOnLayerA() As Double
    f_spheresOnLayerA = noMicronsForX * noMicronsForY - Round((noMicronsForX / 2) - 0.5, 0)
End Function

Public Function f_spheresOnLayerB() As Double
    f_spheresOnLayerB = (noMicronsForX - 1) * noMicronsForY - Round((noMicronsForX / 2) - 0.5, 0)
End Function

Public Function f_layerCountA() As Double
    f_layerCountA = Round((noMicronsForZ / 2) + 0.5, 0)
End Function

Public Function f_layerCountB() As Double
    f_layerCountB = Round((noMicronsForZ / 2) - 0.5, 0)
End Function

Public Function f_totalNoSpheres() As Double
    f_totalNoSpheres = (f_spheresOnLayerA * f_layerCountA) + (f_spheresOnLayerB * f_layerCountB)
End Function

Public Function f_hcpMakeVolumeOne(area As Double, thickness As Double, diameter As Double) As Double
    Dim modMicronX, modMicronY, modMicronZ
    X = Sqrt(area)
    Y = Sqrt(area)
    Z = thickness
    V = X * Y * Z
    f_hcpMakeVolumeOne = V
End Function

Public Function f_hcpMakeVolumeTwo(diameter As Double) As Double
    modX = noMicronsForX * (f_eqHeight(diameter))
    modY = noMicronsForY * (diameter)
    modZ = noMicronsForZ * (f_tetrahedronHeight(diameter))
    modV = modX * modY * modZ
    modA = modX * modY
    remX = (f_eqHeight(diameter)) * (noMicronsIntoX - noMicronsForX)
    remY = diameter * (noMicronsIntoY - noMicronsForY)
    remZ = (f_tetrahedronHeight(diameter)) * (noMicronsIntoZ - noMicronsForZ)
    remV = V - modV
    remH = remV / modA
    f_hcpMakeVolumeTwo = modV
End Function

Public Function f_hcpMakeVolumeThree(diameter As Double) As Double
    boxX = modX
    boxY = modY
    boxZ = modZ + remH
    boxV = boxX * boxY * boxZ
    f_hcpMakeVolumeThree = boxV
End Function

Public Function f_totalSphereVolume(noSpheres As Double, sphereDiameter As Double) As Double
    Dim totalSphereVolume
    totalSphereVolume = (4 / 3) * Pi * ((sphereDiameter / 2) ^ 3)
    f_totalSphereVolume = noSpheres * totalSphereVolume
    'Debug.Print "totalSphereVolume " & totalSphereVolume
End Function

Public Function f_SphereVolume(sphereDiameter As Double) As Double
    Dim totalSphereVolume
    totalSphereVolume = (4 / 3) * Pi * ((sphereDiameter / 2) ^ 3)
    f_SphereVolume = totalSphereVolume
End Function

Public Function f_hcpMicronFractionRatio(diameter As Double) As Double
    f_hcpFractionRatioOne = f_totalSphereVolume(f_totalNoSpheres, diameter) / V
    f_hcpFractionRatioTwo = f_totalSphereVolume(f_totalNoSpheres, diameter) / modV
    f_hcpFractionRatioThree = (f_totalSphereVolume(f_totalNoSpheres, diameter) / boxV) * voidVolumeReduction
    revOne = 1 - f_hcpFractionRatioOne
    revTwo = 1 - f_hcpFractionRatioTwo
    revThree = 1 - f_hcpFractionRatioThree
    nanVolOne = V * revOne
    nanVolTwo = modV * revTwo
    nanVolThree = boxV * revThree
    f_hcpMicronFractionRatio = f_hcpFractionRatioThree
    Debug.Print "NANO:diameter: " & diameter
    Debug.Print "MICRON:f_hcpFractionRatioTwo: " & f_totalSphereVolume(f_totalNoSpheres, diameter) & " / " & modV & " = " & f_hcpFractionRatioTwo
End Function

Public Function f_hcpNanoFractionRatio(diameter As Double) As Double
    f_hcpFractionRatioOne = f_totalSphereVolume(f_totalNoSpheres, diameter) / nanoVolOne
    f_hcpFractionRatioTwo = f_totalSphereVolume(f_totalNoSpheres, diameter) / nanoVolTwo
    f_hcpFractionRatioThree = (f_totalSphereVolume(f_totalNoSpheres, diameter) / nanoVolThree)
    f_hcpNanoFractionRatio = f_hcpFractionRatioThree
    Debug.Print "NANO:diameter: " & diameter
    Debug.Print "NANO:f_hcpFractionRatioTwo: " & f_totalSphereVolume(f_totalNoSpheres, diameter) & " / " & nanoVolTwo & " = " & f_hcpFractionRatioTwo
End Function

Public Function f_assembleSpheres(diameter As Double)
    f_spheresIntoX (f_eqHeight(diameter))
    f_spheresIntoY (diameter)
    f_spheresIntoZ (f_tetrahedronHeight(diameter))
    f_spheresOnLayerA
    f_spheresOnLayerB
    f_layerCountA
    f_layerCountB
    f_totalNoSpheres
End Function

Public Function f_openExcelObject(excelobjectLocation)
    Set excelobject = GetObject(excelobjectLocation)
    Set excelobjectActive = excelobject.ActiveSheet
    If excelobjectActive Is Nothing Then
        Debug.Print "Cannot find the ExcelobjectActive: " & excelobject.ActiveSheet
        excelobjectOpen = False
        Return
    Else
        excelobjectOpen = True
    End If
End Function

Public Function f_closeExcelObject()
    If excelobjectOpen = True Then
        excelobject.Save
        excelobject.Close
        Shell "TASKKILL /F /IM Excel.exe", vbHide
    End If
End Function

Public Function f_mixingRatios()
    Dim MicronVolume As Double
    Dim NanoVolume As Double
    Dim MicronNum As Double
    Dim NanoNum As Double
    Dim MicronVolumeCHK As Double
    Dim NanoVolumeCHK As Double
    MicronVolumeCHK = Val(hcpVolumeFractionForm.TextBox_micronVolume.Text) * Val(hcpVolumeFractionForm.TextBox_totalMicrons.Text)
    NanoVolumeCHK = Val(hcpVolumeFractionForm.TextBox_nanoVolume.Text) * Val(hcpVolumeFractionForm.TextBox_totalNanos.Text)
    MicronVolume = Val(hcpVolumeFractionForm.TextBox_totalMicronVolume.Text) / Val(hcpVolumeFractionForm.TextBox_totalNanoVolume.Text)
    NanoVolume = Val(hcpVolumeFractionForm.TextBox_totalNanoVolume.Text) / Val(hcpVolumeFractionForm.TextBox_totalNanoVolume.Text)
    MicronNum = Val(hcpVolumeFractionForm.TextBox_totalMicrons.Text) / Val(hcpVolumeFractionForm.TextBox_totalMicrons.Text)
    NanoNum = Val(hcpVolumeFractionForm.TextBox_totalNanos.Text) / Val(hcpVolumeFractionForm.TextBox_totalMicrons.Text)
    hcpVolumeFractionForm.TextBox_RatioMicrons.Text = MicronVolume
    hcpVolumeFractionForm.TextBox_RatioNanos.Text = NanoVolume
End Function

Public Function f_CreateNewPart()
    If MsgBox("Would you like to create a new part?", vbYesNo) = vbYes Then
        Set swApp = CreateObject("SldWorks.Application")
        Set part = swApp.NewPart
    End If
End Function

Public Function f_SetActiveDocument()
    Set swApp = Application.SldWorks
    If hcpVolumeFractionForm.Visible Then
    Else
        hcpVolumeFractionForm.Show
    End If
End Function

Function resetForm()
    hcpVolumeFractionForm.TextBox_diameter = 37000
    hcpVolumeFractionForm.TextBox_area = 191.43
    f_iniParameters
End Function
