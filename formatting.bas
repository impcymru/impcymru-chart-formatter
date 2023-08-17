
Function RGBColourQualitative(i As Variant)

    If IsImpCymru.CurrentUserIsImprovementCymru() Then
        RGBColourQualitative = ImpCymruColours.RGBImpCymruColourQualitative(i)
        Exit Function
    End If
    
    Select Case i
        Case 1
            RGBColourQualitative = RGB(27, 158, 119)
        Case 2
            RGBColourQualitative = RGB(217, 95, 2)
        Case 3
            RGBColourQualitative = RGB(117, 112, 179)
        Case 4
            RGBColourQualitative = RGB(231, 112, 179)
        Case Else
            RGBColourQualitative = RGB(50 + Rnd(100), 50 + Rnd(100), 50 + Rnd(100))
    End Select

End Function

Function RGBColourTitle()

    If IsImpCymru.CurrentUserIsImprovementCymru() Then
        RGBColourTitle = ImpCymruColours.RGBImpCymruColourPrimary("Navy")
        Exit Function
    End If
    
    RGBColourTitle = RGB(35, 35, 35)

End Function

Function RGBColourRunChartCentreLine()
    
    If IsImpCymru.CurrentUserIsImprovementCymru() Then
        RGBColourRunChartCentreLine = ImpCymruColours.RGBImpCymruColourAnalytical("GoldenHamster")
        Exit Function
    End If
    
    RGBColourRunChartCentreLine = RGB(0, 0, 0)
    
End Function


Function RGBColourSPCCentreLine()
    
    If IsImpCymru.CurrentUserIsImprovementCymru() Then
        RGBColourSPCCentreLine = ImpCymruColours.RGBImpCymruColourAnalytical("GoldenHamster")
        Exit Function
    End If
    
    RGBColourSPCCentreLine = RGBColourRunChartCentreLine()
    
End Function


Function RGBColourSPCControlLimit()
    
    If IsImpCymru.CurrentUserIsImprovementCymru() Then
        RGBColourSPCControlLimit = RGB(180, 180, 180)
        Exit Function
    End If
    
    RGBColourSPCControlLimit = RGB(123, 125, 125)
    
End Function




