Function RGBImpCymruColourQualitative(i As Variant)
        
    Select Case i
        Case 1
            RGBImpCymruColourQualitative = RGBImpCymruColourAnalytical("NightTrain")
        Case 2
            RGBImpCymruColourQualitative = RGBImpCymruColourAnalytical("ValentineHeart")
        Case 3
            RGBImpCymruColourQualitative = RGBImpCymruColourAnalytical("GoldenHamster")
        Case 4
            RGBImpCymruColourQualitative = RGBImpCymruColourAnalytical("Berry")
        Case Else
            RGBImpCymruColourQualitative = RGB(50 + Rnd(100), 50 + Rnd(100), 50 + Rnd(100))
    End Select

End Function


Function RGBImpCymruColourPrimary(i As Variant)
    ' These are taken from the Primary colour palette of the Improvement Cymru brand guidelines

    Select Case i
        Case "Navy"
            RGBImpCymruColourPrimary = RGB(27, 87, 104)
        Case "Teal"
            RGBImpCymruColourPrimary = RGB(0, 154, 158)
        Case "Green"
            RGBImpCymruColourPrimary = RGB(27, 87, 104)
        Case "Purple"
            RGBImpCymruColourPrimary = RGB(87, 60, 114)
        Case "Orange"
            RGBImpCymruColourPrimary = RGB(206, 133, 1)
        Case "Pink"
            RGBImpCymruColourPrimary = RGB(173, 79, 132)
    End Select

End Function


Function RGBImpCymruColourAnalytical(i As Variant)

    ' https://www.hsluv.org/
    ' https://contrastchecker.com/
    ' TODO: Target WCAG AA against white backgrounds
    ' Or provide &"AA" versions of the colours?
    
    Select Case i
        Case "NightTrain"
            RGBImpCymruColourAnalytical = RGB(74, 121, 134)
            ' WCAG 2.0 1.4.3 'AA' compliant
        Case "ValentineHeart"
            RGBImpCymruColourAnalytical = RGB(190, 114, 157)
        Case "Berry"
            RGBImpCymruColourAnalytical = RGB(157, 15, 78)
            ' WCAG 2.0 1.4.3 'AA' compliant
            ' WCAG 2.0 1.4.6 'AAA' compliant
        Case "GoldenHamster"
            RGBImpCymruColourAnalytical = RGB(216, 159, 62)
    End Select

End Function
