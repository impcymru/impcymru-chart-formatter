
Option Explicit

'Callback for reformatChartButton onAction
Sub ReformatChart(control As IRibbonControl)
    ImpCymruFormatActiveChart
End Sub


Sub ImpCymruFormatActiveChart()
    Dim cht As Chart

    Set cht = ActiveChart
    
    If cht Is Nothing Then
        ' No chart selected
        Exit Sub
    End If
    
    Dim chtSeries As Series
    
    cht.ChartArea.RoundedCorners = False
    cht.ChartArea.Format.Line.Visible = False
    ' TODO: Consider transparency, e.g. given hexagons on ImpCymru PowerPoint templates
    ' Possibly layering white, slightly transparent fill on both ChartArea and PlotArea
    ' Maybe also chart title backgrounds etc.
    
    MakeChart2D cht
    MakePieIntoBar cht
    
    Dim isSPCWithLinesCounter As Integer
    isSPCWithLinesCounter = 0
    
    Dim isLineChartCounter As Integer
    isLineChartCounter = 0
    
    Dim isBarOrColumnChartCounter As Integer
    isBarOrColumnChartCounter = 0
    
    Dim isRunChartCounter As Integer
    isRunChartCounter = 0

    
    For Each chtSeries In cht.SeriesCollection
    
        If chtSeries.Name = "CL" Or _
            chtSeries.Name = "Center" Or _
            chtSeries.Name = "Centre" Or _
            chtSeries.Name = "Median" Then
            isRunChartCounter = isRunChartCounter + 1
        End If
    
        If chtSeries.Name = "UCL" Or _
            chtSeries.Name = "CL" Or _
            chtSeries.Name = "Center" Or _
            chtSeries.Name = "Centre" Or _
            chtSeries.Name = "LCL" Then
            isSPCWithLinesCounter = isSPCWithLinesCounter + 1
        End If
        
        If chtSeries.Type = xlLine Then
            isLineChartCounter = isLineChartCounter + 1
        End If
        
        If chtSeries.Type = xlBar Or chtSeries.Type = xlColumn Then
            isBarOrColumnChartCounter = isBarOrColumnChartCounter + 1
        End If
        
    
    Next chtSeries
    
    If isSPCWithLinesCounter = 3 And isLineChartCounter = 4 Then
        FormatSPCWithLines cht
    ElseIf isRunChartCounter = 1 And isLineChartCounter = 2 Then
        FormatRunChart cht
    ElseIf isLineChartCounter > 0 Then
        FormatLineChart cht
    End If

    
    'TODO: Check if this is better to handle with cht.ChartType
    If isBarOrColumnChartCounter > 0 Then
        If isLineChartCounter = 0 Then
            ReorientBarOrColumnChart cht
        End If
    
        FormatBarOrColumnChart cht
    End If

    SetChartSize cht
    
    If cht.ChartType = xlPie Or _
        cht.ChartType = xlPieExploded Or _
        cht.ChartType = xlDoughnut Or _
        cht.ChartType = xlDoughnutExploded Then
        FormatPieChart cht
    End If
    
    'TODO - support Paretos
    
    'TODO - Support a subtitle
    ' This has implications both for formatting the titles themselves
    ' And for chart layout
    ' Probably needs a special named textbox
    
    FormatChartTitles cht
    FormatGridLines cht
    
    FormatChartLayout cht

    FormatDateAxes cht
    FormatPercentAxes cht
    'TODO - FormatNumberAxes e.g. "10,000" What about currencies?
    
    AddCaption cht
End Sub

Sub SetChartSize(cht As Chart)

    Dim minChartHeightPoints As Long
    minChartHeightPoints = Application.CentimetersToPoints(10)
    
    Dim maxChartWidthPoints As Long
    maxChartWidthPoints = Application.CentimetersToPoints(21)

    
    If cht.ChartArea.Height < minChartHeightPoints Then
        cht.ChartArea.Height = minChartHeightPoints
    End If
    
    Dim minChartWidthPoints As Long
    minChartWidthPoints = cht.ChartArea.Height
    
    Dim maxPointsCount As Double
    Dim seriesPointsCount As Double
    maxPointsCount = 1
    
    Dim chtSeries As Series
    
    For Each chtSeries In cht.SeriesCollection
        If chtSeries.Type = xlLine Or chtSeries.Type = xlColumn Or chtSeries.Type = xlBar Then
            seriesPointsCount = chtSeries.Points.Count
            
            'HACK
            If chtSeries.Type = xlColumn Then
                seriesPointsCount = 0.8 * seriesPointsCount
            End If
            
            If seriesPointsCount > maxPointsCount Then
                maxPointsCount = seriesPointsCount
            End If
        End If
    Next chtSeries
    
    
    If cht.ChartType = xlBarStacked Or _
        cht.ChartType = xlBarClustered Or _
        cht.ChartType = xlBarStacked100 Or _
        cht.ChartType = xl3DBarStacked100 Then
        
        cht.ChartArea.Width = maxChartWidthPoints * 0.8
        cht.ChartArea.Height = Application.CentimetersToPoints(3) + _
            (maxPointsCount * cht.SeriesCollection.Count * Application.CentimetersToPoints(0.7))
        
        
        
        
        Exit Sub
    End If
        
    
    Dim chartWidthPoints As Long
    chartWidthPoints = Application.CentimetersToPoints(3) + (maxPointsCount * Application.CentimetersToPoints(1.8))
    
    If chartWidthPoints < minChartWidthPoints Then
        chartWidthPoints = minChartWidthPoints
    ElseIf chartWidthPoints > maxChartWidthPoints Then
        chartWidthPoints = maxChartWidthPoints
    End If
    
    cht.ChartArea.Width = chartWidthPoints
    

End Sub

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



Sub MakeChart2D(cht As Chart)
    Select Case cht.ChartType
        Case xl3DBarClustered
            cht.ChartType = xlBarClustered
        Case xl3DBarStacked
            cht.ChartType = xlBarStacked
        Case xl3DColumnClustered
            cht.ChartType = xlColumnClustered
        Case xl3DColumnStacked
            cht.ChartType = xlColumnStacked
        Case xl3DPie
            cht.ChartType = xlPie
        Case xl3DPieExploded
            cht.ChartType = xlPie 'Not exploded
        Case xl3DLine
            cht.ChartType = xlLine
    End Select
End Sub

Sub MakePieIntoBar(cht As Chart)
    ' Pie charts shouldn't have many different segments
    If cht.ChartType = xlPie Then
        If UBound(cht.SeriesCollection(1).XValues) > 4 Then
            cht.ChartType = xlBarClustered
        End If
    End If
End Sub

Sub FormatPercentAxes(cht As Chart)
    Dim ax As Axis
    
        
    If Not cht.HasAxis(xlValue) Then
        Exit Sub
    End If
    
    
    Set ax = cht.Axes(xlValue)

    If ax.TickLabelPosition = xlTickLabelPositionNone Then
        Exit Sub
    End If
    
    
    If Right(ax.TickLabels.NumberFormat, 1) <> "%" Then
        Exit Sub
    End If
    
    Dim initialRange As Double
    initialRange = ax.MaximumScale - ax.MinimumScale
    
    If ax.MinimumScale > 0 And (ax.MinimumScale - (0.5 * initialRange)) < 0 Then
        ax.MinimumScale = 0
    End If
    
    If ax.MaximumScale < 1 And (ax.MaximumScale + (0.5 * initialRange)) > 1 Then
        ax.MaximumScale = 1
    End If


    'VBA Mod only deals in Integers
    If ((100 * 100 * ax.MajorUnit) Mod 100) = 0 Then
        
        ax.TickLabels.NumberFormat = "0%"
    End If
    

End Sub

Sub FormatDateAxes(cht As Chart)

    Dim ax As Axis
    
    If Not cht.HasAxis(xlCategory) Then
        Exit Sub
    End If
    
    Set ax = cht.Axes(xlCategory)
    
    On Error Resume Next
    If ax.CategoryType = xlCategoryScale Then
        On Error GoTo 0
        Exit Sub
    End If
    
    Dim minScale As Variant
    minScale = Null

    minScale = ax.MinimumScale
    On Error GoTo 0
    
    If IsNull(minScale) Then
        Exit Sub
    End If
    
    If ax.BaseUnit = xlDays Then
        ax.TickLabels.NumberFormat = "dd mmm yyyy"
    ElseIf ax.BaseUnit = xlMonths Then
        ax.TickLabels.NumberFormat = "mmm yyyy"
    ElseIf ax.BaseUnit = xlYears Then
        ax.TickLabels.NumberFormat = "yyyy"
    End If

End Sub

Sub AddCaption(cht As Chart)

    Dim captionName As String
    captionName = "Caption"
    
    Dim shp As Shape
    
    For Each shp In cht.Shapes
        If shp.Name = captionName Then
            FormatCaption shp
            Exit Sub
        End If
    Next shp

    Dim txtBox As Shape
    Set txtBox = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, 8, cht.ChartArea.Height - 20, cht.ChartArea.Width - (2 * 8), 16)
    
    txtBox.Name = captionName
    txtBox.TextFrame.Characters.Text = "Created by Improvement Cymru"
    FormatCaption txtBox

End Sub

Sub FormatCaption(txtBox As Shape)
    txtBox.TextFrame.HorizontalAlignment = xlHAlignRight
    txtBox.TextFrame.VerticalAlignment = xlVAlignCenter
    txtBox.TextFrame.Characters.Font.Name = "Arial"
    txtBox.TextFrame.Characters.Font.Color = RGB(120, 120, 120)
    txtBox.TextFrame.Characters.Font.Size = 8
    txtBox.Top = txtBox.Parent.ChartArea.Height - 20
    txtBox.Left = 8
    txtBox.Width = txtBox.Parent.ChartArea.Width - (2 * 8)
    txtBox.Height = 16
    
End Sub

Sub FormatChartLayout(cht As Chart)

    ' Pie charts are special cased because of the way that the data labels work
    If cht.ChartType = xlPie Or cht.ChartType = xlPieExploded Then
        FormatPieChartLayout cht
        Exit Sub
    End If
       
    If cht.HasTitle Then
        cht.PlotArea.Height = cht.ChartArea.Height - (cht.ChartTitle.Height + cht.ChartTitle.Top) - 38
        cht.PlotArea.Top = cht.ChartTitle.Height + cht.ChartTitle.Top + 12
    Else
        cht.PlotArea.Top = 15
        cht.PlotArea.Height = cht.ChartArea.Height - 48
    End If
    
    Dim rightMargin As Long
    cht.PlotArea.Left = 15
    rightMargin = 25
    
    If cht.Axes(xlValue).HasTitle Then
        If cht.Axes(xlValue).AxisTitle.Left < (0.5 * cht.ChartArea.Width) Then
            cht.Axes(xlValue).AxisTitle.Left = 5
            cht.PlotArea.Left = cht.PlotArea.Left + cht.Axes(xlValue).AxisTitle.Left + cht.Axes(xlValue).AxisTitle.Width - 5
        Else
            cht.Axes(xlValue).AxisTitle.Left = cht.ChartArea.Width - cht.Axes(xlValue).AxisTitle.Width - 10
            rightMargin = rightMargin + cht.Axes(xlValue).AxisTitle.Width - 10
        End If
        
    End If
    
    cht.PlotArea.Width = cht.ChartArea.Width - cht.PlotArea.Left - rightMargin
    
End Sub

Sub FormatPieChartLayout(cht As Chart)

    If cht.HasTitle Then
        cht.PlotArea.Height = cht.ChartArea.Height - (cht.ChartTitle.Height + cht.ChartTitle.Top) - 100
        cht.PlotArea.Top = cht.ChartTitle.Height + cht.ChartTitle.Top + 30
    Else
        cht.PlotArea.Top = 30
        cht.PlotArea.Height = cht.ChartArea.Height - 100
    End If
        

    cht.PlotArea.Left = cht.ChartArea.Width / 2 - cht.PlotArea.Width / 2
End Sub

Sub FormatChartTitles(cht As Chart)

    If cht.HasTitle Then
        cht.ChartTitle.Format.TextFrame2.TextRange.Font.Name = "Arial"
        cht.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = msoFalse
        cht.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 16
        cht.ChartTitle.HorizontalAlignment = xlHAlignLeft
        cht.ChartTitle.Left = 15
        cht.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGBImpCymruColourPrimary("Navy")
    End If
    

    Dim ax As Axis

    For Each ax In cht.Axes()
        If ax.HasTitle Then
    
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Name = "Arial"
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoFalse
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 9
            ax.AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
            ax.AxisTitle.Orientation = xlHorizontal
        End If
    
        If ax.TickLabelPosition <> xlTickLabelPositionNone Then
            ax.TickLabels.Font.Name = "Arial"
            ax.TickLabels.Font.Bold = msoFalse
            ax.TickLabels.Font.Size = 9
            ax.TickLabels.Format.Fill.Visible = msoFalse
            ax.TickLabels.Font.Color = RGB(60, 60, 60)
            
            
            ' Force TickLabels to be horizontal if they're either:
            ' - Set to a custom angle less than -45deg or greater than 15deg
            ' (The sign of custom angles is inverted, apparently.)
            ' - Set to a XlTickLabelOrientation constant
            '
            ' This allows *some* leeway for cramming labels in that are slightly rotated from horizontal
            
            If ax.TickLabels.Orientation < -15 Or ax.TickLabels.Orientation > 45 Then
                ax.TickLabels.Orientation = xlTickLabelOrientationHorizontal
            End If
        End If
        
    Next ax


End Sub

Sub FormatGridLines(cht As Chart)
    Dim ax As Axis
    
    For Each ax In cht.Axes()
    
        If ax.MajorGridlines.Format.Line.Visible = msoTrue Then
            ax.MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 230, 230)
            ax.MajorGridlines.Format.Line.Weight = 0.5
            ax.MajorGridlines.Format.Line.DashStyle = msoLineSolid
            ax.MajorTickMark = xlTickMarkOutside
        End If
        
        ax.MinorGridlines.Format.Line.Visible = msoFalse
        ax.MinorTickMark = xlTickMarkNone
        
        Dim gridlineCount As Long
        
        ' TODO - Handle dates
        If ax.Type = xlValue Then
            gridlineCount = 1 + (ax.MaximumScale - ax.MinimumScale) / ax.MajorUnit
            
            If gridlineCount > 8 Then 'HACK: This is quite crude!
                ax.MajorUnit = ax.MajorUnit * 2
                FormatGridLines cht
            End If
            
        End If
    
    Next ax
    
    
End Sub

Sub ReorientBarOrColumnChart(cht As Chart)
    ' This tries to decide if a vertical bar chart should be a horizontal bar chart
    
    ' TODO: Cope with nested
    Dim srs As Series
    Set srs = cht.SeriesCollection(1)
    
    Dim maxLength As Long
    Dim totalLength As Long
    Dim meanLength As Double
    
    maxLength = 0
    totalLength = 0
    
    
    Dim xv As Variant
    For Each xv In srs.XValues
        If Len(xv) > maxLength Then
            maxLength = Len(xv)
        End If
        
        totalLength = totalLength + Len(xv)
    
    Next xv

    meanLength = totalLength / (UBound(srs.XValues))
    
    If meanLength > 10 Or maxLength > 15 Then
        If cht.ChartType = xlColumnStacked Then
            cht.ChartType = xlBarStacked
        End If
        
        If cht.ChartType = xlColumnClustered Then
            cht.ChartType = xlBarClustered
        End If
        
    End If

End Sub

Sub FormatPieChart(cht As Chart)
    
    cht.ChartType = xlPie 'Rather than xlPieExploded or a Doughnut
    
    Dim pg As ChartGroup
    
    For Each pg In cht.PieGroups
        pg.FirstSliceAngle = 0
    Next pg
    
    Dim srs As Series
    
    Set srs = cht.SeriesCollection(1)
   
    
    srs.ApplyDataLabels (xlDataLabelsShowLabelAndPercent)
    
    srs.DataLabels.Font.Name = "Arial"
    srs.DataLabels.Font.Size = 10

    Dim pnt As Point
 
    Dim i As Long
    For i = 1 To srs.Points.Count
        Set pnt = srs.Points(i)
        
        pnt.Format.Fill.ForeColor.RGB = RGBImpCymruColourQualitative(i)
        pnt.DataLabel.Position = xlLabelPositionOutsideEnd
        pnt.DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = pnt.Format.Fill.ForeColor.RGB
 
 
        With pnt.DataLabel.Format.TextFrame2.TextRange.Font.Glow
           .Color.RGB = RGB(255, 255, 255)
           .Transparency = 0
           .Radius = 10
        End With
    
    Next i
    
    If cht.HasLegend Then
        cht.Legend.Delete
    End If

End Sub

Sub FormatBarOrColumnChart(cht As Chart)
    
    ' Bar charts should almost always start at zero
    ' otherwise their area/length is misleading
    Dim ax As Axis
    Set ax = cht.Axes(xlValue)
    If ax.MinimumScale > 0 Then
        ax.MinimumScale = 0
    ElseIf ax.MaximumScale < 0 Then
        ax.MaximumScale = 0
    End If

    Dim chtSeries As Series
    
    Dim dataSeriesIndex As Long
    dataSeriesIndex = 1
    For Each chtSeries In cht.SeriesCollection
        If chtSeries.Type = xlBar Or chtSeries.Type = xlColumn Then
            FormatBarOrColumnSeries chtSeries, dataSeriesIndex
            dataSeriesIndex = dataSeriesIndex + 1
        End If
    
    Next chtSeries
    
    If dataSeriesIndex = 2 Then
        'Only one bar series
        
        For Each chtSeries In cht.SeriesCollection
            If chtSeries.Type = xlBar Or chtSeries.Type = xlColumn Then
            
                ' This helps with Data Label formatting
                If chtSeries.ChartType = xlBarStacked Then
                    chtSeries.ChartType = xlBarClustered
                ElseIf chtSeries.ChartType = xlColumnStacked Then
                    chtSeries.ChartType = xlColumnClustered
                End If
            
                FormatOtherTypeBarOrColumnSeries chtSeries
            End If
        Next chtSeries
        
    End If
    
        
    If cht.SeriesCollection.Count = 1 And cht.HasLegend Then
        cht.Legend.Delete
    End If
    
    For Each chtSeries In cht.SeriesCollection
        If chtSeries.Type = xlBar Or chtSeries.Type = xlColumn Then
            FormatDataLabelsForBarOrColumnSeries chtSeries
        End If
    Next chtSeries
    
    ' This tries to implement ONS guidance
    ' Single-series charts should have gaps that are considerably narrower than the bars
    ' Clustered bar charts should have gaps that are 'slightly wider than a single bar'
    
    Dim TargetGapWidth As Long
    TargetGapWidth = 40
    
    If cht.SeriesCollection.Count > 1 Then
        TargetGapWidth = 105
    End If
    
    Dim barGrp As ChartGroup
    
    
    If cht.ChartType <> xlBarStacked And cht.ChartType <> xlColumnStacked And _
        cht.ChartType <> xlBarStacked100 And cht.ChartType <> xlColumnStacked100 Then
    
        For Each barGrp In cht.BarGroups
            barGrp.GapWidth = TargetGapWidth
            barGrp.Overlap = 0
        Next barGrp
        
        For Each barGrp In cht.ColumnGroups
            barGrp.GapWidth = TargetGapWidth
            barGrp.Overlap = 0
        Next barGrp
    Else
          For Each barGrp In cht.BarGroups
            barGrp.GapWidth = TargetGapWidth
            barGrp.Overlap = 100
        Next barGrp
        
        For Each barGrp In cht.ColumnGroups
            barGrp.GapWidth = TargetGapWidth
            barGrp.Overlap = 100
        Next barGrp
    
    End If

End Sub

Sub FormatDataLabelsForBarOrColumnSeries(chtSeries As Series)

    Dim pnt As Point
    
    If chtSeries.HasDataLabels Then
        chtSeries.DataLabels.Font.Name = "Arial"
    End If
    
    Dim pntColour As Variant
    
    For Each pnt In chtSeries.Points
        If pnt.HasDataLabel Then
            pnt.DataLabel.Font.Name = "Arial"
            pnt.DataLabel.Position = xlLabelPositionInsideEnd
            pnt.DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                
            pntColour = pnt.Format.Fill.ForeColor.RGB

            With pnt.DataLabel.Format.TextFrame2.TextRange.Font.Glow
                .Color.RGB = pntColour
                .Transparency = 0
                .Radius = 10
            End With
                

            If chtSeries.Type = xlBar And pnt.DataLabel.Width > pnt.Width Then
            
                If chtSeries.ChartType = xlBarClustered Then
                    pnt.DataLabel.Position = xlLabelPositionOutsideEnd
                ElseIf chtSeries.ChartType = xlBarStacked Then
                    pnt.DataLabel.Position = xlLabelPositionInsideBase
                End If
                
                pnt.DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = pntColour

                With pnt.DataLabel.Format.TextFrame2.TextRange.Font.Glow
                   .Color.RGB = RGB(255, 255, 255)
                   .Transparency = 0
                   .Radius = 10
                End With

            End If
            
            If chtSeries.Type = xlColumn And pnt.DataLabel.Height > pnt.Height Then

                If chtSeries.ChartType = xlColumnClustered Then
                    pnt.DataLabel.Position = xlLabelPositionOutsideEnd
                ElseIf chtSeries.ChartType = xlColumnStacked Then
                    pnt.DataLabel.Position = xlLabelPositionInsideBase
                End If

                pnt.DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = pntColour
                
                With pnt.DataLabel.Format.TextFrame2.TextRange.Font.Glow
                   .Color.RGB = RGB(255, 255, 255)
                   .Transparency = 0
                   .Radius = 10
                End With
                
            End If
        End If
    Next pnt

End Sub

Sub FormatOtherTypeBarOrColumnSeries(chtSeries As Series)
    ' This tries to find Other or Missing bars and re colour them
    
    Dim i As Long
    For i = LBound(chtSeries.XValues) To UBound(chtSeries.XValues)
        
        If _
            chtSeries.XValues(i) = "Other" Or _
            chtSeries.XValues(i) = "Missing" Or _
            chtSeries.XValues(i) = "NA" Or _
            chtSeries.XValues(i) = "Unknown" Then
            chtSeries.Points(i).Format.Fill.ForeColor.RGB = RGB(100, 100, 100)
        End If
    Next i

End Sub

Sub FormatBarOrColumnSeries(chtSeries As Series, Optional dataSeriesIndex As Variant)

    chtSeries.Format.Fill.ForeColor.RGB = RGBImpCymruColourQualitative(dataSeriesIndex)
    chtSeries.Format.Line.Visible = msoFalse

End Sub

Sub FormatLineChart(cht As Chart)
    Dim dataSeriesIndex As Long
    dataSeriesIndex = 1
    
    Dim chtSeries As Series

    For Each chtSeries In cht.SeriesCollection
        If chtSeries.Type = xlLine Then
            FormatBasicLineChartSeries chtSeries, dataSeriesIndex
            dataSeriesIndex = dataSeriesIndex + 1
        End If
    
    Next chtSeries
    
    If cht.SeriesCollection.Count = 1 And cht.HasLegend Then
        cht.Legend.Delete
    End If
    
End Sub

Sub FormatRunChart(cht As Chart)
    
    Dim chtSeries As Series

    For Each chtSeries In cht.SeriesCollection
        If chtSeries.Name = "CL" Or _
            chtSeries.Name = "Center" Or _
            chtSeries.Name = "Centre" Or _
            chtSeries.Name = "Median" Then
            FormatRunChartCL chtSeries
        Else
            FormatBasicLineChartSeries chtSeries, 1
            chtSeries.PlotOrder = 2
        End If
    Next chtSeries
    
    If cht.HasLegend Then
        cht.Legend.Delete
    End If

End Sub


Sub FormatRunChartCL(chtSeries As Series)

    Dim PntVisible() As MsoTriState
    Dim PntDashStyle() As MsoLineDashStyle
    
    ReDim PntVisible(1 To chtSeries.Points.Count)
    ReDim PntDashStyle(1 To chtSeries.Points.Count)

    Dim i As Long

    For i = LBound(PntVisible) To UBound(PntVisible)
        PntVisible(i) = chtSeries.Points(i).Format.Line.Visible
        PntDashStyle(i) = chtSeries.Points(i).Format.Line.DashStyle
    Next i
    
    chtSeries.Format.Line.ForeColor.RGB = RGBImpCymruColourAnalytical("GoldenHamster")
    chtSeries.Format.Line.Weight = 1.5
    chtSeries.MarkerStyle = xlMarkerStyleNone

    
    For i = LBound(PntVisible) To UBound(PntVisible)
        chtSeries.Points(i).Format.Line.Visible = PntVisible(i)
        If PntVisible(i) = msoTrue Then
            chtSeries.Points(i).Format.Line.DashStyle = PntDashStyle(i)
        End If
    Next i
    
End Sub


Sub FormatSPCWithLines(cht As Chart)

    Dim chtSeries As Series

    For Each chtSeries In cht.SeriesCollection
    
        If chtSeries.Name = "UCL" Then
            FormatSPCWithLinesUCLOrLCL chtSeries
        ElseIf chtSeries.Name = "LCL" Then
            FormatSPCWithLinesUCLOrLCL chtSeries
        ElseIf chtSeries.Name = "CL" Or _
            chtSeries.Name = "Center" Or _
            chtSeries.Name = "Centre" Then
            FormatSPCWithLinesCL chtSeries
        Else
            FormatBasicLineChartSeries chtSeries, 1
            chtSeries.PlotOrder = 4
        End If
        
    Next chtSeries
    
    If cht.HasLegend Then
        cht.Legend.Delete
    End If

End Sub

Sub FormatBasicLineChartSeries(chtSeries As Series, Optional dataSeriesIndex As Variant)
    Dim PntVisible() As MsoTriState
    Dim PntDashStyle() As MsoLineDashStyle
    
    ReDim PntVisible(1 To chtSeries.Points.Count)
    ReDim PntDashStyle(1 To chtSeries.Points.Count)

    ' TODO: Respect individually recoloured points?

    Dim i As Long
    For i = LBound(PntVisible) To UBound(PntVisible)
        PntVisible(i) = chtSeries.Points(i).Format.Line.Visible
        PntDashStyle(i) = chtSeries.Points(i).Format.Line.DashStyle
    Next i
    
    chtSeries.MarkerStyle = xlMarkerStyleCircle
    chtSeries.MarkerSize = 7
    chtSeries.MarkerForegroundColorIndex = xlColorIndexNone
    
    chtSeries.MarkerBackgroundColor = RGB(255, 255, 255)
    chtSeries.Format.Line.ForeColor.RGB = RGBImpCymruColourQualitative(dataSeriesIndex)
    
    chtSeries.Format.Line.Weight = 3
    
    If chtSeries.HasDataLabels Then
        chtSeries.DataLabels.Font.Name = "Arial"
    End If
    
    
    
    For i = LBound(PntVisible) To UBound(PntVisible)
        chtSeries.Points(i).Format.Line.Visible = PntVisible(i)
        If PntVisible(i) = msoTrue Then
            chtSeries.Points(i).Format.Line.DashStyle = PntDashStyle(i)
        End If
        
        If chtSeries.Points(i).HasDataLabel Then
            chtSeries.Points(i).DataLabel.Font.Name = "Arial"
            chtSeries.Points(i).DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = chtSeries.Format.Line.ForeColor.RGB

            With chtSeries.Points(i).DataLabel.Format.TextFrame2.TextRange.Font.Glow
               .Color.RGB = RGB(255, 255, 255)
               .Transparency = 0
               .Radius = 10
            End With


        End If
        
    Next i

End Sub


Sub FormatSPCWithLinesUCLOrLCL(chtSeries As Series)
    
    Dim PntVisible() As MsoTriState
    Dim PntDashStyle() As MsoLineDashStyle
    
    ReDim PntVisible(1 To chtSeries.Points.Count)
    ReDim PntDashStyle(1 To chtSeries.Points.Count)

    Dim i As Long
    For i = LBound(PntVisible) To UBound(PntVisible)
        PntVisible(i) = chtSeries.Points(i).Format.Line.Visible
        PntDashStyle(i) = chtSeries.Points(i).Format.Line.DashStyle
    Next i
    
    chtSeries.Format.Line.ForeColor.RGB = RGB(180, 180, 180)
    chtSeries.Format.Line.Weight = 1.5
    chtSeries.MarkerStyle = xlMarkerStyleNone
    
    For i = LBound(PntVisible) To UBound(PntVisible)
        chtSeries.Points(i).Format.Line.Visible = PntVisible(i)
        If PntVisible(i) = msoTrue Then
            chtSeries.Points(i).Format.Line.DashStyle = PntDashStyle(i)
        End If
    Next i

End Sub

Sub FormatSPCWithLinesCL(chtSeries As Series)

    Dim PntVisible() As MsoTriState
    Dim PntDashStyle() As MsoLineDashStyle
    
    ReDim PntVisible(1 To chtSeries.Points.Count)
    ReDim PntDashStyle(1 To chtSeries.Points.Count)

    Dim i As Long

    For i = LBound(PntVisible) To UBound(PntVisible)
        PntVisible(i) = chtSeries.Points(i).Format.Line.Visible
        PntDashStyle(i) = chtSeries.Points(i).Format.Line.DashStyle
    Next i
    
    chtSeries.Format.Line.ForeColor.RGB = RGBImpCymruColourAnalytical("GoldenHamster")
    chtSeries.Format.Line.Weight = 1.5
    chtSeries.MarkerStyle = xlMarkerStyleNone

    
    For i = LBound(PntVisible) To UBound(PntVisible)
        chtSeries.Points(i).Format.Line.Visible = PntVisible(i)
        If PntVisible(i) = msoTrue Then
            chtSeries.Points(i).Format.Line.DashStyle = PntDashStyle(i)
        End If
    Next i

End Sub


