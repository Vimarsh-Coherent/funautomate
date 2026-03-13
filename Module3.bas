Attribute VB_Name = "Module3"

Function AddDoughnutChart(Rng1 As Range, ShapeIndex As Integer)

    With pptSlide.Shapes.AddChart2(251, Type:=xlDoughnut, Top:=pptSlide.Shapes(ShapeIndex).Top, Width:=pptSlide.Shapes(ShapeIndex).Width, Height:=pptSlide.Shapes(ShapeIndex).Height - 50).Chart
        Application.Wait (Now + TimeValue("0:00:2"))

        'Changing the source data
        '.ChartData.Workbook.Worksheets(1).Activate
        .ChartData.Workbook.Worksheets(1).UsedRange.ClearContents
        Application.Wait (Now + TimeValue("0:00:1"))
        
        Rng1.Copy
        Err.Clear
        On Error Resume Next
        .ChartData.Workbook.Worksheets(1).Cells(1, 1).PasteSpecial ppPasteText
        If Err.Description <> "" Then
            Rng1.Copy
            .ChartData.Workbook.Worksheets(1).Cells(1, 1).PasteSpecial ppPasteText
        End If
        On Error GoTo 0
        
        Application.Wait (Now + TimeValue("0:00:1"))
        .SetSourceData Source:="='Sheet1'!" & .ChartData.Workbook.Worksheets(1).Cells(1, 1).CurrentRegion.Address, PlotBy:=xlColumns
        Application.Wait (Now + TimeValue("0:00:1"))
        .FullSeriesCollection(1).XValues = "=Sheet1!$A$2:$A$" & .ChartData.Workbook.Worksheets(1).Cells(.ChartData.Workbook.Worksheets(1).Rows.Count, "A").End(xlUp).Row
        
        .ChartData.Workbook.RefreshAll
        Application.Wait (Now + TimeValue("0:00:1"))
        
        .ChartData.Workbook.Close
        Application.Wait (Now + TimeValue("0:00:1"))
        
        'Chart Titel
        '.HasTitle = True
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 12
        '.SetElement (msoElementChartTitleCenteredOverlay)
        .SetElement msoElementChartTitleCenteredOverlay
        
        .FullSeriesCollection(1).Select
        .ChartGroups(1).DoughnutHoleSize = 50
        With .FullSeriesCollection(1)
            .ApplyDataLabels
            .Select
            '.DataLabels.Position = xlLabelPositionAbove
            .DataLabels.ShowValue = False
            .DataLabels.ShowPercentage = True
            .DataLabels.NumberFormat = "0.0%"
            .DataLabels.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .DataLabels.Format.TextFrame2.TextRange.Font.Bold = msoTrue
            .DataLabels.Format.TextFrame2.TextRange.Font.Size = 8
            .DataLabels.Format.TextFrame2.VerticalAnchor = msoAnchorTop
            .DataLabels.Format.TextFrame2.HorizontalAnchor = msoAnchorNone
            .HasLeaderLines = True
        End With
        
        '''''''''
        'Make Data point as XX for Sample
        If Sheet15.Range("B3").Value = "sample" Or Sheet15.Range("B3").Value = "report storybording" Then
            For i = 1 To .FullSeriesCollection(1).Points.Count
                .FullSeriesCollection(1).Points(i).DataLabel.Formula = "xx.x%"
            Next i
        End If
        
        '''''''
        
        .HasLegend = False
    End With
    
End Function

