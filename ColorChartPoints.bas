Sub ColorChartPointsToTarget(myseries As series, target As Double, Optional UpperColor As XlRgbColor = rgbGreen, Optional LowerColor As XlRgbColor = rgbRed)

Dim mypoint As Point

For i = 1 To myseries.Points.Count
If (myseries.Values(i) > target) Then
myseries.Points.Item(i).Format.Fill.ForeColor.RGB = UpperColor
Else
myseries.Points.Item(i).Format.Fill.ForeColor.RGB = LowerColor
End If


Next i


End Sub

Public Sub Colorpoints()

ColorChartPointsToTarget ActiveChart.SeriesCollection(1), 20, rgbBlue, rgbDarkMagenta

End Sub
