Attribute VB_Name = "ClrActiveChartSeries"
Option Explicit
Sub ClrActiveChartSeries()

'Purpose:   Remove all series from the active chart

Define_Variables:

  Dim s                                         'iterative series variable
  
Delete_All_Series_From_Active_Chart:
  
  For Each s In ActiveChart.SeriesCollection
      s.Delete
  Next s
  
End Sub
