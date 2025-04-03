Attribute VB_Name = "mod_Chart"
Option Explicit

' 2019/11/27 --- adjustment of chart's graph position and x, y scale

Sub adjustChartGraph()
    Dim Q0, Q1, E0, E1, SwQ0, SwQ1, IQ As Double
    
    ' IQ -- Initial Q
    ' SafeYield�̴�. �����
    
    Q0 = Range("D3").Value
    Q1 = Range("D7").Value
    
    E0 = Range("F35").Value
    E1 = Range("F32").Value
    
    SwQ0 = Range("F3").Value
    SwQ1 = Range("F7").Value
    
    IQ = Range("M51").Value
    
    Call setAxisScale("Chart 5", Q0, Q1, SwQ0, SwQ1)
    Call setAxisScale("Chart 7", Q0, Q1, SwQ0, SwQ1)
    
    Call setAxisScale_Efficiency("Chart 8", Q0, Q1, E0, E1)
    
    Call SetGONGBEON
End Sub


Sub SetChartTitleText(ByVal i As Integer)
    
    Call SetGONGBEON
    
    ActiveSheet.ChartObjects("Chart 7").Activate
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "�����(��/day)(W-" & CStr(i) & ")"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "�����(��/day)(W-" & CStr(i) & ")"
    
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "��������Ϸ�(day/��)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "��������Ϸ�(day/��)"
    
    ActiveSheet.ChartObjects("Chart 5").Activate
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "�����(��/day)(W-" & CStr(i) & ")"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "�����(��/day)(W-" & CStr(i) & ")"
    
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "��������Ϸ�(day/��)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "��������Ϸ�(day/��)"
    
    ActiveSheet.ChartObjects("Chart 9").Activate
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "�����(Q)(W-" & CStr(i) & ")"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "�����(Q)(W-" & CStr(i) & ")"
    
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "�������Ϸ�(Sw)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "�������Ϸ�(Sw)"
    
End Sub



Public Sub SetGONGBEON()
    Dim gong As Integer
                  
    If ActiveSheet.name = "Input" Then
        gong = Val(CleanString(Range("J48").Value))
        Range("i54").Value = "W-" & gong
    End If
End Sub



Function determinX(ByVal x0 As Double, ByVal x1 As Double) As Double
    determinX = (x1 - x0) / 3
End Function

Function determinY(ByVal y0 As Double, ByVal y1 As Double) As Double
    'determiney ���� - 2020-6-21
    'y0 = Round(y0 / 10, 0) * 10
    'y1 = Round(y1 / 10, 0) * 10
    
    determinY = (y1 - y0) / 3
End Function

Sub setAxisScale_Efficiency(strName As String, ByVal x0 As Double, ByVal x1 As Double, ByVal y0 As Double, ByVal y1 As Double)
    Dim dresx, dresy As Double
    Dim xMax, xMin, yMax, yMin As Double
    
    dresx = determinX(x0, x1)
    
    xMin = (x0 - dresx)
    xMax = (x1 + dresx)
    
    yMin = WorksheetFunction.RoundDown(y0, -1) - 20
    yMax = WorksheetFunction.RoundUp(y1, -1) + 10
    
    ActiveSheet.ChartObjects(strName).Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MinimumScale = xMin
    ActiveChart.Axes(xlCategory).MaximumScale = xMax
    
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = yMin
    ActiveChart.Axes(xlValue).MaximumScale = yMax
    
    Call setAxisUnit(strName, xMin, xMax, yMin, yMax)
End Sub

Sub setAxisScale(strName As String, ByVal x0 As Double, ByVal x1 As Double, ByVal y0 As Double, ByVal y1 As Double)
    Dim dresx, dresy As Double
    Dim xMax, xMin, yMax, yMin As Double
    
    dresx = determinX(x0, x1)
    dresy = determinY(y0 * 1000, y1 * 1000)
    
    xMin = (x0 - dresx)
    xMax = (x1 + dresx)
    
    yMin = (y0 * 1000 - dresy) / 1000
    yMax = (y1 * 1000 + dresy) / 1000
    
    ActiveSheet.ChartObjects(strName).Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MinimumScale = xMin
    ActiveChart.Axes(xlCategory).MaximumScale = xMax
    
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = yMin
    ActiveChart.Axes(xlValue).MaximumScale = yMax
    
    Call setAxisUnit(strName, xMin, xMax, yMin, yMax)
End Sub

Sub setAxisUnit(strName As String, ByVal x0 As Double, ByVal x1 As Double, ByVal y0 As Double, ByVal y1 As Double)
    Dim dresx, dresy As Double
    Dim xMax, xMin, yMax, yMin As Double
    
    ActiveSheet.ChartObjects(strName).Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MajorUnit = (x1 - x0) / 10
    
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MajorUnit = (y1 - y0) / 4
End Sub
