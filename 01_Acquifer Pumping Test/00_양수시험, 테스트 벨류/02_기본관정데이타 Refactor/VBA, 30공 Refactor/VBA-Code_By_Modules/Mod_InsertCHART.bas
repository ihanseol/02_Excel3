Attribute VB_Name = "Mod_InsertCHART"
Option Explicit

Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)


Sub DeleteAllCharts()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    
    Set ws = ThisWorkbook.Worksheets("AggChart")
    
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj
End Sub

Sub DeleteAllImages(ByVal singleWell As Integer)
    Dim ws As Worksheet
    Dim sh As Shape
    
    Set ws = ThisWorkbook.Worksheets("AggChart")
    
    If singleWell = 999 Then
        For Each sh In ws.Shapes
            If sh.Type = msoPicture Then
                sh.Delete
            End If
        Next sh
    End If
    
End Sub

Sub WriteAllCharts(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
'AggChart ChartImport

    Dim fName, source_name As String
    Dim nofwell, i As Integer
    
    nofwell = GetNumberOfWell()
    If ActiveSheet.name <> "AggChart" Then Sheets("AggChart").Select
    
    ' Call DeleteAllCharts
    
    
    If isSingleWellImport Then
        Call DeleteAllImages(singleWell)
    Else
        Call DeleteAllImages(999)
    End If
    
    
    source_name = ActiveWorkbook.name
    
    For i = 1 To nofwell
    
        ' isSingleWellImport = True ---> SingleWell Import
        ' isSingleWellImport = False ---> AllWell Import
        
        If isSingleWellImport Then
            If i = singleWell Then
                GoTo SINGLE_ITERATION
            Else
                GoTo NEXT_ITERATION
            End If
        End If
        
SINGLE_ITERATION:
        Call Write_InsertChart(i, source_name)
        
NEXT_ITERATION:
    Next i
    
End Sub

Sub Write_InsertChart(well As Integer, source_name As String)
    Dim fName As String
    Dim imagePath As String

    imagePath = Environ("TEMP") & "\tempChartImage.png"

    fName = "A" & CStr(well) & "_ge_OriginalSaveFile.xlsm"
    If Not IsWorkBookOpen(fName) Then
        MsgBox "Please open the yangsoo data ! " & fName
        Exit Sub
    End If

    Call SaveAndInsertChart(well, source_name, "Chart 5", "d" & CStr(3 + 16 * (well - 1)))
    Call SaveAndInsertChart(well, source_name, "Chart 7", "j" & CStr(3 + 16 * (well - 1)))
    Call SaveAndInsertChart(well, source_name, "Chart 9", "p" & CStr(3 + 16 * (well - 1)))
End Sub


Sub SaveAndInsertChart(well As Integer, source_name As String, chartName As String, targetRange As String)
    Dim imagePath As String
    Dim fName As String
    Dim targetCell As Range
    Dim picWidth As Double, picHeight As Double
    
    imagePath = Environ("TEMP") & "\tempChartImage.png"
    fName = "A" & CStr(well) & "_ge_OriginalSaveFile.xlsm"

    Windows(fName).Activate
    Worksheets("Input").ChartObjects(chartName).Activate
    ActiveChart.Export fileName:=imagePath, FilterName:="PNG"
    
    With ActiveChart.Parent
        picWidth = .Width
        picHeight = .height
    End With
    

    Windows(source_name).Activate
    Set targetCell = Sheets("AggChart").Range(targetRange)
        
    
    Sleep (1000)
    Sheets("AggChart").Shapes.AddPicture _
        fileName:=imagePath, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=targetCell.Left, _
        Top:=targetCell.Top, _
        Width:=picWidth, _
        height:=picHeight
        
End Sub


Sub ActivateChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject

    Set ws = ThisWorkbook.Worksheets("Input")
    Set chartObj = ws.ChartObjects("Chart 5")
    chartObj.Activate
End Sub


