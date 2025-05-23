Attribute VB_Name = "modWell"

Function CheckWorkbookNameWithRegex(ByVal WB_NAME As String) As Boolean
    Dim regex As Object
    Dim pattern As String
    Dim match As Object

    ' Create the regex object
    Set regex = CreateObject("VBScript.RegExp")

    ' Define the pattern
    ' \bA(1|[2-9]|1[0-9]|2[0-9]|30)_ge_OriginalSaveFile
    pattern = "\bA(1|[2-9]|1[0-9]|2[0-9]|30)_ge_OriginalSaveFile.xlsm"

    ' Configure the regex object
    With regex
        .pattern = pattern
        .IgnoreCase = True
        .Global = False
    End With

    ' Check for the pattern
    If regex.test(WB_NAME) Then
        Set match = regex.Execute(WB_NAME)
        Debug.Print "The workbook name contains the pattern: " & match(0).value
        CheckWorkbookNameWithRegex = True
    Else
        Debug.Print "The workbook name does not contain the pattern."
        CheckWorkbookNameWithRegex = False
    End If
End Function

Function IsOpenedYangSooFiles() As Boolean
'
' 양수일보파일, A1_ge_OriginalSaveFile 이 열려있어서
' 양수일보의 갯수가, 관정의 갯수와 같으면 True
' 그렇지 않으면 False
'
    Dim fileName, WBNAME As String
    Dim nof_yangsoo As Integer
    Dim nofwell As Integer
    
    nof_yangsoo = 0
    nofwell = sheets_count()
    
    For Each Workbook In Application.Workbooks
        WBNAME = Workbook.name
        If StrComp(ThisWorkbook.name, WBNAME, vbTextCompare) = 0 Then
        ' 이름이 thisworkbook.name 과 같다면 , 다음분기로
            GoTo NEXT_ITERATION
        End If
        
        If CheckWorkbookNameWithRegex(WBNAME) Then
            nof_yangsoo = nof_yangsoo + 1
        End If
        
NEXT_ITERATION:
    Next
    
    If nof_yangsoo = nofwell Then
        IsOpenedYangSooFiles = True
    Else
        IsOpenedYangSooFiles = False
    End If

End Function


Sub PressAll_Button()
' Push All Button
' Fx - Collect Data
' Fx - Formula
' ImportAll, Collect Each Well
' Agg2
' Agg1
' AggStep
' AggChart
' AggWhpa
'
    If Not IsOpenedYangSooFiles() Then
        Popup_MessageBox ("YangSoo File is Does not match with number of well")
        Exit Sub
    End If

    Call Popup_MessageBox("YangSoo, modAggFX - get Data from YangSoo ilbo ...")
        
    Sheets("YangSoo").Visible = True
    Sheets("YangSoo").Select
    Call GetBaseDataFromYangSoo(999, False)
    Sheets("YangSoo").Visible = False
    
    Call Popup_MessageBox("YangSoo, Aggregate2 - ImportWellSpec ...")
    

    Sheets("Aggregate2").Visible = True
    Sheets("Aggregate2").Select
    Call modAgg2.GROK_ImportWellSpec(999, False)
    Sheets("Aggregate2").Visible = False
    
    Call Popup_MessageBox("YangSoo, Aggregate1 - AggregateOne_Import ...")
    

    Sheets("Aggregate1").Visible = True
    Sheets("Aggregate1").Select
    Call modAgg1.ImportAggregateData(999, False)
    Sheets("Aggregate1").Visible = False
    
    Call Popup_MessageBox("YangSoo, AggStep - Import StepTest Data ...")
     
    Sheets("AggStep").Visible = True
    Sheets("AggStep").Select
    Call modAggStep.WriteStepTestData(999, False)
    Sheets("AggStep").Visible = False
    
    
   ' <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><>
   ' Reduce OverHead,
    
     If Sheets("Well").CheckBox_GetChart.value Then
         Call Popup_MessageBox("YangSoo, AggChart - Chart Import...")
        
         Sheets("AggChart").Visible = True
         Sheets("AggChart").Select
         Call modAggChart.WriteAllCharts(999, False)
         Sheets("AggChart").Visible = False
    End If
        
    ' <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><>
    
    Call Popup_MessageBox("Import All QT ...")
    Call modWell.ImportAll_QT
    
    Call Popup_MessageBox("ImportAll Each Well Spec ...")
    Call modWell.ImportAll_EachWellSpec
    
    Call Popup_MessageBox("ImportWell MainWellPage ...")
    Call modWell.ImportWell_MainWellPage("_ALL_")
    
    Call Popup_MessageBox("Push Drastic Index ...")
    Call modWell.PushDrasticIndex
    
    

End Sub


Sub ImportSingleWell_Main(ByVal WellNumber As Integer)
'
' Just in case , Single Well Import
' Action --> FX Sheet Press SingleWellImport, ImportMainWell Page Setting
' 2024/4/7

    ' Sub GetBaseDataFromYangSoo(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
    Call Popup_MessageBox(" GetBaseDataFromYangSoo W-" & WellNumber)
    Sheets("YangSoo").Activate
    Call modAggFX_A.GetBaseDataFromYangSoo(WellNumber, True)
    
    
    Call Popup_MessageBox(" ImportWell_MainWellPage W-" & WellNumber)
    Sheets("Well").Activate
    Call ImportWell_MainWellPage("_SINGLE_", WellNumber)
        
    ' 2024/4/7 - is Import Chart YES
    
    If Sheets("Well").CheckBox_GetChart.value Then
        Call Popup_MessageBox(" Import Charts W-" & WellNumber)
        Sheets("AggChart").Activate
        Call modAggChart.WriteAllCharts(WellNumber, True)
        
        Sheets("AggChart").Visible = False
        Sheets("Well").Select
    End If
    
    Call Popup_MessageBox("ImportAll Each Well Spec ...")
    Call modWell.ImportAll_EachWellSpec
    
    Call Popup_MessageBox("Push Drastic Index ...")
    Call modWell.PushDrasticIndex
    
End Sub



Function RemoveSheetIfExists(shname As String) As Boolean
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(shname)
    If Not ws Is Nothing Then sheetExists = True
    On Error GoTo 0

    If sheetExists Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        RemoveSheetIfExists = True
        Exit Function
    Else
        RemoveSheetIfExists = False
        Exit Function
    End If
End Function


Public Sub AddWell_CopyOneSheet()
    Dim n_sheets    As Integer
    
    n_sheets = sheets_count()
    
    '2020/5/30 관정리스트의 목록삽입해주는 부분 추가
    InsertOneRow (n_sheets)
    
    If (n_sheets = 1) Then
        Sheets("1").Select
        Sheets("1").Copy Before:=Sheets("Q1")
        Call DeleteCommandButton
    Else
        Sheets("2").Select
        Sheets("2").Copy Before:=Sheets("Q1")
    End If
    
    ActiveSheet.name = CStr(n_sheets + 1)
    Range("b2").value = "W-" & (n_sheets + 1)
    Range("e15").value = CStr(n_sheets + 1)
    
    '2022/6/9 일
    Range("i2") = "A" & CStr(n_sheets + 1) & "_ge_OriginalSaveFile.xlsm"
    
    If n_sheets = 1 Then
        Call ChangeCellData(n_sheets + 1, 1)
    Else
        Call ChangeCellData(n_sheets + 1, 2)
    End If
    
    Sheets("Well").Select
End Sub



' --------------------------------------------------------------------------------------------------------------


Sub DeleteCommandButton()
    ActiveSheet.Shapes.Range(Array("CommandButton2")).Delete
End Sub



Sub InsertOneRow(ByVal n_sheets As Integer)
    n_sheets = n_sheets + 4
    Rows(n_sheets & ":" & n_sheets).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Rows(CStr(n_sheets - 1) & ":" & CStr(n_sheets - 1)).Select
    Selection.Copy
    Rows(CStr(n_sheets) & ":" & CStr(n_sheets)).Select
    ActiveSheet.Paste
    
    Application.CutCopyMode = False
End Sub

Sub ChangeCellData(ByVal nsheet As Integer, ByVal nselect As Integer)
    ' change sheet data direct to well sheet data value
    ' https://stackoverflow.com/questions/18744537/vba-setting-the-formula-for-a-cell
    
    Range("C2, C3, C4, C5, C6, C7, C8, C15, C16, C17, C18, C19, E17, F21").Select
    
    nsheet = nsheet + 3
    Selection.Replace What:=CStr(nselect + 3), Replacement:=CStr(nsheet), LookAt:=xlPart, _
                      SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                      ReplaceFormat:=False
    
    
    
    ' minhwasoo, 2023-10-13
    ' block, this code ....
    ' Range("E21").Select
    ' Range("E21").formula = "=Well!" & Cells(nsheet, "I").Address
End Sub



' --------------------------------------------------------------------------------------------------------------


Sub JojungButton()
    Dim nofwell As Integer

    TurnOffStuff

    nofwell = sheets_count()
    Call JojungSheetData
    Call make_wellstyle
    Call DecorateWellBorder(nofwell)
    
    Worksheets("1").Range("E21") = "=Well!" & Cells(5 + GetNumberOfWell(), "I").Address
    
    TurnOnStuff
End Sub

Sub Make_OneButton()
    Dim i, nofwell As Integer
    Dim response As VbMsgBoxResult
    
    nofwell = GetNumberOfWell()
    
    If nofwell = 1 Then Exit Sub
    
    response = MsgBox("Do you deletel all water well?", vbYesNo)
    If response = vbYes Then
         For i = 2 To nofwell
             RemoveSheetIfExists (CStr(i))
        Next i
        
        Sheets("Well").Activate
        Rows("5:" & CStr(nofwell + 3)).Select
        Selection.Delete Shift:=xlUp
        
        For i = 1 To 12
            If Not RemoveSheetIfExists("p" & CStr(i)) Then Exit For
        Next i
        
        Call DecorateWellBorder(1)
        Range("A1").Select
    End If
End Sub


Sub DeleteLast()
' delete last

    Dim nofwell As Integer
    'nofwell = GetNumberOfWell()
    nofwell = sheets_count()
    
    If nofwell = 1 Then
        MsgBox "Last is not delete ... ", vbOK
        Exit Sub
    End If
    
    Rows(nofwell + 3).Delete
    Call DeleteWorksheet(nofwell)
    Call DecorateWellBorder(nofwell - 1)
End Sub



Sub DeleteWorksheet(well As Integer)
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(CStr(well)).Delete
    Application.DisplayAlerts = True
End Sub


Sub DecorateWellBorder(ByVal nofwell As Integer)
    Sheets("Well").Activate
    Range("A2:R" & CStr(nofwell + 3)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    Range("D15").Select
End Sub




Sub getDuoSolo(ByVal nofwell As Integer, ByRef nDuo As Integer, ByRef nSolo As Integer)
    Dim page, quotient, remainder As Integer
    
    quotient = WorksheetFunction.quotient(nofwell, 2)
    remainder = nofwell Mod 2
    
    If remainder = 0 Then
        nDuo = quotient
        nSolo = 0
    Else
        nDuo = quotient
        nSolo = 1
    End If

End Sub


Sub ImportAll_EachWellSpec()
'
' 각관정을 순회하면서, 관정데이타를 각관정에 써준다.
'
    Dim nofwell, i  As Integer
    ' Dim obj As New Class_Boolean

    nofwell = sheets_count()
    
    BaseData_ETC_02.TurnOffStuff
    
    For i = 1 To nofwell
        Sheets(CStr(i)).Activate
        Call modWell_Each.ImportWellSpecFX(i)
    Next i
    
    Sheets("Well").Activate
    
    BaseData_ETC_02.TurnOnStuff
End Sub


Sub ImportAll_EachWellSpec_OLD()
'
' 각관정을 순회하면서, 관정데이타를 각관정에 써준다.
'
    Dim nofwell, i  As Integer
    Dim obj As New Class_Boolean

    nofwell = sheets_count()
    
    BaseData_ETC_02.TurnOffStuff
    
    For i = 1 To nofwell
        Sheets(CStr(i)).Activate
        Call Module_ImportWellSpec.ImportWellSpec_OLD(i, obj)
        If obj.result Then Exit For
    Next i
    
    Sheets("Well").Activate
    
    BaseData_ETC_02.TurnOnStuff
End Sub


Function ReplaceMultipleSpacesFunction(ByVal val_str As String) As String
  Dim cleanedString As String
  Dim regex As Object
  
  Set regex = CreateObject("VBScript.RegExp")
  regex.pattern = "\s+"
  regex.Global = True

  cleanedString = regex.Replace(val_str, " ")
  ReplaceMultipleSpacesFunction = cleanedString
End Function



Function AddressReducer(ByVal val_address As String) As String
' Reduce Address
' val --> Address String
'

    Dim Address, FinalAddress As String
    Dim i As Integer
    Dim AddressArray As Variant
    
        
    Address = Trim(val_address)
        
    If InStr(1, val_address, "번지") > 0 And InStr(1, val_address, "호") > 0 Then
          Address = ReplaceMultipleSpacesFunction(Address)
          Address = Replace(Address, "번지 ", "-")
          Address = Left(Address, Len(Address) - 1)  ' remove 호
    Else
        Address = Replace(val_address, "번지", "")
    End If
        
    Address = Replace(Address, "특별자치", "")
    Address = Replace(Address, "광역", "")
    Address = Replace(Address, "충청남도", "")
    Address = Replace(Address, "충남", "")
    AddressArray = Split(Address, " ")
    
    For i = 0 To UBound(AddressArray)
        
        If Right(AddressArray(i), 1) = "도" Then
            GoTo NextIteration
        Else
            FinalAddress = FinalAddress & " " & AddressArray(i)
        End If
NextIteration:
    Next i
    
    AddressReducer = FinalAddress
    
End Function



Sub ImportWell_MainWellPage(Optional ByVal mode As String = "_ALL_", Optional ByVal WellNumber As Integer)
'
' mode = _SINGLE_
' mode = _ALL_
'
' import Sheets("Well") Page
'
    Dim fName As String
    Dim nofwell, i, j, start_index   As Integer
    
    Dim Address, AddressValue, Company, FinalAddress As String
    Dim address_array As Variant
    Dim simdo, diameter, Q, Hp As Double
    ' 2025/4/9 - 펌프심도, 토출관, 양수능력 추가
    Dim pump_simdo, tochul, pump_capa As Double
    
    nofwell = sheets_count()
    Sheets("Well").Select
    
    Dim wsYangSoo, wsWell, wsRecharge As Worksheet
    Set wsYangSoo = Worksheets("YangSoo")
    Set wsWell = Worksheets("Well")
    Set wsRecharge = Worksheets("Recharge")
    
    '2024,12,25 - Add Title
    wsWell.Range("D1").value = wsYangSoo.Cells(5, "AR").value
    
    Call TurnOffStuff
               
    If mode = "_ALL_" Then
        For i = 1 To nofwell
            '2025/3/5
            AddressValue = wsYangSoo.Cells(4 + i, "ao").value
            FinalAddress = AddressReducer(AddressValue)
            
            simdo = wsYangSoo.Cells(4 + i, "i").value
            diameter = wsYangSoo.Cells(4 + i, "g").value
            Q = wsYangSoo.Cells(4 + i, "k").value
            Hp = wsYangSoo.Cells(4 + i, "m").value
            
            ' 2025/4/9 - 펌프심도, 토출관, 양수능력 추가
            pump_simdo = wsYangSoo.Cells(4 + i, "as").value
            tochul = wsYangSoo.Cells(4 + i, "at").value
            pump_capa = wsYangSoo.Cells(4 + i, "au").value
            
            
            wsWell.Cells(3 + i, "d").value = FinalAddress
            wsWell.Cells(3 + i, "g").value = diameter
            wsWell.Cells(3 + i, "h").value = simdo
            wsWell.Cells(3 + i, "i").value = Q
            wsWell.Cells(3 + i, "j").value = Q
            wsWell.Cells(3 + i, "l").value = Hp
            
            ' 2025/4/9 - 펌프심도, 토출관, 양수능력 추가
            wsWell.Cells(3 + i, "m").value = pump_simdo
            wsWell.Cells(3 + i, "n").value = tochul
            wsWell.Cells(3 + i, "k").value = pump_capa
            
        Next i
        
        Company = wsYangSoo.Range("AP5").value
    Else
        ' SingleWell Import, MainWellPage
        AddressValue = wsYangSoo.Cells(4 + WellNumber, "ao").value
        FinalAddress = AddressReducer(AddressValue)
        
        simdo = wsYangSoo.Cells(4 + WellNumber, "i").value
        diameter = wsYangSoo.Cells(4 + WellNumber, "g").value
        Q = wsYangSoo.Cells(4 + WellNumber, "k").value
        Hp = wsYangSoo.Cells(4 + WellNumber, "m").value
        
        ' 2025/4/9 - 펌프심도, 토출관, 양수능력 추가
        pump_simdo = wsYangSoo.Cells(4 + WellNumber, "as").value
        tochul = wsYangSoo.Cells(4 + WellNumber, "at").value
        pump_capa = wsYangSoo.Cells(4 + WellNumber, "au").value
            
        
        wsWell.Cells(3 + WellNumber, "d").value = FinalAddress
        wsWell.Cells(3 + WellNumber, "g").value = diameter
        wsWell.Cells(3 + WellNumber, "h").value = simdo
        wsWell.Cells(3 + WellNumber, "i").value = Q
        wsWell.Cells(3 + WellNumber, "j").value = Q
        wsWell.Cells(3 + WellNumber, "l").value = Hp
        
        ' 2025/4/9 - 펌프심도, 토출관, 양수능력 추가
        wsWell.Cells(3 + WellNumber, "m").value = pump_simdo
        wsWell.Cells(3 + WellNumber, "n").value = tochul
        wsWell.Cells(3 + WellNumber, "k").value = pump_capa
    
        Company = wsYangSoo.Range("AP" & (4 + WellNumber)).value
    End If

    wsRecharge.Range("B32").value = Company
    
    Application.CutCopyMode = False
    Call TurnOnStuff
End Sub




Sub DuplicateBasicWellData()
' 2024/6/24 - dupl, duplicate basic well data ...
' 기본관정데이타 복사하는것
' 관정을 순회하면서, 거기에서 데이터를 가지고 오는데 …
' 와파 , 장축부, 단축부
' 유향, 거리, 관정높이, 지표수표고 이렇게 가지고 오면 될듯하다.

' k6 - 장축부 / long axis
' k7 - 단축부 / short axis
' k12 - degree of flow
' k13 - well distance
' k14 - well height
' k15 - surfacewater height

    Dim nofwell, i  As Integer
    Dim obj As New Class_Boolean
    Dim WB_NAME As String
    Dim weather_station, river_section As String
    

    nofwell = sheets_count()
     
    WB_NAME = mod_DuplicatetWellSpec.GetOtherFileName
    
    If WB_NAME = "NOTHING" Then
        MsgBox "기본관정데이타를 복사해야 하므로, 기본관정데이터를 열어두시기 바랍니다. ", vbOK
        Exit Sub
    Else
        BaseData_ETC_02.TurnOffStuff
        
        Call mod_DuplicatetWellSpec.Duplicate_WATER(ThisWorkbook.name, WB_NAME)
        Call mod_DuplicatetWellSpec.Duplicate_WELL_MAIN(ThisWorkbook.name, WB_NAME, nofwell)
        weather_station = Replace(Sheets("Well").Range("F4").value, "기상청", "")
        river_section = Sheets("Well").Range("E4").value
        
        ' 2024/6/27 일, 새로 추가해준 방법으로 복사해줌 ...
        ThisWorkbook.Sheets("Recharge").Range("b32") = Range("B4").value
        
        ' 각 관정별 데이터 복사
        For i = 1 To nofwell
            Sheets(CStr(i)).Activate
            Call mod_DuplicatetWellSpec.DuplicateWellSpec(ThisWorkbook.name, WB_NAME, i, obj)
            
            If obj.result Then Exit For
        Next i
        
        Worksheets("Well").Activate
        
        'WSet Button, CommandButton14
        For i = 1 To nofwell
            Cells(i + 3, "E").formula = "=Recharge!$I$24"
            Cells(i + 3, "F").formula = "=All!$B$2"
            Cells(i + 3, "O").formula = "=ROUND(water!$F$7, 1)"
            
            Cells(i + 3, "B").formula = "=Recharge!$B$32"
        Next i
        
        Sheets("Well").Activate
        BaseData_ETC_02.TurnOnStuff
    End If
    
     ' 대권역, 중권역 세팅
     Sheets("Recharge").Range("I24") = river_section
     
     ' 2024/7/9 Add, Company Name Setting
     Sheets("Recharge").Range("B32") = Sheets("YangSoo").Range("AP5")
     
     
    ' 기상청 데이타, 다시 불러오기
    If Not BaseData_ETC.CheckSubstring(Sheets("All").Range("T5").value, weather_station) Then
         Call modProvince.ResetWeatherData(weather_station)
    End If
    
    Call modWell.PushDrasticIndex

End Sub


Sub ImportAll_QT()
'
' 양수정의 수질변화기록
'
    Dim i, nof_p As Integer
    Dim qt As String
    
    nof_p = GetNumberOf_P
    
    For i = 1 To nof_p
        Sheets("p" & i).Activate
        qt = determin_Q_Type
        
        Application.Run "modWaterQualityTest.GetWaterSpecFromYangSoo_" & qt
        If qt = "Q1" Then
            Call modWaterQualityTest.QualityTest_Delete2880(i)
        End If
        
    Next i
End Sub


Function determin_Q_Type() As String
' 이것은, p1, p2, p3 가 어떤 타입인지 체크하는부분
' 즉 Q1, Q2, Q3 인지 알아내는것
' D12 --- q1
' G12 --- q2
' J12 --- q3

    If Range("J12").value <> "" Then
        determin_Q_Type = "Q3"
    ElseIf Range("G12").value <> "" Then
        determin_Q_Type = "Q2"
    Else
        determin_Q_Type = "Q1"
    End If

End Function

Function GetNumberOf_P()
    Dim nofwell, i, nof_p As Integer

    nofwell = sheets_count()
    nof_p = 0
    
    For Each sheet In Worksheets
        If Left(sheet.name, 1) = "p" And ConvertToLongInteger(Right(sheet.name, 1)) <> 0 Then
            nof_p = nof_p + 1
        End If
    Next

    GetNumberOf_P = nof_p
End Function


Sub PushDrasticIndex()

    Call BaseData_DrasticIndex.main_drasticindex
    Call BaseData_DrasticIndex.print_drastic_string
    
End Sub
