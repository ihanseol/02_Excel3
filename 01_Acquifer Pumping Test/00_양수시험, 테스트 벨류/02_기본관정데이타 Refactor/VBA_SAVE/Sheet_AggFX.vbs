VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_AggFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton1_Click()
'Hide Aggregate

    Sheets("YangSoo").Visible = False
    Sheets("Well").Select
End Sub

Private Sub CommandButton2_Click()
'Collect Data
    
    Call GetBaseDataFromYangSoo(999, False)
End Sub

Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub



Private Sub CommandButton4_Click()
'single well import

Dim singleWell  As Integer
Dim WB_NAME As String


WB_NAME = GetOtherFileName
'MsgBox WB_NAME
    
'If Workbook Is Nothing Then
'    GetOtherFileName = "Empty"
'Else
'    GetOtherFileName = Workbook.name
'End If
    
If WB_NAME = "Empty" Then
    MsgBox "WorkBook is Empty"
    Exit Sub
Else
    singleWell = CInt(ExtractNumberFromString(WB_NAME))
'   MsgBox (SingleWell)
End If

Call GetBaseDataFromYangSoo(singleWell, True)

End Sub


Sub GetBaseDataFromYangSoo(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
' isSingleWellImport = True ---> SingleWell Import
' isSingleWellImport = False ---> AllWell Import
'
' SingleWell --> ImportWell Number
' 999 & False --> �������� ����Ʈ
'
    Dim fName As String
    Dim nofwell, i As Integer
    Dim rngString As String
        
    Dim natural() As Double ' �ڿ�����, natural depth
    Dim stable() As Double  ' ��������, stable depth
    Dim recover() As Double ' ȸ������, recover depth
    Dim Sw() As Double ' ����ȸ���� - �������� - ȸ������
    
    Dim delta_h() As Double ' deltah : �������Ϸ�
    
    Dim radius() As Double ' ���ݰ�
    Dim Rw() As Double      ' ���ݰ� / 2000
    
    Dim well_depth() As Double     ' �����ɵ�, well depth
    Dim casing() As Double  ' ���̽̽ɵ�
    
    Dim Q() As Double       '�����ȹ��
    Dim delta_s() As Double
    Dim hp() As Double
    
    Dim daeSoo() As Double  ' ����� �β�
    
    Dim T1() As Double      ' T1
    Dim T2() As Double      ' T2
    Dim TA() As Double      ' TA - (T1+T2)/2, TAverage
    
    Dim S1() As Double      ' S1
    Dim S2() As Double      ' S2 - ��Ų���� �ؼ�, s��
    
    Dim T0() As Double         ' �ܰ��������� T��
    Dim S0() As Double         ' S��, 0.005: �Ǿд����, 0.001: ���������, 0.1: ����������
    Dim ER_MODE() As String    ' ����ݰ�������� ����
    Dim ER1() As Double
    Dim ER2() As Double
    Dim ER3() As Double
    
    Dim K() As Double
    Dim time_() As Double   ' �����������޽ð�
    
    Dim shultze() As Double
    Dim webber() As Double
    Dim jacob() As Double
    
    Dim skin() As Double ' skin factor
    Dim er() As Double   ' effective radius, ��ȿ�칰�ݰ�
        
    
    Dim qh() As Double ' �Ѱ�����
    Dim qg() As Double ' ��ä����
    Dim q1() As Double ' Q1
    
    
    Dim sd1() As Double ' 1�ܰ� �������Ϸ�
    Dim sd2() As Double ' 4�ܰ� �������Ϸ�
    
    
    Dim C() As Double
    Dim B() As Double
    
    Dim ratio() As Double
    
    
 ' --------------------------------------------------------------------------------------

    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
    
 ' --------------------------------------------------------------------------------------
    
    ReDim natural(1 To nofwell)
    ReDim stable(1 To nofwell)
    ReDim recover(1 To nofwell)
    ReDim delta_h(1 To nofwell)
    ReDim Sw(1 To nofwell)
    
    
    ReDim radius(1 To nofwell)
    ReDim Rw(1 To nofwell)
    
    ReDim well_depth(1 To nofwell)
    ReDim casing(1 To nofwell)
    
    ReDim Q(1 To nofwell)
    ReDim delta_s(1 To nofwell)
    ReDim hp(1 To nofwell)
    
    ReDim daeSoo(1 To nofwell)
    
    ReDim T1(1 To nofwell)
    ReDim T2(1 To nofwell)
    ReDim TA(1 To nofwell)
    
    ReDim S1(1 To nofwell)
    ReDim S2(1 To nofwell)
    
    ReDim K(1 To nofwell)
    ReDim time_(1 To nofwell)
    
    ReDim shultze(1 To nofwell)
    ReDim webber(1 To nofwell)
    ReDim jacob(1 To nofwell)
    
    ReDim skin(1 To nofwell)
    ReDim er(1 To nofwell)
        
    ReDim ER1(1 To nofwell)
    ReDim ER2(1 To nofwell)
    ReDim ER3(1 To nofwell)
    
    ReDim qh(1 To nofwell)
    ReDim qg(1 To nofwell)
    
    
    ReDim sd1(1 To nofwell)
    ReDim sd2(1 To nofwell)
    ReDim q1(1 To nofwell)
    
    ReDim C(1 To nofwell)
    ReDim B(1 To nofwell)
    
    ReDim ratio(1 To nofwell)
    
    
    ReDim T0(1 To nofwell)         ' �ܰ��������� T��
    ReDim S0(1 To nofwell)         ' S��, 0.005: �Ǿд����, 0.001: ���������, 0.1: ����������
    ReDim ER_MODE(1 To nofwell)
    
    Dim wb As Workbook
    Dim wsInput As Worksheet
    Dim wsSkinFactor As Worksheet
    Dim wsSafeYield As Worksheet
    

    
    If isSingleWellImport Then
        rngString = "A" & (singleWell + 5 - 1) & ":AN" & (singleWell + 5 - 1)
        Call EraseCellData(rngString)
    Else
        rngString = "A5:AN37"
        Call EraseCellData(rngString)
    End If
        
    
    For i = 1 To nofwell
        
        If Not isSingleWellImport Or (isSingleWellImport And i = singleWell) Then
            GoTo SINGLE_ITERATION
        Else
            GoTo NEXT_ITERATION
        End If
        
SINGLE_ITERATION:
        
        fName = "A" & CStr(i) & "_ge_OriginalSaveFile.xlsm"
        If Not IsWorkBookOpen(fName) Then
            MsgBox "Please open the yangsoo data ! " & fName
            Exit Sub
        End If
        
        Set wb = Workbooks(fName)
        Set wsInput = wb.Worksheets("Input")
        Set wsSkinFactor = wb.Worksheets("SkinFactor")
        Set wsSafeYield = wb.Worksheets("SafeYield")
        
        
        Q(i) = wsInput.Range("m51").value
        hp(i) = wsInput.Range("i48").value
        
        natural(i) = wsInput.Range("m48").value
        stable(i) = wsInput.Range("m49").value
        radius(i) = wsInput.Range("m44").value
        Rw(i) = radius(i) / 2000
        
        well_depth(i) = wsInput.Range("m45").value
        casing(i) = wsInput.Range("i52").value
        
        
        C(i) = wsInput.Range("A31").value
        B(i) = wsInput.Range("B31").value
        
        
        
        recover(i) = wsSkinFactor.Range("c10").value
        Sw(i) = stable(i) - recover(i)
        
        delta_h(i) = wsSkinFactor.Range("b16").value
        delta_s(i) = wsSkinFactor.Range("b4").value
        
        daeSoo(i) = wsSkinFactor.Range("c16").value
        
        '----------------------------------------------------------------------------------
        
        T0(i) = wsSkinFactor.Range("d4").value
        S0(i) = wsSkinFactor.Range("f4").value
        ER_MODE(i) = wsSkinFactor.Range("h10").value
        
        T1(i) = wsSkinFactor.Range("d5").value
        T2(i) = wsSkinFactor.Range("h13").value
        TA(i) = (T1(i) + T2(i)) / 2
        
        S1(i) = wsSkinFactor.Range("e10").value
        S2(i) = wsSkinFactor.Range("i16").value
        
        K(i) = wsSkinFactor.Range("e16").value
        time_(i) = wsSkinFactor.Range("h16").value
        
        shultze(i) = wsSkinFactor.Range("c13").value
        webber(i) = wsSkinFactor.Range("c18").value
        jacob(i) = wsSkinFactor.Range("c23").value
        
        skin(i) = wsSkinFactor.Range("g6").value
        er(i) = wsSkinFactor.Range("c8").value
        
        
        ' �����, 1��, 2��, 3���� ��ȿ�칰�ݰ�
        ER1(i) = wsSkinFactor.Range("K8").value
        ER2(i) = wsSkinFactor.Range("K9").value
        ER3(i) = wsSkinFactor.Range("K10").value
        
        '----------------------------------------------------------------------------------
        
        qh(i) = wsSafeYield.Range("b13").value
        qg(i) = wsSafeYield.Range("b7").value
        
        sd1(i) = wsSafeYield.Range("b3").value
        sd2(i) = wsSafeYield.Range("b4").value
        q1(i) = wsSafeYield.Range("b2").value
        
        ratio(i) = wsSafeYield.Range("b11").value
        
        '*****************************************************************************************
        
        Cells(4 + i, "a").value = "W-" & CStr(i)
        Cells(4 + i, "b").value = natural(i)
        Cells(4 + i, "c").value = stable(i)
        
        Cells(4 + i, "d").value = recover(i)
        Cells(4 + i, "d").NumberFormat = "0.00"
        
        Cells(4 + i, "e").value = Sw(i)
        Cells(4 + i, "e").NumberFormat = "0.00"
        
        Cells(4 + i, "f").value = delta_h(i)
        Cells(4 + i, "f").NumberFormat = "0.00"
        
        Cells(4 + i, "g").value = radius(i)
        Cells(4 + i, "h").value = Rw(i)
        Cells(4 + i, "i").value = well_depth(i)
        Cells(4 + i, "j").value = casing(i)
        Cells(4 + i, "k").value = Q(i)
        
        Cells(4 + i, "l").value = delta_s(i)
        Cells(4 + i, "l").NumberFormat = "0.00"
        
        Cells(4 + i, "m").value = hp(i)
        Cells(4 + i, "n").value = daeSoo(i)
        
        Cells(4 + i, "o").value = T1(i)
        Cells(4 + i, "o").NumberFormat = "0.0000"
         
        Cells(4 + i, "p").value = T2(i)
        Cells(4 + i, "p").NumberFormat = "0.0000"
         
        Cells(4 + i, "q").value = TA(i)
        Cells(4 + i, "q").NumberFormat = "0.0000"
        
        Cells(4 + i, "r").value = S1(i)
        
        Cells(4 + i, "s").value = S2(i)
        Cells(4 + i, "s").NumberFormat = "0.0000000"
        
        Cells(4 + i, "t").value = K(i)
        Cells(4 + i, "t").NumberFormat = "0.0000"
        
        Cells(4 + i, "u").value = time_(i)
        
        Cells(4 + i, "v").value = shultze(i)
        Cells(4 + i, "v").NumberFormat = "0.0"
        
        Cells(4 + i, "w").value = webber(i)
        Cells(4 + i, "w").NumberFormat = "0.0"
        
        Cells(4 + i, "x").value = jacob(i)
        Cells(4 + i, "x").NumberFormat = "0.0"
        
        
        
        Cells(4 + i, "y").value = Format(skin(i), "0.0000")
        
        Cells(4 + i, "z").value = er(i)
        Cells(4 + i, "z").NumberFormat = "0.0000"
        
        Cells(4 + i, "aa").value = Format(qh(i), "0.")
        Cells(4 + i, "ab").value = Format(qg(i), "0.00")
        Cells(4 + i, "ac").value = Format(q1(i), "0.")
        
        Cells(4 + i, "ad").value = Format(sd1(i), "0.00")
        Cells(4 + i, "ae").value = Format(sd2(i), "0.00")
        
        Cells(4 + i, "af").value = C(i)
        Cells(4 + i, "ag").value = B(i)
        
        Cells(4 + i, "ah").value = ratio(i)
        Cells(4 + i, "ah").NumberFormat = "0.0%"
        
        
        ' 2023/09/22 ���� �߰��� ���� ...
        Cells(4 + i, "AI").value = Format(T0(i), "0.0000")
        Cells(4 + i, "AJ").value = Format(S0(i), "0.0000")
        Cells(4 + i, "AK").value = ER_MODE(i)
        
        Cells(4 + i, "AL").value = Format(ER1(i), "0.0000")
        Cells(4 + i, "AM").value = Format(ER2(i), "0.0000")
        Cells(4 + i, "AN").value = Format(ER3(i), "0.0000")
        
NEXT_ITERATION:

    Next i
End Sub


Public Sub MyDebug(sPrintStr As String, Optional bClear As Boolean = False)
   If bClear = True Then
      Application.SendKeys "^g^{END}", True

      DoEvents '  !!! DoEvents is VERY IMPORTANT here !!!

      Debug.Print String(30, vbCrLf)
   End If

   Debug.Print sPrintStr
End Sub


'0 : skin factor, cell, C8
'1 : Re1,         cell, E8
'2 : Re2,         cell, H8
'3 : Re3,         cell, G10

Function DetermineEffectiveRadius(ERMode As String) As Integer
    Dim er, r As String
    
    er = ERMode
    'MsgBox er
    r = Mid(er, 5, 1)
    
    If r = "F" Then
        DetermineEffectiveRadius = erRE0
    Else
        DetermineEffectiveRadius = val(r)
    End If
End Function



Private Sub CommandButton3_Click()
' Formula Button

Dim formula1, formula2 As String
Dim nofwell As String
Dim i As Integer
Dim T, Q, radius, skin, er As Double
Dim T0, S0 As Double
Dim ERMode As String
Dim ER1, ER2, ER3, B, S1 As Double

' Save array to a file
Dim filePath As String
Dim FileNum As Integer



nofwell = GetNumberOfWell()
Sheets("YangSoo").Select


filePath = ThisWorkbook.Path & "\" & ActiveSheet.name & ".csv"
FileNum = FreeFile

Open filePath For Output As FileNum

    
Call MyDebug("Formula SkinFactor ... ", True)

Debug.Print "************************************************************************************************************************************************************************************************"
Print #FileNum, "************************************************************************************************************************************************************************************************"

For i = 1 To nofwell

    T = Format(Cells(4 + i, "o").value, "0.0000")
    Q = Cells(4 + i, "k").value
    
    T0 = Format(Cells(4 + i, "AI").value, "0.0000")
    S0 = Format(Cells(4 + i, "AJ").value, "0.0000")
    S1 = Cells(4 + i, "R").value
            
    delta_s = Format(Cells(4 + i, "l").value, "0.00")
    radius = Format(Cells(4 + i, "h").value, "0.000")
    skin = Cells(4 + i, "y").value
    er = Cells(4 + i, "z").value
    
    
    B = Format(Cells(4 + i, "AG").value, "0.0000")
    ER1 = Cells(4 + i, "AL").value
    ER2 = Cells(4 + i, "AM").value
    ER3 = Cells(4 + i, "AN").value
    
    
    Select Case DetermineEffectiveRadius(Cells(4 + i, "AK").value)
    ' ����� 1��
    Case erRE1
        formula1 = "W-" & i & "ȣ��~~r _{e-" & i & "} `=~ sqrt {{2.25 TIMES  " & T0 & " TIMES  0.0833333} over {" & S0 & " TIMES  10 ^{5.46 TIMES  " & T & " TIMES  " & B & "}}} `=~" & ER1 & "m"
        formula2 = "erRE1, ����� 1��"
        
    ' ����� 2��
    Case erRE2
        formula1 = "W-" & i & "ȣ��~~r _{e-" & i & "} `=~ sqrt {{2.25 TIMES  " & T0 & " TIMES  0.0833333} over {" & S0 & " TIMES  10 ^{4 pi TIMES " & T & " TIMES  " & B & "}}} `=~" & ER2 & "m"
        formula2 = "erRE2, ����� 2��"
    ' ����� 3��
    Case erRE3
        formula1 = "W-" & i & "ȣ��~~r _{e-" & i & "} `=~" & radius & " TIMES  sqrt {{" & S1 & "} over {" & S0 & "}} `=~" & ER3 & "m"
        formula2 = "erRE3, ����� 3��"
        
    Case Else
        ' ��Ų���
        formula1 = "W-" & i & "ȣ��~~ sigma  _{w-" & i & "} = {2 pi  TIMES  " & T & " TIMES  " & delta_s & " } over {" & Q & "} -1.15 TIMES  log {2.25 TIMES  " & T & " TIMES  (1/1440)} over {0.0005 TIMES  (" & radius & " TIMES  " & radius & ")} =`" & skin
        ' ��ȿ�칰�ݰ�
        formula2 = "W-" & i & "ȣ��~~r _{e-" & i & "} `=~r _{w} e ^{- sigma  _{w-" & i & "}} =" & radius & " TIMES e ^{-(" & skin & ")} =" & er & "m"
    End Select
    
    Debug.Print formula1
    Debug.Print formula2
    Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
    Print #FileNum, formula1
    Print #FileNum, formula2
    Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    

Next i

Debug.Print "************************************************************************************************************************************************************************************************"
Print #FileNum, "************************************************************************************************************************************************************************************************"


Call FormulaChwiSoo(FileNum)
' 3-7, ���������

Call FormulaRadiusOfInfluence(FileNum)
' �������ݰ�
End Sub


Sub FormulaChwiSoo(FileNum As Integer)
' 3-7, ���������

    Dim formula As String
    Dim nofwell As String
    Dim i As Integer
    Dim q1, S1, S2, res As Double
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
        
        
    For i = 1 To 3
        Debug.Print " "
        Print #FileNum, " "
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    For i = 1 To nofwell
        q1 = Cells(4 + i, "ac").value
 
        S1 = Format(Cells(4 + i, "ad").value, "0.00")
        S2 = Format(Cells(4 + i, "ae").value, "0.00")
        
        res = Cells(4 + i, "ab").value
        formula = "W-" & i & "ȣ��~~Q _{ & 2} `��" & q1 & "` TIMES  `(` {" & S2 & "} over {" & S1 & "} `) ^{2/3} `��" & res & "��/��"
        
        Debug.Print formula
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
        Print #FileNum, formula
        Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    
End Sub


Sub FormulaRadiusOfInfluence(FileNum As Integer)
' �������ݰ�

    Dim formula1, formula2, formula3 As String
    ' ����, ����, ������
    
    Dim nofwell As String
    Dim i As Integer
    Dim shultze, webber, jacob, T, K, S, time_, delta_h As String
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
        
        
    For i = 1 To 3
        Debug.Print " "
        Print #FileNum, " "
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    
    For i = 1 To nofwell
        shultze = CStr(Cells(4 + i, "v").value)
        webber = CStr(Cells(4 + i, "w").value)
        jacob = CStr(Cells(4 + i, "x").value)
        
        T = CStr(Format(Cells(4 + i, "q").value, "0.0000"))
        S = CStr(Format(Cells(4 + i, "s").value, "0.0000000"))
        K = CStr(Format(Cells(4 + i, "t").value, "0.0000"))
    
        delta_h = CStr(Cells(4 + i, "f").value)
        time_ = CStr(Format(Cells(4 + i, "u").value, "0.0000"))
        
        
        ' Cells(4 + i, "y").value = Format(skin(i), "0.0000")
        
        formula1 = "W-" & i & "ȣ��~~R _{W-" & i & "} ``=`` sqrt {6 TIMES  " & delta_h & " TIMES  " & K & " TIMES  " & time_ & "/" & S & "} ``=~" & shultze & "m"
        formula2 = "W-" & i & "ȣ��~~R _{W-" & i & "} ``=3`` sqrt {" & delta_h & " TIMES " & K & " TIMES " & time_ & "/" & S & "} `=`" & webber & "`m"
        formula3 = "W-" & i & "ȣ��~~r _{0(W-" & i & ")} `=~ sqrt {{2.25 TIMES  " & T & " TIMES  " & time_ & "} over {" & S & "}} `=~" & jacob & "m"
        
        Debug.Print ""
        Debug.Print formula1
        Debug.Print formula2
        Debug.Print formula3
        Debug.Print ""
        
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
        
         Print #FileNum, " "
         Print #FileNum, formula1
         Print #FileNum, formula2
         Print #FileNum, formula3
         Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
End Sub






