

'<><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><>
' 2024,12,25 Add import "Title "
'
Sub FXSAVE_GetBaseDataFromYangSoo(ByVal WellNumber As Integer, ByVal isSingleWellImport As Boolean)
    Dim nofwell As Integer
    Dim i As Integer
    Dim rngString As String

    ' Arrays to store data
    Dim dataArrays As Variant
    dataArrays = Array("natural", "stable", "recover", "delta_h", "Sw", "radius", _
                       "Rw", "well_depth", "casing", "Q", "delta_s", "hp", _
                       "daeSoo", "T1", "T2", "TA", "S1", "S2", "K", "time_", _
                       "shultze", "webber", "jacob", "skin", "er", "ER1", _
                       "ER2", "ER3", "qh", "qg", "sd1", "sd2", "q1", "C", _
                       "B", "ratio", "T0", "S0", "ER_MODE", "Address", "Company", "S3", "Title")

    ' Check if all well data should be imported
    nofwell = GetNumberOfWell()
    If Not isSingleWellImport And WellNumber = 999 Then
        rngString = "A5:AR37"
    Else
       rngString = "A" & (WellNumber + 5 - 1) & ":AR" & (WellNumber + 5 - 1)
    End If

    Call EraseCellData(rngString)

    ' Loop through each well
    For i = 1 To nofwell
        ' Import data for all wells or only for the specified single well
        If Not isSingleWellImport Or (isSingleWellImport And i = WellNumber) Then
            ImportDataForWell i, dataArrays
        End If
    Next i
End Sub

Sub FXSAVE_ImportDataForWell(ByVal wellIndex As Integer, ByVal dataArrays As Variant)
    Dim fName As String
    Dim wb As Workbook
    Dim wsInput As Worksheet
    Dim wsSkinFactor As Worksheet
    Dim wsSafeYield As Worksheet
    Dim dataIdx As Integer
    Dim cellOffset As Integer
    Dim dataCell As Range

    ' Open the workbook
    fName = "A" & CStr(wellIndex) & "_ge_OriginalSaveFile.xlsm"
    If Not IsWorkBookOpen(fName) Then
        MsgBox "Please open the yangsoo data! " & fName
        Exit Sub
    End If
    Set wb = Workbooks(fName)

    ' Loop through data arrays and import values
    For dataIdx = LBound(dataArrays) To UBound(dataArrays)
        SetDataArrayValues wb, wellIndex, dataArrays(dataIdx)
    Next dataIdx

    ' Close workbook
    ' wb.Close SaveChanges:=False
End Sub


Sub FXSAVE_SetDataArrayValues(ByVal wb As Workbook, ByVal wellIndex As Integer, ByVal dataArrayName As String)
    Dim wsInput As Worksheet
    Dim wsSkinFactor As Worksheet
    Dim wsSafeYield As Worksheet
    Dim dataCell As Range


    Dim dataRanges() As Variant
    Dim addresses() As Variant
    Dim i As Integer

    ' Set references to worksheets
    Set wsInput = wb.Worksheets("Input")
    Set wsSkinFactor = wb.Worksheets("SkinFactor")
    Set wsSafeYield = wb.Worksheets("SafeYield")

    ' Define data ranges for each dataArrayName
    dataRanges = Array(wsInput.Range("m51"), wsInput.Range("i48"), _
                        wsInput.Range("m48"), wsInput.Range("m49"), _
                        wsInput.Range("m44"), wsSkinFactor.Range("e4"), _
                        wsInput.Range("m45"), wsInput.Range("i52"), _
                        wsInput.Range("A31"), wsInput.Range("B31"), _
                        wsSkinFactor.Range("c10"), wsSkinFactor.Range("c11"), _
                        wsSkinFactor.Range("b16"), wsSkinFactor.Range("b4"), _
                        wsSkinFactor.Range("c16"), wsSkinFactor.Range("d4"), _
                        wsSkinFactor.Range("f4"), wsSkinFactor.Range("h10"), _
                        wsSkinFactor.Range("d5"), wsSkinFactor.Range("h13"), _
                        wsSkinFactor.Range("d16"), wsSkinFactor.Range("e10"), _
                        wsSkinFactor.Range("i16"), wsSkinFactor.Range("e16"), _
                        wsSkinFactor.Range("h16"), wsSkinFactor.Range("c13"), _
                        wsSkinFactor.Range("c18"), wsSkinFactor.Range("c23"), _
                        wsSkinFactor.Range("g6"), wsSkinFactor.Range("c8"), _
                        wsSkinFactor.Range("k8"), wsSkinFactor.Range("k9"), _
                        wsSkinFactor.Range("k10"), wsSafeYield.Range("b13"), _
                        wsSafeYield.Range("b7"), wsSafeYield.Range("b3"), _
                        wsSafeYield.Range("b4"), wsSafeYield.Range("b2"), _
                        wsSafeYield.Range("b11"), wsInput.Range("i46"), _
                        wsInput.Range("i47"), wsSkinFactor.Range("i13"), _
                        wsInput.Range("i44"))

    ' Array of data addresses
    addresses = Array("Q", "hp", "natural", "stable", "radius", "Rw", _
                        "well_depth", "casing", "C", "B", "recover", "Sw", _
                        "delta_h", "delta_s", "daeSoo", "T0", "S0", "ER_MODE", _
                        "T1", "T2", "TA", "S1", "S2", "K", "time_", "shultze", _
                        "webber", "jacob", "skin", "er", "ER1", "ER2", "ER3", _
                        "qh", "qg", "sd1", "sd2", "q1", "ratio", "Address", _
                        "Company", "S3", "Title")

    ' Find index of dataArrayName in addresses array
    For i = LBound(addresses) To UBound(addresses)
        If addresses(i) = dataArrayName Then
            Set dataCell = dataRanges(i)
            Exit For
        End If
    Next i

    ' Check if dataArrayName is found
    If Not dataCell Is Nothing Then
        SetCellValueForWell wellIndex, dataCell, dataArrayName
    Else
        MsgBox "Data array name not found: " & dataArrayName
    End If
End Sub


' 2024-8-22 : �����������޽ð�, time_ : 0.0000 �� ����

Sub FXSAVE_SetCellValueForWell(ByVal wellIndex As Integer, ByVal dataCell As Range, ByVal dataArrayName As String)
    Dim WellData As Variant
    Dim numberFormats As Object
    Set numberFormats = CreateObject("Scripting.Dictionary")

    ' Define number formats for each dataArrayName
    With numberFormats
        .Add "recover", "0.00"
        .Add "Sw", "0.00"
        .Add "S2", "0.0000000"
        .Add "S3", "0.00"
        .Add "T1", "0.0000"
        .Add "T2", "0.0000"
        .Add "TA", "0.0000"
        .Add "qh", "0."
        .Add "qg", "0.00"
        .Add "q1", "0.00"
        .Add "sd1", "0.00"
        .Add "sd2", "0.00"
        .Add "skin", "0.0000"
        .Add "er", "0.0000"
        .Add "ratio", "0.0%"
        .Add "T0", "0.0000"
        .Add "S0", "0.0000"
        .Add "delta_s", "0.00"
        .Add "time_", "0.0000"
        .Add "shultze", "0.00"
        .Add "webber", "0.00"
        .Add "jacob", "0.00"

    End With

    ' Get value from dataCell
    WellData = dataCell.value

    Cells(4 + wellIndex, 1).value = "W-" & wellIndex

    ' Set value and number format based on dataArrayName
    With Cells(4 + wellIndex, GetColumnIndex(dataArrayName))
        .value = WellData
        If numberFormats.Exists(dataArrayName) Then
            .numberFormat = numberFormats(dataArrayName)
        End If
    End With
End Sub



Function FXSAVE_GetColumnIndex(ByVal columnName As String) As Integer
    ' Define array to store column indices
    Dim columnIndices As Variant
    columnIndices = Array( _
        11, 13, 2, 3, 7, 8, 9, 10, _
        32, 33, 4, 5, 6, 12, 14, _
        35, 36, 37, 15, 16, 17, 18, _
        19, 20, 21, 22, 23, 24, 25, _
        26, 38, 39, 40, 27, 28, 30, _
        31, 29, 34, 41, 42, 43, 44 _
    )

    ' Define array to store column names
    Dim columnNames As Variant
    columnNames = Array( _
        "Q", "hp", "natural", "stable", "radius", "Rw", "well_depth", "casing", _
        "C", "B", "recover", "Sw", "delta_h", "delta_s", "daeSoo", _
        "T0", "S0", "ER_MODE", "T1", "T2", "TA", "S1", _
        "S2", "K", "time_", "shultze", "webber", "jacob", "skin", _
        "er", "ER1", "ER2", "ER3", "qh", "qg", "sd1", _
        "sd2", "q1", "ratio", "Address", "Company", "S3", "Title" _
    )

    ' Find index of columnName in columnNames array
    Dim index As Integer
    index = Application.match(columnName, columnNames, 0)

    ' Check if columnName exists in columnNames array
    If IsNumeric(index) Then
        GetColumnIndex = columnIndices(index - 1)
    Else
        ' Return -1 if columnName is not found
        GetColumnIndex = -1
    End If
End Function



' in here by refctor by  openai
' replace GetBaseDataFromYangSoo Module
'
'<><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><>

Public Sub FXSAVE_MyDebug(sPrintStr As String, Optional bClear As Boolean = False)
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

Function FXSAVE_DetermineEffectiveRadius(ERMode As String) As Integer
    Dim Er, R As String

    Er = ERMode
    'MsgBox er
    R = Mid(Er, 5, 1)

    If R = "F" Then
        DetermineEffectiveRadius = erRE0
    Else
        DetermineEffectiveRadius = val(R)
    End If
End Function


Function FXSAVE_CheckFileExistence(filePath As String) As Boolean

    If Dir(filePath) <> "" Then
        CheckFileExistence = True
    Else
        CheckFileExistence = False
    End If

End Function


Sub FXSAVE_WriteFormula()

    Dim formula1, formula2 As String
    Dim nofwell As Integer
    Dim i As Integer
    Dim T, Q, Radius, Skin, Er As Double
    Dim T0, S0 As Double
    Dim ERMode As String
    Dim ER1, ER2, ER3, B, S1 As Double

    ' Save array to a file
    Dim filePath As String
    Dim FileNum As Integer


    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select

    filePath = ThisWorkbook.Path & "\" & ActiveSheet.name & ".csv"
    DeleteFileIfExists filePath
    FileNum = FreeFile

    Open filePath For Output As FileNum

    Call MyDebug("Formula SkinFactor ... ", True)

    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"

    ' ��Ų���
    Call FormulaSkinFactorAndER("SKIN", FileNum)

    ' ��ȿ�칰�ݰ�
    Call FormulaSkinFactorAndER("ER", FileNum)


    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"


    Call FormulaChwiSoo(FileNum)
    ' 3-7, ���������

    Call FormulaRadiusOfInfluence(FileNum)
    ' �������ݰ�

    Close FileNum

End Sub



Sub FXSAVE_FormulaSkinFactorAndER(ByVal Mode As String, ByVal FileNum As Integer)
    Dim formula1, formula2 As String
    Dim nofwell As Integer
    Dim i As Integer
    Dim T, Q, Radius, Skin, Er As Double
    Dim T0, S0 As Double
    Dim ERMode As String
    Dim ER1, ER2, ER3, B, S1 As Double


    Call MyDebug("Formula SkinFactor ... ", True)

    nofwell = GetNumberOfWell()

    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"


    For i = 1 To nofwell
        T = format(Cells(4 + i, "o").value, "0.0000")
        Q = Cells(4 + i, "k").value

        T0 = format(Cells(4 + i, "AI").value, "0.0000")
        S0 = format(Cells(4 + i, "AJ").value, "0.0000")
        S1 = Cells(4 + i, "R").value

        delta_s = format(Cells(4 + i, "l").value, "0.00")
        Radius = format(Cells(4 + i, "h").value, "0.000")
        Skin = format(Cells(4 + i, "y").value, "0.0000")
        Er = format(Cells(4 + i, "z").value, "0.0000")


        B = format(Cells(4 + i, "AG").value, "0.0000")
        ER1 = format(Cells(4 + i, "AL").value, "0.0000")
        ER2 = format(Cells(4 + i, "AM").value, "0.0000")
        ER3 = format(Cells(4 + i, "AN").value, "0.0000")


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
            formula1 = "W-" & i & "ȣ��~~r _{e-" & i & "} `=~" & Radius & " TIMES  sqrt {{" & S1 & "} over {" & S0 & "}} `=~" & ER3 & "m"
            formula2 = "erRE3, ����� 3��"

        Case Else
            ' ��Ų���
            formula1 = "W-" & i & "ȣ��~~ sigma  _{w-" & i & "} = {2 pi  TIMES  " & T & " TIMES  " & delta_s & " } over {" & Q & "} -1.15 TIMES  log {2.25 TIMES  " & T & " TIMES  (1/1440)} over {" & S0 & " TIMES  (" & Radius & " TIMES  " & Radius & ")} =`" & Skin
            ' ��ȿ�칰�ݰ�
            formula2 = "W-" & i & "ȣ��~~r _{e-" & i & "} `=~r _{w} e ^{- sigma  _{w-" & i & "}} =" & Radius & " TIMES e ^{-(" & Skin & ")} =" & Er & "m"
        End Select


        If Mode = "SKIN" Then
            Debug.Print formula1
            Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

            Print #FileNum, formula1
            Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Else
            Debug.Print formula2
            Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

            Print #FileNum, formula2
            Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        End If
    Next i

End Sub


Sub FXSAVE_DeleteFileIfExists(filePath As String)
    If Len(Dir(filePath)) > 0 Then ' Check if file exists
        On Error Resume Next
        Kill filePath ' Attempt to delete the file

        On Error GoTo 0
        If Len(Dir(filePath)) > 0 Then
            MsgBox "Unable to delete file '" & filePath & "'. The file may be in use.", vbExclamation
        Else
            ' MsgBox "File '" & filePath & "' has been deleted.", vbInformation

        End If
    Else
        ' MsgBox "File '" & filePath & "' does not exist.", vbExclamation
    End If
End Sub



Sub FXSAVE_FormulaChwiSoo(FileNum As Integer)
' 3-7, ���������

    Dim formula As String
    Dim nofwell As String
    Dim i As Integer
    Dim Q1, S1, S2, res As Double

    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select


    For i = 1 To 3
        Debug.Print " "
        Print #FileNum, " "
    Next i

    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"

    For i = 1 To nofwell
        Q1 = Cells(4 + i, "ac").value

        S1 = format(Cells(4 + i, "ad").value, "0.00")
        S2 = format(Cells(4 + i, "ae").value, "0.00")
        res = format(Cells(4 + i, "ab").value, "0.00")

        'W-1ȣ��~~Q _{& 2} =100 TIMES  ( {8.71} over {4.09} ) ^{2/3} =165.52��/��
        'W-1ȣ��~~Q _{& 2} =100 TIMES  ( {8.71} over {4.09} ) ^{2/3} =`165.52��/��

        formula = "W-" & i & "ȣ��~~Q_{& 2} = " & Q1 & " TIMES ({" & S2 & "} over {" & S1 & "}) ^{2/3} = `" & res & " ��/��"

        Debug.Print formula
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

        Print #FileNum, formula
        Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

    Next i

    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"

End Sub


Sub FXSAVE_FormulaRadiusOfInfluence(FileNum As Integer)
' �������ݰ�

    Call FormulaSUB_ROI("SCHULTZE", FileNum)
    Call FormulaSUB_ROI("WEBBER", FileNum)
    Call FormulaSUB_ROI("JCOB", FileNum)

End Sub




Sub FXSAVE_FormulaSUB_ROI(Mode As String, FileNum As Integer)
  Dim formula1, formula2, formula3 As String
    ' ����, ����, ������

    Dim nofwell As String
    Dim i As Integer
    Dim Shultze, Webber, Jacob, T, K, S, time_, delta_h As String

    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select


    For i = 1 To 3
        Debug.Print " "
        Print #FileNum, " "
    Next i

    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"


    For i = 1 To nofwell
        schultze = CStr(format(Cells(4 + i, "v").value, "0.0"))
        Webber = CStr(format(Cells(4 + i, "w").value, "0.0"))
        Jacob = CStr(format(Cells(4 + i, "x").value, "0.0"))

        T = CStr(format(Cells(4 + i, "q").value, "0.0000"))
        S = CStr(format(Cells(4 + i, "s").value, "0.0000000"))
        K = CStr(format(Cells(4 + i, "t").value, "0.0000"))

        delta_h = CStr(format(Cells(4 + i, "f").value, "0.00"))
        time_ = CStr(format(Cells(4 + i, "u").value, "0.0000"))


        ' Cells(4 + i, "y").value = Format(skin(i), "0.0000")

        formula1 = "W-" & i & "ȣ��~~R _{W-" & i & "} ``=`` sqrt {6 TIMES  " & delta_h & " TIMES  " & K & " TIMES  " & time_ & "/" & S & "} ``=~" & schultze & "m"
        formula2 = "W-" & i & "ȣ��~~R _{W-" & i & "} ``=3`` sqrt {" & delta_h & " TIMES " & K & " TIMES " & time_ & "/" & S & "} `=`" & Webber & "`m"
        formula3 = "W-" & i & "ȣ��~~r _{0(W-" & i & ")} `=~ sqrt {{2.25 TIMES  " & T & " TIMES  " & time_ & "} over {" & S & "}} `=~" & Jacob & "m"


        Select Case Mode
            Case "SCHULTZE"
                Debug.Print formula1
                Print #FileNum, formula1

            Case "WEBBER"
                Debug.Print formula2
                Print #FileNum, formula2

            Case "JCOB"
                Debug.Print formula3
                Print #FileNum, formula3
        End Select

        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

    Next i

    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"

End Sub



