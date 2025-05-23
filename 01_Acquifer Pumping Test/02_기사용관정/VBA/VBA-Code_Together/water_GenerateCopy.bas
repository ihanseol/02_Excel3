
' *********************************************************************
' * water_GenerationCopy                                              *
' *********************************************************************
' *  ShortCut of this sheet                                           *
' *                                                                   *
' *********************************************************************
' * Ctrl+I , insert Row                                               *
' * 2025/5/23                                                         *
' *********************************************************************
' * Ctrl+Shift+I , Initial Clear                                      *
' * SubModuleInitialClear()                                           *
' * 2025/5/23                                                         *
' *********************************************************************
' * TransferWellData                                                  *
' * Ctrl+R , Transfer Well Data                                       *
' * =D2&" "&E2&" 번지"                                                 *
' * 2025/5/15                                                         *
' *********************************************************************
' * Ctrl+Q,  Toggle SS and AA Sheet                                   *
' * 2025/5/15                                                         *
' *********************************************************************
' *  Ctrl+P , ExportDataSheet()                                       *
' *  2025/5/15                                                        *
' *********************************************************************
' * Ctrl+D , Toggle OX, Toggle SINGO, HEOGA' Toggle SS and AA Sheet   *
' * 2025/5/15                                                         *
' *********************************************************************
' * Ctrl+Shift+F , Finalize Active Sheet                              *
' * 2025/5/23                                                         *
' *********************************************************************
' * Ctrl+Shift+C , Main, Generate Copy                                *
' * MainMoudleGenerateCopy()                                          *
' * 2025/5/23                                                         *
' *********************************************************************
' * Ctrl+Shift+D , Delete                                             *
' * SubModuleCleanCopySection()                                       *
' * 2025/5/23                                                         *
' *********************************************************************
' * Ctrl+Shift+T , Toggle Sheet Show and Hide hidden Sheet            *
' * ShowHiddenSheet()                                                 *
' * 2025/5/23                                                         *
' *********************************************************************


Option Explicit

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long




' *********************************************************************
' * Ctrl+Shift+T , Toggle Sheet Show and Hide hidden Sheet            *
' * ShowHiddenSheet()                                                 *
' * 2025/5/23                                                         *
' *********************************************************************

Sub ShowHiddenSheet()

  If Sheets("ref").Visible Then
        Sheets("ref").Visible = False
        Sheets("ref1").Visible = False
        Sheets("ss_out").Visible = False
        Sheets("aa_out").Visible = False
        Sheets("ii_out").Visible = False
    Else
        Sheets("ref").Visible = True
        Sheets("ref1").Visible = True
        Sheets("ss_out").Visible = True
        Sheets("aa_out").Visible = True
        Sheets("ii_out").Visible = True
    End If

End Sub


' *********************************************************************
' * Ctrl+Q,  Toggle SS and AA Sheet                                   *
' * 2025/5/15                                                         *
' *********************************************************************
Sub SS_Active()
    If ActiveSheet.Name = "ss" Then
        Sheet2_aa.Activate
    Else
        Sheet5_ss.Activate
    End If
End Sub

Private Function lastRowByKey(cell As String) As Long
    lastRowByKey = Range(cell).End(xlDown).row
End Function


Private Function lastRowByRowsCount(cell As String) As Long
    lastRowByRowsCount = Cells(Rows.Count, cell).End(xlUp).row
End Function

Public Sub clearRowA()
    
End Sub

Private Function lastRowByFindAll() As Long
    Dim lastrow As Long
    
    lastrow = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    lastRowByFind = lastrow
End Function


' 여기서 검색시, "AA" 같은 경우에는, 셀의 텍스트 데이타 뿐만 아니라 ...
' =SUMIF(SS_INSIDE_AREA,"O",$L$2:$L$2) 의 경우에서 처럼, 이런것도 검색이 되기에,
' 일단은 Ctrl+F 로 검색을 해보는것을 추천한다.
' 이것은, 엑셀의 검색을 이용해서, 서치하는 함수이기 때문이다.

Private Function lastRowByFind(ByVal str As String) As Long
    Dim lastrow As Long
    
    lastrow = Cells.Find(str, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    lastRowByFind = lastrow
End Function


Private Sub DoCopy(lastrow As Long)
    Range("F2:H" & lastrow).Select
    Selection.Copy
    
    Range("n2").Select
    ActiveSheet.Paste
    
    
    ' 물량
    Range("L2:L" & lastrow).Select
    Selection.Copy
    
    Range("q2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("k2:k" & lastrow).Select
    Selection.Copy
    
    Range("r2").Select
    ActiveSheet.Paste
    
    Range("N14").Select
    Application.CutCopyMode = False
End Sub



' return letter of range ...
Function Alpha_Column(Cell_Add As Range) As String
    Dim No_of_Rows As Integer
    Dim No_of_Cols As Integer
    Dim Num_Column As Integer
    
    No_of_Rows = Cell_Add.Rows.Count
    No_of_Cols = Cell_Add.Columns.Count
    
    If ((No_of_Rows <> 1) Or (No_of_Cols <> 1)) Then
        Alpha_Column = ""
        Exit Function
    End If
    
    Num_Column = Cell_Add.Column
    If Num_Column < 26 Then
        Alpha_Column = Chr(64 + Num_Column)
    Else
        Alpha_Column = Chr(Int(Num_Column / 26) + 64) & Chr((Num_Column Mod 26) + 64)
    End If
End Function



' *********************************************************************
' *  Ctrl+P , ExportDataSheet()                                       *
' *  2025/5/15                                                        *
' *********************************************************************
Sub ExportDataSheet()
    Call ExportDataWorksheet("ss_out")
    Call ExportDataWorksheet("aa_out")
End Sub

' *********************************************************************
' *  Ctrl+D , Toggle OX, Toggle SINGO, HEOGA' Toggle SS and AA Sheet  *
' *  2025/5/15                                                        *
' *********************************************************************
Sub ToggleOX()
    Dim activeCellColumn, activeCellRow As String
    Dim row As Long
    Dim col As Long
    Dim lastrow As Long
    Dim cp, fillRange As String
    

    activeCellColumn = Split(ActiveCell.address, "$")(1)
    activeCellRow = Split(ActiveCell.address, "$")(2)
  
    row = ActiveCell.row
    col = ActiveCell.Column
    
    Debug.Print Alpha_Column(ActiveCell)
    
    
    '2024-12-25, Add Compute Q
    If activeCellColumn = "L" Then
        Popup_MessageBox ("Calculation Compute Q .... ")
        Call water_q.ComputeQ
        Sheets("ss").Activate
    End If
    
    
    
    If activeCellColumn = "S" Then
        If ActiveCell.Value = "O" Then
            ActiveCell.Value = "X"
        Else
            ActiveCell.Value = "O"
        End If
    End If
    

    If activeCellColumn = "B" Then
        If ActiveCell.Value = "신고공" Then
            ActiveCell.Value = "허가공"
            With Selection.Font
                .Color = -16776961
                .TintAndShade = 0
            End With
            Selection.Font.Bold = True
        Else
            ActiveCell.Value = "신고공"
             With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            Selection.Font.Bold = False
        End If
    End If
    
    If activeCellColumn = "D" Then
        cp = Replace(ActiveCell.address, "$", "")
        lastrow = lastRowByKey(ActiveCell.address)
        
        fillRange = "D" & Range(cp).row & ":D" & lastrow
        
        Range(cp).Select
        Selection.Copy
        Range(fillRange).Select
        
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        Range(cp).Select
        Application.CutCopyMode = False
    End If
    
    If activeCellColumn = "C" Then
        cp = Replace(ActiveCell.address, "$", "")
        lastrow = lastRowByKey(ActiveCell.address)
        
        fillRange = "C" & Range(cp).row & ":C" & lastrow
        
        Range(cp).Select
        Selection.Copy
        Range(fillRange).Select
        
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        Range(cp).Select
        Application.CutCopyMode = False
    End If
    
       
    ' 2024,12,22 - toggle address format
    If activeCellColumn = "M" Then
      Call AddressReset(ActiveSheet.Name)
    End If
    
        
    If ActiveSheet.Name = "ss" And activeCellColumn = "K" Then
        UserForm_SS.Show
    End If
    
    If ActiveSheet.Name = "aa" And activeCellColumn = "K" Then
        UserForm_AA.Show
    End If
    
    If ActiveSheet.Name = "ii" And activeCellColumn = "K" Then
        UserForm_II.Show
    End If
    
    
    ' 2025, 5, 15 - Export UsedWell DataSheet
    If (ActiveSheet.Name = "ss" Or ActiveSheet.Name = "aa") And activeCellColumn = "N" Then
      ' ExportAllUsedWellData
      Call MakeOutSheet
    End If
    
End Sub


' *********************************************************************
' * 2025/5/15                                                         *
' * Make Export DataSheet,  SS, AA, II                                *
' *********************************************************************
Sub MakeOutSheet()
    Dim i As Integer
    Dim nSS, nAA, nII As Integer
   
' *********************************************************************
' *   Initial Clear & Number Setting                                  *
' *********************************************************************
   
    ' Range("W6").Value = Sheets("ss").Range("ss_in_count").Value
    
    nSS = Sheets("ss").Range("ss_in_count").Value + Sheets("ss").Range("ss_out_count").Value
    nAA = Sheets("ss").Range("aa_in_count").Value + Sheets("ss").Range("aa_out_count").Value
    nII = Sheets("ss").Range("ii_in_count").Value + Sheets("ss").Range("ii_out_count").Value
    
    Sheets("ss_out").Activate
    ActiveSheet.Range("A2:Z300").Select
    Selection.ClearContents
    
    For i = 1 To nSS
        Cells(i + 1, "A").Value = "S-" & i
    Next i
    
    Sheets("aa_out").Activate
    ActiveSheet.Range("A2:Z300").Select
    Selection.ClearContents
        
    For i = 1 To nAA
        Cells(i + 1, "A").Value = "A-" & i
    Next i
    
    
    Sheets("ii_out").Activate
    ActiveSheet.Range("A2:Z300").Select
    Selection.ClearContents
    Range("A1").Select

    For i = 1 To nII
        Cells(i + 1, "A").Value = "I-" & i
    Next i
    
' *********************************************************************
' *   Initial Clear & Number Setting                                  *
' *********************************************************************

    Sheets("ss").Activate
    If nSS >= 3 Then
        ActiveSheet.Range("m2:s" & CStr(nSS + 1)).Select
        Selection.Copy
        Range("A1").Select
        Sheets("ss_out").Activate
        ActiveSheet.Range("b2").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
        Range("A1").Select
    End If
    
    Sheets("aa").Activate
    
    If nAA >= 3 Then
        ActiveSheet.Range("m2:s" & CStr(nAA + 1)).Select
        Selection.Copy
        Range("A1").Select
        Sheets("aa_out").Activate
        ActiveSheet.Range("b2").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
        Range("A1").Select
    End If
    
    Sheets("ii").Activate
    
    If nII > 2 Then
        ActiveSheet.Range("m2:s" & CStr(nII + 1)).Select
        Selection.Copy
        Range("A1").Select
        Sheets("ii_out").Activate
        ActiveSheet.Range("b2").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
        Range("A1").Select
    End If

End Sub


' Ctrl+R , Transfer Well Data
' =D2&" "&E2&" 번지"
Sub ToggleAddressFormatString()

    Dim activeCellColumn, activeCellRow As String
    Dim row As Long
    Dim col As Long
    Dim lastrow As Long
    Dim cp, fillRange As String
    Dim MainSheet, TargetSheet As String
    
    activeCellColumn = Split(ActiveCell.address, "$")(1)
    activeCellRow = Split(ActiveCell.address, "$")(2)
  
 
    If lastrow = 1048577 Or Range("E" & (lastrow - 1)).Value = "생활용" Then
        lastrow = 2
    End If
    
    
    Range("E" & lastrow).Select
    ActiveSheet.Paste
    
       
    AddressReset (MainSheet)
    AddressReset (TargetSheet)

End Sub



' *********************************************************************
' * 2025/5/15                                                         *
' * TransferWellData                                                  *
' * Ctrl+R , Transfer Well Data                                       *
' * =D2&" "&E2&" 번지"                                                 *
' *********************************************************************
Sub TransferWellData()

    Dim activeCellColumn, activeCellRow As String
    Dim row As Long
    Dim col As Long
    Dim lastrow As Long
    Dim cp, fillRange As String
    Dim MainSheet, TargetSheet As String
    
    activeCellColumn = Split(ActiveCell.address, "$")(1)
    activeCellRow = Split(ActiveCell.address, "$")(2)
  
    row = ActiveCell.row
    col = ActiveCell.Column
    
    MainSheet = ActiveSheet.Name
    
    If MainSheet = "aa" Then
        TargetSheet = "ss"
    ElseIf MainSheet = "ss" Then
        TargetSheet = "aa"
    Else
        Exit Sub
    End If
    
    fillRange = "E" & row & ":J" & row
    Range(fillRange).Select
    Selection.Cut
    
    Sheets(TargetSheet).Activate
    lastrow = lastRowByKey("E1") + 1
    
    If lastrow = 1048577 Or Range("E" & (lastrow - 1)).Value = "생활용" Then
        lastrow = 2
    End If
    
    
    Range("E" & lastrow).Select
    ActiveSheet.Paste
    
       
    AddressReset (MainSheet)
    AddressReset (TargetSheet)
    
    Range("E2").Select
    Selection.End(xlDown).Select


End Sub

' =D2&" "&E2&" 번지"
Sub AddressReset(Optional ByVal shName As String = "option")
    Dim lastrow As Long
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    sheetExists = False
    
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = shName Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    
    If Not sheetExists Then
        shName = ActiveSheet.Name
    End If
    
    
    Sheets(shName).Activate
    
    lastrow = lastRowByKey("M2")
    
    On Error Resume Next
    If CheckSubstring(Range("M2"), "번지") Then
        Range("M2").Formula = "=D2&"" ""&E2"
    Else
        Range("M2").Formula = "=D2&"" ""&E2&""번지"" "
    End If
    
    
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M" & lastrow)
    
    Range("E2").Select
    Selection.End(xlDown).Select
  
End Sub

Sub test()
    Dim lastrow As Long
    
    
    lastrow = lastRowByKey("E1") + 1
        
   If lastrow = 1048577 Or ActiveCell.Value = "생활용" Then
        lastrow = 2
    End If
    
    Range("o2").Value = "ll"
End Sub



Sub ClearSystemClipboard()
    OpenClipboard 0
    EmptyClipboard
    CloseClipboard
End Sub


' *********************************************************************
' * Ctrl+Shift+C , Main, Generate Copy                                *
' * MainMoudleGenerateCopy()                                          *
' * 2025/5/23                                                         *
' *********************************************************************
Sub MainMoudleGenerateCopy()
    Dim lastrow As Long
        
    Call ClearSystemClipboard
    Call TurnOffStuff
    
    Sheets("ss").Activate
    lastrow = lastRowByKey("A1")
    Call DoCopy(lastrow)
    
    Sheets("aa").Activate
    lastrow = lastRowByKey("A1")
    Call DoCopy(lastrow)
    
    Sheets("ii").Activate
    lastrow = lastRowByKey("A1")
    Call DoCopy(lastrow)
    
    Sheets("ss").Activate
    
    Call TurnOnStuff
    
End Sub


' *********************************************************************
' * Ctrl+Shift+I , Initial Clear                                      *
' * SubModuleInitialClear()                                           *
' * 2025/5/23                                                         *
' *********************************************************************
Sub SubModuleInitialClear()
    Dim lastrow As Long
    Dim userChoice As VbMsgBoxResult
    
    lastrow = lastRowByKey("A1")
  
    userChoice = MsgBox("Do you want to continue?", vbOKCancel, "Confirmation")

    If userChoice <> vbOK Then
        Exit Sub
    End If
    
    Range("e2:j" & lastrow).Select
    Selection.ClearContents
    Range("n2:r" & lastrow).Select
    Selection.ClearContents
    
    If lastrow >= 23 Then
        Rows("23:" & lastrow).Select
        Selection.Delete Shift:=xlUp
    End If
    
    
    If (ActiveSheet.Name = "ii") Then
        Range("l2").Value = 0
    End If
    
    Range("m2").Select
End Sub

Sub SaveBackupToDocuments()
    Dim documentsPath As String
    Dim backupPath As String
    
    On Error GoTo ErrorHandler
    
    documentsPath = Environ("USERPROFILE") & "\Documents\"
    backupPath = documentsPath & "backup.xlsm"
    ThisWorkbook.SaveCopyAs backupPath
    ' MsgBox "Backup saved to: " & backupPath, vbInformation, "Backup Complete"
    Exit Sub
ErrorHandler:
    ' MsgBox "Error saving backup: " & Err.Description, vbCritical, "Backup Error"
End Sub


' *********************************************************************
' * Ctrl+Shift+F , Finalize Active Sheet                              *
' * 2025/5/23                                                         *
' *********************************************************************
Sub Finallize()
    Dim lastrow As Long
    Dim delStartRow, delEndRow, delAddressStart As Long
    Dim userChoice As VbMsgBoxResult
    Dim shName As String
       
    ' 2025/5/22, Just in cas finalize then backkup my file ...
    Call SaveBackupToDocuments
    ' shName = ActiveSheet.Name
    
    If Range("L2").Value = 0 Then
        Call Popup_MessageBox("Calculation Compute Q .... ")
        Call water_q.ComputeQ
    End If
    
    lastrow = lastRowByKey("A1")
    delStartRow = lastRowByKey("D1") + 1
    
    delAddressStart = lastRowByKey("E1") + 1
    
    
    Select Case ActiveSheet.Name
    
        Case "ss"
            delEndRow = lastRowByFind("구분") - 4
            
        Case "aa"
            delEndRow = lastRowByFind("유역내") - 4
        
        Case "ii"
            delEndRow = lastRowByFind("유역내") - 6
    
    End Select
    
    '
    'if q is 0 then this section is not have water resource so clear next well
    '
    If Range("L2").Value = 0 Then
        delStartRow = 3
        delEndRow = lastRowByKey("L1")
    Else
        delStartRow = delAddressStart
    End If
    
    
    userChoice = MsgBox("Do you want to continue?", vbOKCancel, "Confirmation")

    If userChoice <> vbOK Then
        Exit Sub
    End If
    
    If delStartRow = 1048577 Or lastrow = 2 Or (delEndRow - delStartRow <= 2) Then
        Exit Sub
    Else
        Rows(delStartRow & ":" & delEndRow).Select
        Selection.Delete Shift:=xlUp
        Range("A2").Select
    End If
      
End Sub



' *********************************************************************
' * Ctrl+Shift+D , Delete                                             *
' * SubModuleCleanCopySection()                                       *
' * 2025/5/23                                                         *
' *********************************************************************
Sub SubModuleCleanCopySection()
    Dim lastrow As Long
        
    Call TurnOffStuff
    
    Sheets("ss").Activate
    lastrow = lastRowByKey("A1")
    Range("n2:r" & lastrow).Select
    Selection.ClearContents
    Range("a1").Select
    
    Sheets("aa").Activate
    lastrow = lastRowByKey("A1")
    Range("n2:r" & lastrow).Select
    Selection.ClearContents
    Range("a1").Select
    
    Call TurnOnStuff
    
    Sheets("ss").Activate
End Sub


' *********************************************************************
' * Ctrl+I , insert Row                                               *
' * 2025/5/23                                                         *
' *********************************************************************
Sub insertRow()
    Dim lastrow As Long, i As Long, j As Long
    Dim selection_origin, selection_target As String
    Dim AddingRowCount As Long
    
    'lastRow = lastRowByKey("A1")

    AddingRowCount = 10

    lastrow = lastRowByRowsCount("A")
    
    Rows(CStr(lastrow + 1) & ":" & CStr(lastrow + AddingRowCount)).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    
    i = lastRowByKey("A1")
    j = i + AddingRowCount
    
    selection_origin = "A" & i & ":D" & i
    selection_target = "A" & i & ":D" & j
    
    Range(selection_origin).Select
    Selection.AutoFill Destination:=Range(selection_target), Type:=xlFillDefault
 
    selection_origin = "K" & i & ":M" & i
    selection_target = "K" & i & ":M" & j

    Range(selection_origin).Select
    Selection.AutoFill Destination:=Range(selection_target), Type:=xlFillDefault
    
    Range("S" & i).Select
    Selection.AutoFill Destination:=Range("S" & i & ":S" & j), Type:=xlFillDefault
    
    Application.CutCopyMode = False
    
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
End Sub


Public Sub TurnOffStuff()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
End Sub

Public Sub TurnOnStuff()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub






