Private Sub CommandButton1_Click()
   Sheets("water").Visible = False
    Sheets("Well").Select
End Sub

Private Sub CommandButton2_Click()
    Dim WB_NAME, cpRange  As String

    If Workbooks.count = 1 Then
        MsgBox "Please Open �������, ���� ", vbOKOnly
        Exit Sub
    End If
   
    ' ������� ������ �ҷ����� ���� ����
    WB_NAME = Sheet4_Water.GetOtherFileName
    
    cpRange = GetCopyPoint(WB_NAME)
    Call CopyFromGWAN_JUNG(WB_NAME, cpRange)
    Call FormulaInjection
    
End Sub

' 2024-01-14
' inject formula ...

Private Sub FormulaInjection()
    Dim nofwell, i As Integer
    
    nofwell = GetNumberOfWell()
    For i = 4 To nofwell + 3
        Sheets("Well").Cells(i, "O").formula = "=ROUND(water!$F$7, 1)"
    Next i

End Sub



'�⺻���������͸� ������ �´�.
Function GetOtherFileName() As String
    Dim Workbook As Workbook
    Dim workbookNames As String
    Dim i As Long

    workbookNames = ""
    
    For Each Workbook In Application.Workbooks
        If StrComp(ThisWorkbook.name, Workbook.name, vbTextCompare) = 0 Then
            GoTo NEXT_ITERATION
        End If
        
        If CheckSubstring(Workbook.name, "����") Then
            Exit For
        End If
        
NEXT_ITERATION:
    Next
    
    GetOtherFileName = Workbook.name
End Function


'
'Function lastRowByKey(cell As String) As Long
'    lastRowByKey = Range(cell).End(xlDown).Row
'End Function


Function GetCopyPoint(ByVal fName As String) As String

  Dim ip1, ip2 As Integer

  ip1 = Workbooks(fName).Worksheets("ss").Range("b1").End(xlDown).row + 4
  ip2 = ip1 + 2
  
  GetCopyPoint = "B" & ip1 & ":J" & ip2
  ThisWorkbook.Activate

End Function


Sub CopyFromGWAN_JUNG(ByVal fName As String, ByVal cpRange As String)

    Workbooks(fName).Worksheets("ss").Activate
    Workbooks(fName).Worksheets("ss").Range(cpRange).Select
    Selection.Copy
    
    ThisWorkbook.Sheets("water").Activate
    
    Range("d6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub

Sub ListOpenWorkbookNames()
    Dim Workbook As Workbook
    Dim workbookNames As String
    Dim i As Long
        
    workbookNames = ""
    
    For Each Workbook In Application.Workbooks
        workbookNames = workbookNames & Workbook.name & vbCrLf
    Next
    
    Cells(1, 1).value = workbookNames
End Sub

