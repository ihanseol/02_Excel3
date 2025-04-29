Option Explicit

Private Sub CommandButton_CB1_Click()
Top:
    On Error GoTo ErrorCheck
    Call set_CB1
    Exit Sub
    
ErrorCheck:
    GoTo Top
End Sub

Private Sub CommandButton_CB2_Click()
Top:
    On Error GoTo ErrorCheck
    Call set_CB2
    Exit Sub
    
ErrorCheck:
    GoTo Top
End Sub

' Chart Button
Private Sub CommandButton_Chart_Click()
    Dim gong        As Integer
    Dim KeyCell     As Range
    
    Call adjustChartGraph
    
    Set KeyCell = Range("J48")
    
    gong = Val(CleanString(KeyCell.Value))
    Call mod_Chart.SetChartTitleText(gong)
    
    Call mod_INPUT.Step_Pumping_Test
    Call mod_INPUT.Vertical_Copy
End Sub

Private Sub CommandButton_ClearReport_Click()
    
    DeleteStepSheet
       
End Sub

Sub DeleteStepSheet()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    
    On Error Resume Next
    Set ws1 = Sheets("Step")
    Set ws2 = Sheets("out")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        Application.DisplayAlerts = False
        Sheets("Step").Delete
        Application.DisplayAlerts = True
    Else
        Debug.Print "Sheet 'Step' does not exist."
    End If
    
    If Not ws2 Is Nothing Then
        Application.DisplayAlerts = False
        Sheets("out").Delete
        Application.DisplayAlerts = True
    Else
        Debug.Print "Sheet 'out' does not exist."
    End If
End Sub



Private Sub CommandButton_ResetScreenSize_Click()
    Call ResetScreenSize
End Sub

Private Sub CommandButton_STEP_Click()
    Call Make_Step_Document
End Sub

Private Sub CommandButton_2880_Click()
    'Call make_long_document
    Call Make2880_Document
End Sub

Private Sub CommandButton_1440_Click()
    Call Make2880_Document
    Call make1440sheet
End Sub

Private Sub CommandButton8_Click()
    Call set_CB_ALL
End Sub

Private Sub CommandButton1_Click()

    Sheets("장회").Visible = True
    Sheets("장회14").Visible = True
    Sheets("단계").Visible = True
    Sheets("장기28").Visible = True
    Sheets("장기14").Visible = True
    Sheets("회복").Visible = True
    Sheets("회복12").Visible = True

End Sub



Private Sub CommandButton2_Click()

    Sheets("장회").Visible = False
    Sheets("장회14").Visible = False
    Sheets("단계").Visible = False
    Sheets("장기28").Visible = False
    Sheets("장기14").Visible = False
    Sheets("회복").Visible = False
    Sheets("회복12").Visible = False

End Sub



Sub ColoringTestTime()
' change current setting by yangsoo day, 2880 or 1440
' change background color by current selecttion
    
    Call TurnOffStuff
    Sheets("SkinFactor").Activate
    
    shW_aSkinFactor.Range("C10:D11").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
    
    If mod_INPUT.gblTestTime = 2880 Then
        
        shW_aSkinFactor.Range("C10:C11").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 13500415
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    Else

        shW_aSkinFactor.Range("D10:D11").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 13500415
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    
    End If
    
    shW_aSkinFactor.Range("H10").Select
    
    Sheets("Input").Activate
    Call TurnOnStuff
    
End Sub

Private Sub OptionButton1_Click()
    mod_INPUT.gblTestTime = 2880
    'MsgBox "2880"
    shW_aSkinFactor.Range("C9").Value = 2880
    Call ColoringTestTime
    Call mod_W1.Restore2880
End Sub

Private Sub OptionButton2_Click()
    mod_INPUT.gblTestTime = 1440
    'MsgBox "1440"
    shW_aSkinFactor.Range("C9").Value = 1440
    Call ColoringTestTime
    Call mod_W1.Delete2880
End Sub


