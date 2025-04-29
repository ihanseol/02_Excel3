Option Explicit

Private Sub Workbook_Open()
      

    With shW_LongTEST.ComboBox1
        .AddItem "60"
        .AddItem "75"
        .AddItem "90"
        .AddItem "105"
        .AddItem "120"
        .AddItem "140"
        .AddItem "160"
        .AddItem "180"
        .AddItem "240"
        .AddItem "300"
        .AddItem "360"
        .AddItem "420"
        .AddItem "480"
        .AddItem "540"
        .AddItem "600"
        .AddItem "660"
        .AddItem "720"
        .AddItem "780"
        .AddItem "840"
        .AddItem "900"
        .AddItem "960"
        .AddItem "1020"
        .AddItem "1080"
        .AddItem "1140"
        .AddItem "1200"
        .AddItem "1260"
        .AddItem "1320"
        .AddItem "1380"
        .AddItem "1440"
        .AddItem "1500"
    End With
   
    Call initDictionary
    ' Call GotoTopPosition
    
    mod_INPUT.gblTestTime = shW_aSkinFactor.Range("C9").Value
    
    If mod_INPUT.gblTestTime = 2880 Then
        shInput.OptionButton1.Value = True
    Else
        shInput.OptionButton2.Value = True
        mod_INPUT.gblTestTime = 1440
    End If
    
    ' 2025/4/29
    ' depends on yangsoo test time, sheet on off
    
    sh01_StepSelect.name = "Step.Select"
           
    If mod_INPUT.gblTestTime = 2880 Then
        sh02_JanggiSelect1.name = "J1440"
        sh03_RecoverSelect1.name = "R120"
        
        sh02_JanggiSelect.name = "Janggi.Select"
        sh03_RecoverSelect.name = "Recover.Select"
        
        
        sh02_JanggiSelect.Visible = True
        sh03_RecoverSelect.Visible = True
        sh02_JanggiSelect1.Visible = False
        sh03_RecoverSelect1.Visible = False
    Else
        sh02_JanggiSelect.name = "J2880"
        sh03_RecoverSelect.name = "R360"
        
        sh02_JanggiSelect1.name = "Janggi.Select"
        sh03_RecoverSelect1.name = "Recover.Select"
        
        sh02_JanggiSelect.Visible = False
        sh03_RecoverSelect.Visible = False
        sh02_JanggiSelect1.Visible = True
        sh03_RecoverSelect1.Visible = True
        
    End If
       
    'shInput.Frame1.Controls("optionbutton1").Value = True
    
End Sub


'Private Sub GotoTopPosition()
'
'    Dim sht As Worksheet
'
'    For Each sht In Application.Worksheets
'        sht.Activate
'        Application.GoTo Reference:=Range("a1"), Scroll:=True
'    Next sht
'
'    shInput.Activate
'
'End Sub
'
