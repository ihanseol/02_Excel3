' ***************************************************************
' UserForm_SS
'
' ***************************************************************

' Optionbutton1 - ������
' Optionbutton2 - �Ϲݿ�
' Optionbutton3 - û�ҿ�
' Optionbutton4 - �ι�����
' Optionbutton5 - �б���
' Optionbutton6 - �������ÿ�
' Optionbutton7 - ���̻����
' Optionbutton8 - ���Ȱ���
' Optionbutton9 - ��Ÿ
' Optionbutton10 - �����
' Optionbutton11 - �����ó���
' Optionbutton12 - �����
' Optionbutton13 - �ҹ��

Private Sub CommandButton1_Click()
    Dim i As Integer
    Dim options() As Variant
    Dim selectedOption As String
    
    ' Assign captions to an array
    options = Array("������", "�Ϲݿ�", "û�ҿ�", "�ι�����", "�б���", "�������ÿ�", "���̻����", "���Ȱ���", "��Ÿ", "�����", "�����ó���", "�����", "�ҹ��")
    
    ' Loop through OptionButtons to find the selected one
    For i = 0 To 12
        If Controls("OptionButton" & i + 1).Value Then
            selectedOption = options(i)
            Exit For
        End If
    Next i
    
    ' Set the value of the active cell
    If selectedOption <> "" Then
        ActiveCell.Value = selectedOption
        Unload Me
    Else
        MsgBox "Please select an option."
    End If
End Sub

Private Sub CommandButton2_Click()
  Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Me.StartUpPosition = 0
    
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
   OptionButton1.Value = True
End Sub

'Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = 27 Then
'        Unload Me
'    End If
'End Sub
