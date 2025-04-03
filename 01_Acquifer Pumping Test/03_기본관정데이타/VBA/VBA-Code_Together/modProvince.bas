Function GetProvince_Dictionary(city As String) As String
    Dim cityProvinceMap As Object
    Set cityProvinceMap = CreateObject("Scripting.Dictionary")
    
    ' ��û��
    cityProvinceMap.Add "����", "��û��"
    cityProvinceMap.Add "��õ", "��û��"
    cityProvinceMap.Add "û��", "��û��"
    cityProvinceMap.Add "��ǳ��", "��û��"
    cityProvinceMap.Add "����", "��û��"
    cityProvinceMap.Add "����", "��û��"
    cityProvinceMap.Add "�ݻ�", "��û��"
    cityProvinceMap.Add "����", "��û��"
    cityProvinceMap.Add "�ο�", "��û��"
    cityProvinceMap.Add "����", "��û��"
    cityProvinceMap.Add "õ��", "��û��"
    cityProvinceMap.Add "ȫ��", "��û��"

    ' ������
    cityProvinceMap.Add "���ǻ�", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "��ȭ", "������"
    cityProvinceMap.Add "��ɵ�", "������"
    cityProvinceMap.Add "��õ", "������"
    cityProvinceMap.Add "����õ", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "��õ", "������"
    cityProvinceMap.Add "����", "������"

    ' ������
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "�����", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "�ϰ���", "������"
    cityProvinceMap.Add "����õ", "������"
    cityProvinceMap.Add "��ô", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "����", "������"
    cityProvinceMap.Add "������", "������"
    cityProvinceMap.Add "ö��", "������"
    cityProvinceMap.Add "��õ", "������"
    cityProvinceMap.Add "�¹�", "������"
    cityProvinceMap.Add "ȫõ", "������"

    ' ����
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "��â", "����"
    cityProvinceMap.Add "��â��", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "�ξ�", "����"
    cityProvinceMap.Add "��â��", "����"
    cityProvinceMap.Add "�ӽ�", "����"
    cityProvinceMap.Add "���", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "������", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "�����", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "������", "����"
    cityProvinceMap.Add "��õ", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "������", "����"
    cityProvinceMap.Add "�ϵ�", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "�־�", "����"
    cityProvinceMap.Add "����", "����"
    cityProvinceMap.Add "÷ö��", "����"
    cityProvinceMap.Add "������", "����"
    cityProvinceMap.Add "�س�", "����"
    cityProvinceMap.Add "��굵", "����"

    ' ���
    cityProvinceMap.Add "�뱸", "���"
    cityProvinceMap.Add "�뱸(��)", "���"
    cityProvinceMap.Add "���", "���"
    cityProvinceMap.Add "�λ�", "���"
    cityProvinceMap.Add "���ֽ�", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "��ȭ", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "�ȵ�", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "��õ", "���"
    cityProvinceMap.Add "�︪��", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "�Ǽ�", "���"
    cityProvinceMap.Add "û�۱�", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "��â", "���"
    cityProvinceMap.Add "���ؽ�", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "�о�", "���"
    cityProvinceMap.Add "��â��", "���"
    cityProvinceMap.Add "��û", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "�Ƿɱ�", "���"
    cityProvinceMap.Add "����", "���"
    cityProvinceMap.Add "â��", "���"
    cityProvinceMap.Add "�뿵", "���"
    cityProvinceMap.Add "�Ծ籺", "���"
    cityProvinceMap.Add "��õ", "���"

    ' ���ֵ�
    cityProvinceMap.Add "���", "���ֵ�"
    cityProvinceMap.Add "������", "���ֵ�"
    cityProvinceMap.Add "����", "���ֵ�"
    cityProvinceMap.Add "����2", "���ֵ�"
    cityProvinceMap.Add "������", "���ֵ�"
    
    ' Return the province if found
    If cityProvinceMap.Exists(city) Then
        GetProvince_Dictionary = cityProvinceMap(city)
    Else
        GetProvince_Dictionary = "Not in list"
    End If
    
    ' Clean up
    Set cityProvinceMap = Nothing
End Function



Function GetProvince_Vlookup(city As String) As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim result As Variant
    
    ' Set the worksheet and range
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to the name of your sheet
    Set rng = ws.Range("A1:B100") ' Adjust the range to match your data

    ' Use VLOOKUP to find the province
    result = Application.WorksheetFunction.VLookup(city, rng, 2, False)
    
    ' Return the result or handle the error
    If IsError(result) Then
        GetProvince_Vlookup = "Not in list"
    Else
        GetProvince_Vlookup = result
    End If
End Function


Function GetProvince_Collection(city As String) As String
    Dim cityProvinceArray As Collection
    Dim cityProvince As Variant
    Dim i As Integer
    
    ' Initialize the collection with city-province pairs
    Set cityProvinceArray = New Collection
    
    ' Add city-province pairs to the collection
    cityProvinceArray.Add Array("����", "��û��")
    cityProvinceArray.Add Array("��õ", "��û��")
    cityProvinceArray.Add Array("û��", "��û��")
    cityProvinceArray.Add Array("��ǳ��", "��û��")
    cityProvinceArray.Add Array("����", "��û��")
    cityProvinceArray.Add Array("����", "��û��")
    cityProvinceArray.Add Array("�ݻ�", "��û��")
    cityProvinceArray.Add Array("����", "��û��")
    cityProvinceArray.Add Array("�ο�", "��û��")
    cityProvinceArray.Add Array("����", "��û��")
    cityProvinceArray.Add Array("õ��", "��û��")
    cityProvinceArray.Add Array("ȫ��", "��û��")
    cityProvinceArray.Add Array("���ǻ�", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("��ȭ", "������")
    cityProvinceArray.Add Array("��ɵ�", "������")
    cityProvinceArray.Add Array("��õ", "������")
    cityProvinceArray.Add Array("����õ", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("��õ", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("�����", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("�ϰ���", "������")
    cityProvinceArray.Add Array("����õ", "������")
    cityProvinceArray.Add Array("��ô", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("����", "������")
    cityProvinceArray.Add Array("������", "������")
    cityProvinceArray.Add Array("ö��", "������")
    cityProvinceArray.Add Array("��õ", "������")
    cityProvinceArray.Add Array("�¹�", "������")
    cityProvinceArray.Add Array("ȫõ", "������")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("��â", "����")
    cityProvinceArray.Add Array("��â��", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("�ξ�", "����")
    cityProvinceArray.Add Array("��â��", "����")
    cityProvinceArray.Add Array("�ӽ�", "����")
    cityProvinceArray.Add Array("���", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("������", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("�����", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("������", "����")
    cityProvinceArray.Add Array("��õ", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("������", "����")
    cityProvinceArray.Add Array("�ϵ�", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("�־�", "����")
    cityProvinceArray.Add Array("����", "����")
    cityProvinceArray.Add Array("÷ö��", "����")
    cityProvinceArray.Add Array("������", "����")
    cityProvinceArray.Add Array("�س�", "����")
    cityProvinceArray.Add Array("��굵", "����")
    cityProvinceArray.Add Array("�뱸", "���")
    cityProvinceArray.Add Array("�뱸(��)", "���")
    cityProvinceArray.Add Array("���", "���")
    cityProvinceArray.Add Array("�λ�", "���")
    cityProvinceArray.Add Array("���ֽ�", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("��ȭ", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("�ȵ�", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("��õ", "���")
    cityProvinceArray.Add Array("�︪��", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("�Ǽ�", "���")
    cityProvinceArray.Add Array("û�۱�", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("��â", "���")
    cityProvinceArray.Add Array("���ؽ�", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("�о�", "���")
    cityProvinceArray.Add Array("��â��", "���")
    cityProvinceArray.Add Array("��û", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("�Ƿɱ�", "���")
    cityProvinceArray.Add Array("����", "���")
    cityProvinceArray.Add Array("â��", "���")
    cityProvinceArray.Add Array("�뿵", "���")
    cityProvinceArray.Add Array("�Ծ籺", "���")
    cityProvinceArray.Add Array("��õ", "���")
    cityProvinceArray.Add Array("���", "���ֵ�")
    cityProvinceArray.Add Array("������", "���ֵ�")
    cityProvinceArray.Add Array("����", "���ֵ�")
    cityProvinceArray.Add Array("����2", "���ֵ�")
    cityProvinceArray.Add Array("������", "���ֵ�")
    
    ' Loop through the collection to find the city and return the corresponding province
    For Each cityProvince In cityProvinceArray
        If cityProvince(0) = city Then
            GetProvince_Collection = cityProvince(1)
            Exit Function
        End If
    Next cityProvince
    
    ' If city is not found in the collection, return "Not in list"
    GetProvince_Collection = "Not in list"
End Function



Sub importRainfall()
    Dim myArray As Variant
    Dim rng As Range

    Select Case UCase(Range("T6").value)
        Case "SEJONG", "HONGSUNG"
            Exit Sub
    End Select

    Dim indexString As String
    indexString = "data_" & UCase(Range("T6").value)

    On Error Resume Next
    myArray = Application.Run(indexString)
    On Error GoTo 0

    If Not IsArray(myArray) Then
        MsgBox "An error occurred while fetching data.", vbExclamation
        Exit Sub
    End If

    Set rng = ThisWorkbook.ActiveSheet.Range("B5:N34")
    rng.value = myArray

    Range("B2").value = Range("T5").value & "���û"
End Sub

'
' 2025/03/02 �泲���� ��ư �߰�
'
'

Sub importRainfall_button(ByVal AREA As String)
    Dim myArray As Variant
    Dim rng As Range

    Dim indexString As String
    indexString = "data_" & UCase(AREA)


    Select Case UCase(AREA)
        Case "SEJONG", "HONGSUNG"
            Exit Sub
            
        Case "BORYUNG"
            Range("S5").value = "��û��"
            Range("T5").value = "����"
        
        Case "DAEJEON"
            Range("S5").value = "��û��"
            Range("T5").value = "����"
        
        Case "SEOSAN"
            Range("S5").value = "��û��"
            Range("T5").value = "����"
        
        Case "BUYEO"
            Range("S5").value = "��û��"
            Range("T5").value = "�ο�"
        
        Case "CHEONAN"
            Range("S5").value = "��û��"
            Range("T5").value = "õ��"
        
        Case "CHEONGJU"
            Range("S5").value = "��û��"
            Range("T5").value = "û��"
        
        Case "GEUMSAN"
            Range("S5").value = "��û��"
            Range("T5").value = "�ݻ�"
             
        
        Case "SEOUL"
            Range("S5").value = "������"
            Range("T5").value = "����"
            
        Case "SUWON"
            Range("S5").value = "������"
            Range("T5").value = "����"
            
        Case "INCHEON"
            Range("S5").value = "������"
            Range("T5").value = "��õ"
            
    End Select



    On Error Resume Next
    myArray = Application.Run(indexString)
    On Error GoTo 0

    If Not IsArray(myArray) Then
        MsgBox "An error occurred while fetching data.", vbExclamation
        Exit Sub
    End If

    Set rng = ThisWorkbook.ActiveSheet.Range("B5:N34")
    rng.value = myArray

    Range("B2").value = Range("T5").value & "���û"
End Sub




Function GetProvince_Case(city As String) As String
    Select Case city
        ' ��û��
        Case "����", "��õ", "û��", "��ǳ��", "����", "����", "�ݻ�", "����", "�ο�", "����", "õ��", "ȫ��"
            GetProvince_Case = "��û��"
        ' ������
        Case "���ǻ�", "����", "��ȭ", "��ɵ�", "��õ", "����õ", "����", "����", "��õ", "����"
            GetProvince_Case = "������"
        ' ������
        Case "����", "�����", "����", "�ϰ���", "����õ", "��ô", "����", "����", "����", "����", "������", "ö��", "��õ", "�¹�", "ȫõ"
            GetProvince_Case = "������"
        ' ����
        Case "����", "��â", "��â��", "����", "����", "�ξ�", "��â��", "�ӽ�", "���", "����", "����", "������", "����", "�����", "����", "����", "������", "��õ", "����", "������", "�ϵ�", "����", "�־�", "����", "÷ö��", "������", "�س�", "��굵"
            GetProvince_Case = "����"
        ' ���
        Case "�뱸", "�뱸(��)", "���", "�λ�", "���ֽ�", "����", "����", "��ȭ", "����", "�ȵ�", "����", "����", "��õ", "�︪��", "����", "�Ǽ�", "û�۱�", "����", "����", "��â", "���ؽ�", "����", "�о�", "��â��", "��û", "����", "�Ƿɱ�", "����", "â��", "�뿵", "�Ծ籺", "��õ"
            GetProvince_Case = "���"
        ' ���ֵ�
        Case "���", "������", "����", "����2", "������"
            GetProvince_Case = "���ֵ�"
        ' Default case
        Case Else
            GetProvince_Case = "Not in list"
    End Select
End Function



Sub ResetWeatherData(ByVal AREA As String)

    Dim Province As String
    
    Sheets("All").Activate
'    Range("S5") = "��û��"
'    Range("T5") = "û��"
    
    Province = GetProvince_Case(AREA)

    If CheckSubstring(Province, "Not in list") Then
        Popup_MessageBox (" Province is Not in list .... ")
        Exit Sub
    End If

    Range("S5") = Province
    Range("T5") = AREA
    
    
    Popup_MessageBox ("Clear Contents")
    Range("b5:n34").ClearContents
    
    Popup_MessageBox (" Load 30 year Weather Data ")
    Call modProvince.importRainfall
    

End Sub


Sub test()

    Call ResetWeatherData("����")

End Sub

