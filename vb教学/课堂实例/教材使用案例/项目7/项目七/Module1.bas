Attribute VB_Name = "Module1"
Public Sub TxTOut(D() As Integer)
    Dim C()     '���嶯̬��
    Dim Whole As String
    ReDim C(UBound(D))      '����̬��������ڴ�ռ�
    
    For i = 1 To UBound(D)  '���鸴��
        C(i) = D(i)
    Next i
    
    For i = 1 To UBound(D)  '���鸴��
        Whole = Whole + Str(C(i))
    Next i
    
    Form1.TxtText.Text = Form1.TxtText.Text + Whole + Chr(13) + Chr(10)
    
    
End Sub
