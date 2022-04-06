Attribute VB_Name = "Module1"
Public Sub TxTOut(D() As Integer)
    Dim C()     '定义动态数
    Dim Whole As String
    ReDim C(UBound(D))      '给动态数组分配内存空间
    
    For i = 1 To UBound(D)  '数组复制
        C(i) = D(i)
    Next i
    
    For i = 1 To UBound(D)  '数组复制
        Whole = Whole + Str(C(i))
    Next i
    
    Form1.TxtText.Text = Form1.TxtText.Text + Whole + Chr(13) + Chr(10)
    
    
End Sub
