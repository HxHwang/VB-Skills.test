Attribute VB_Name = "Module1"
Public Function ChooseMobile(price As Integer, mlab As String) As String
    Select Case price
        Case 500
        Select Case mlab
            Case "Ħ������"
                If cboType.ListIndex = 0 Then
                    If mfunc = ", ����" Then
                        mobile = mlab + "L6"
                    Else
                        mobile = mlab + "c168"
                    End If
                Else
                    mobile = "��"
                End If
            Case "ŵ����"
                Select Case Form1.cboType.ListIndex
                    Case 0
                        If mfunc = ", ����" Then
                            mobile = mlab + "6020"
                        Else
                            mobile = mlab + "6030"
                        End If
                    Case 1
                        If mfunc = ", ����" Then
                            mobile = "��"
                        Else
                            mobile = mlab + "6060"
                        End If
                    Case Else
                    mobile = "��"
                End Select
            Case Else
                mobile = "������"
            End Select
        Case Else
            mobile = "������"
    End Select
    ChooseMobile = mobile
End Function

