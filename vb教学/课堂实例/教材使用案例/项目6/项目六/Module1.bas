Attribute VB_Name = "Module1"
Public Function ChooseMobile(price As Integer, mlab As String) As String
    Select Case price
        Case 500
        Select Case mlab
            Case "摩托罗拉"
                If cboType.ListIndex = 0 Then
                    If mfunc = ", 拍照" Then
                        mobile = mlab + "L6"
                    Else
                        mobile = mlab + "c168"
                    End If
                Else
                    mobile = "无"
                End If
            Case "诺基亚"
                Select Case Form1.cboType.ListIndex
                    Case 0
                        If mfunc = ", 拍照" Then
                            mobile = mlab + "6020"
                        Else
                            mobile = mlab + "6030"
                        End If
                    Case 1
                        If mfunc = ", 拍照" Then
                            mobile = "无"
                        Else
                            mobile = mlab + "6060"
                        End If
                    Case Else
                    mobile = "无"
                End Select
            Case Else
                mobile = "进货中"
            End Select
        Case Else
            mobile = "进货中"
    End Select
    ChooseMobile = mobile
End Function

