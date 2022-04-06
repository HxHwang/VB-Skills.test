VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   2400
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim m As Integer, n As Integer
    Dim i As Integer, sum As Long
    m = Val(Text1.Text)
    n = Val(Text2.Text)
    sum = 0
    If m < n Then
        For i = m To n
            If i Mod 7 = 0 Then ' 如果能被7整除，那么sum就累加这个数字
                sum = sum + i
            End If
        Next i
    Else
        MsgBox "m必须小于n，也就是左边的文本框的值必须小于右边文本框的值！"
        
    End If
    Print sum
End Sub

Private Sub Form_Load()
'载入的时候，清空文本框内容
Text1.Text = ""
Text2.Text = ""
End Sub
