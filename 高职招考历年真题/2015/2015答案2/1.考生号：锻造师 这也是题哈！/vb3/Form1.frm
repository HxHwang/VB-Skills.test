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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "����"
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
            If i Mod 7 = 0 Then ' ����ܱ�7��������ôsum���ۼ��������
                sum = sum + i
            End If
        Next i
    Else
        MsgBox "m����С��n��Ҳ������ߵ��ı����ֵ����С���ұ��ı����ֵ��"
        
    End If
    Print sum
End Sub

Private Sub Form_Load()
'�����ʱ������ı�������
Text1.Text = ""
Text2.Text = ""
End Sub
