VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�Ի������"
   ClientHeight    =   1965
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtPass 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "ȷ��"
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���������룺"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1080
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Form_Load()
    i = 3
End Sub

Private Sub OKButton_Click()
    Dim mybt As Integer
    i = i - 1
    If TxtPass.Text = "12345" Then
        Form1.Show
        Unload Dialog
    Else
        If i = 0 Then
            MsgBox "3�����������������Ȩʹ�ü�������", vbOKOnly + vbCritical, _
            "���棡"
            Unload Dialog
        Else
            mybt = MsgBox("�����������,�㻹��" + Str(i) + "�λ���", vbOKCancel _
            + vbCritical, "�������")
            If mybt = 1 Then
                Dialog.Show
            Else
                Unload Dialog
            End If
        End If
    End If
End Sub

