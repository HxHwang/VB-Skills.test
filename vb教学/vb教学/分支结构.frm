VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "�ж�"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "������ɼ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = Val(Text1.Text)
If a >= 85 Then
Label2.Caption = "��ĳɼ����㣡"
ElseIf a >= 70 Then
Label2.Caption = "��ĳɼ����ã�"
ElseIf a >= 60 Then
Label2.Caption = "��ĳɼ��ϸ�"
Else
Label2.Caption = "��ĳɼ����ϸ�"
End If
End Sub

