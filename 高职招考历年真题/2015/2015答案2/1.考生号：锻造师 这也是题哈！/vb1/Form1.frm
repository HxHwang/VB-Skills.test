VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�ַ�����"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ַ�����"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim n As Integer
n = Len(Text1.Text)
Text2.Text = Val(n)
End Sub

Private Sub Command2_Click()
    End
End Sub
