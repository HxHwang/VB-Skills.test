VERSION 5.00
Begin VB.Form Text1 
   Caption         =   "Form1"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   735
      Left            =   2520
      TabIndex        =   7
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ж�������"
      Height          =   735
      Left            =   600
      TabIndex        =   6
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "c="
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "b="
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "a="
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Text1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Integer, b As Integer, c As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
If a + b > c And a + c > b And b + c > a Then
    Print "���Թ���������"
Else
    Print "���ܹ���������"
End If
End Sub

Private Sub Command2_Click()
    End
End Sub
