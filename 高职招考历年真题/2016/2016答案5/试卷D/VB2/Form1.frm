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
      Caption         =   "�Ƚ�"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "�������ǣ�"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "��������������"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Single, b As Single, c As Single, max As Single
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
max = a
If b > max Then
    max = b
End If
If c > max Then
    max = c
End If
Label2.Caption = "�������ǣ�" & max
End Sub
