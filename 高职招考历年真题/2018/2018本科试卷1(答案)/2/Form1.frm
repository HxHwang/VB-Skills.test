VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   3855
   End
   Begin VB.ListBox List1 
      Height          =   2040
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʾ"
      Height          =   735
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = "��ѡ����ǣ�" & List1.ListIndex + 1 & "��  �ǣ�" & List1.Text
End Sub

Private Sub Form_Load()
List1.AddItem "ˮ����"
List1.AddItem "ˮ����"
List1.AddItem "ˮ����"
List1.AddItem "ˮ����"
List1.AddItem "ˮ����"
List1.AddItem "ˮ����"
End Sub
