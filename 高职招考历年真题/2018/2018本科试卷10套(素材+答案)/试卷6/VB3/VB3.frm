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
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "�ʼ��������ˣ�"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    x = Val(Text1.Text)
    If x <= 1000 Then
        s = 8
    Else
        s = 8 + (x - 1000) * 4 \ 500
        If (x - 1000) Mod 500 <> 0 Then
            s = s + 4
        End If
    End If
    MsgBox "���ʣ�" & s & "Ԫ"
End Sub
