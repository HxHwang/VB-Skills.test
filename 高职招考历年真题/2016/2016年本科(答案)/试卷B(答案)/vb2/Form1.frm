VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "�ж���ż��"
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "��������ǣ�"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim x As Integer
    x = InputBox("������һ����������")
    If x Mod 2 = 0 Then
        Label2.Caption = "ż��"
    Else
        Label2.Caption = "����"
    End If
End Sub
