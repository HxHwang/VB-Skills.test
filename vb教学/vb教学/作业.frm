VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   5385
   StartUpPosition =   1  '����������
   Begin VB.CheckBox Check3 
      Caption         =   "����"
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Caption         =   "�Ǿ�"
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�ھ�"
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.OptionButton Option4 
      Caption         =   "�����"
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Ӣ��"
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "�¹�"
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "����"
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label6 
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label5 
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   3600
      TabIndex        =   11
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label4 
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   3000
      TabIndex        =   10
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Ԥ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Ԥ��Ĺ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   480
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
Private Sub Check1_Click()
If Check1.Value = 1 Then
Label4.Caption = "�ھ�"
Else
Label4.Caption = ""
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Label5.Caption = "�Ǿ�"
Else
Label5.Caption = ""
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Label6.Caption = "����"
Else
Label6.Caption = ""
End If
End Sub

Private Sub Option1_Click()
Label3.Caption = "�����ӱ��챭���п��ܻ�ã�"
End Sub
Private Sub Option2_Click()
Label3.Caption = "�¹��ӱ��챭���п��ܻ�ã�"
End Sub

Private Sub Option3_Click()
Label3.Caption = "Ӣ���ӱ��챭���п��ܻ�ã�"
End Sub

Private Sub Option4_Click()
Label3.Caption = "������ӱ��챭���п��ܻ�ã�"
End Sub
