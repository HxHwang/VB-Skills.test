VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   7305
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame2 
      Caption         =   "����ɼ�"
      Height          =   3735
      Left            =   3960
      TabIndex        =   1
      Top             =   600
      Width           =   2895
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   3
         Left            =   1080
         TabIndex        =   14
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   13
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "������"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Ӣ�"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "��ѧ��"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "���ģ�"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ͳ�Ʒ�ʽ"
      Height          =   3735
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   2655
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   3000
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�ܳɼ�"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ƽ���ɼ�"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "��߳ɼ�"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   2280
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, sum
Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
  sum = Val(Text2(0).Text)
  For i = 1 To 3
   If Val(Text2(i).Text) > sum Then
     sum = Val(Text2(i).Text)
    End If
  Next i
Case 1
  For i = 0 To 3
   sum = sum + Val(Text2(i).Text)
  Next i
  sum = sum / 4
Case 2
  For i = 0 To 3
    sum = sum + Val(Text2(i).Text)
  Next i
End Select
Label1.Caption = Option1(Index).Caption & ":"
Text1.Text = Str$(sum)
End Sub
