VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   6945
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   4080
      List            =   "Form1.frx":000A
      TabIndex        =   15
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox TxtShow 
      Height          =   3255
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   840
      Width           =   4095
   End
   Begin VB.Frame Frame3 
      Caption         =   "����"
      Height          =   1695
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
      Begin VB.CheckBox ChkLove3 
         Caption         =   "�鷨"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox ChkLove5 
         Caption         =   "����"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox ChkLove4 
         Caption         =   "����"
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox ChkLove6 
         Caption         =   "����"
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox ChkLove2 
         Caption         =   "����"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox ChkLove1 
         Caption         =   "����"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�Ļ��̶�"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2295
      Begin VB.OptionButton OptSch4 
         Caption         =   "��ѧ"
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton OptSch3 
         Caption         =   "����"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton OptSch2 
         Caption         =   "����"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton OptSch1 
         Caption         =   "Сѧ"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.TextBox TxtName 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   300
      Left            =   3240
      TabIndex        =   16
      Top             =   360
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sname As String, Sman As String, Sschoole As String
Dim Slove As String

Private Sub CmdOk_Click()
Sname = TxtName.Text
Sman = Combo1.Text
If OptSch1.Value = True Then
Sschoole = "Сѧ"
ElseIf OptSch2.Value = True Then
Sschoole = "����"
ElseIf OptSch3.Value = True Then
Sschoole = "����"
Else
Sschoole = "��ѧ"
End If

Slove = ""
If ChkLove1.Value = 1 Then
Slove = Slove + " ����"
End If

If ChkLove2.Value = 1 Then
Slove = Slove + " ����"
End If

If ChkLove3.Value = 1 Then
Slove = Slove + " ����"
End If

If ChkLove4.Value = 1 Then
Slove = Slove + " �鷨"
End If

If ChkLove5.Value = 1 Then
Slove = Slove + " ����"
End If

If ChkLove6.Value = 1 Then
Slove = Slove + " ����"
End If

TxtShow.Text = TxtShow.Text + Sname + "  " + Sman + "  " + Sschoole + "  " + Slove + Chr(13) + Chr(10)


End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
End Sub

