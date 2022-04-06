VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   5310
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdOk 
      Caption         =   "确定"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox TxtShow 
      Height          =   3975
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      Caption         =   "爱好"
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
      Begin VB.CheckBox ChkLove3 
         Caption         =   "书法"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox ChkLove5 
         Caption         =   "音乐"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox ChkLove4 
         Caption         =   "看书"
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox ChkLove6 
         Caption         =   "旅游"
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox ChkLove2 
         Caption         =   "篮球"
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox ChkLove1 
         Caption         =   "足球"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "文化程度"
      Height          =   1455
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   1095
      Begin VB.OptionButton OptSch3 
         Caption         =   "中职"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.OptionButton OptSch2 
         Caption         =   "高中"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton OptSch1 
         Caption         =   "初中"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "性别"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1095
      Begin VB.OptionButton OptMan2 
         Caption         =   "女"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   615
      End
      Begin VB.OptionButton OptMan1 
         Caption         =   "男"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.TextBox TxtName 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "姓名："
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
If OptMan1.Value = True Then
Sman = "男"
Else
Sman = "女"
End If

If OptSch1.Value = True Then
Sschoole = "初中"
ElseIf OptSch1.Value = True Then
Sschoole = "高中"
Else
Sschoole = "中职"
End If

Slove = ""
If ChkLove1.Value = 1 Then
Slove = Slove + " 足球"
End If

If ChkLove2.Value = 1 Then
Slove = Slove + " 篮球"
End If

If ChkLove3.Value = 1 Then
Slove = Slove + " 看书"
End If

If ChkLove4.Value = 1 Then
Slove = Slove + " 书法"
End If

If ChkLove5.Value = 1 Then
Slove = Slove + " 音乐"
End If

If ChkLove6.Value = 1 Then
Slove = Slove + " 旅游"
End If

TxtShow.Text = TxtShow.Text + Sname + "  " + Sman + "  " + Sschoole + "  " + Slove + Chr(13) + Chr(10)


End Sub
