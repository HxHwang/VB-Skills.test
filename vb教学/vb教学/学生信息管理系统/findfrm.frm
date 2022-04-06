VERSION 5.00
Begin VB.Form findfrm 
   Caption         =   "Form1"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   8085
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtfind 
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "选择查询条件"
      Height          =   2055
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2535
      Begin VB.OptionButton optname 
         Caption         =   "按姓名"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton optid 
         Caption         =   "按学号"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
   End
End
Attribute VB_Name = "findfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim ans As String
If Trim(txtfind.Text) = "" Then
  MsgBox "请先输入查询条件！", vbOKOnly + vbCritical, "查询"
Else
  If optid.Value = True Then
   editfrm.Datastudent.Recordset.FindFirst " 学号='" & Trim(txtfind.Text) & "'"
  Else
   editfrm.Datastudent.Recordset.FindFirst "姓名='" & Trim(txtfind.Text) & "'"
  End If
  If editfrm.Datastudent.Recordset.NoMatch Then
    MsgBox "没有找到该同学！", vbOKOnly + vbInformation, "查询结果"
    txtfind.Text = ""
  Else
    ans = MsgBox("已找到该同学！", vbYesNo + vbQuestion, "查询结果")
      'If ans = vbYes Then
       'editfrm.cmdedit.Enabled = True
      'Else
       'editfrm.cmdedit.Enabled = False
      'End If
    editfrm.Show
    findfrm.Hide
  End If
End If

   
End Sub

Private Sub Command2_Click()
findfrm.Hide
mainfrm.Show
End Sub
