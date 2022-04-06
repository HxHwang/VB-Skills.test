VERSION 5.00
Begin VB.Form editfrm 
   Caption         =   "修改学生信息"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   8280
   StartUpPosition =   3  '窗口缺省
   Begin VB.Data Datastudent 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "F:\hgy\教案\教案08-09\学生成绩管理系统\student.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "studentinfo"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "返回"
      Height          =   495
      Left            =   7080
      TabIndex        =   17
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "查找"
      Height          =   495
      Left            =   7080
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "删除"
      Height          =   495
      Left            =   7080
      TabIndex        =   15
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "修改"
      Height          =   495
      Left            =   7080
      TabIndex        =   14
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "学生信息"
      Height          =   4455
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6495
      Begin VB.ComboBox Comboxb 
         DataField       =   "性别"
         DataSource      =   "Datastudent"
         Height          =   300
         ItemData        =   "editfrm.frx":0000
         Left            =   1200
         List            =   "editfrm.frx":000A
         TabIndex        =   18
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txthome 
         DataField       =   "生源地"
         DataSource      =   "Datastudent"
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtbj 
         DataField       =   "班级"
         DataSource      =   "Datastudent"
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtnj 
         DataField       =   "年级"
         DataSource      =   "Datastudent"
         Height          =   375
         Left            =   4560
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtage 
         DataField       =   "出生日期"
         DataSource      =   "Datastudent"
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtname 
         DataField       =   "姓名"
         DataSource      =   "Datastudent"
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtid 
         DataField       =   "学号"
         DataSource      =   "Datastudent"
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "出生日期："
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "性别："
         Height          =   495
         Left            =   480
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "姓名："
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "学号："
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "生源地:"
         Height          =   615
         Left            =   3840
         TabIndex        =   3
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "班级:"
         Height          =   495
         Left            =   3840
         TabIndex        =   2
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "年级："
         Height          =   495
         Left            =   3840
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "editfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChangeEnabled(choose As Boolean)
Dim i As Integer
txtid.Enabled = choose
txtname.Enabled = choose
Comboxb.Enabled = choose
txtage.Enabled = choose
txtnj.Enabled = choose
txtbj.Enabled = choose
txthome.Enabled = choose

End Sub



Private Sub cmddelete_Click()
Datastudent.Recordset.Delete

End Sub

Private Sub cmdedit_Click()
If cmdedit.Caption = "修改" Then
  cmdedit.Caption = "保存"
  Datastudent.Recordset.Edit
  Call ChangeEnabled(True)
Else
  Datastudent.Recordset.Update
  MsgBox "该学生信息已修改！", vbOKOnly + vbInformation, "提示"
  Call ChangeEnabled(False)
End If

End Sub

Private Sub cmdexit_Click()
editfrm.Hide
mainfrm.Show
End Sub

Private Sub cmdfind_Click()
findfrm.Show
editfrm.Hide
End Sub


Private Sub Form_Load()
Call ChangeEnabled(False)
End Sub
