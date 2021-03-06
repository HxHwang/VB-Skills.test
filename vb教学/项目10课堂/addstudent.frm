VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form addstudentfrm 
   Caption         =   "添加学生信息"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   8280
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5160
      Top             =   4320
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\fjlg\Desktop\项目10课堂\学生成绩信息库.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\fjlg\Desktop\项目10课堂\学生成绩信息库.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "学生信息表"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "返回"
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "添加"
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "学生信息"
      Height          =   4455
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   3735
      Begin VB.ComboBox Comboxb 
         DataField       =   "性别"
         DataSource      =   "Adodc1"
         Height          =   300
         ItemData        =   "addstudent.frx":0000
         Left            =   1200
         List            =   "addstudent.frx":000A
         TabIndex        =   10
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtage 
         DataField       =   "出生日期"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtname 
         DataField       =   "姓名"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtid 
         DataField       =   "学号"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "出生日期："
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "性别："
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "姓名："
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "学号："
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "addstudentfrm"
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

Private Sub cmdadd_Click()
Dim num, name As String
Call ChangeEnabled(True)
num = Trim(txtid.Text)
name = Trim(txtname.Text)
If Len(num) <> 5 Or IsNumeric(num) = False Then
  MsgBox "学号输入有误，请输入5位数！", vbOKOnly + vbInformation, "提示"
  txtid.SetFocus
Else
  If Trim(txtname.Text) = "" Then
    MsgBox "姓名不能为空，请重输！", vbOKOnly + vbInformation, "提示"
    txtname.SetFocus
  Else
   If cmdadd.Caption = "添加" Then
      cmdadd.Caption = "保存"
      Adodc1.Recordset.AddNew
  
   Else
      cmdadd.Caption = "添加"
      Adodc1.Recordset.Fields(0) = txtid.Text
      Adodc1.Recordset.Fields(1) = txtname.Text
      Adodc1.Recordset.Fields(2) = Comboxb.Text
      Adodc1.Recordset.Fields(3) = txtage.Text
    
      Adodc1.Recordset.Update
      MsgBox "添加成功！", vbOKOnly + vbInformation, "添加记录"
      infofrm.Show
      addstudentfrm.Hide
    End If
  End If
End If

  ' Datastudent.Recordset.MoveLast
'cmdok.Enabled = True
End Sub
Private Sub cmdexit_Click()
mainfrm.Show
addstudentfrm.Hide

End Sub



Private Sub Form_Load()
Call ChangeEnabled(False)
End Sub
