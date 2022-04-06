VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   Caption         =   "学生成绩查询系统"
   ClientHeight    =   6135
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   6975
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox TxtScore 
      Alignment       =   2  'Center
      DataField       =   "成绩"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   3
      Top             =   4560
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "学生成绩信息库.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "学生成绩表"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "学生成绩统计图"
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   6975
      Begin VB.PictureBox PicScore 
         BackColor       =   &H8000000E&
         Height          =   4455
         Left            =   120
         ScaleHeight     =   4395
         ScaleWidth      =   6555
         TabIndex        =   2
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "100"
         Height          =   180
         Left            =   6600
         TabIndex        =   9
         Top             =   4800
         Width           =   270
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "90"
         Height          =   180
         Left            =   5286
         TabIndex        =   8
         Top             =   4800
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "80"
         Height          =   180
         Left            =   3972
         TabIndex        =   7
         Top             =   4800
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "70"
         Height          =   180
         Left            =   2658
         TabIndex        =   6
         Top             =   4800
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "60"
         Height          =   180
         Left            =   1344
         TabIndex        =   5
         Top             =   4800
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   4800
         Width           =   90
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6480
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0000
            Key             =   "login"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":031C
            Key             =   "logout"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0638
            Key             =   "Reg"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0954
            Key             =   "mng"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0C70
            Key             =   "borrow"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0F8C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   930
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1640
      ButtonWidth     =   1773
      ButtonHeight    =   1482
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "登录"
            Key             =   "login"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "注册"
            Key             =   "reg"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "注销"
            Key             =   "logout"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "系统员管理"
            Key             =   "mng"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "用户查询"
            Key             =   "borrow"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "成绩统计图"
            Key             =   "graphic"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileLog 
         Caption         =   "登录(&I)..."
      End
      Begin VB.Menu mnuReg 
         Caption         =   "注册(&R)..."
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "注销(&O)"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuGen 
      Caption         =   "功能(&G)"
      Begin VB.Menu mnuMng 
         Caption         =   "系统员管理(&M)"
      End
      Begin VB.Menu Mnustudent 
         Caption         =   "显示学生信息"
      End
      Begin VB.Menu Mnuaddstudent 
         Caption         =   "添加学生信息"
      End
      Begin VB.Menu mnuBorrow 
         Caption         =   "成绩查询(&B)"
      End
      Begin VB.Menu Show 
         Caption         =   "显示学生成绩统计图(&S)"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuAboutit 
         Caption         =   "关于 学生成绩查询系统(&A)"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Score() As Single
Dim i As Integer
Dim NumScore(4) As Integer
Dim NumMax As Integer
    
Private Sub Graph()
    PicScore.Scale (0, NumMax + 1)-(5, 0)
    PicScore.FillStyle = 0
    For i = 0 To 4
        PicScore.FillColor = QBColor(i)
        PicScore.Line (i, NumScore(i))-(i + 1, 0), , B
    Next i
End Sub


Private Sub mnuAboutit_Click()
    MsgBox "       项目十" + Chr$(13) + Chr$(10) + "——基于数据库的学生成绩查询系统" + Chr$(13) + Chr$(10) + "             2008.4", 0, "关于学生成绩管理系统"
End Sub

Private Sub Mnuaddstudent_Click()
Unload Main
infofrm.Show
End Sub

Private Sub mnuBorrow_Click()
    If Admin = False Then
        Unload Main
        Search.Show
    Else
        MsgBox "请先登录系统！", vbInformation, "登录"
    End If
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuFileLog_Click()
    Unload Main
    Login.Show
End Sub

Private Sub mnuLogout_Click()
    If MsgBox("是否真的注销当前用户？", vbCritical + vbYesNo, "注销") = vbYes Then
        End
    End If
End Sub

Private Sub mnuMng_Click()
    If Admin = True Then
        Unload Main
        UserManage.Show
    Else
        MsgBox "请以管理员身份登录系统！", vbInformation, "登录"
    End If
End Sub

Private Sub mnuReg_Click()
    Reg = 2
    Unload Main
    Register.Show
End Sub

Private Sub Show_Click()

    i = 0
    Data1.Recordset.MoveFirst
    Do While Data1.Recordset.EOF = False
        i = i + 1
        Data1.Recordset.MoveNext
    Loop
    ReDim Score(i) As Single
    
    For i = 0 To 4
        NumScore(i) = 0
    Next i
    
    i = 0
    Data1.Recordset.MoveFirst
    Do While Data1.Recordset.EOF = False
        i = i + 1
        Data1.Recordset.MoveNext
        Score(i) = Val(txtscore.Text)
        Select Case Score(i)
            Case Is < 60
                NumScore(0) = NumScore(0) + 1
            Case Is >= 90
                NumScore(4) = NumScore(4) + 1
            Case Is > 80
                NumScore(3) = NumScore(3) + 1
            Case Is >= 70
                NumScore(2) = NumScore(2) + 1
            Case Is >= 60
                NumScore(1) = NumScore(1) + 1
        End Select
    Loop
    
    NumMax = 0
    For i = 0 To 4
        If NumScore(i) > NumMax Then
            NumMax = NumScore(i)
        End If
    Next i
    Call Graph
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "login"
        mnuFileLog_Click
    Case "logout"
        mnuLogout_Click
    Case "reg"
        mnuReg_Click
    Case "mng"
        mnuMng_Click
    Case "borrow"
        mnuBorrow_Click
    Case "graphic"
        Show_Click
    End Select
End Sub
