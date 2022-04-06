VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "数组操作器"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   7170
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5220
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2008-6-1"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:14"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "复制"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "剪切"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "粘贴"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0224
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtText 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   480
      Width           =   7095
   End
   Begin VB.Menu MnuFont 
      Caption         =   "显示"
      Begin VB.Menu MnuFontStyle 
         Caption         =   "样式"
         Begin VB.Menu MnuFontStyle1 
            Caption         =   "宋体"
            Checked         =   -1  'True
         End
         Begin VB.Menu MnuFontStyle2 
            Caption         =   "黑体"
         End
      End
      Begin VB.Menu MnuFontSeprator 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFontSize 
         Caption         =   "大小"
         Begin VB.Menu MnuFontSize1 
            Caption         =   "16"
            Checked         =   -1  'True
         End
         Begin VB.Menu MnuFontSize2 
            Caption         =   "24"
         End
      End
      Begin VB.Menu MnuColor 
         Caption         =   "颜色"
         Begin VB.Menu MnuColorRed 
            Caption         =   "红色"
            Checked         =   -1  'True
         End
         Begin VB.Menu MnuColorBlue 
            Caption         =   "蓝色"
         End
         Begin VB.Menu MnuColorGreen 
            Caption         =   "绿色"
         End
      End
   End
   Begin VB.Menu MnuArray 
      Caption         =   "数组"
      Begin VB.Menu MnuIni 
         Caption         =   "赋值"
         Begin VB.Menu MnuRnd 
            Caption         =   "随机数(&R)"
         End
         Begin VB.Menu MnuNum 
            Caption         =   "序号数(&N)"
         End
      End
      Begin VB.Menu MnuAdd 
         Caption         =   "求和"
      End
      Begin VB.Menu MnuMul 
         Caption         =   "求积"
      End
      Begin VB.Menu MnuSqu 
         Caption         =   "平方"
      End
   End
   Begin VB.Menu MnuEdit1 
      Caption         =   "编辑"
      Visible         =   0   'False
      Begin VB.Menu MnuEditCopy 
         Caption         =   "复制"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuEditCut 
         Caption         =   "剪切"
         Shortcut        =   ^X
      End
      Begin VB.Menu MnuEditPaste 
         Caption         =   "粘贴"
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A(1 To 10) As Integer   '数组定义
Dim i As Integer

Private Sub MnuRnd_Click()
    For i = 1 To 10           '数组赋初值为0~100内的整数
        A(i) = Int(Rnd * 100)
    Next i
    '显示输出
    TxtText.Text = TxtText.Text + "随机数:"
    Call TxTOut(A())

End Sub

Private Sub MnuNum_Click()
    For i = 1 To 10     '数组赋初值
        A(i) = i
    Next i
    '显示输出
    TxtText.Text = TxtText.Text + "序数:"
    Call TxTOut(A())
End Sub

Private Sub MnuAdd_Click()
    Dim sum As Integer
    sum = 0
    For i = 1 To 10 '数组求和
        sum = sum + A(i)
    Next i
    TxtText.Text = TxtText.Text + "和=" + Str(sum) + Chr$(13) + Chr$(10)
End Sub

Private Sub MnuMul_Click()
    Dim Mul As Double
    Mul = 1
    For Each X In A '使用For Each…Next数组求积
        Mul = Mul * X
    Next X
    TxtText.Text = TxtText.Text + "积=" + Str(Mul) + Chr$(13) + Chr$(10)
End Sub

Private Sub MnuSqu_Click()
    Dim B(10) As Integer
    For i = 1 To 10  '数组复制
        B(i) = A(i) ^ 2
    Next i
    '显示输出平方值
    TxtText.Text = TxtText.Text + "平方:"
    Call TxTOut(B())
End Sub

Private Sub mnuFontStyle1_Click()
     '单击【宋体】菜单，文本框中文字的样式为宋体
     TxtText.FontName = "宋体"
     '单击【宋体】菜单，在其前面添加选中符号
     MnuFontStyle1.Checked = True
     '单击【宋体】菜单，去掉【黑体】菜单前面的选中符号
     MnuFontStyle2.Checked = False
     
     StatusBar1.Panels(3).Text = "宋体"
End Sub

Private Sub mnuFontStyle2_Click()
     '单击【黑体】菜单，文本框中文字的样式为黑体
     TxtText.FontName = "黑体"
     '单击【黑体】菜单，去掉【宋体】菜单前面的选中符号
     MnuFontStyle1.Checked = False
     '单击【黑体】菜单，在其前面添加选中符号
     MnuFontStyle2.Checked = True
     
     StatusBar1.Panels(3).Text = "黑体"
End Sub

Private Sub mnuFontSize1_Click()
     '单击【16】菜单，文本框中文字的大小为16
     TxtText.FontSize = 16
     MnuFontSize1.Checked = True
     MnuFontSize2.Checked = False
End Sub

Private Sub mnuFontSize2_Click()
     '单击【24】菜单，文本框中文字的大小为24
     TxtText.FontSize = 24
     MnuFontSize1.Checked = False
     MnuFontSize2.Checked = True
End Sub

Private Sub mnuColorRed_Click()
    '将字体变为红色
    TxtText.ForeColor = vbRed
    MnuColorRed.Checked = True
    MnuColorBlue.Checked = False
    MnuColorGreen.Checked = False
End Sub

Private Sub mnuColorBlue_Click()
    '将字体变为蓝色
    TxtText.ForeColor = vbBlue
    MnuColorRed.Checked = False
    MnuColorBlue.Checked = True
    MnuColorGreen.Checked = False
    
End Sub

Private Sub mnuColorGreen_Click()
    '将字体变为绿色
    TxtText.ForeColor = vbGreen
    MnuColorRed.Checked = False
    MnuColorBlue.Checked = False
    MnuColorGreen.Checked = True
End Sub

Private Sub txtText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '判断单击的是否是鼠标右键
    If Button = 2 Then
    '单击鼠标右键，显示【颜色】菜单的子菜单
        TxtText.Enabled = False
        TxtText.Enabled = True
        PopupMenu MnuEdit1
    End If
End Sub

Private Sub MnuEditCopy_Click()
    Clipboard.Clear
    Clipboard.SetText (TxtText.SelText)
End Sub

Private Sub MnuEditCut_Click()
    Clipboard.Clear
    Clipboard.SetText (TxtText.SelText)
    TxtText.SelText = ""
End Sub

Private Sub MnuEditPaste_Click()
    TxtText.SelText = Clipboard.GetText
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        Call MnuEditCopy_Click
    Case 2
        Call MnuEditCut_Click
    Case 3
        Call MnuEditPaste_Click
  End Select
End Sub

