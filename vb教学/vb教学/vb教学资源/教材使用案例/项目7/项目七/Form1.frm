VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "数组操作器"
   ClientHeight    =   5595
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   7170
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   5100
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   5070
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.TextBox Text1 
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
Dim i As Integer
Dim a(9) As Integer
Dim sum As Integer
Dim mul As Double


Private Sub Form_Load()
Randomize
StatusBar2.Panels(1).Text = Date
StatusBar2.Panels(2).Text = Time$
End Sub

Private Sub MnuAdd_Click()
For i = 0 To 9
  sum = sum + a(i)
Next i
Text1.Text = Text1.Text & "和=" & sum & vbCrLf

End Sub

Private Sub MnuMul_Click()
mul = 1
For i = 0 To 9
  mul = mul * a(i)
Next i
Text1.Text = Text1.Text & "积=" & mul & vbCrLf
End Sub

Private Sub MnuNum_Click()

Text1.Text = Text1.Text & "序号数："
For i = 0 To 9
  Text1.Text = Text1.Text & (i + 1) & "  "
  a(i) = i + 1
Next i
Text1.Text = Text1.Text & vbCrLf
End Sub

Private Sub MnuRnd_Click()
Text1.Text = Text1.Text & "随机数："
For i = 0 To 9
  a(i) = CInt(Rnd * 100)
  Text1.Text = Text1.Text & a(i) & "  "
Next i
Text1.Text = Text1.Text & vbCrLf
End Sub

Private Sub MnuSqu_Click()
Text1.Text = Text1.Text & "平方数："
For i = 0 To 9
  a(i) = a(i) * a(i)
  Text1.Text = Text1.Text & a(i) & "  "
Next i
Text1.Text = Text1.Text & vbCrLf
End Sub

