VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9465
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   17340
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   17340
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   8970
      Width           =   17340
      _ExtentX        =   30586
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   17340
      _ExtentX        =   30586
      _ExtentY        =   741
      ButtonWidth     =   609
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
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   6480
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   10575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   18495
   End
   Begin VB.Menu Mnuxianshi 
      Caption         =   "显示"
      Begin VB.Menu Mnuyangshi 
         Caption         =   "样式"
         Begin VB.Menu Mnusongti 
            Caption         =   "宋体"
            Checked         =   -1  'True
         End
         Begin VB.Menu Mnuheiti 
            Caption         =   "黑体"
         End
         Begin VB.Menu Mnufontname 
            Caption         =   "更多字体..."
         End
      End
      Begin VB.Menu Mnufenge 
         Caption         =   "-"
      End
      Begin VB.Menu Mnudaxiao 
         Caption         =   "大小"
         Begin VB.Menu Mnufontsize16 
            Caption         =   "16"
            Checked         =   -1  'True
         End
         Begin VB.Menu Mnufontsize24 
            Caption         =   "24"
         End
      End
      Begin VB.Menu Mnufontcolor 
         Caption         =   "颜色"
         Begin VB.Menu Mnucolorred 
            Caption         =   "红色"
         End
         Begin VB.Menu Mnucolorgreen 
            Caption         =   "绿色"
         End
         Begin VB.Menu Mnucolorblue 
            Caption         =   "蓝色"
         End
         Begin VB.Menu Mnufontcolormore 
            Caption         =   "更多颜色..."
         End
      End
   End
   Begin VB.Menu Mnushuzu 
      Caption         =   "数组"
      Begin VB.Menu Mnufuzhi 
         Caption         =   "赋值"
         Begin VB.Menu Mnurandom 
            Caption         =   "随机数"
         End
         Begin VB.Menu Mnuxuhao 
            Caption         =   "序号数"
         End
      End
      Begin VB.Menu Mnusum 
         Caption         =   "求和"
      End
      Begin VB.Menu Mnumul 
         Caption         =   "求积"
      End
      Begin VB.Menu Mnusqu 
         Caption         =   "平方"
      End
   End
   Begin VB.Menu Mnuedit 
      Caption         =   "编辑"
      Visible         =   0   'False
      Begin VB.Menu Mnucopy 
         Caption         =   "复制"
         Shortcut        =   ^C
      End
      Begin VB.Menu Mnucut 
         Caption         =   "剪切"
         Shortcut        =   ^X
      End
      Begin VB.Menu Mnupaste 
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

Private Sub Form_Load()
Form1.Width = Screen.Width - 500
Form1.Height = Screen.Height - 1000
Form1.WindowState = 2
Text1.Width = Form1.Width
Text1.Height = Form1.Height - 500
StatusBar1.Panels(1).Text = Date & "       " & Time$
End Sub


Private Sub Mnucopy_Click()
Clipboard.Clear
Clipboard.SetText (Text1.SelText)
End Sub

Private Sub Mnucut_Click()
Clipboard.Clear
Clipboard.SetText (Text1.SelText)
Text1.SelText = ""
End Sub


Private Sub Mnupaste_Click()
Text1.SelText = Clipboard.GetText
End Sub

Private Sub Mnufontname_Click()
CommonDialog1.Flags = cdlCFScreenFonts
CommonDialog1.ShowFont
Text1.Font = CommonDialog1.FontName
Text1.FontSize = CommonDialog1.FontSize
Text1.Font.Bold = CommonDialog1.FontBold
Text1.Font.Italic = CommonDialog1.FontItalic


End Sub

Private Sub Mnuheiti_Click()
Text1.FontName = "黑体"
Mnuheiti.Checked = True
Mnusongti.Checked = False
StatusBar1.Panels(2).Text = "字体：黑体"
End Sub



Private Sub Mnusongti_Click()
Text1.FontName = "宋体"
Mnusongti.Checked = True
Mnuheiti.Checked = False
StatusBar1.Panels(2).Text = "字体：宋体"
End Sub




Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
  Text1.Enabled = False
  Text1.Enabled = True
  PopupMenu Mnuedit
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
  Case 1
    Call Mnucopy_Click
  Case 2
    Call Mnucut_Click
  Case 3
    Call Mnupaste_Click
End Select
End Sub
