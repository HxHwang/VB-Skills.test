VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6330
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "新建"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "打开"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "保存"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "剪切"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "复制"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "粘贴"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "字体"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtText 
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   480
      Width           =   6255
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2820
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2004-2-4"
            Object.ToolTipText     =   "系统日期"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:13"
            Object.ToolTipText     =   "系统时间"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "文本的字体"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":099A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":131C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":250C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2EA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3908
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileNew 
         Caption         =   "新建(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSeprator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "打开(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "另存为(&A)"
      End
      Begin VB.Menu mnuSeprator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "退出(&Q)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditCut 
         Caption         =   "剪切(&C)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "复制(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "粘贴(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSeprator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAll 
         Caption         =   "全选"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuSet 
      Caption         =   "设置(&U)"
      Begin VB.Menu mnuSetFont 
         Caption         =   "字体…"
      End
      Begin VB.Menu mnuSetColor 
         Caption         =   "颜色…"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
'初始化窗体的位置及大小
With Form1
.Left = 0
.Top = 0
.Width = Screen.Width
.Height = Screen.Height - 400
End With
'初始化文本框的位置及大小
With txtText
.Left = Form1.ScaleLeft
.Top = Form1.ScaleTop + 450
.Width = Form1.ScaleWidth
.Height = Form1.ScaleHeight - 800
End With
Form1.Caption = "Untitled"
'菜单【编辑】的子菜单不可用
mnuEditCut.Enabled = False
mnuEditCopy.Enabled = False
mnuEditPaste.Enabled = False
mnuEditAll.Enabled = False
End Sub

Private Sub mnuEditAll_Click()
'文本框中的内容全被选中
txtText.SelStart = 0
txtText.SelLength = Len(txtText.text)
End Sub
Private Sub mnuEditCopy_Click()
'复制文本框中被选中的内容
Clipboard.SetText txtText.SelText
End Sub
Private Sub mnuEditCut_Click()
'剪切文本框中被选中的内容
Clipboard.SetText txtText.SelText
txtText.SelText = ""
End Sub
Private Sub mnuEditPaste_Click()
'粘贴被复制或被剪切的内容
txtText.SelText = Clipboard.GetText()
End Sub
Private Sub mnuFileNew_Click()
NewFile
End Sub
Private Sub mnuFileOpen_Click()
OpenFile
End Sub
Private Sub mnuFileQuit_Click()
Dim s As Integer
'显示退出消息框
s = MsgBox("是否保存文件?", vbYesNoCancel + vbInformation, "退出")
'根据所单击的按钮执行不同的操作
Select Case s
'单击按钮“是”，保存文件，然后退出程序
Case vbYes
mnuFileSave_Click
GoTo ss
'单击按钮“否”，不保存文件，直接退出程序
Case vbNo
GoTo ss
'单击按钮“取消”，回到主程序
Case vbCancel
Exit Sub
End Select
'退出程序
ss:
Unload Form1
End Sub

Private Sub mnuFileSave_Click()
SaveFile
End Sub
Private Sub mnuFileSaveAs_Click()
SaveAsFile
End Sub
Private Sub mnuSetColor_Click()
On Error GoTo ErrHandler
CommonDialog1.CancelError = True
'设置默认的颜色对话框
CommonDialog1.Flags = &H1&
'显示颜色对话框
CommonDialog1.ShowColor
'改变文本框中字体的颜色
txtText.ForeColor = CommonDialog1.Color
ErrHandler:
Exit Sub
End Sub
Private Sub mnuSetFont_Click()
On Error GoTo ErrHandler
CommonDialog1.CancelError = True
CommonDialog1.Flags = cdlCFBoth
'显示字体对话框
CommonDialog1.ShowFont
'将字体对话框中所选的字体赋给文本框与字体有关的属性
txtText.Font.Name = CommonDialog1.FontName
txtText.Font.Size = CommonDialog1.FontSize
txtText.Font.Bold = CommonDialog1.FontBold
txtText.Font.Italic = CommonDialog1.FontItalic
'在状态栏的第3个窗格中显示所选字体的名称
StatusBar1.Panels.Item(3).text = "字体：" + CommonDialog1.FontName
ErrHandler:
Exit Sub
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
mnuFileNew_Click
Case 2
mnuFileOpen_Click
Case 3
mnuFileSave_Click
Case 4
mnuEditCut_Click
Case 5
mnuEditCopy_Click
Case 6
mnuEditPaste_Click
Case 7
mnuSetFont_Click
End Select
End Sub
Private Sub txtText_Change()
dirty = True
'菜单【编辑】的子菜单可用
mnuEditCut.Enabled = True
mnuEditCopy.Enabled = True
mnuEditPaste.Enabled = True
mnuEditAll.Enabled = True
End Sub

