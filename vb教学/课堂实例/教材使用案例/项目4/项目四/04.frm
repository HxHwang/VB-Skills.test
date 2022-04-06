VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   4500
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "颜色"
      Height          =   2295
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
      Begin VB.CommandButton CmdColor 
         Caption         =   "更多颜色"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton OptGreen 
         Caption         =   "绿色"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton OptBlue 
         Caption         =   "蓝色"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton OptRed 
         Caption         =   "红色"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "字体样式"
      Height          =   2295
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
      Begin VB.CheckBox ChkItalic 
         Caption         =   "斜体"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   735
      End
      Begin VB.CheckBox ChkUnderline 
         Caption         =   "下划线"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.CheckBox ChkBold 
         Caption         =   "粗线"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "字体大小"
      Height          =   2295
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
      Begin VB.CommandButton CmdSize 
         Caption         =   "更多字体"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton Opt20 
         Caption         =   "20"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   495
      End
      Begin VB.OptionButton Opt16 
         Caption         =   "16"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   735
      End
      Begin VB.OptionButton Opt12 
         Caption         =   "12"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.TextBox Txt 
      Height          =   1095
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "04.frx":0000
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkbold_Click()
'将文本框中的字体的粗体特性
    If ChkBold.Value = vbChecked Then
       Txt.FontBold = True
    Else
       Txt.FontBold = False
    End If
End Sub

Private Sub ChkItalic_Click()
'将文本框中的字体的斜体特性
    If ChkItalic.Value = vbChecked Then
       Txt.FontItalic = True
    Else
       Txt.FontItalic = False
    End If
End Sub

Private Sub chkunderline_Click()
'将文本框中的字体下划线特性
    If ChkUnderline.Value = vbChecked Then
       Txt.FontUnderline = True
    Else
       Txt.FontUnderline = False
    End If
End Sub

Private Sub Cmd_Click()


End Sub

Private Sub Command1_Click()


End Sub

Private Sub CmdColor_Click()
    CommonDialog1.Flags = &H1&
    CommonDialog1.ShowColor
    Txt.ForeColor = CommonDialog1.Color
End Sub

Private Sub CmdSize_Click()
    CommonDialog1.Flags = cdlCFScreenFonts
    CommonDialog1.ShowFont
    Txt.Font = CommonDialog1.FontName
    Txt.FontSize = CommonDialog1.FontSize
End Sub

Private Sub Opt12_Click()
'将文本框中的字体大小变为12号
    Txt.FontSize = 12
End Sub

Private Sub Opt16_Click()
'将文本框中的字体大小变为16号
    Txt.FontSize = 16
End Sub

Private Sub Opt20_Click()
'将文本框中的字体大小变为20号
    Txt.FontSize = 20
End Sub

Private Sub Form_Load()
'设置文本框字体的初始特性
    Txt.FontBold = True
    Txt.ForeColor = vbRed
    Txt.FontSize = 12

End Sub

Private Sub OptBlue_Click()
'将文本框中的字体颜色变为蓝色
    Txt.ForeColor = vbBlue
End Sub

Private Sub OptGreen_Click()
'将文本框中的字体颜色变为绿色
    Txt.ForeColor = vbGreen
End Sub

Private Sub OptRed_Click()
'将文本框中的字体颜色变为红色
    Txt.ForeColor = vbRed
End Sub
