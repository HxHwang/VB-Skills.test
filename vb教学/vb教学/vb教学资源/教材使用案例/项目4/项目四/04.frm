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
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "��ɫ"
      Height          =   2295
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
      Begin VB.CommandButton CmdColor 
         Caption         =   "������ɫ"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton OptGreen 
         Caption         =   "��ɫ"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton OptBlue 
         Caption         =   "��ɫ"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton OptRed 
         Caption         =   "��ɫ"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "������ʽ"
      Height          =   2295
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
      Begin VB.CheckBox ChkItalic 
         Caption         =   "б��"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   735
      End
      Begin VB.CheckBox ChkUnderline 
         Caption         =   "�»���"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.CheckBox ChkBold 
         Caption         =   "����"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�����С"
      Height          =   2295
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
      Begin VB.CommandButton CmdSize 
         Caption         =   "��������"
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
'���ı����е�����Ĵ�������
    If ChkBold.Value = vbChecked Then
       Txt.FontBold = True
    Else
       Txt.FontBold = False
    End If
End Sub

Private Sub ChkItalic_Click()
'���ı����е������б������
    If ChkItalic.Value = vbChecked Then
       Txt.FontItalic = True
    Else
       Txt.FontItalic = False
    End If
End Sub

Private Sub chkunderline_Click()
'���ı����е������»�������
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
'���ı����е������С��Ϊ12��
    Txt.FontSize = 12
End Sub

Private Sub Opt16_Click()
'���ı����е������С��Ϊ16��
    Txt.FontSize = 16
End Sub

Private Sub Opt20_Click()
'���ı����е������С��Ϊ20��
    Txt.FontSize = 20
End Sub

Private Sub Form_Load()
'�����ı�������ĳ�ʼ����
    Txt.FontBold = True
    Txt.ForeColor = vbRed
    Txt.FontSize = 12

End Sub

Private Sub OptBlue_Click()
'���ı����е�������ɫ��Ϊ��ɫ
    Txt.ForeColor = vbBlue
End Sub

Private Sub OptGreen_Click()
'���ı����е�������ɫ��Ϊ��ɫ
    Txt.ForeColor = vbGreen
End Sub

Private Sub OptRed_Click()
'���ı����е�������ɫ��Ϊ��ɫ
    Txt.ForeColor = vbRed
End Sub
