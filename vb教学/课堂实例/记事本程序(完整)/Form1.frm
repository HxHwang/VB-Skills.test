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
   StartUpPosition =   3  '����ȱʡ
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
            Object.ToolTipText     =   "�½�"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "��"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "����"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "����"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "����"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "ճ��"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "����"
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
            Object.ToolTipText     =   "ϵͳ����"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:13"
            Object.ToolTipText     =   "ϵͳʱ��"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ı�������"
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
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileNew 
         Caption         =   "�½�(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSeprator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "��(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "���Ϊ(&A)"
      End
      Begin VB.Menu mnuSeprator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "�˳�(&Q)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditCut 
         Caption         =   "����(&C)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "����(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "ճ��(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSeprator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAll 
         Caption         =   "ȫѡ"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuSet 
      Caption         =   "����(&U)"
      Begin VB.Menu mnuSetFont 
         Caption         =   "���塭"
      End
      Begin VB.Menu mnuSetColor 
         Caption         =   "��ɫ��"
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
'��ʼ�������λ�ü���С
With Form1
.Left = 0
.Top = 0
.Width = Screen.Width
.Height = Screen.Height - 400
End With
'��ʼ���ı����λ�ü���С
With txtText
.Left = Form1.ScaleLeft
.Top = Form1.ScaleTop + 450
.Width = Form1.ScaleWidth
.Height = Form1.ScaleHeight - 800
End With
Form1.Caption = "Untitled"
'�˵����༭�����Ӳ˵�������
mnuEditCut.Enabled = False
mnuEditCopy.Enabled = False
mnuEditPaste.Enabled = False
mnuEditAll.Enabled = False
End Sub

Private Sub mnuEditAll_Click()
'�ı����е�����ȫ��ѡ��
txtText.SelStart = 0
txtText.SelLength = Len(txtText.text)
End Sub
Private Sub mnuEditCopy_Click()
'�����ı����б�ѡ�е�����
Clipboard.SetText txtText.SelText
End Sub
Private Sub mnuEditCut_Click()
'�����ı����б�ѡ�е�����
Clipboard.SetText txtText.SelText
txtText.SelText = ""
End Sub
Private Sub mnuEditPaste_Click()
'ճ�������ƻ򱻼��е�����
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
'��ʾ�˳���Ϣ��
s = MsgBox("�Ƿ񱣴��ļ�?", vbYesNoCancel + vbInformation, "�˳�")
'�����������İ�ťִ�в�ͬ�Ĳ���
Select Case s
'������ť���ǡ��������ļ���Ȼ���˳�����
Case vbYes
mnuFileSave_Click
GoTo ss
'������ť���񡱣��������ļ���ֱ���˳�����
Case vbNo
GoTo ss
'������ť��ȡ�������ص�������
Case vbCancel
Exit Sub
End Select
'�˳�����
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
'����Ĭ�ϵ���ɫ�Ի���
CommonDialog1.Flags = &H1&
'��ʾ��ɫ�Ի���
CommonDialog1.ShowColor
'�ı��ı������������ɫ
txtText.ForeColor = CommonDialog1.Color
ErrHandler:
Exit Sub
End Sub
Private Sub mnuSetFont_Click()
On Error GoTo ErrHandler
CommonDialog1.CancelError = True
CommonDialog1.Flags = cdlCFBoth
'��ʾ����Ի���
CommonDialog1.ShowFont
'������Ի�������ѡ�����帳���ı����������йص�����
txtText.Font.Name = CommonDialog1.FontName
txtText.Font.Size = CommonDialog1.FontSize
txtText.Font.Bold = CommonDialog1.FontBold
txtText.Font.Italic = CommonDialog1.FontItalic
'��״̬���ĵ�3����������ʾ��ѡ���������
StatusBar1.Panels.Item(3).text = "���壺" + CommonDialog1.FontName
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
'�˵����༭�����Ӳ˵�����
mnuEditCut.Enabled = True
mnuEditCopy.Enabled = True
mnuEditPaste.Enabled = True
mnuEditAll.Enabled = True
End Sub

