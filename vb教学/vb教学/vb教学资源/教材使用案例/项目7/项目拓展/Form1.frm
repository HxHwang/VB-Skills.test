VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "���������"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   7200
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5220
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2008-8-12"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "1:07"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   873
      ButtonWidth     =   767
      ButtonHeight    =   714
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�½�"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "��"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "����"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "����"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "����"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "ճ��"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "����"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":05E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":06F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":080A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":160A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtText 
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.Menu MnuFile 
      Caption         =   "�ļ�"
      Begin VB.Menu MnuFileNew 
         Caption         =   "�½�(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuSeprator1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFileOpen 
         Caption         =   "��(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuFileSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuFileSaveAs 
         Caption         =   "���Ϊ(&A)"
      End
      Begin VB.Menu MnuSeprator2 
         Caption         =   "-"
      End
      Begin VB.Menu MunFileQuit 
         Caption         =   "�˳�(&Q)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu MnuEditCut 
         Caption         =   "����(&X)"
         Shortcut        =   ^X
      End
      Begin VB.Menu MnuEditCopy 
         Caption         =   "����(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuEditPaste 
         Caption         =   "ճ��(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu MnuSeprator3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEditAll 
         Caption         =   "ȫѡ"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu MnuSet 
      Caption         =   "����(&U)"
      Begin VB.Menu MnuSetFont 
         Caption         =   "����..."
      End
      Begin VB.Menu MnuSetColor 
         Caption         =   "��ɫ..."
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
     With TxtText
        .Left = Form1.ScaleLeft
        .Top = Form1.ScaleTop + 450
        .Width = Form1.ScaleWidth
        .Height = Form1.ScaleHeight - 800
     End With
     Form1.Caption = "Untitled"
     '���༭���˵����Ӳ˵�������
     MnuEditCut.Enabled = False
     MnuEditCopy.Enabled = False
     MnuEditPaste.Enabled = False
     MnuEditAll.Enabled = False
End Sub

Private Sub mnuEditCopy_Click()
     '�����ı����б�ѡ�е�����
     Clipboard.SetText TxtText.SelText
End Sub

Private Sub mnuEditCut_Click()
     '�����ı����б�ѡ�е�����
     Clipboard.SetText TxtText.SelText
     TxtText.SelText = ""
End Sub

Private Sub mnuEditPaste_Click()
     'ճ�������ƻ򱻼��е�����
     TxtText.SelText = Clipboard.GetText()
End Sub

Private Sub mnuEditAll_Click()
     '�ı����е�����ȫ��ѡ��
     TxtText.SelStart = 0
     TxtText.SelLength = Len(TxtText.Text)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
     Select Case Button.Index
     Case 4
         mnuEditCut_Click
     Case 5
         mnuEditCopy_Click
     Case 6
         mnuEditPaste_Click
     End Select
End Sub

Private Sub txtText_Change()
     '���༭���˵����Ӳ˵�����
     MnuEditCut.Enabled = True
     MnuEditCopy.Enabled = True
     MnuEditPaste.Enabled = True
     MnuEditAll.Enabled = True
End Sub

Private Sub txtText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 2 Then
     '��������Ҽ�����ʾ����ɫ���˵����Ӳ˵�
            TxtText.Enabled = False
            TxtText.Enabled = True
            PopupMenu MnuEdit
     End If
End Sub

