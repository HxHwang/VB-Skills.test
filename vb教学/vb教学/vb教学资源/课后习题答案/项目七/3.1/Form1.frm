VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   5445
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   5175
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4020
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2008-8-11"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "1:12"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "红色"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "蓝色"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7B16
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuFont 
      Caption         =   "字体"
      Begin VB.Menu MunFontSize 
         Caption         =   "大小"
         Begin VB.Menu MunFontSize1 
            Caption         =   "14"
         End
         Begin VB.Menu MunFontSize2 
            Caption         =   "16"
         End
      End
      Begin VB.Menu MunFontColor 
         Caption         =   "颜色"
         Begin VB.Menu MunFontColor1 
            Caption         =   "红色"
         End
         Begin VB.Menu MunFontColor2 
            Caption         =   "蓝色"
         End
      End
   End
   Begin VB.Menu MunLook 
      Caption         =   "查看"
      Begin VB.Menu MunTool 
         Caption         =   "工具栏"
         Checked         =   -1  'True
      End
      Begin VB.Menu MunBar 
         Caption         =   "状态栏"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MunBar_Click()
    If StatusBar1.Visible Then
        StatusBar1.Visible = False
        MunBar.Checked = False
    Else
        StatusBar1.Visible = True
        MunBar.Checked = True
    End If
End Sub

Private Sub MunFontColor1_Click()
    Text1.ForeColor = vbRed
    StatusBar1.Panels(3).Text = "红色"
End Sub

Private Sub MunFontColor2_Click()
    Text1.ForeColor = vbBlue
    StatusBar1.Panels(3).Text = "蓝色"
End Sub

Private Sub MunFontSize1_Click()
    Text1.FontSize = 14
End Sub

Private Sub MunFontSize2_Click()
    Text1.FontSize = 16
End Sub

Private Sub MunTool_Click()
    If Toolbar1.Visible Then
        Toolbar1.Visible = False
        MunTool.Checked = False
    Else
        Toolbar1.Visible = True
        MunTool.Checked = True
    End If
    
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        Call MunFontColor1_Click
    Case 2
        Call MunFontColor2_Click
  End Select
End Sub
