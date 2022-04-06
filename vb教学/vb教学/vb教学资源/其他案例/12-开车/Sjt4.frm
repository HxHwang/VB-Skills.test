VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   3555
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton C2 
      Caption         =   "停止"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton C1 
      Caption         =   "开始"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   240
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      Picture         =   "Sjt4.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub C1_Click()
Timer1.Enabled = True
End Sub

Private Sub C2_Click()
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    P1.Left = P1.Left + 20
End Sub

