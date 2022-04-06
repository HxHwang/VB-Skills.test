VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   3030
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdQuit 
      Caption         =   "退出"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdLarge 
      Caption         =   "放大"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdSmall 
      Caption         =   "缩小"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   855
   End
   Begin VB.Image imgcat 
      Height          =   3375
      Left            =   240
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i As Double
Dim j As Double
Private Sub cmdLarge_Click()
imgcat.Left = imgcat.Left - imgcat.Width * 0.05
imgcat.Top = imgcat.Top - imgcat.Height * 0.05
imgcat.Width = imgcat.Width + imgcat.Width * 0.1
imgcat.Height = imgcat.Height + imgcat.Height * 0.1
End Sub

Private Sub cmdSmall_Click()
imgcat.Left = imgcat.Left + imgcat.Width * 0.05
imgcat.Top = imgcat.Top + imgcat.Height * 0.05
imgcat.Width = imgcat.Width - imgcat.Width * 0.1
imgcat.Height = imgcat.Height - imgcat.Height * 0.1
End Sub
