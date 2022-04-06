VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2355
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   2355
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox P2 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   1560
      Picture         =   "sjt3.frx":0000
      ScaleHeight     =   420
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   3120
   End
   Begin VB.CommandButton C1 
      Caption         =   "·¢Éä"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   975
   End
   Begin VB.PictureBox P1 
      BorderStyle     =   0  'None
      FillStyle       =   3  'Vertical Line
      Height          =   495
      Left            =   1560
      Picture         =   "sjt3.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   3840
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
Private Sub Timer1_Timer()
    Static a%
    a = a + 1
    If P1.Top > P2.Top + P2.Height Then
        P1.Move P1.Left, P1.Top - 5 - a, P1.Width, P1.Height
    Else
        Timer1.Enabled = False
    End If
End Sub

