VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   3795
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   360
   End
   Begin VB.CommandButton C1 
      Caption         =   "¿ª³µ"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.PictureBox P2 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2880
      Picture         =   "sjt5.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   720
      Top             =   360
   End
   Begin VB.PictureBox P1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1440
      Picture         =   "sjt5.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a%, b As Boolean

Private Sub C1_Click()
    Timer2.Enabled = True
    b = True
End Sub

Private Sub Timer1_Timer()
    a = a + 1
    If a > 6 Then
        a = 1
    End If
    Select Case a
        Case 1
            P1.Picture = LoadPicture(App.Path + "\" + "»ÆµÆ.ico")
        Case 2, 3
            P1.Picture = LoadPicture(App.Path + "\" + "ºìµÆ.ico")
        Case 4, 5, 6
            P1.Picture = LoadPicture(App.Path + "\" + "ÂÌµÆ.ico")
            If b Then Timer2.Enabled = b
    End Select

End Sub

Private Sub Timer2_Timer()
   If (a < 4) And (P2.Left > P1.Left And P2.Left < P1.Left + P1.Width) Or P2.Left <= 100 Then
        Timer2.Enabled = False
    Else
        P2.Move P2.Left - 10, P2.Top, P2.Width, P2.Height
    End If

End Sub

