VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   360
      ScaleHeight     =   2355
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldx As Integer
Dim oldy As Integer


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  oldx = X
  oldy = Y
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Abs(X - oldx) > Abs(Y - oldy) Then radius = Abs(Y - oldy) Else radius = Abs(X - oldx)
     Picture1.Circle (oldx, oldy), radius
End Sub
