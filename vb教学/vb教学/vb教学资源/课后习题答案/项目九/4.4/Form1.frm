VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   7305
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4680
      Top             =   4320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cx As Double
Dim Cy As Double


Private Sub Form_Load()
    Form1.Show
    Form1.Scale (0, 1.5)-(2 * 3.1415926, -1.5)
End Sub

Private Sub Timer1_Timer()
    Static i As Integer
    i = i + 1
    Cx = i * 3.1415926 / 100
    Cy = Sin(Cx)
    PSet (Cx, Cy)
    If i = 200 Then
        Timer1.Enabled = False
    End If
    
End Sub
