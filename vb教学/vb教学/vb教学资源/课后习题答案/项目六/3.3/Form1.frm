VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   6645
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function findmax(a() As Integer)
 Dim start As Integer, finish As Integer, i As Integer
 start = LBound(a)
 finish = UBound(a)
 Max = a(start)
 For i = start To finish
  If a(i) > Max Then Max = a(i)
 Next i
 findmax = Max
 
End Function


Private Sub Form_Load()

End Sub
