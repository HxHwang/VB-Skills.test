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
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
Dim i, j, m, n As Integer
Dim s As String
Form1.AutoRedraw = True
For i = 1 To 9
    For j = 1 To i
        s = i * j
        Print j & "¡Á" & i & "=" & s,
    Next j
    Print
Next i
For m = 1 To 9
    For n = 1 To 10 - m
        s = m * n
        Print n & "¡Á" & m & "=" & s,
    Next n
    Print
Next m
End Sub
