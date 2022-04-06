VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "移动窗体"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   2880
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Left = Left - 100
Top = Top + 100
Form1.Caption = "(" & Left & "," & Top & ")"
End Sub

Private Sub Form_Load()
Left = Screen.Width - Form1.Width
Top = 0
End Sub
