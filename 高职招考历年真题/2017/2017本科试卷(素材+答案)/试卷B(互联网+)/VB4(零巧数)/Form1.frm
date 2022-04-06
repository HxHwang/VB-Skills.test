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
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "求零巧数"
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i
Dim bai As Integer
For i = 1000 To 9999
    bai = i \ 100 Mod 10
    If bai = 0 Then
        'left取i左边的1位，right取i右边的2位
        If (Left(i, 1) & Right(i, 2)) * 9 = i Then
            Print i
        End If
    End If
Next i
End Sub
