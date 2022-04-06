VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "守形数"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "求守形数"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i As Long
    Dim s As Long
    Print "1000以内的守形数有："
    For i = 1 To 1000 '1000以内的数
        s = i * i
        If s Mod 10 = i Or s Mod 100 = i Or s Mod 1000 = i Then
            Print i
        End If
    Next i
End Sub
