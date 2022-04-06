VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "降序数"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "求降序数"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i As Integer
    Dim t As Integer
    Dim a As Integer
    Dim b As Integer
    Print "所有2位的降序数有："
    For i = 10 To 99  '对所有两位数进行循环判断
        a = i \ 10 '求出十位上的数
        b = i Mod 10 '求出个位上的数
        If a >= b Then
            Print i;
            t = t + 1
            If t = 10 Then
                Print  '换行
                t = 0 '计数器重置为0
            End If
        End If
    Next i
End Sub
