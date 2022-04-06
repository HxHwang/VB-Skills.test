VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   5655
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "打印"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "请输入一个整数【1-20】"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim num As Integer
Dim i As Integer, j As Integer
num = Val(Text1)
For i = 1 To num
    Print Space(1);
    For j = 1 To num
        Select Case j
            Case Is <= i: Print "1" & Space(1);
            Case Else: Print "0" & Space(1);
        End Select
    Next j
    Print
Next i
End Sub
