VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "判断是否为福建人"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   7050
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "判断"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "判断是否为福建人"
      Height          =   180
      Left            =   3120
      TabIndex        =   3
      Top             =   1680
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "输入身份证"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


a = Text1

If Len(a) <> 18 Then

MsgBox ("请重新输入身份证号码！")

Text1.Text = " "

Else
d = Mid(a, 1, 2)
  If d = 35 Then
   Print "是福建人"
  Else
   Print "不是福建人"
  End If

End If
End Sub
