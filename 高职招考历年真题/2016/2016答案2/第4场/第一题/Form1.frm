VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "右取(&R)"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "左取(&L)"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim str As String
str = Text1
If Option1 = True Then
    Label1 = Left(str, 3)
Else
    Label1 = Right(str, 3)
End If

End Sub
