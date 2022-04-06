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
      Left            =   1680
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "右取(&R)"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "左取(&L)"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim str As String
str = Text1.Text
If Option1.Value = True Then
    Label1.Caption = Left(str, 3)
Else
    Label1.Caption = Right(str, 3)
End If



End Sub
