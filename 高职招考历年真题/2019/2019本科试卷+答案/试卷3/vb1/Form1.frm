VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   4950
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton Option2 
      Caption         =   "字符（&C）"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "好(&Y)"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&K)"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   1215
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer ' 提高作用域
Private Sub Command1_Click()

If Option1.Value = True Then
    n = Int(Rnd * 26 + 65)
    Text1.Text = n
    Label1.Caption = String(n, "好")
Else
    Label1.Caption = Chr(n)
End If
End Sub
