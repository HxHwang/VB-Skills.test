VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "字符串比较"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "比较"
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "第二个"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "第一个"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Len(Trim(Text1.Text)) > Len(Trim(Text2.Text)) Then
        Print "第一个串更长"
    Else
        Print "第二个串长"
    End If
End Sub
