VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   11310
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算运费"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "运费："
      Height          =   180
      Left            =   3240
      TabIndex        =   5
      Top             =   2280
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "重量："
      Height          =   180
      Left            =   3240
      TabIndex        =   3
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "票价："
      Height          =   180
      Left            =   3240
      TabIndex        =   2
      Top             =   840
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Val(Text2) < 20 Then
    Text3 = 0
Else
    Text3 = (Val(Text2) - 20) * Val(Text1) * 0.015
End If
End Sub
