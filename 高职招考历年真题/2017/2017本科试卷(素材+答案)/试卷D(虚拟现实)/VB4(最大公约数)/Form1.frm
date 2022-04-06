VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "最大公约数与最小公倍数"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "最小公倍数为"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "最大公约数为"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "请输入两个整数"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim m As Integer, n As Integer
Dim max As Integer
Dim a As Integer, b As Integer, gbs As Integer
max = 0
m = Val(Text1.Text)
n = Val(Text2.Text)
For i = 1 To m
    If m Mod i = 0 And n Mod i = 0 Then
        If i > max Then
            max = i
            a = m \ max
            b = n \ max
            gbs = a * b * max
        End If
    End If
Next i
Label3.Caption = max
Label5.Caption = gbs
Print zxgbs
End Sub

Private Sub Command2_Click()
End Sub
'最小公倍数
'12 和 8
'12 = 4 * 3
'8 =  4 * 2
'最小公倍数 = 2 * 3 * 4
'因为4为最大公约数所以可以利用4求出最小公倍数


