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
   Begin VB.CommandButton Command2 
      Caption         =   "计算"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "最小值"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim min As Integer, max As Integer
Private Sub Command1_Click()
Dim x As Integer, y As Integer
x = Val(InputBox("请输入第一数字", "判断最小数", "10"))
y = Val(InputBox("请输入第二数字", "判断最小数", "30"))
min = IIf(x > y, y, x) 'min会取得最小值
max = IIf(x > y, x, y) 'max会取得最大值
Print min
Label1.Caption = min
Label2.Caption = max
End Sub

Private Sub Command2_Click()
Dim sum As Integer
sum = 0
For i = min To max
    sum = sum + i
Next i
Print sum
End Sub
