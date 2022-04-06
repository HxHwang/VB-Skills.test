VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   5190
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "给二维数据赋初值"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   2880
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i, j, a(1 To 5, 1 To 5) As Integer
For i = 1 To 5
 For j = 1 To i
 a(i, j) = i
 Next
Next

For i = 1 To 5
 For j = 1 To 5
 Print a(i, j);
 Next
 Print
Next
End Sub
