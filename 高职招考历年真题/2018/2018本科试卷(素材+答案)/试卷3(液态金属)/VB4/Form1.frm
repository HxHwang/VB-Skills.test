VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6465
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "优秀率"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   2520
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim a(1 To 4, 1 To 4) As Integer
    Dim s As Integer
    Randomize
    Print "科目"; "  学生1"; "  学生2"; " 学生3"; " 学生4"; " 优秀率"
    For i = 1 To 4
        Print "科目"; Trim(i);
        s = 0
        For j = 1 To 4
            a(i, j) = Int(Rnd * 50 + 50)
            Print a(i, j); "   ";
            If a(i, j) >= 90 Then s = s + 1
        Next j
        Print Trim((s / 4) * 100); "%"
    Next i
End Sub
