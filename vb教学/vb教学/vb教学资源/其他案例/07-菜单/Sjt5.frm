VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1740
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   4050
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Menu Edit 
      Caption         =   "编辑"
      Begin VB.Menu Cut 
         Caption         =   "剪切"
      End
      Begin VB.Menu Copy 
         Caption         =   "复制"
      End
      Begin VB.Menu Paste 
         Caption         =   "粘贴"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim which As Integer

Private Sub copy_Click()
    If which = 1 Then
        Text3.Text = Text1.Text
    ElseIf which = 2 Then
        Text3.Text = Text2.Text
    End If
End Sub

Private Sub cut_Click()
    If which = 1 Then
        Text3.Text = Text1.Text
        Text1.Text = ""
    ElseIf which = 2 Then
        Text3.Text = Text2.Text
        Text2.Text = ""
    End If
End Sub

Private Sub edit_Click()
    If which = 1 Then
        If Text1.Text = "" Then
            Cut.Enabled = False
            Copy.Enabled = False
        Else
            Cut.Enabled = True
            Copy.Enabled = True
        End If
    ElseIf which = 2 Then
        If Text2.Text = "" Then
            Cut.Enabled = False
            Copy.Enabled = False
        Else
            Cut.Enabled = True
            Copy.Enabled = True
        End If
    End If
    If Text3.Text = "" Then
        Paste.Enabled = False
    Else
        Paste.Enabled = True
    End If
End Sub

Private Sub paste_Click()
    If which = 1 Then
        Text1.Text = Text3.Text
    ElseIf which = 2 Then
        Text2.Text = Text3.Text
    End If
End Sub

Private Sub Text1_GotFocus()   '本过程的作用是：当焦点在Text1中时，which = 1
    which = 1
End Sub

Private Sub Text2_GotFocus()   '本过程的作用是：当焦点在Text2中时，which = 2
    which = 2
End Sub

