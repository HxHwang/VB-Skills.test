VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   8400
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   5400
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   6120
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Height          =   3870
      Left            =   3960
      TabIndex        =   2
      Top             =   1200
      Width           =   3135
   End
   Begin VB.DirListBox Dir1 
      Height          =   3450
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fname As String
Dim yfile As String

Private Sub Command1_Click()

Kill yfile
File1.Refresh
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive

End Sub

Private Sub File1_Click()
fname = File1.FileName
If Right(Dir1.Path, 1) = "\" Then
   yfile = Dir1.Path & fname
Else
   yfile = Dir1.Path & "\" & fname
End If
End Sub

Private Sub Form_Load()
Drive1.Drive = "D"
End Sub

