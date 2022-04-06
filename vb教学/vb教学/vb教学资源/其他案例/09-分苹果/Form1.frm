VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   5760
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "分苹果大赛"
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      Begin VB.CommandButton Command2 
         Caption         =   "加一个"
         Height          =   495
         Index           =   1
         Left            =   3840
         TabIndex        =   12
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "减一个"
         Height          =   495
         Index           =   1
         Left            =   2760
         TabIndex        =   11
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "加一个"
         Height          =   495
         Index           =   0
         Left            =   1320
         TabIndex        =   10
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "减一个"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   3600
         Width           =   975
      End
      Begin VB.PictureBox Picsmile 
         AutoSize        =   -1  'True
         Height          =   2070
         Index           =   1
         Left            =   2880
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   2010
         ScaleWidth      =   1545
         TabIndex        =   5
         Top             =   600
         Width           =   1605
      End
      Begin VB.PictureBox Piccry 
         AutoSize        =   -1  'True
         Height          =   2025
         Index           =   1
         Left            =   2880
         Picture         =   "Form1.frx":A392
         ScaleHeight     =   1965
         ScaleWidth      =   1650
         TabIndex        =   4
         Top             =   600
         Width           =   1710
      End
      Begin VB.PictureBox Piccry 
         AutoSize        =   -1  'True
         Height          =   2025
         Index           =   0
         Left            =   600
         Picture         =   "Form1.frx":14DB8
         ScaleHeight     =   1965
         ScaleWidth      =   1650
         TabIndex        =   3
         Top             =   600
         Width           =   1710
      End
      Begin VB.PictureBox Picsmile 
         AutoSize        =   -1  'True
         Height          =   2070
         Index           =   0
         Left            =   600
         Picture         =   "Form1.frx":1F7DE
         ScaleHeight     =   2010
         ScaleWidth      =   1545
         TabIndex        =   2
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label Label3 
         Caption         =   "0"
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   8
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "0"
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   7
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Kite"
         Height          =   495
         Left            =   2880
         TabIndex        =   6
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Tom"
         Height          =   615
         Left            =   600
         TabIndex        =   1
         Top             =   3000
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click(Index As Integer)

If Index = 0 Then
    If Val(Label3(0).Caption) > 1 Then
            Command1(0).Enabled = True
            Label3(0).Caption = Label3(0).Caption - 1
    Else
         If Val(Label3(0).Caption) = 1 Then
            Label3(0).Caption = Label3(0).Caption - 1
         End If
            Command1(0).Enabled = False
    End If
Else
    If Val(Label3(1).Caption) > 1 Then
            Command1(1).Enabled = True
            Label3(1).Caption = Label3(1).Caption - 1
    Else
        If Val(Label3(1).Caption) = 1 Then
            Label3(1).Caption = Label3(1).Caption - 1
        End If
            Command1(1).Enabled = False
    End If
    
End If
If Val(Label3(0).Caption) > Val(Label3(1).Caption) Then
   Picsmile(0).Visible = True
   Piccry(0).Visible = False
   Piccry(1).Visible = True
   Picsmile(1).Visible = False
   Else
   If Val(Label3(0).Caption) < Val(Label3(1).Caption) Then
      Piccry(0).Visible = True
      Picsmile(0).Visible = False
      Picsmile(1).Visible = True
      Piccry(1).Visible = False
    Else
     Picsmile(0).Visible = True
     Piccry(0).Visible = False
     Picsmile(1).Visible = True
     Piccry(1).Visible = False
    End If
End If

End Sub
Private Sub Command2_Click(Index As Integer)

If Index = 0 Then
        Label3(0).Caption = Label3(0).Caption + 1
        Command1(0).Enabled = True
Else:
        Label3(1).Caption = Label3(1).Caption + 1
        Command1(1).Enabled = True
End If

If Val(Label3(0).Caption) > Val(Label3(1).Caption) Then
     Picsmile(0).Visible = True
     Piccry(0).Visible = False
     Piccry(1).Visible = True
      Picsmile(1).Visible = False
Else
     If Val(Label3(0).Caption) < Val(Label3(1).Caption) Then
            Piccry(0).Visible = True
            Picsmile(0).Visible = False
            Picsmile(1).Visible = True
            Piccry(1).Visible = False
     Else
            Picsmile(0).Visible = True
            Piccry(0).Visible = False
            Picsmile(1).Visible = True
            Piccry(1).Visible = False
     End If
End If

End Sub


Private Sub Form_Load()

Picsmile(0).Visible = True
Picsmile(1).Visible = True
Command1(0).Enabled = False
Command1(1).Enabled = False

End Sub

