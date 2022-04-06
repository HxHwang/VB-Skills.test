VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   4455
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   1320
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "退出"
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton CmdReset 
      Caption         =   "重新设置"
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton CmdPause 
      Caption         =   "暂停"
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton CmdStar 
      Caption         =   "开始"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   975
   End
   Begin VB.HScrollBar HscSec 
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.HScrollBar HscMin 
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.HScrollBar HscHour 
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.ComboBox CboSec 
      Height          =   300
      Left            =   3360
      TabIndex        =   5
      Text            =   "00"
      Top             =   240
      Width           =   855
   End
   Begin VB.ComboBox CboMin 
      Height          =   300
      Left            =   1920
      TabIndex        =   3
      Text            =   "00"
      Top             =   240
      Width           =   855
   End
   Begin VB.ComboBox CboHour 
      Height          =   300
      Left            =   480
      TabIndex        =   1
      Text            =   "00"
      Top             =   240
      Width           =   855
   End
   Begin VB.Label LblShow 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1440
      TabIndex        =   10
      Top             =   1440
      Width           =   1800
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "剩余时间"
      Height          =   180
      Left            =   1920
      TabIndex        =   9
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "秒："
      Height          =   180
      Left            =   3000
      TabIndex        =   4
      Top             =   300
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "分："
      Height          =   180
      Left            =   1560
      TabIndex        =   2
      Top             =   300
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "时："
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Hours As Integer
Dim Minutes As Integer
Dim Seconds As Integer
Dim Time As Date

Private Sub Mydisplay()
    Hours = Val(CboHour.Text)
    Minutes = Val(CboMin.Text)
    Seconds = Val(CboSec.Text)
    Time = TimeSerial(Hours, Minutes, Seconds)
Label1.Caption = Format$(Time, "hh") & ":" & Format$(Time, "nn") _
& ":" & Format$(Time, "ss")
End Sub

Private Sub CmdPaust_Click()

End Sub

Private Sub Form_Load()
    Timer1.Interval = 500
    Hours = 0
    Minutes = 0
    Seconds = 0
    Time = 0
End Sub

Private Sub cmdStar_Click()
    Timer1.Interval = 500
    Timer1.Enabled = True
    CmdReset.Enabled = False
End Sub

Private Sub cmdPause_Click()
    Timer1.Interval = 0
    CmdReset.Enabled = True
End Sub

Private Sub cmdReset_Click()
    Hours = 0
    Minutes = 0
    Seconds = 0
    Time = 0
    CboHour.Text = " "
    CboMin.Text = " "
    CboSec.Text = " "
End Sub

Private Sub cmdQuit_Click()
    End
End Sub


Private Sub cbohour_Change()
    Mydisplay
End Sub

Private Sub cbomin_Change()
    Mydisplay
End Sub

Private Sub cbosec_Change()
    Mydisplay
End Sub

Private Sub hscHour_Change()
CboHour.Text = HscHour.Value
End Sub

Private Sub hscMin_Change()
CboMin.Text = HscMin.Value
End Sub

Private Sub hscSec_Change()
CboSec.Text = HscSec.Value
End Sub


Private Sub Timer1_Timer()
    'Count down loop
    Timer1.Enabled = False
    If (Format$(Time, "hh") & ":" & Format$(Time, "nn") & ":" _
& Format$(Time, "ss")) <> "00:00:00" Then
        Time = DateAdd("s", -1, Time)
        LblShow.Visible = False
        LblShow.Caption = Format$(Time, "hh") & ":" & _
Format$(Time, "nn") & ":" & Format$(Time, "ss")
        LblShow.Visible = True
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
        Beep
        Beep
        CmdReset.Enabled = True
    End If
End Sub


