VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   6255
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      Caption         =   "查找"
      Height          =   375
      Left            =   4680
      TabIndex        =   19
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "修改"
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "删除"
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "添加"
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   3840
      Width           =   1050
   End
   Begin VB.TextBox Text3 
      DataField       =   "性别"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "最后一个"
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "下一个"
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "上一个"
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "第一个"
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   2880
      Width           =   975
   End
   Begin VB.Data Data1 
      BOFAction       =   1  'BOF
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "F:\hgy\vb\data控件\student.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      EOFAction       =   1  'EOF
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "studentinfo"
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      DataField       =   "班级"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      DataField       =   "年级"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      DataField       =   "出生日期"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      DataField       =   "姓名"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "学号"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "班级："
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "年级："
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "出生日期："
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "性别："
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "姓名："
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "学号："
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

  Data1.Recordset.MoveFirst

End Sub

Private Sub Command2_Click()

If Data1.Recordset.AbsolutePosition <= 0 Then
   MsgBox "已经是第一条记录！", vbOKOnly + vbInformation, "提示"
   Data1.Recordset.MoveLast
Else
   Data1.Recordset.MovePrevious
End If
End Sub

Private Sub Command3_Click()
If Data1.Recordset.AbsolutePosition >= Data1.Recordset.RecordCount - 1 Then
  MsgBox "已经是最后一条记录！", vbOKOnly + vbInformation, "提示"
  Data1.Recordset.MoveFirst
Else
  Data1.Recordset.MoveNext
End If
End Sub

Private Sub Command4_Click()
Data1.Recordset.MoveLast
End Sub

Private Sub Command5_Click()
If Command5.Caption = "添加" Then
 Data1.Recordset.AddNew
 Command5.Caption = "保存"
Else
 Data1.Recordset.Update
 Command5.Caption = "添加"
End If
End Sub

Private Sub Command6_Click()
Data1.Recordset.Delete
Data1.Recordset.MoveNext
End Sub

Private Sub Command7_Click()
Data1.Recordset.FindFirst "学号='06003'"
End Sub

Private Sub Command8_Click()

If Command8.Caption = "修改" Then
 Data1.Recordset.Edit
 Command8.Caption = "保存"
Else
 Data1.Recordset.Update
 MsgBox "该学生信息已修改！", vbOKOnly + vbInformation, "提示"
 Command8.Caption = "修改"
End If

End Sub

Private Sub Form_Click()
Text7.Text = Data1.Recordset.AbsolutePosition
End Sub

