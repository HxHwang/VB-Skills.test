VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form addresult 
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   9270
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   3240
      Top             =   5640
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=student.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=student.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "courseinfo"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "���"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "����"
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "ѧ����Ϣ"
      Height          =   4455
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   4560
         TabIndex        =   10
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         DataSource      =   "Adodc2"
         Height          =   300
         Left            =   1440
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox text1 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   4560
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox text2 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "������"
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "�ɼ���"
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "�γ�����"
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "ѧ�ţ�"
         Height          =   495
         Left            =   3720
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   5040
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=student.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=student.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "result"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "addresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim courseid As String
Private Sub cmdadd_Click()
If Trim(text1.Text) = "" Then
  MsgBox "ѧ�Ų���Ϊ�գ������䣡", vbOKOnly + vbInformation, "��ʾ"
Else
  Adodc1.Recordset.Find "�γ̱�� = '" & courseid & "'"
  If Adodc1.Recordset.EOF Then
      Adodc1.Recordset.AddNew
      Adodc1.Recordset.Fields("ѧ��") = Trim(text1.Text)
      Adodc1.Recordset.Fields("�γ̱��") = courseid
      Adodc1.Recordset.Fields("�ɼ�") = Val(text2.Text)
      Adodc1.Recordset.Update
      MsgBox "��ӳɹ���", vbOKOnly + vbInformation, "��Ӽ�¼"
      resultfrm.Show
      addresult.Hide
   Else
    MsgBox "�ü�¼�ظ��������䣡", vbOKOnly + vbInformation, "��ʾ"
  End If
End If
End Sub

Private Sub Combo1_Click()
Adodc2.Refresh
Adodc2.Recordset.Find "�γ���='" & Trim(Combo1.Text) & "'"
courseid = Adodc2.Recordset.Fields("�γ̱��")

End Sub

Private Sub Form_Load()
Dim i As Integer
Adodc2.Refresh
For i = 1 To Adodc2.Recordset.RecordCount
 Combo1.AddItem Adodc2.Recordset.Fields("�γ���")
 Adodc2.Recordset.MoveNext
Next i
End Sub
