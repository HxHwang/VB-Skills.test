VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9e.ocx"
Begin VB.Form xtjj2 
   Caption         =   "��Ƶ����"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   11970
   LinkTopic       =   "Form7"
   ScaleHeight     =   7335
   ScaleWidth      =   11970
   StartUpPosition =   3  '����ȱʡ
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      _cx             =   5953
      _cy             =   4895
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "xtjj2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.ShockwaveFlash1.Movie = App.Path & "\3.2.swf"
Me.WindowState = 2
End Sub

Private Sub Form_Resize()
Me.ShockwaveFlash1.Width = Me.Width
Me.ShockwaveFlash1.Height = Me.Height
End Sub

