VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Galaxy Viewer"
   ClientHeight    =   3885
   ClientLeft      =   1800
   ClientTop       =   2550
   ClientWidth     =   4590
   Icon            =   "galaxy.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   4590
   Begin VB.ComboBox cboSelectText 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   2955
   End
   Begin VB.TextBox txtContent 
      Height          =   1155
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   120
      Width           =   2955
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   3060
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   2460
      Width           =   1095
   End
   Begin VB.ListBox lstPlanets 
      Height          =   3570
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line linAccent 
      BorderColor     =   &H80000002&
      BorderWidth     =   5
      X1              =   60
      X2              =   4560
      Y1              =   3780
      Y2              =   3780
   End
   Begin VB.Image imgLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   3480
      Picture         =   "galaxy.frx":030A
      Stretch         =   -1  'True
      Top             =   1740
      Width           =   1065
   End
   Begin VB.Image imgShow 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   1740
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'一个小程序演示了列表框、图片框Image、下拉列表框的使用
Option Explicit

Private Sub cboSelectText_Change()
Dim rs 'returned string from GetInfo function
Dim Choice As Long

' *** get the text of the list
Choice = Form1.lstPlanets.ListIndex
' *** call the function and get the display text
rs = GetInfo(Choice)
' *** display the text
Form1.txtContent.Text = rs
End Sub

Private Sub cboSelectText_Click()
Dim rs
Dim Choice As Long

' *** get the text of the list
Choice = Form1.lstPlanets.ListIndex
' *** call the function and get the display text
rs = GetInfo(Choice)
' *** display the text
Form1.txtContent.Text = rs

End Sub


Private Sub cmdAbout_Click()
'---------------------------------
'Creates a simple About box using
'the MsgBox function.
'---------------------------------
    MsgBox ("Created by: Your Name Here"), , "About"
End Sub

Private Sub cmdQuit_Click()
'---------------------------------
'Exit the program first UnLoading
'the form using UnLoad Me and then
'the End statement.
'---------------------------------
    Unload Me
    End
End Sub

Private Sub Form_Load()
' *** clear the combo and list boxes
Form1.lstPlanets.Clear
Form1.cboSelectText.Clear

'---------------------------------
'Add items to the List box using
'AddItem method.
'---------------------------------
    Form1.lstPlanets.AddItem "Asteroids"
    Form1.lstPlanets.AddItem "Earth"
    Form1.lstPlanets.AddItem "Jupiter"
    Form1.lstPlanets.AddItem "Mars"
    Form1.lstPlanets.AddItem "Mercury"
    Form1.lstPlanets.AddItem "Meteor"
    Form1.lstPlanets.AddItem "Neptune"
    Form1.lstPlanets.AddItem "Pluto"
    Form1.lstPlanets.AddItem "Saturn"
    Form1.lstPlanets.AddItem "Space Craft"
    Form1.lstPlanets.AddItem "Sun"
    Form1.lstPlanets.AddItem "Uranus"
    Form1.lstPlanets.AddItem "Venus"
    
'----------------------------------
'Add items to the Combo Box using
'AddItem method.
'----------------------------------
    Form1.cboSelectText.AddItem "General Info."
    '
    ' *** The ItemData area of a list box is a great place to
    '     store numbers. These could be a record number from a
    '     database. Here we are using the number as a pointer
    '     that we can select the text with.
    '
    Form1.cboSelectText.ItemData(Form1.cboSelectText.NewIndex) = 1
    Form1.cboSelectText.AddItem "Statistics"
    Form1.cboSelectText.ItemData(Form1.cboSelectText.NewIndex) = 2
    Form1.cboSelectText.AddItem "History"
    Form1.cboSelectText.ItemData(Form1.cboSelectText.NewIndex) = 3
End Sub


Private Sub lstPlanets_Click()
'--------------------------------------------------
'Path-stores path to picture files; cuts down on
'the repetition.
'Choice-string value passed to the GetInfo function
'in the Description.Bas module; indicates what text
'to return.
'--------------------------------------------------
Dim Path As String
Dim Choice As String
'------------------------------------------------------
'Set the default value for the ComboBox "cboSelectText"
'to 0 "General Info."
'------------------------------------------------------
cboSelectText.ListIndex = 0
    
  Select Case lstPlanets.ListIndex
      Case 0
        imgShow.Picture = LoadPicture(App.Path + "\" + "asteroid.bmp")
        txtContent.Text = GetInfo(0)
      Case 1
        imgShow.Picture = LoadPicture(App.Path + "\" + "earth.bmp")
        txtContent.Text = GetInfo(1)
      Case 2
        imgShow.Picture = LoadPicture(App.Path + "\" + "jupiter.bmp")
        txtContent.Text = GetInfo(2)
      Case 3
        imgShow.Picture = LoadPicture(App.Path + "\" + "mars.bmp")
        txtContent.Text = GetInfo(3)
      Case 4
        imgShow.Picture = LoadPicture(App.Path + "\" + "merc.bmp")
        txtContent.Text = GetInfo(4)
      Case 5
        imgShow.Picture = LoadPicture(App.Path + "\" + "meteor.bmp")
        txtContent.Text = GetInfo(5)
      Case 6
        imgShow.Picture = LoadPicture(App.Path + "\" + "neptune.bmp")
        txtContent.Text = GetInfo(6)
      Case 7
        imgShow.Picture = LoadPicture(App.Path + "\" + "pluto.bmp")
        txtContent.Text = GetInfo(7)
      Case 8
        imgShow.Picture = LoadPicture(App.Path + "\" + "saturn.bmp")
        txtContent.Text = GetInfo(8)
      Case 9
        imgShow.Picture = LoadPicture(App.Path + "\" + "craft.bmp")
        txtContent.Text = GetInfo(9)
      Case 10
        imgShow.Picture = LoadPicture(App.Path + "\" + "sun.bmp")
        txtContent.Text = GetInfo(10)
      Case 11
        imgShow.Picture = LoadPicture(App.Path + "\" + "uranus.bmp")
        txtContent.Text = GetInfo(11)
      Case 12
        imgShow.Picture = LoadPicture(App.Path + "\" + "venus.bmp")
        txtContent.Text = GetInfo(12)
    End Select
End Sub

