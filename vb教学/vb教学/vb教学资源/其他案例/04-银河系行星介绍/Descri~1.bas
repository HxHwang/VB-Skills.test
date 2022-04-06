Attribute VB_Name = "ProcessRequest"
Option Explicit
Public Function GetInfo(Choice As Long) As String

'-------------------------------------------------
'Module:Description.Bas                 7/25/97
'-------------------------------------------------
'Author:Burt Abreu
'
'Purpose:Function recieves choice from
'lstPlanets_Click event and then processes
'using a simple Select Case decision structure;
'returning descriptive text to the txtContent
'Text property.
'-------------------------------------------------
  
'-------------------------------------------------
'Create and initialize NewLine variable which will
'use the ascii values for carriage return and line
'feed to create the functionality of a <Enter> key
'on the typewriter. Those of you using VB 5.0 can
'use the new constant vbCrLf for this.
'-------------------------------------------------
  Dim NewLine As String
  NewLine = Chr(13) + Chr(10)
  
  Select Case Choice
    Case 0 ' "asteroid"
    '--------------------------------------------
    'Process users combo box selection and assign
    'appropriate text to txtContent Text property
    'using a nested Select Case Statement.
    '--------------------------------------------
       
         Select Case Form1.cboSelectText.ListIndex
           Case 0
            '
            ' *** The return string of the function will be the text
            '     for the list box
            GetInfo = "Here you can write some interesting" _
            & " general information about asteroids that" _
            & " will introduce the subject."
           Case 1
            GetInfo = "Here you can write some statistics" _
            & " about asteroids, their size, weight etc."
           Case 2
            GetInfo = "Here you see the NewLine in action; we use it to" _
            & " create, what else, a new line; rather than wrapping" _
            & " the text." & NewLine & NewLine _
            & "Question: What's the name of the oldest known asteroid?" & NewLine _
            & NewLine _
            & "Answer: Rip Van Twinkle."
         End Select
         Exit Function  ' we got what we wanted
    Case 1 '"earth"
      GetInfo = "Write small blurb about earth here."
    Case 2 '"jupiter"
      GetInfo = "Write small blurb about jupiter here."
    Case 3 '"mars"
      GetInfo = "Write small blurb about mars here."
    Case 4 '"mercury"
      GetInfo = "Write small blurb about mercury here."
    Case 5 ' "meteor"
      GetInfo = "Write small blurb about meteors here."
    Case 6 ' "neptune"
      GetInfo = "Write small blurb about neptune here."
    Case 7 '"pluto"
      GetInfo = "Write small blurb about pluto here."
    Case 8 '"saturn"
      GetInfo = "Write small blurb about saturn here."
    Case 9 '"space craft"
      GetInfo = "Write small blurb about space craft here."
    Case 10 '"sun"
      GetInfo = "Write small blurb about sun here."
    Case 11 '"uranus"
      GetInfo = "Write small blurb about uranus here."
    Case 12 '"venus"
      GetInfo = "Write small blurb about venus here."
    End Select
End Function
