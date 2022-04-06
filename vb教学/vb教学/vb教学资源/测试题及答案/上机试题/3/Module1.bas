Attribute VB_Name = "Module1"
Option Explicit

Sub putdata(t_FileName As String, T_Str As Variant)
    Dim sFile As String
    sFile = "\" & t_FileName
    Open App.Path & sFile For Output As #1
    Print #1, T_Str
    Close #1
End Sub

Function isprime(t_I As Integer) As Boolean
   Dim J As Integer
   isprime = False
   For J = 2 To t_I / 2
      If t_I Mod J = 0 Then Exit For
   Next J
   If J > t_I / 2 Then isprime = True
End Function

