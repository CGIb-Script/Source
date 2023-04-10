VERSION 5.00
Begin VB.Form CGIForm 
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5025
   Icon            =   "CGIForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox RequestText 
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   240
   End
End
Attribute VB_Name = "CGIForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)

End Sub

Private Sub Timer1_Timer()

If SetTime = 60 Then
 g = FreeFile
 Open App.Path & "\longtime.log" For Append As #g
 Print #g, Now & ", " & Environ("PATH_TRANSLATED") & ", " & Environ("PATH_TRANSLATED")
 Close #g
End If

SetTime = SetTime + 1
If Fix(SetTime) >= Fix(SetTimeIntervall) Then
 Timer1.Enabled = False
 Send "Status: 200 OK"
 Send "Content-type: text/html" & vbCrLf
 If TimeOutScriptData = "" Then Send "time out..."
 If TimeOutScriptData <> "" Then Send "<script>" & TimeOutScriptData & "</script>"
 Close
 Unload CGIForm
 End
End If
End Sub
