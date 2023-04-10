VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function Decode(ByVal sStr As String) As String
 On Error Resume Next
 Dim l As Long, sUTF8 As String, iChar As Integer, iChar2 As Integer
 For l = 1 To Len(sStr)
 iChar = Asc(Mid(sStr, l, 1))
 If iChar > 127 Then
 If Not iChar And 32 Then ' 2 chars
 iChar2 = Asc(Mid(sStr, l + 1, 1))
 sUTF8 = sUTF8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
 l = l + 1
 Else
 Dim iChar3 As Integer
 iChar2 = Asc(Mid(sStr, l + 1, 1))
 iChar3 = Asc(Mid(sStr, l + 2, 1))
 sUTF8 = sUTF8 & ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
 l = l + 2
 End If
 Else
 sUTF8 = sUTF8 & Chr$(iChar)
 End If
 Next l
 Decode = sUTF8
End Function
Public Function Encode(ByVal sStr As String) As String
 On Error Resume Next
 For l& = 1 To Len(sStr)
 lChar& = AscW(Mid(sStr, l&, 1))
 If lChar& < 128 Then
 sUTF8$ = sUTF8$ + Mid(sStr, l&, 1)
 ElseIf ((lChar& > 127) And (lChar& < 2048)) Then
 sUTF8$ = sUTF8$ + Chr(((lChar& \ 64) Or 192))
 sUTF8$ = sUTF8$ + Chr(((lChar& And 63) Or 128))
 Else
 sUTF8$ = sUTF8$ + Chr(((lChar& \ 144) Or 234))
 sUTF8$ = sUTF8$ + Chr((((lChar& \ 64) And 63) Or 128))
 sUTF8$ = sUTF8$ + Chr(((lChar& And 63) Or 128))
 End If
 Next l&
 Encode = sUTF8$
End Function
