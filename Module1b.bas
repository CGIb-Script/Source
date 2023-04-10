Attribute VB_Name = "Module1"
Public AllSite As String
Public ScriptNow As String

Public Sub Send(ByVal data As String)
ScriptNow = ScriptNow & data & vbCrLf
End Sub

Public Function SCrop(ByVal ifi As String, ByVal startStr1 As String, ByVal endStr1 As String) As String
On Error Resume Next
Dim tempVar As String
Dim s As Long
Dim e As Long
Dim h As Long
Dim sS As Long
Dim vV As Long

keresh = Len(startStr1)
s = InStr(1, ifi, startStr1)
e = InStr(s, ifi, endStr1)
sS = s
vV = e - (sS)
tempVar = Mid(ifi, sS, vV + Len(endStr1))
If Err.Description <> "" Then tempVar = ""
SCrop = tempVar

End Function
Public Function SSearch(ByVal ifi As String, ByVal searchStr As String) As String
On Error Resume Next
Dim tempVar As String
Dim s As String
Dim jo As Integer
  jo = 0
  For t = 1 To Len(ifi)
   s = Mid(ifi, t, Len(searchStr))
   If s = searchStr Then jo = jo + 1
   Next
 
tempVar = Str(jo)
SSearch = Trim(tempVar)
End Function

