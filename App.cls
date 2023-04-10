VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const sBASE_64_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Function Enc(ByVal asContents As String)



        Dim lnPosition
        Dim lsResult
        Dim Char1
        Dim Char2
        Dim Char3
        Dim Char4
        Dim Byte1
        Dim Byte2
        Dim Byte3
        Dim SaveBits1
        Dim SaveBits2
        Dim lsGroupBinary
        Dim lsGroup64
        
           
        If Len(asContents) Mod 3 > 0 Then asContents = asContents & String(3 - (Len(asContents) Mod 3), " ")
        lsResult = ""
        
        For lnPosition = 1 To Len(asContents) Step 3
            lsGroup64 = ""
            lsGroupBinary = Mid(asContents, lnPosition, 3)
    
            Byte1 = Asc(Mid(lsGroupBinary, 1, 1)): SaveBits1 = Byte1 And 3
            Byte2 = Asc(Mid(lsGroupBinary, 2, 1)): SaveBits2 = Byte2 And 15
            Byte3 = Asc(Mid(lsGroupBinary, 3, 1))
    
            Char1 = Mid(sBASE_64_CHARACTERS, ((Byte1 And 252) \ 4) + 1, 1)
            Char2 = Mid(sBASE_64_CHARACTERS, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
            Char3 = Mid(sBASE_64_CHARACTERS, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)
            Char4 = Mid(sBASE_64_CHARACTERS, (Byte3 And 63) + 1, 1)
            lsGroup64 = Char1 & Char2 & Char3 & Char4
            
            lsResult = lsResult + lsGroup64
        Next
        
        Enc = lsResult
    End Function
    



Function ArrivalAddress() As String
On Error Resume Next
Dim tempVar As String

tempVar = Environ("REMOTE_ADDR")
ArrivalAddress = tempVar
' arrival
End Function

Function ArrivalUrl() As String
On Error Resume Next
Dim tempVar As String

tempVar = Environ("HTTP_REFERER")
ArrivalUrl = tempVar
' arrival
End Function

Public Function FileName() As String
On Error Resume Next
Dim tempVar As String
Dim t2 As String
Dim a1 As String
Dim szaM As Integer

tempVar = Environ("PATH_TRANSLATED")
t2 = Environ("PATH_TRANSLATED")



If KIT <> "CMD" Then
 For t = 1 To Len(tempVar)
  a1 = Mid(Right(tempVar, t), 1, 1)
    If a1 = "\" Then szaM = t: GoTo Folytat
 Next
 
Folytat:
  tempVar = Mid(t2, Len(t2) - szaM + 2, 100)
FileName = tempVar
End If

If KIT = "CMD" Then
h = SSearch(cbFile, "\")
FileName = Split(cbFile, "\")(h)
End If
End Function
Function URLPath() As String
On Error Resume Next
Dim tempVar As String

tempVar = Environ("PATH_INFO")

If Mid(tempVar, 1, 1) = "/" Then tempVar = Mid(tempVar, 2, 10000)
URLPath = tempVar

' Environ("PATH_TRANSLATED")
' arrival
End Function


Function URL() As String
On Error Resume Next
Dim tempVar As String

tempVar = Environ("PATH_INFO")

If Mid(tempVar, 1, 1) = "/" Then tempVar = Mid(tempVar, 2, 10000)
u1 = tempVar


Dim t2 As String
Dim a1 As String
Dim szaM As Integer

tempVar = Environ("PATH_TRANSLATED")
t2 = Environ("PATH_TRANSLATED")

 For t = 1 To Len(tempVar)
  a1 = Mid(Right(tempVar, t), 1, 1)
    If a1 = "\" Then szaM = t: GoTo Folytat
 Next
 
Folytat:
  tempVar = Mid(t2, Len(t2) - szaM + 2, 100)
u2 = tempVar

u3 = Replace(u1, "/", "\")
u3 = Replace(u3, "\" & u2, "")

u4 = Environ("SERVER_NAME")
u5 = "http://" & u4 & "/" & u3
u5 = Replace(u5, "\", "/")
URL = u5


End Function
Function CacheStatus() As String
On Error Resume Next
CacheStatus = iniFileCache
End Function


Public Function RequestURI() As String
On Error Resume Next
Dim tempVar As String
tempVar = Environ("REQUEST_URI")
RequestURI = Mid(tempVar, 2, 100000)
End Function
Function DomainName() As String
On Error Resume Next
Dim tempVar As String

tempVar = Environ("SERVER_NAME")
DomainName = tempVar
End Function
Public Function Path() As String
On Error Resume Next

If KIT <> "CMD" Then
 Dim tempVar As String
 Dim t2 As String
 Dim a1 As String
 Dim szaM As Integer

 tempVar = Environ("PATH_TRANSLATED")
 t2 = Environ("PATH_TRANSLATED")

 For t = 1 To Len(tempVar)
  a1 = Mid(Right(tempVar, t), 1, 1)
    If a1 = "\" Then szaM = t: GoTo Folytat
 Next
End If

 
Folytat:
If KIT <> "CMD" Then
  tempVar = Mid(t2, 1, Len(t2) - szaM)
  Path = tempVar
End If



'If KIT = "CMD" Then Path = cbFile

If KIT = "CMD" Then
 h = SSearch(cbFile, "\")
 fileNeve = Split(cbFile, "\")(h)
 Path = Replace(cbFile, "\" & fileNeve, "")
End If



End Function
Function QueryString() As String
On Error Resume Next
Dim tempVar
tempVar = Environ("QUERY_STRING")
QueryString = tempVar
End Function

Function GetCgiVars(ByVal keres As String) As String
On Error Resume Next
Dim tempVar
tempVar = Environ(keres)
GetCgiVars = tempVar
End Function


Function GetPostData() As String
On Error Resume Next
GetPostData = sFormData
End Function



Private Sub UserControl_Initialize()

End Sub

